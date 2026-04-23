"""日次バッチCSV生成：DBから20件選定してCSVとClaude読込用MDを作る。

各候補エントリはAtomPub存在確認を行い、404の場合は excluded=1 に
マークして代替を選出する（削除済み記事の混入を避けるため）。
"""

from __future__ import annotations

import csv
import sqlite3
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Callable


CSV_HEADER = [
    "entry_id",
    "title",
    "category",
    "url",
    "style_summary",
    "prompt",
    "alt",
    "attempt",
    "previous_ng_tags",
    "previous_ng_note",
]


def select_batch_entry_ids(
    conn: sqlite3.Connection, batch_size: int
) -> list[str]:
    ng_quota = max(1, batch_size // 5)
    ng_rows = conn.execute(
        """SELECT entry_id FROM articles
        WHERE status='ng_pending'
        ORDER BY last_ng_date ASC
        LIMIT ?""",
        (ng_quota,),
    ).fetchall()
    ng_ids = [r[0] for r in ng_rows]

    remaining = batch_size - len(ng_ids)
    pending_rows = conn.execute(
        """SELECT entry_id FROM articles
        WHERE status='pending' AND excluded=0
        ORDER BY published_date ASC
        LIMIT ?""",
        (remaining,),
    ).fetchall()
    pending_ids = [r[0] for r in pending_rows]

    return ng_ids + pending_ids


def select_batch_with_existence_check(
    conn: sqlite3.Connection,
    batch_size: int,
    *,
    exists_fn: Callable[[str], bool],
) -> list[str]:
    """存在確認を通しながら batch_size 件を確保する。

    - 各候補で exists_fn(entry_id) を呼ぶ
    - 404 の場合は articles.excluded=1, excluded_reason='deleted_from_hatena'
      に更新して代替を選出
    - 代替候補が尽きた場合は取れた分だけ返す
    """
    selected: list[str] = []

    # 初期候補
    candidates = select_batch_entry_ids(conn, batch_size)

    while len(selected) < batch_size and candidates:
        eid = candidates.pop(0)
        if eid in selected:
            continue
        if exists_fn(eid):
            selected.append(eid)
            continue

        # 404 → excluded マーク
        now_iso = datetime.now(timezone.utc).astimezone().isoformat(
            timespec="seconds"
        )
        conn.execute(
            """UPDATE articles
            SET excluded=1,
                excluded_reason='deleted_from_hatena',
                status='excluded',
                last_ng_date=?
            WHERE entry_id=?""",
            (now_iso, eid),
        )
        conn.commit()

        # 補充: 既に selected / 今の候補キュー / excluded になった記事を除いて
        # 足りない分を取り直す
        need = batch_size - len(selected) - len(candidates)
        if need <= 0:
            continue
        # 余裕を持って need*2+1件 取得して、未選択のものを追加
        refill = select_batch_entry_ids(conn, need * 2 + 1)
        for rid in refill:
            if rid not in selected and rid not in candidates:
                candidates.append(rid)
                if len(candidates) + len(selected) >= batch_size:
                    break

    return selected


def generate_daily_batch(
    db_path: Path,
    out_dir: Path,
    *,
    batch_size: int = 20,
    batch_date: str | None = None,
    exists_fn: Callable[[str], bool] | None = None,
) -> tuple[Path, Path]:
    """対象記事を選定し、prompts.csvテンプレと entries_for_claude/<date>.md を生成。

    Args:
        exists_fn: entry_id を受けて存在確認する関数。None ならスキップ。

    Returns:
        (prompts_csv_path, entries_md_path)
    """
    batch_date = batch_date or date.today().isoformat()
    out_dir.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(db_path)
    try:
        if exists_fn is None:
            entry_ids = select_batch_entry_ids(conn, batch_size)
        else:
            entry_ids = select_batch_with_existence_check(
                conn, batch_size, exists_fn=exists_fn
            )
        if not entry_ids:
            raise RuntimeError("選定対象がありません。全完了または在庫切れ。")

        rows: list[dict[str, object]] = []
        md_sections: list[str] = []
        for eid in entry_ids:
            row = conn.execute(
                """SELECT title, category, published_date, url FROM articles
                WHERE entry_id=?""",
                (eid,),
            ).fetchone()
            if not row:
                continue
            title, category, published, url = row
            url = url or ""

            # 前回NG情報
            prev_ng = conn.execute(
                """SELECT ng_tags, ng_note, attempt FROM generations
                WHERE entry_id=? AND decision='ng'
                ORDER BY attempt DESC LIMIT 1""",
                (eid,),
            ).fetchone()
            prev_tags = prev_ng[0] if prev_ng else ""
            prev_note = prev_ng[1] if prev_ng else ""
            next_attempt = (prev_ng[2] if prev_ng else 0) + 1

            rows.append(
                {
                    "entry_id": eid,
                    "title": title or "",
                    "category": category or "",
                    "url": url,
                    "style_summary": "",
                    "prompt": "",
                    "alt": "",
                    "attempt": next_attempt,
                    "previous_ng_tags": prev_tags,
                    "previous_ng_note": prev_note,
                }
            )

            conn.execute(
                """UPDATE articles SET status='in_batch', last_batch_date=?
                WHERE entry_id=?""",
                (batch_date, eid),
            )

            md_sections.append(
                f"## entry_id: {eid}\n"
                f"- title: {title}\n"
                f"- category: {category}\n"
                f"- published: {published}\n"
            )

        conn.commit()

        csv_path = out_dir / f"{batch_date}_prompts.csv"
        with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_HEADER)
            writer.writeheader()
            writer.writerows(rows)

        md_path = out_dir.parent / "entries_for_claude" / f"{batch_date}.md"
        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text("\n".join(md_sections), encoding="utf-8")

        conn.execute(
            """INSERT OR REPLACE INTO batches
            (batch_date, entry_ids, prompts_csv, status)
            VALUES (?, ?, ?, 'generating')""",
            (batch_date, ",".join(entry_ids), str(csv_path)),
        )
        conn.commit()

        return csv_path, md_path
    finally:
        conn.close()


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--batch-size", type=int, default=20)
    parser.add_argument(
        "--batch-date",
        default=None,
        help="YYYY-MM-DD (default: today)。batches/generations の日付キーに使われる。",
    )
    parser.add_argument(
        "--skip-existence-check",
        action="store_true",
        help="AtomPub存在確認を省く（テスト/オフライン用）",
    )
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    out_dir = root / "data" / "batches"

    exists_fn: Callable[[str], bool] | None = None
    if not args.skip_existence_check:
        from scripts.hatena_client import entry_exists, load_env

        env_file = Path(
            r"C:\Users\jsber\OneDrive\Documents\Claude_task"
            r"\youtube-slides\食事指導シリーズ\_shared\.env"
        )
        api_key = load_env(env_file)["HATENA_API_KEY"]
        exists_fn = lambda eid: entry_exists(eid, api_key)  # noqa: E731

    csv_path, md_path = generate_daily_batch(
        db_path,
        out_dir,
        batch_size=args.batch_size,
        batch_date=args.batch_date,
        exists_fn=exists_fn,
    )
    print(f"バッチCSV: {csv_path}")
    print(f"Claude参照MD: {md_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
