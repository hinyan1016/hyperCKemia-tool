"""日次バッチCSV生成：DBから20件選定してCSVとClaude読込用MDを作る。"""

from __future__ import annotations

import csv
import sqlite3
import sys
from datetime import date
from pathlib import Path


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


def generate_daily_batch(
    db_path: Path,
    out_dir: Path,
    *,
    batch_size: int = 20,
    batch_date: str | None = None,
) -> tuple[Path, Path]:
    """対象記事を選定し、prompts.csvテンプレと entries_for_claude/<date>.md を生成。

    Returns:
        (prompts_csv_path, entries_md_path)
    """
    batch_date = batch_date or date.today().isoformat()
    out_dir.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(db_path)
    try:
        entry_ids = select_batch_entry_ids(conn, batch_size)
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
    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    out_dir = root / "data" / "batches"
    csv_path, md_path = generate_daily_batch(db_path, out_dir, batch_size=20)
    print(f"バッチCSV: {csv_path}")
    print(f"Claude参照MD: {md_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
