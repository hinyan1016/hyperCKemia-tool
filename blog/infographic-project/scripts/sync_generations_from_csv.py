"""prompts.csv の prompt/alt/style_summary を generations テーブルに UPSERT する。

process_decisions._handle_ok は generations 行が必須（alt_text 参照）のため、
Claude がプロンプトを記入した直後にこのスクリプトで同期する。

同期の方針:
- (entry_id, attempt) で既存行を検索
- 既に decision が set されている（approved / ng / skip）行は触らない
- まだ decision が空の既存行は prompt/alt_text/style_summary を更新
- 存在しなければ新規 INSERT（created_at も記録）

既存の手動INSERT運用（Day 2試行）を恒久化する位置づけ。
"""

from __future__ import annotations

import csv
import sqlite3
import sys
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path


@dataclass
class SyncResult:
    inserted: int = 0
    updated: int = 0
    skipped_processed: int = 0
    skipped_empty: int = 0


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def sync_from_csv(db_path: Path, prompts_csv: Path) -> SyncResult:
    result = SyncResult()
    conn = sqlite3.connect(db_path)
    try:
        with prompts_csv.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))

        for row in rows:
            eid = (row.get("entry_id") or "").strip()
            prompt = (row.get("prompt") or "").strip()
            alt = (row.get("alt") or "").strip()
            style = (row.get("style_summary") or "").strip()
            try:
                attempt = int(row.get("attempt") or 1)
            except ValueError:
                attempt = 1

            if not eid or not prompt:
                result.skipped_empty += 1
                continue

            existing = conn.execute(
                """SELECT gen_id, decision FROM generations
                WHERE entry_id=? AND attempt=?""",
                (eid, attempt),
            ).fetchone()

            if existing:
                gen_id, decision = existing
                if decision:
                    # 既に approved/ng などの判定済。上書きしない
                    result.skipped_processed += 1
                    continue
                conn.execute(
                    """UPDATE generations
                    SET prompt=?, alt_text=?, style_summary=?
                    WHERE gen_id=?""",
                    (prompt, alt, style, gen_id),
                )
                result.updated += 1
            else:
                conn.execute(
                    """INSERT INTO generations
                    (entry_id, attempt, prompt, alt_text, style_summary, created_at)
                    VALUES (?, ?, ?, ?, ?, ?)""",
                    (eid, attempt, prompt, alt, style, _now()),
                )
                result.inserted += 1

        conn.commit()
    finally:
        conn.close()

    return result


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--date",
        default=None,
        help="YYYY-MM-DD (default: today)。指定日の prompts.csv を同期。",
    )
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    batch_date = args.date or date.today().isoformat()
    prompts_csv = root / "data" / "batches" / f"{batch_date}_prompts.csv"

    if not prompts_csv.exists():
        print(f"[ERROR] バッチCSVが見つかりません: {prompts_csv}", file=sys.stderr)
        return 2

    result = sync_from_csv(db_path, prompts_csv)
    print(
        f"sync結果: INSERT={result.inserted} UPDATE={result.updated} "
        f"既判定スキップ={result.skipped_processed} 空行スキップ={result.skipped_empty}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
