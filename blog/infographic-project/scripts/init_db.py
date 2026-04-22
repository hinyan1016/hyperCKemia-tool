"""SQLite state.sqlite の初期化と targets_v2.csv の投入。"""

from __future__ import annotations

import csv
import sqlite3
import sys
from pathlib import Path


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS articles (
  entry_id TEXT PRIMARY KEY,
  title TEXT,
  category TEXT,
  published_date TEXT,
  body_html_path TEXT,
  body_html_before TEXT,
  word_count INTEGER,
  has_infographic INTEGER DEFAULT 0,
  excluded INTEGER DEFAULT 0,
  excluded_reason TEXT,
  status TEXT DEFAULT 'pending',
  last_batch_date TEXT,
  last_ng_date TEXT,
  edited_at TEXT
);

CREATE TABLE IF NOT EXISTS generations (
  gen_id INTEGER PRIMARY KEY,
  entry_id TEXT,
  attempt INTEGER,
  prompt TEXT,
  alt_text TEXT,
  style_summary TEXT,
  image_path TEXT,
  decision TEXT,
  ng_tags TEXT,
  ng_note TEXT,
  fotolife_url TEXT,
  created_at TEXT,
  processed_at TEXT,
  FOREIGN KEY (entry_id) REFERENCES articles(entry_id)
);

CREATE TABLE IF NOT EXISTS batches (
  batch_date TEXT PRIMARY KEY,
  entry_ids TEXT,
  prompts_csv TEXT,
  dashboard_html TEXT,
  decisions_csv TEXT,
  status TEXT
);

CREATE TABLE IF NOT EXISTS processed_decisions (
  decisions_csv_path TEXT,
  entry_id TEXT,
  decision TEXT,
  processed_at TEXT,
  PRIMARY KEY (decisions_csv_path, entry_id)
);
"""


def initialize_schema(db_path: Path) -> None:
    """state.sqlite を作成、スキーマを適用。既存DBは破壊しない。"""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(SCHEMA_SQL)
        conn.commit()
    finally:
        conn.close()


def load_targets_csv(db_path: Path, csv_path: Path) -> int:
    """targets_v2.csv の行を articles テーブルに INSERT OR IGNORE で投入。"""
    conn = sqlite3.connect(db_path)
    inserted = 0
    try:
        with csv_path.open(encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                cur = conn.execute(
                    """INSERT OR IGNORE INTO articles
                    (entry_id, title, category, published_date, word_count, status)
                    VALUES (?, ?, ?, ?, ?, 'pending')""",
                    (
                        row["entry_id"],
                        row["title"],
                        row["category"],
                        row["published"],
                        int(row.get("word_count") or 0),
                    ),
                )
                if cur.rowcount > 0:
                    inserted += 1
        conn.commit()
    finally:
        conn.close()
    return inserted


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    csv_path = root / "data" / "targets_v2.csv"

    initialize_schema(db_path)
    print(f"スキーマ初期化: {db_path}")

    if csv_path.exists():
        count = load_targets_csv(db_path, csv_path)
        print(f"articles テーブルに {count} 件投入")
    else:
        print(f"[警告] {csv_path} がありません。inventory スキャンを先に実行してください。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
