from __future__ import annotations

import sqlite3
import tempfile
from pathlib import Path

from scripts.init_db import initialize_schema, load_targets_csv


def test_schema_creates_four_tables():
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.sqlite"
        initialize_schema(db_path)
        conn = sqlite3.connect(db_path)
        try:
            tables = {
                row[0]
                for row in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }
        finally:
            conn.close()
        assert tables == {
            "articles",
            "generations",
            "batches",
            "processed_decisions",
        }


def test_load_targets_marks_articles_pending():
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.sqlite"
        csv_path = Path(tmpdir) / "targets.csv"
        csv_path.write_text(
            "entry_id,title,category,published,word_count,reason,url\n"
            "123,テスト記事,脳神経内科,2025-03-07T04:20:53+09:00,8582,no_image_in_head,https://x.y/z\n",
            encoding="utf-8-sig",
        )
        initialize_schema(db_path)
        count = load_targets_csv(db_path, csv_path)
        assert count == 1
        conn = sqlite3.connect(db_path)
        try:
            row = conn.execute(
                "SELECT entry_id, status, title FROM articles WHERE entry_id='123'"
            ).fetchone()
        finally:
            conn.close()
        assert row == ("123", "pending", "テスト記事")


def test_load_targets_idempotent():
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.sqlite"
        csv_path = Path(tmpdir) / "targets.csv"
        csv_path.write_text(
            "entry_id,title,category,published,word_count,reason,url\n"
            "123,テスト,cat,2025-01-01T00:00:00+09:00,1000,no_image,https://x\n",
            encoding="utf-8-sig",
        )
        initialize_schema(db_path)
        load_targets_csv(db_path, csv_path)
        load_targets_csv(db_path, csv_path)
        conn = sqlite3.connect(db_path)
        try:
            count = conn.execute("SELECT COUNT(*) FROM articles").fetchone()[0]
        finally:
            conn.close()
        assert count == 1
