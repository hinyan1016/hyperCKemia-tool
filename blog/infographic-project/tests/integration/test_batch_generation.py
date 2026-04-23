from __future__ import annotations

import csv
import sqlite3
import tempfile
from pathlib import Path

import pytest

from scripts.generate_batch_csv import generate_daily_batch
from scripts.init_db import initialize_schema


def _insert_article(
    conn: sqlite3.Connection,
    entry_id: str,
    *,
    status: str = "pending",
    published: str = "2025-03-01T00:00:00+09:00",
    title: str = "title",
    url: str = "",
) -> None:
    conn.execute(
        """INSERT INTO articles
        (entry_id, title, category, published_date, url, status, excluded, word_count)
        VALUES (?, ?, 'cat', ?, ?, ?, 0, 1000)""",
        (entry_id, title, published, url, status),
    )


@pytest.mark.integration
def test_generate_batch_picks_oldest_pending_first() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db_path = tmp / "s.sqlite"
        out_dir = tmp / "batches"
        initialize_schema(db_path)

        conn = sqlite3.connect(db_path)
        try:
            _insert_article(conn, "A", published="2025-01-01T00:00:00+09:00")
            _insert_article(conn, "B", published="2025-02-01T00:00:00+09:00")
            _insert_article(conn, "C", published="2025-03-01T00:00:00+09:00")
            conn.commit()
        finally:
            conn.close()

        csv_path, md_path = generate_daily_batch(
            db_path, out_dir, batch_size=2, batch_date="2026-04-23"
        )
        with csv_path.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        assert [r["entry_id"] for r in rows] == ["A", "B"]

        conn = sqlite3.connect(db_path)
        try:
            statuses = dict(
                conn.execute("SELECT entry_id, status FROM articles").fetchall()
            )
        finally:
            conn.close()
        assert statuses == {
            "A": "in_batch",
            "B": "in_batch",
            "C": "pending",
        }


@pytest.mark.integration
def test_generate_batch_uses_db_url_not_date_guess() -> None:
    """CSVのURL列はarticles.urlから読む（公開日からの推測ではない）。"""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db_path = tmp / "s.sqlite"
        out_dir = tmp / "batches"
        initialize_schema(db_path)

        real_url = "https://hinyan1016.hatenablog.com/entry/2025/03/07/044435"
        conn = sqlite3.connect(db_path)
        try:
            _insert_article(
                conn, "X",
                published="2025-03-07T04:44:35+09:00",
                url=real_url,
            )
            conn.commit()
        finally:
            conn.close()

        csv_path, _ = generate_daily_batch(
            db_path, out_dir, batch_size=1, batch_date="2026-04-23"
        )
        with csv_path.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        assert rows[0]["url"] == real_url


@pytest.mark.integration
def test_generate_batch_url_empty_when_db_url_missing() -> None:
    """articles.urlが空ならCSVのurlも空（日付推測でフォールバックしない）。"""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db_path = tmp / "s.sqlite"
        out_dir = tmp / "batches"
        initialize_schema(db_path)

        conn = sqlite3.connect(db_path)
        try:
            _insert_article(
                conn, "X",
                published="2025-03-07T04:44:35+09:00",
                url="",
            )
            conn.commit()
        finally:
            conn.close()

        csv_path, _ = generate_daily_batch(
            db_path, out_dir, batch_size=1, batch_date="2026-04-23"
        )
        with csv_path.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        assert rows[0]["url"] == ""


@pytest.mark.integration
def test_generate_batch_with_existence_check_skips_deleted_and_refills() -> None:
    """exists_fn で404の記事は excluded になり、代替で batch_size を満たす。"""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db_path = tmp / "s.sqlite"
        out_dir = tmp / "batches"
        initialize_schema(db_path)

        conn = sqlite3.connect(db_path)
        try:
            for i, eid in enumerate(["A", "B", "C", "D", "E"]):
                _insert_article(
                    conn, eid,
                    published=f"2025-0{i + 1}-01T00:00:00+09:00",
                )
            conn.commit()
        finally:
            conn.close()

        # B だけ "削除済み" として404扱い
        deleted = {"B"}

        def exists_fn(eid: str) -> bool:
            return eid not in deleted

        csv_path, _ = generate_daily_batch(
            db_path, out_dir,
            batch_size=3, batch_date="2026-04-23",
            exists_fn=exists_fn,
        )
        with csv_path.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        entry_ids = [r["entry_id"] for r in rows]
        assert "B" not in entry_ids
        assert len(entry_ids) == 3
        # A, C, D が選ばれるはず (B がスキップされて D が補充)
        assert entry_ids == ["A", "C", "D"]

        conn = sqlite3.connect(db_path)
        try:
            b_row = conn.execute(
                "SELECT excluded, excluded_reason, status FROM articles WHERE entry_id='B'"
            ).fetchone()
        finally:
            conn.close()
        assert b_row == (1, "deleted_from_hatena", "excluded")


@pytest.mark.integration
def test_generate_batch_includes_ng_revision() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db_path = tmp / "s.sqlite"
        out_dir = tmp / "batches"
        initialize_schema(db_path)

        conn = sqlite3.connect(db_path)
        try:
            for i, eid in enumerate(["A", "B", "C", "D"]):
                _insert_article(
                    conn,
                    eid,
                    published=f"2025-0{i + 1}-01T00:00:00+09:00",
                )
            _insert_article(
                conn, "NG1", status="ng_pending",
                published="2024-12-01T00:00:00+09:00",
            )
            conn.execute(
                "UPDATE articles SET last_ng_date='2026-04-22' WHERE entry_id='NG1'"
            )
            conn.commit()
        finally:
            conn.close()

        csv_path, _ = generate_daily_batch(
            db_path, out_dir, batch_size=5, batch_date="2026-04-23"
        )
        with csv_path.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        entry_ids = [r["entry_id"] for r in rows]
        assert "NG1" in entry_ids
        assert entry_ids.index("NG1") == 0
        assert "A" in entry_ids and "B" in entry_ids
