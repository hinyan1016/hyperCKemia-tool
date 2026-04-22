from __future__ import annotations

import csv
import sqlite3
from pathlib import Path

import pytest

from scripts.init_db import initialize_schema
from scripts.process_decisions import process_decisions


def _seed_article(conn: sqlite3.Connection, eid: str) -> None:
    conn.execute(
        """INSERT INTO articles
        (entry_id, title, category, published_date, status, excluded)
        VALUES (?, 'タイトル', 'カテゴリ', '2025-01-01T00:00:00+09:00', 'in_batch', 0)""",
        (eid,),
    )


def _seed_generation(
    conn: sqlite3.Connection, eid: str, attempt: int = 1
) -> None:
    conn.execute(
        """INSERT INTO generations (entry_id, attempt, alt_text, created_at)
        VALUES (?, ?, 'alt', '2026-04-23T00:00:00+09:00')""",
        (eid, attempt),
    )


def _write_decisions_csv(
    path: Path, rows: list[tuple[str, str, str, str]]
) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["entry_id", "decision", "ng_tags", "ng_note"])
        w.writerows(rows)


def _make_fake_clients():
    class FakeClient:
        def __init__(self) -> None:
            self.upload_count = 0
            self.put_count = 0

        def upload(self, image_path: Path, api_key: str, **_: object) -> str:
            self.upload_count += 1
            return f"https://cdn-ak.f.st-hatena.com/fake/{image_path.stem}.png"

        def fetch(self, entry_id: str, api_key: str, **_: object) -> str:
            return (
                '<?xml version="1.0" encoding="utf-8"?>'
                '<entry xmlns="http://www.w3.org/2005/Atom">'
                f'<title>t{entry_id}</title>'
                '<content type="text/html">&lt;p&gt;本文&lt;/p&gt;&lt;!-- 目次 --&gt;</content>'
                '<category term="c" />'
                '</entry>'
            )

        def put(self, entry_id: str, xml: str, api_key: str, **_: object) -> str:
            self.put_count += 1
            return "OK"

    return FakeClient()


@pytest.mark.integration
def test_ok_decision_uploads_and_updates(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        _seed_article(conn, "E1")
        _seed_generation(conn, "E1")
        conn.commit()
    finally:
        conn.close()

    images_in = tmp_path / "images" / "in"
    images_in.mkdir(parents=True)
    (images_in / "E1.png").write_bytes(b"fake-png")
    archive = tmp_path / "images" / "archive"
    archive.mkdir()

    decisions = tmp_path / "dec.csv"
    _write_decisions_csv(decisions, [("E1", "ok", "", "")])

    fake = _make_fake_clients()
    result = process_decisions(
        db_path,
        decisions,
        images_in,
        archive,
        api_key="k",
        upload_fn=fake.upload,
        fetch_fn=fake.fetch,
        put_fn=fake.put,
    )
    assert result.uploaded == 1
    assert result.ng_recorded == 0
    assert not (images_in / "E1.png").exists()
    assert (archive / "E1.png").exists()

    conn = sqlite3.connect(db_path)
    try:
        row = conn.execute(
            "SELECT status FROM articles WHERE entry_id='E1'"
        ).fetchone()
    finally:
        conn.close()
    assert row == ("done",)


@pytest.mark.integration
def test_idempotent_same_csv_twice(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        _seed_article(conn, "E1")
        _seed_generation(conn, "E1")
        conn.commit()
    finally:
        conn.close()

    images_in = tmp_path / "images" / "in"
    images_in.mkdir(parents=True)
    (images_in / "E1.png").write_bytes(b"fake-png")
    archive = tmp_path / "images" / "archive"
    archive.mkdir()

    decisions = tmp_path / "dec.csv"
    _write_decisions_csv(decisions, [("E1", "ok", "", "")])

    fake = _make_fake_clients()

    r1 = process_decisions(
        db_path, decisions, images_in, archive, api_key="k",
        upload_fn=fake.upload, fetch_fn=fake.fetch, put_fn=fake.put,
    )
    r2 = process_decisions(
        db_path, decisions, images_in, archive, api_key="k",
        upload_fn=fake.upload, fetch_fn=fake.fetch, put_fn=fake.put,
    )
    assert r1.uploaded == 1
    assert r2.uploaded == 0
    assert r2.already_processed == 1
    # INV-4: 2回目は put_fn / upload_fn を呼ばない
    assert fake.put_count == 1
    assert fake.upload_count == 1


@pytest.mark.integration
def test_ng_decision_records_and_marks_revision(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        _seed_article(conn, "E2")
        _seed_generation(conn, "E2")
        conn.commit()
    finally:
        conn.close()

    images_in = tmp_path / "images" / "in"
    images_in.mkdir(parents=True)
    (images_in / "E2.png").write_bytes(b"fake")
    archive = tmp_path / "images" / "archive"
    archive.mkdir()

    decisions = tmp_path / "dec.csv"
    _write_decisions_csv(
        decisions, [("E2", "ng", "medical_error,typo", "神経名の綴り誤り")]
    )

    fake = _make_fake_clients()
    result = process_decisions(
        db_path, decisions, images_in, archive, api_key="k",
        upload_fn=fake.upload, fetch_fn=fake.fetch, put_fn=fake.put,
    )
    assert result.ng_recorded == 1

    conn = sqlite3.connect(db_path)
    try:
        article = conn.execute(
            "SELECT status FROM articles WHERE entry_id='E2'"
        ).fetchone()
        assert article == ("ng_pending",)
        gen = conn.execute(
            "SELECT decision, ng_tags, ng_note FROM generations WHERE entry_id='E2'"
        ).fetchone()
    finally:
        conn.close()
    assert gen == ("ng", "medical_error,typo", "神経名の綴り誤り")
    assert not (images_in / "E2.png").exists()
