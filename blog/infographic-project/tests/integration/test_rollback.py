from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from scripts.init_db import initialize_schema
from scripts.rollback import rollback_entry


@pytest.mark.integration
def test_rollback_restores_previous_body(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            """INSERT INTO articles
            (entry_id, title, status, body_html_before)
            VALUES ('E1', 't', 'done', '<p>元の本文</p>')"""
        )
        conn.commit()
    finally:
        conn.close()

    puts: list[tuple[str, str]] = []

    def fake_put(entry_id: str, xml: str, api_key: str, **_: object) -> str:
        puts.append((entry_id, xml))
        return "OK"

    def fake_fetch(entry_id: str, api_key: str, **_: object) -> str:
        return (
            '<entry xmlns="http://www.w3.org/2005/Atom">'
            '<title>t</title>'
            '<category term="cat" />'
            '<content type="text/html">&lt;p&gt;新本文&lt;/p&gt;</content>'
            '</entry>'
        )

    rollback_entry(
        db_path, "E1", api_key="k", fetch_fn=fake_fetch, put_fn=fake_put
    )
    assert len(puts) == 1
    assert "<p>元の本文</p>" in puts[0][1]

    # ロールバック後は body_html_before がNULL、status が rolled_back
    conn = sqlite3.connect(db_path)
    try:
        row = conn.execute(
            "SELECT status, body_html_before FROM articles WHERE entry_id='E1'"
        ).fetchone()
    finally:
        conn.close()
    assert row == ("rolled_back", None)


@pytest.mark.integration
def test_rollback_raises_without_backup(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            "INSERT INTO articles (entry_id, title, status) VALUES ('E2', 't', 'done')"
        )
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(RuntimeError, match="body_html_before"):
        rollback_entry(
            db_path, "E2", api_key="k",
            fetch_fn=lambda *a, **k: "<entry></entry>",
            put_fn=lambda *a, **k: "OK",
        )
