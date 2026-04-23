from __future__ import annotations

import csv
import sqlite3
from pathlib import Path

from scripts.init_db import initialize_schema
from scripts.sync_generations_from_csv import sync_from_csv


HEADER = [
    "entry_id", "title", "category", "url",
    "style_summary", "prompt", "alt",
    "attempt", "previous_ng_tags", "previous_ng_note",
]


def _write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=HEADER)
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in HEADER})


def _setup(tmp_path: Path) -> tuple[Path, Path]:
    db = tmp_path / "state.sqlite"
    initialize_schema(db)
    csv_path = tmp_path / "prompts.csv"
    return db, csv_path


def test_insert_when_missing(tmp_path: Path) -> None:
    db, csv_path = _setup(tmp_path)
    _write_csv(csv_path, [
        {"entry_id": "e1", "prompt": "p1", "alt": "a1", "style_summary": "s1", "attempt": "1"},
    ])

    result = sync_from_csv(db, csv_path)

    assert result.inserted == 1
    assert result.updated == 0
    assert result.skipped_processed == 0

    conn = sqlite3.connect(db)
    row = conn.execute(
        "SELECT prompt, alt_text, style_summary, attempt FROM generations WHERE entry_id='e1'"
    ).fetchone()
    conn.close()
    assert row == ("p1", "a1", "s1", 1)


def test_update_when_undecided(tmp_path: Path) -> None:
    db, csv_path = _setup(tmp_path)
    conn = sqlite3.connect(db)
    conn.execute(
        """INSERT INTO generations (entry_id, attempt, prompt, alt_text, style_summary)
        VALUES ('e1', 1, 'old', 'old_alt', 'old_style')"""
    )
    conn.commit()
    conn.close()

    _write_csv(csv_path, [
        {"entry_id": "e1", "prompt": "new_p", "alt": "new_alt", "style_summary": "new_s", "attempt": "1"},
    ])
    result = sync_from_csv(db, csv_path)

    assert result.inserted == 0
    assert result.updated == 1

    conn = sqlite3.connect(db)
    row = conn.execute(
        "SELECT prompt, alt_text, style_summary FROM generations WHERE entry_id='e1'"
    ).fetchone()
    conn.close()
    assert row == ("new_p", "new_alt", "new_s")


def test_skip_when_already_processed(tmp_path: Path) -> None:
    db, csv_path = _setup(tmp_path)
    conn = sqlite3.connect(db)
    conn.execute(
        """INSERT INTO generations
        (entry_id, attempt, prompt, alt_text, style_summary, decision)
        VALUES ('e1', 1, 'orig', 'orig_alt', 'orig_s', 'approved')"""
    )
    conn.commit()
    conn.close()

    _write_csv(csv_path, [
        {"entry_id": "e1", "prompt": "should_not_overwrite", "alt": "x", "style_summary": "y", "attempt": "1"},
    ])
    result = sync_from_csv(db, csv_path)

    assert result.skipped_processed == 1
    assert result.updated == 0
    assert result.inserted == 0

    conn = sqlite3.connect(db)
    row = conn.execute(
        "SELECT prompt, alt_text FROM generations WHERE entry_id='e1'"
    ).fetchone()
    conn.close()
    assert row == ("orig", "orig_alt")


def test_different_attempt_creates_new_row(tmp_path: Path) -> None:
    db, csv_path = _setup(tmp_path)
    conn = sqlite3.connect(db)
    conn.execute(
        """INSERT INTO generations
        (entry_id, attempt, prompt, alt_text, style_summary, decision)
        VALUES ('e1', 1, 'first', 'first_alt', 'first_s', 'ng')"""
    )
    conn.commit()
    conn.close()

    _write_csv(csv_path, [
        {"entry_id": "e1", "prompt": "second", "alt": "second_alt", "style_summary": "second_s", "attempt": "2"},
    ])
    result = sync_from_csv(db, csv_path)

    assert result.inserted == 1

    conn = sqlite3.connect(db)
    rows = conn.execute(
        "SELECT attempt, prompt FROM generations WHERE entry_id='e1' ORDER BY attempt"
    ).fetchall()
    conn.close()
    assert rows == [(1, "first"), (2, "second")]


def test_empty_prompt_skipped(tmp_path: Path) -> None:
    db, csv_path = _setup(tmp_path)
    _write_csv(csv_path, [
        {"entry_id": "e1", "prompt": "", "alt": "", "style_summary": "", "attempt": "1"},
        {"entry_id": "e2", "prompt": "valid", "alt": "a", "style_summary": "s", "attempt": "1"},
    ])
    result = sync_from_csv(db, csv_path)

    assert result.skipped_empty == 1
    assert result.inserted == 1

    conn = sqlite3.connect(db)
    count = conn.execute("SELECT COUNT(*) FROM generations").fetchone()[0]
    conn.close()
    assert count == 1
