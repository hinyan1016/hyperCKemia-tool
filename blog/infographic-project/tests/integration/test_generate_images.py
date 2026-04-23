from __future__ import annotations

import csv
import sqlite3
import tempfile
from pathlib import Path

import pytest

from scripts.generate_images import generate_images_for_batch
from scripts.init_db import initialize_schema


PNG_STUB = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\xff\xff"
    b"\x3f\x00\x05\xfe\x02\xfe\xdc\xccY\xe7\x00\x00\x00\x00IEND\xaeB`\x82"
)


def _write_prompts_csv(
    path: Path,
    rows: list[dict[str, str]],
) -> None:
    header = [
        "entry_id", "title", "category", "url",
        "style_summary", "prompt", "alt",
        "attempt", "previous_ng_tags", "previous_ng_note",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=header)
        w.writeheader()
        for row in rows:
            filled = {k: row.get(k, "") for k in header}
            w.writerow(filled)


def _insert_generation(
    conn: sqlite3.Connection, entry_id: str, attempt: int = 1
) -> None:
    conn.execute(
        """INSERT INTO generations (entry_id, attempt, prompt, alt_text, style_summary)
        VALUES (?, ?, 'p', 'alt', 'style')""",
        (entry_id, attempt),
    )


def _insert_article(conn: sqlite3.Connection, eid: str) -> None:
    conn.execute(
        """INSERT INTO articles
        (entry_id, title, category, published_date, url, status, excluded, word_count)
        VALUES (?, 't', 'c', '2025-03-01', 'u', 'in_batch', 0, 1000)""",
        (eid,),
    )


@pytest.mark.integration
def test_generate_images_writes_png_and_updates_db() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db = tmp / "s.sqlite"
        initialize_schema(db)
        conn = sqlite3.connect(db)
        try:
            _insert_article(conn, "A")
            _insert_article(conn, "B")
            _insert_generation(conn, "A")
            _insert_generation(conn, "B")
            conn.commit()
        finally:
            conn.close()

        csv_path = tmp / "batches" / "2026-04-23_prompts.csv"
        _write_prompts_csv(csv_path, [
            {"entry_id": "A", "prompt": "first prompt"},
            {"entry_id": "B", "prompt": "second prompt"},
        ])
        images_in = tmp / "images" / "in"

        calls: list[tuple[str, str, str, str]] = []

        def fake_gen(prompt, *, model, size, quality):
            calls.append((prompt, model, size, quality))
            return PNG_STUB

        result = generate_images_for_batch(
            csv_path, images_in, db,
            generator=fake_gen,
        )

        assert result.total == 2
        assert result.generated == 2
        assert result.failed == 0
        assert (images_in / "A.png").exists()
        assert (images_in / "B.png").exists()
        assert (images_in / "A.png").read_bytes() == PNG_STUB
        assert len(calls) == 2

        conn = sqlite3.connect(db)
        try:
            paths = dict(conn.execute(
                "SELECT entry_id, image_path FROM generations"
            ).fetchall())
        finally:
            conn.close()
        assert str(images_in / "A.png") in paths["A"]
        assert str(images_in / "B.png") in paths["B"]


@pytest.mark.integration
def test_generate_images_skips_existing_file() -> None:
    """既に images/in/<eid>.png があればAPIを呼ばない（idempotency）。"""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db = tmp / "s.sqlite"
        initialize_schema(db)
        conn = sqlite3.connect(db)
        try:
            _insert_article(conn, "A")
            _insert_generation(conn, "A")
            conn.commit()
        finally:
            conn.close()

        csv_path = tmp / "batches" / "2026-04-23_prompts.csv"
        _write_prompts_csv(csv_path, [
            {"entry_id": "A", "prompt": "some prompt"},
        ])
        images_in = tmp / "images" / "in"
        images_in.mkdir(parents=True)
        (images_in / "A.png").write_bytes(b"existing-bytes")

        calls: list[str] = []

        def fake_gen(prompt, *, model, size, quality):
            calls.append(prompt)
            return PNG_STUB

        result = generate_images_for_batch(
            csv_path, images_in, db,
            generator=fake_gen,
        )

        assert result.generated == 0
        assert result.skipped_existing == 1
        assert calls == []
        # 既存バイト列が保持されている
        assert (images_in / "A.png").read_bytes() == b"existing-bytes"


@pytest.mark.integration
def test_generate_images_skips_empty_prompt() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db = tmp / "s.sqlite"
        initialize_schema(db)

        csv_path = tmp / "batches" / "2026-04-23_prompts.csv"
        _write_prompts_csv(csv_path, [
            {"entry_id": "A", "prompt": ""},
            {"entry_id": "B", "prompt": "   "},
        ])

        def fake_gen(prompt, *, model, size, quality):
            pytest.fail("should not be called for empty prompts")

        result = generate_images_for_batch(
            csv_path, tmp / "images" / "in", db,
            generator=fake_gen,
        )
        assert result.skipped_empty_prompt == 2
        assert result.generated == 0


@pytest.mark.integration
def test_generate_images_retries_then_fails() -> None:
    """失敗時はリトライしてから failed 計上、次行へ進む。"""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db = tmp / "s.sqlite"
        initialize_schema(db)
        conn = sqlite3.connect(db)
        try:
            _insert_article(conn, "A")
            _insert_article(conn, "B")
            _insert_generation(conn, "A")
            _insert_generation(conn, "B")
            conn.commit()
        finally:
            conn.close()

        csv_path = tmp / "batches" / "2026-04-23_prompts.csv"
        _write_prompts_csv(csv_path, [
            {"entry_id": "A", "prompt": "bad prompt"},
            {"entry_id": "B", "prompt": "good prompt"},
        ])
        images_in = tmp / "images" / "in"

        call_log: list[str] = []

        def fake_gen(prompt, *, model, size, quality):
            call_log.append(prompt)
            if "bad" in prompt:
                raise RuntimeError("api err")
            return PNG_STUB

        log_path = tmp / "logs" / "image_gen.log"
        result = generate_images_for_batch(
            csv_path, images_in, db,
            generator=fake_gen,
            max_retries=1,
            retry_delay_s=0,
            log_path=log_path,
        )
        assert result.generated == 1
        assert result.failed == 1
        assert result.failures[0][0] == "A"
        # A が 2回試行されている (初回 + 1リトライ)
        assert call_log.count("bad prompt") == 2
        assert (images_in / "B.png").exists()
        assert not (images_in / "A.png").exists()
        # エラーログ書き出し
        assert log_path.exists()
        assert "A" in log_path.read_text(encoding="utf-8")
