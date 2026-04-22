from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from scripts.build_dashboard import build_dashboard


SAMPLE_PROMPTS_CSV = (
    "entry_id,title,category,url,style_summary,prompt,alt,attempt,previous_ng_tags,previous_ng_note\n"
    "111,片頭痛の5つのこと,神経,https://x/y,青系,img prompt,alt text,1,,\n"
    "222,ヘパリンの使い方,循環器,https://x/y2,緑系,img prompt2,alt text2,1,,\n"
)


@pytest.mark.unit
def test_dashboard_includes_all_entries() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        csv_path = tmp / "prompts.csv"
        csv_path.write_text(SAMPLE_PROMPTS_CSV, encoding="utf-8-sig")
        images_dir = tmp / "images" / "in"
        images_dir.mkdir(parents=True)
        (images_dir / "111.png").write_bytes(b"fake")

        out_html = tmp / "review.html"
        build_dashboard(csv_path, images_dir, out_html)

        html = out_html.read_text(encoding="utf-8")
        assert "片頭痛の5つのこと" in html
        assert "ヘパリンの使い方" in html
        assert 'data-entry-id="111"' in html
        assert 'data-entry-id="222"' in html


@pytest.mark.unit
def test_dashboard_marks_missing_images() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        csv_path = tmp / "prompts.csv"
        csv_path.write_text(SAMPLE_PROMPTS_CSV, encoding="utf-8-sig")
        images_dir = tmp / "images" / "in"
        images_dir.mkdir(parents=True)
        # 111.png のみ配置、222.png は未到着

        out_html = tmp / "review.html"
        build_dashboard(csv_path, images_dir, out_html)

        html = out_html.read_text(encoding="utf-8")
        assert 'data-status="awaiting_image"' in html


@pytest.mark.unit
def test_dashboard_has_export_button_and_ng_tags() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        csv_path = tmp / "prompts.csv"
        csv_path.write_text(SAMPLE_PROMPTS_CSV, encoding="utf-8-sig")
        images_dir = tmp / "images" / "in"
        images_dir.mkdir(parents=True)

        out_html = tmp / "review.html"
        build_dashboard(csv_path, images_dir, out_html)

        html = out_html.read_text(encoding="utf-8")
        assert 'id="export-button"' in html
        assert "medical_error" in html
        assert "typo" in html
        assert "design" in html
        assert "mismatch" in html
