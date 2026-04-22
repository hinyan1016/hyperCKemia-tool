from __future__ import annotations

from pathlib import Path

import pytest

from scripts.inventory import detect_v2, parse_entry_xml


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "entries"


def _body_from_fixture(name: str) -> str:
    xml = (FIXTURES / name).read_text(encoding="utf-8")
    entry = parse_entry_xml(xml)
    assert entry is not None, f"fixture parse failed: {name}"
    return entry.body_html


@pytest.mark.unit
@pytest.mark.parametrize(
    "fixture_name, expected_status, expected_reason_part",
    [
        # VZV記事: alt="...にまとめたインフォグラフィック" → alt_marker_infographic
        ("with_infographic_comment.xml", "has_ig_strong", "alt_marker"),
        # ヤンキー記事: figure-image-fotolife クラス → fotolife_figure_head
        ("with_fotolife_figure.xml", "has_ig_likely", "fotolife"),
        # ヘパリン記事: alt="...のまとめ" → alt_marker_summary（strong）
        ("iframe_then_fotolife.xml", "has_ig_strong", "alt_marker_summary"),
        ("no_image.xml", "needs_ig", "no_image"),
        ("only_external_link.xml", "needs_ig", "no_image"),
    ],
)
def test_detect_v2_classifies_correctly(
    fixture_name: str, expected_status: str, expected_reason_part: str
) -> None:
    body = _body_from_fixture(fixture_name)
    status, reason = detect_v2(body)
    assert status == expected_status, (
        f"{fixture_name}: expected {expected_status}, got {status} ({reason})"
    )
    assert expected_reason_part in reason, (
        f"{fixture_name}: reason {reason!r} missing {expected_reason_part!r}"
    )


@pytest.mark.unit
def test_detect_v2_ignores_iframe_content() -> None:
    body = (
        '<p>lead</p>'
        '<iframe width="560" height="315" src="https://youtube.com/...">'
        '<img src="fake.png" width="600"></iframe>'
        '<h2>本文</h2>'
    )
    status, _ = detect_v2(body)
    assert status == "needs_ig"
