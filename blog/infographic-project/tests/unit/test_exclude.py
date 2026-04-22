from __future__ import annotations

import pytest

from scripts.inventory import Entry, should_exclude, strip_tags


def _entry(
    cats: tuple[str, ...] = (),
    body_html: str = "<p>" + "本文" * 300 + "</p>",
) -> Entry:
    return Entry(
        entry_id="1",
        title="t",
        categories=cats,
        published="2025-01-01T00:00:00+09:00",
        body_html=body_html,
        alternate_url="",
    )


@pytest.mark.unit
def test_exclude_diary_category() -> None:
    e = _entry(cats=("日記",))
    excluded, reason = should_exclude(e, strip_tags(e.body_html))
    assert excluded
    assert reason.startswith("category")


@pytest.mark.unit
def test_exclude_short_content() -> None:
    e = _entry(body_html="<p>短すぎる本文</p>")
    excluded, reason = should_exclude(e, strip_tags(e.body_html))
    assert excluded
    assert "short" in reason


@pytest.mark.unit
def test_exclude_index_page() -> None:
    internal_links = "\n".join(
        f'<a href="https://hinyan1016.hatenablog.com/entry/x/{i}">link {i}</a>'
        for i in range(10)
    )
    e = _entry(body_html="<p>" + "目次" * 300 + "</p>" + internal_links)
    excluded, reason = should_exclude(e, strip_tags(e.body_html))
    assert excluded
    assert "index" in reason


@pytest.mark.unit
def test_no_exclude_normal_article() -> None:
    e = _entry(cats=("脳神経内科",))
    excluded, _ = should_exclude(e, strip_tags(e.body_html))
    assert not excluded
