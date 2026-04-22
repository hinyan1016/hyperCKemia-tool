from __future__ import annotations

import pytest

from scripts.html_inserter import insert_infographic


IMG_TAG = '<img src="https://cdn-ak.f.st-hatena.com/images/fotolife/h/x.png" alt="テストalt">'


@pytest.mark.unit
def test_insert_with_toc_marker() -> None:
    body = "<p>リード文</p>\n<!-- 目次 -->\n<h2>セクション1</h2>"
    result = insert_infographic(body, IMG_TAG)
    assert "<!-- インフォグラフィック -->" in result
    assert result.index("<img") < result.index("<!-- 目次 -->")


@pytest.mark.unit
def test_insert_fallback_to_h2() -> None:
    body = "<p>リード文</p>\n<p>続き</p>\n<h2>セクション1</h2>"
    result = insert_infographic(body, IMG_TAG)
    assert result.index("<img") < result.index("<h2>")


@pytest.mark.unit
def test_insert_fallback_to_first_p() -> None:
    body = "<p>リード文だけ</p>\n<p>本文</p>"
    result = insert_infographic(body, IMG_TAG)
    first_p_end = result.index("</p>") + len("</p>")
    rest = result[first_p_end:].lstrip()
    assert rest.startswith("<!-- インフォグラフィック -->")


@pytest.mark.unit
def test_insert_fallback_to_prepend() -> None:
    body = "<div>pタグも見出しもない</div>"
    result = insert_infographic(body, IMG_TAG)
    assert result.startswith("\n<!-- インフォグラフィック -->")


@pytest.mark.unit
def test_insert_preserves_existing_content() -> None:
    body = "<p>リード</p>\n<!-- 目次 -->\n<h2>本文</h2>\n<p>結び</p>"
    result = insert_infographic(body, IMG_TAG)
    assert "<p>リード</p>" in result
    assert "<h2>本文</h2>" in result
    assert "<p>結び</p>" in result
