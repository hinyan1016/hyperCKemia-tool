from __future__ import annotations

import base64

import pytest

from scripts.hatena_client import (
    basic_auth_header,
    build_entry_put_xml,
    make_wsse_header,
)


@pytest.mark.unit
def test_make_wsse_header_format() -> None:
    header = make_wsse_header(
        "testuser",
        "testkey",
        nonce_str="abc",
        created="2026-04-23T00:00:00Z",
    )
    assert header.startswith('UsernameToken Username="testuser",')
    assert 'PasswordDigest="' in header
    assert 'Nonce="' in header
    assert 'Created="2026-04-23T00:00:00Z"' in header


@pytest.mark.unit
def test_make_wsse_header_digest_reproducible() -> None:
    h1 = make_wsse_header(
        "u", "k", nonce_str="xyz", created="2026-01-01T00:00:00Z"
    )
    h2 = make_wsse_header(
        "u", "k", nonce_str="xyz", created="2026-01-01T00:00:00Z"
    )
    assert h1 == h2


@pytest.mark.unit
def test_basic_auth_header() -> None:
    header = basic_auth_header("user", "key")
    assert header.startswith("Basic ")
    encoded = header[len("Basic ") :]
    assert base64.b64decode(encoded).decode() == "user:key"


@pytest.mark.unit
def test_build_entry_put_xml_preserves_title_and_categories() -> None:
    xml = build_entry_put_xml(
        title="テストタイトル & test",
        author="hinyan1016",
        body_html="<p>本文</p>",
        categories=("脳神経内科", "医師向け"),
    )
    assert "<title>テストタイトル &amp; test</title>" in xml
    assert '<category term="脳神経内科"' in xml
    assert '<category term="医師向け"' in xml
    assert "<![CDATA[<p>本文</p>]]>" in xml
    assert "<app:draft>no</app:draft>" in xml
