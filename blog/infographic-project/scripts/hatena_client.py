"""はてなブログ AtomPub / Fotolife クライアント。

既存の blog/_scripts/upload_and_insert_infographic.py のロジックを
再利用可能な形に整理したもの。
"""

from __future__ import annotations

import base64
import hashlib
import random
import re
import string
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


HATENA_ID_DEFAULT = "hinyan1016"
BLOG_DOMAIN_DEFAULT = "hinyan1016.hatenablog.com"


def load_env(env_file: Path) -> dict[str, str]:
    env: dict[str, str] = {}
    with env_file.open(encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def _escape_xml_text(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def make_wsse_header(
    username: str,
    api_key: str,
    *,
    nonce_str: Optional[str] = None,
    created: Optional[str] = None,
) -> str:
    if nonce_str is None:
        nonce_str = "".join(
            random.choices(string.ascii_letters + string.digits, k=40)
        )
    if created is None:
        created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    nonce_bytes = nonce_str.encode("utf-8")
    digest_raw = hashlib.sha1(
        nonce_bytes + created.encode("utf-8") + api_key.encode("utf-8")
    ).digest()
    password_digest = base64.b64encode(digest_raw).decode("utf-8")
    nonce_b64 = base64.b64encode(nonce_bytes).decode("utf-8")
    return (
        f'UsernameToken Username="{username}", '
        f'PasswordDigest="{password_digest}", '
        f'Nonce="{nonce_b64}", '
        f'Created="{created}"'
    )


def basic_auth_header(username: str, api_key: str) -> str:
    token = base64.b64encode(f"{username}:{api_key}".encode()).decode()
    return f"Basic {token}"


def build_entry_put_xml(
    *,
    title: str,
    author: str,
    body_html: str,
    categories: tuple[str, ...],
    draft: bool = False,
) -> str:
    cats_xml = "\n".join(
        f'  <category term="{_escape_xml_text(c)}" />' for c in categories
    )
    draft_value = "yes" if draft else "no"
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<entry xmlns="http://www.w3.org/2005/Atom" '
        'xmlns:app="http://www.w3.org/2007/app">\n'
        f"  <title>{_escape_xml_text(title)}</title>\n"
        f"  <author><name>{_escape_xml_text(author)}</name></author>\n"
        f"  <content type=\"text/html\"><![CDATA[{body_html}]]></content>\n"
        f"{cats_xml}\n"
        "  <app:control>\n"
        f"    <app:draft>{draft_value}</app:draft>\n"
        "  </app:control>\n"
        "</entry>"
    )


def fetch_entry(
    entry_id: str,
    api_key: str,
    *,
    hatena_id: str = HATENA_ID_DEFAULT,
    blog_domain: str = BLOG_DOMAIN_DEFAULT,
) -> str:
    url = (
        f"https://blog.hatena.ne.jp/{hatena_id}/{blog_domain}"
        f"/atom/entry/{entry_id}"
    )
    req = urllib.request.Request(
        url,
        method="GET",
        headers={"Authorization": basic_auth_header(hatena_id, api_key)},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def entry_exists(
    entry_id: str,
    api_key: str,
    *,
    hatena_id: str = HATENA_ID_DEFAULT,
    blog_domain: str = BLOG_DOMAIN_DEFAULT,
) -> bool:
    """AtomPub GET で記事の存在確認。404ならFalse、2xxならTrue。

    他のHTTPエラーは例外を伝播させる（認証失敗などの早期発見のため）。
    """
    try:
        fetch_entry(
            entry_id,
            api_key,
            hatena_id=hatena_id,
            blog_domain=blog_domain,
        )
        return True
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return False
        raise


def put_entry(
    entry_id: str,
    xml_body: str,
    api_key: str,
    *,
    hatena_id: str = HATENA_ID_DEFAULT,
    blog_domain: str = BLOG_DOMAIN_DEFAULT,
) -> str:
    url = (
        f"https://blog.hatena.ne.jp/{hatena_id}/{blog_domain}"
        f"/atom/entry/{entry_id}"
    )
    req = urllib.request.Request(
        url,
        data=xml_body.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": basic_auth_header(hatena_id, api_key),
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def upload_to_fotolife(
    image_path: Path,
    api_key: str,
    *,
    title: str = "infographic",
    hatena_id: str = HATENA_ID_DEFAULT,
) -> str:
    """Fotolife に PNG をアップロードし、画像URLを返す。

    失敗時は 'ERROR:' または 'URL not found:' プレフィクスの文字列を返す。
    """
    image_data = image_path.read_bytes()
    image_b64 = base64.b64encode(image_data).decode("utf-8")

    xml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<entry xmlns="http://purl.org/atom/ns#">\n'
        f"  <title>{_escape_xml_text(title)}</title>\n"
        f'  <content mode="base64" type="image/png">{image_b64}</content>\n'
        "</entry>"
    )

    url = "https://f.hatena.ne.jp/atom/post"
    wsse = make_wsse_header(hatena_id, api_key)

    req = urllib.request.Request(
        url,
        data=xml.encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "X-WSSE": wsse,
            "Authorization": 'WSSE profile="UsernameToken"',
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="replace")
        return f"ERROR:HTTP {e.code}: {err[:500]}"

    m = re.search(r"<hatena:imageurl>([^<]+)</hatena:imageurl>", body)
    if m:
        return m.group(1)
    m = re.search(r"https://cdn-ak\.f\.st-hatena\.com/images/fotolife/[^<\"]+", body)
    if m:
        return m.group(0)
    return f"URL not found:{body[:500]}"
