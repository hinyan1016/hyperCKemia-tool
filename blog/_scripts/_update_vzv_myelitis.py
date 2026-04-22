#!/usr/bin/env python3
"""VZV脊髄炎ブログ記事を更新(インフォグラフィック追加)"""
from __future__ import annotations

import base64
import sys
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Final

ENV_FILE: Final = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env"
)
HATENA_HTML: Final = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\posts\vzv_myelitis_blog_hatena.html"
)
TARGET_URL: Final = "https://hinyan1016.hatenablog.com/entry/2026/04/22/162339"
TITLE: Final = (
    "VZVによるウイルス性脊髄炎とは｜症状・診断・治療・予防をわかりやすく解説"
)
CATEGORIES: Final = ["帯状疱疹", "神経内科", "感染症", "脊髄"]


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def auth_header(env: dict[str, str]) -> str:
    auth_str = f"{env['HATENA_ID']}:{env['HATENA_API_KEY']}"
    return "Basic " + base64.b64encode(auth_str.encode()).decode()


def find_entry_edit_url(env: dict[str, str], target_url: str) -> str | None:
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_url: str | None = (
        f"https://blog.hatena.ne.jp/{hatena_id}/{blog_domain}/atom/entry"
    )
    ns = {"atom": "http://www.w3.org/2005/Atom", "app": "http://www.w3.org/2007/app"}

    page = 0
    while api_url and page < 20:
        page += 1
        req = urllib.request.Request(
            api_url, headers={"Authorization": auth_header(env)}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read().decode("utf-8")
        root = ET.fromstring(body)

        for entry in root.findall("atom:entry", ns):
            alt_link = None
            edit_link = None
            for link in entry.findall("atom:link", ns):
                if link.get("rel") == "alternate":
                    alt_link = link.get("href")
                elif link.get("rel") == "edit":
                    edit_link = link.get("href")
            if alt_link == target_url and edit_link:
                return edit_link

        next_url: str | None = None
        for link in root.findall("atom:link", ns):
            if link.get("rel") == "next":
                next_url = link.get("href")
                break
        api_url = next_url
        print(f"  [page {page}] next = {next_url[:80] + '...' if next_url else 'END'}")
    return None


def build_atom_entry(title: str, content: str, categories: list[str], draft: bool = True) -> str:
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join(f'  <category term="{c}" />' for c in categories)
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">\n'
        f"  <title>{title}</title>\n"
        "  <author><name>hinyan1016</name></author>\n"
        f"  <content type=\"text/html\"><![CDATA[{content}]]></content>\n"
        f"{cat_xml}\n"
        "  <app:control>\n"
        f"    <app:draft>{draft_val}</app:draft>\n"
        "  </app:control>\n"
        "</entry>\n"
    )


def put_entry(edit_url: str, env: dict[str, str], xml_body: str) -> tuple[bool, str]:
    req = urllib.request.Request(
        edit_url,
        data=xml_body.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth_header(env),
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8", errors="replace")
        return True, body
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.read().decode('utf-8', errors='replace')[:800]}"


def main() -> None:
    env = load_env()
    print(f"[INFO] Finding edit URL for {TARGET_URL}...")
    edit_url = find_entry_edit_url(env, TARGET_URL)
    if not edit_url:
        print("[FAIL] Edit URL not found")
        sys.exit(1)
    print(f"[OK] edit_url = {edit_url}")

    content = HATENA_HTML.read_text(encoding="utf-8")
    xml = build_atom_entry(TITLE, content, list(CATEGORIES), draft=True)
    ok, body = put_entry(edit_url, env, xml)
    if ok:
        print("[SUCCESS] entry updated (draft)")
    else:
        print(f"[FAIL] {body}")
        sys.exit(1)


if __name__ == "__main__":
    main()
