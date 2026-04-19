#!/usr/bin/env python3
"""2026年てんかん論文アップデート記事を更新（画像差替え）"""
from __future__ import annotations

import base64
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Final

ENV_FILE: Final = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env"
)
HATENA_HTML: Final = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\posts\epilepsy_2026_papers_hatena.html"
)
TARGET_URL: Final = "https://hinyan1016.hatenablog.com/entry/2026/04/19/001753"
TITLE: Final = (
    "2026年てんかん論文アップデート｜遺伝子治療・新規ASM・SUDEPバイオマーカー"
    "を網羅する重要10本"
)
CATEGORIES: Final = ["医師向け", "脳神経内科", "医学情報"]


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

    # AtomPub paginates by ~7 entries per page; follow rel="next"
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

        # Follow feed-level rel="next" for pagination
        next_url: str | None = None
        for link in root.findall("atom:link", ns):
            if link.get("rel") == "next":
                next_url = link.get("href")
                break
        api_url = next_url
        print(f"  [page {page}] next = {next_url[:80] + '...' if next_url else 'END'}")
    return None


def update_entry(
    env: dict[str, str],
    edit_url: str,
    title: str,
    content: str,
    categories: list[str],
    draft: bool = True,
) -> bool:
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join(f'  <category term="{c}" />' for c in categories)
    xml = f"""<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{cat_xml}
  <app:control>
    <app:draft>{draft_val}</app:draft>
  </app:control>
</entry>"""

    data = xml.encode("utf-8")
    req = urllib.request.Request(
        edit_url,
        data=data,
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth_header(env),
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            print(f"  更新成功: HTTP {resp.status}")
            return True
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"  更新失敗: HTTP {e.code}: {body[:400]}")
        return False


KNOWN_EDIT_URL: Final = (
    "https://blog.hatena.ne.jp/hinyan1016/hinyan1016.hatenablog.com"
    "/atom/entry/17179246901378035186"
)


def main() -> None:
    env = load_env()
    print(f"[1/3] エントリ特定: {TARGET_URL}")
    edit_url = KNOWN_EDIT_URL
    print(f"      edit_url = {edit_url}")

    print("[2/3] HTML読み込み")
    content = HATENA_HTML.read_text(encoding="utf-8")
    print(f"      content length = {len(content)} chars")

    print("[3/3] 更新中（下書きのまま）")
    update_entry(env, edit_url, TITLE, content, list(CATEGORIES), draft=True)


if __name__ == "__main__":
    main()
