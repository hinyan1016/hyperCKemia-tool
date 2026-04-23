#!/usr/bin/env python3
"""
その症状大丈夫#03 力が入らない
- 公開済み記事にYouTube iframe + VideoObject JSON-LDを追加
- ローカル 03_weakness_hatena.html (埋め込み挿入済み) を丸ごとAtomPub PUT
- REQUIRE_DRAFT=False: 公開済み記事を対象
"""

import re
import sys
import hashlib
import base64
import random
import string
import urllib.request
import urllib.error
from datetime import datetime, timezone
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"

HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\symptom-concern-series\03_weakness_hatena.html")
TITLE_KEYWORD = "脳？ 神経？ 筋肉？"
REQUIRE_DRAFT = False
TARGET_URL_FRAGMENT = "2026/04/23/133800"


def auth_header():
    b64 = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()
    return "Basic {}".format(b64)


def find_entry_id():
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(HATENA_ID, BLOG_DOMAIN)
    page = 0
    while url and page < 30:
        req = urllib.request.Request(url, headers={"Authorization": auth_header()})
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
        entries = re.findall(r'<entry>(.*?)</entry>', body, re.DOTALL)
        for e in entries:
            title_m = re.search(r'<title[^>]*>([^<]*)</title>', e)
            if not title_m or TITLE_KEYWORD not in title_m.group(1):
                continue
            draft_m = re.search(r'<app:draft>(yes|no)</app:draft>', e)
            is_draft = (draft_m.group(1) == "yes") if draft_m else False
            if REQUIRE_DRAFT and not is_draft:
                print("      候補スキップ（draft=no要求されたが公開済み）: {}".format(title_m.group(1)))
                continue
            alt_m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', e)
            alt_url = alt_m.group(1) if alt_m else ""
            if TARGET_URL_FRAGMENT and TARGET_URL_FRAGMENT not in alt_url:
                print("      候補スキップ（URLにfragment一致なし）: {} / {}".format(title_m.group(1), alt_url))
                continue
            id_m = re.search(r'<id>tag:blog\.hatena\.ne\.jp[^:]*:[^:]*-(\d+)</id>', e)
            if id_m:
                return id_m.group(1), title_m.group(1), alt_url
            id_m = re.search(r'<id>([^<]+)</id>', e)
            if id_m:
                return id_m.group(1).rsplit("-", 1)[-1], title_m.group(1), alt_url
        next_m = re.search(r'<link rel="next" href="([^"]+)"', body)
        url = next_m.group(1) if next_m else None
        page += 1
    raise RuntimeError("対象記事が見つかりません（TITLE_KEYWORD={}, REQUIRE_DRAFT={}, URL_FRAG={}）".format(TITLE_KEYWORD, REQUIRE_DRAFT, TARGET_URL_FRAGMENT))


def get_entry(entry_id):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, headers={"Authorization": auth_header()})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def put_entry(entry_id, title, content, categories, is_draft):
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)
    xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml += '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">\n'
    xml += '  <title>{}</title>\n'.format(title)
    xml += '  <author><name>{}</name></author>\n'.format(HATENA_ID)
    xml += '  <content type="text/html"><![CDATA[{}]]></content>\n'.format(content)
    xml += '{}\n'.format(cat_xml)
    xml += '  <app:control>\n    <app:draft>{}</app:draft>\n  </app:control>\n'.format("yes" if is_draft else "no")
    xml += '</entry>'

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(
        url,
        data=xml.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth_header(),
        },
    )
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")
    url_m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
    return url_m.group(1) if url_m else "(URL不明)"


def main():
    print("[1/3] AtomPubで対象エントリを検索中...")
    entry_id, title, alt_url = find_entry_id()
    print("      entry_id={} / title={} / url={}".format(entry_id, title, alt_url))

    print("[2/3] 既存エントリからcategoryとdraft状態を取得...")
    current = get_entry(entry_id)
    categories = re.findall(r'<category[^>]*term="([^"]+)"', current)
    draft_m = re.search(r'<app:draft>(yes|no)</app:draft>', current)
    is_draft = (draft_m.group(1) == "yes") if draft_m else False
    DEFAULT_CATS = ["その症状大丈夫", "脳神経内科", "セルフチェック"]
    if not categories:
        categories = DEFAULT_CATS
        print("      categories空→既定値補完: {}".format(categories))
    print("      categories={} / draft={}".format(categories, is_draft))

    html = HTML_FILE.read_text(encoding="utf-8")
    if "<!-- youtube embed -->" not in html or "ldXs_bJaLnc" not in html:
        raise RuntimeError("ローカルHTMLにiframe埋め込みが見つかりません。先に埋め込みを挿入してください。")

    print("[3/3] AtomPub PUT中（draft={}）...".format("yes" if is_draft else "no"))
    result = put_entry(entry_id, title, html, categories, is_draft)
    print("      更新成功: {}".format(result))


if __name__ == "__main__":
    try:
        main()
    except urllib.error.HTTPError as e:
        print("HTTP Error {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:500]))
        sys.exit(1)
