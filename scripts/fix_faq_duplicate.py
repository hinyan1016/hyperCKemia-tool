#!/usr/bin/env python3
"""FAQPage重複エラー修正: 3記事のFAQ JSON-LDを条件付きスクリプトに置換"""

import re
import base64
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"
AUTH = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()

ATOM_NS = "http://www.w3.org/2005/Atom"
APP_NS = "http://www.w3.org/2007/app"

# 修正対象記事のURL
TARGET_URLS = [
    "https://hinyan1016.hatenablog.com/entry/2026/04/17/063136",
    "https://hinyan1016.hatenablog.com/entry/2026/04/16/193734",
    "https://hinyan1016.hatenablog.com/entry/2026/04/16/125321",
]


def api_request(url, method="GET", data=None):
    headers = {
        "Authorization": "Basic {}".format(AUTH),
    }
    if data:
        headers["Content-Type"] = "application/xml; charset=utf-8"
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def find_entries():
    """AtomPub APIから最近のエントリを取得し、対象記事のedit URLを返す"""
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(HATENA_ID, BLOG_DOMAIN)
    body = api_request(url)
    root = ET.fromstring(body)

    results = {}
    for entry in root.findall("{%s}entry" % ATOM_NS):
        # alternate link = 記事の公開URL
        alt_link = None
        edit_link = None
        for link in entry.findall("{%s}link" % ATOM_NS):
            if link.get("rel") == "alternate":
                alt_link = link.get("href")
            elif link.get("rel") == "edit":
                edit_link = link.get("href")

        if alt_link in TARGET_URLS:
            title_el = entry.find("{%s}title" % ATOM_NS)
            title = title_el.text if title_el is not None else "?"
            results[alt_link] = {"edit_url": edit_link, "title": title}

    return results


def replace_faq_in_content(content):
    """静的FAQ JSON-LDを条件付きスクリプトに置換"""
    # パターン: <script type="application/ld+json">...FAQPage...</script>
    pattern = r'<script\s+type=["\']application/ld\+json["\']>\s*(\{[^<]*"@type"\s*:\s*"FAQPage"[^<]*\})\s*</script>'

    def replacer(match):
        json_str = match.group(1).strip()
        # JSON文字列内のシングルクォートをエスケープ（念のため）
        # また & を \\u0026 にエスケープ（JSON内で安全に）
        safe_json = json_str.replace("'", "\\'")
        return "<script>if(location.pathname.indexOf('/entry/')===0){var s=document.createElement('script');s.type='application/ld+json';s.textContent='" + safe_json + "';document.head.appendChild(s);}</script>"

    new_content, count = re.subn(pattern, replacer, content, flags=re.DOTALL)
    return new_content, count


def update_entry(edit_url, title):
    """エントリを取得→FAQ修正→PUT"""
    print("  GET {}".format(edit_url))
    body = api_request(edit_url)

    # contentを抽出
    root = ET.fromstring(body)
    content_el = root.find("{%s}content" % ATOM_NS)
    if content_el is None or not content_el.text:
        print("  [スキップ] contentが空")
        return False

    original = content_el.text
    modified, count = replace_faq_in_content(original)

    if count == 0:
        print("  [スキップ] FAQPage JSON-LDが見つからない（既に修正済み？）")
        return False

    print("  FAQ JSON-LD {}件を条件付きスクリプトに置換".format(count))

    # XMLを再構築してPUT（content全体を丸ごと置換）
    content_el.text = modified

    # ElementTreeのnamespace prefix問題回避
    ET.register_namespace("", ATOM_NS)
    ET.register_namespace("app", APP_NS)

    xml_bytes = ET.tostring(root, encoding="unicode", xml_declaration=True).encode("utf-8")

    print("  PUT {}".format(edit_url))
    api_request(edit_url, method="PUT", data=xml_bytes)
    print("  [成功] 更新完了")
    return True


def main():
    print("=== FAQPage重複エラー修正 ===\n")

    print("1. 対象エントリを検索中...")
    entries = find_entries()
    print("   {}件のエントリを発見\n".format(len(entries)))

    if len(entries) < len(TARGET_URLS):
        missing = set(TARGET_URLS) - set(entries.keys())
        print("   [警告] 以下のエントリが見つかりません:")
        for u in missing:
            print("   - {}".format(u))
        print()

    for url, info in entries.items():
        print("2. 修正中: {}".format(info["title"][:50]))
        try:
            update_entry(info["edit_url"], info["title"])
        except urllib.error.HTTPError as e:
            error_body = e.read().decode("utf-8", errors="replace")
            print("  [失敗] HTTP {}: {}".format(e.code, error_body[:300]))
        print()

    print("=== 完了 ===")


if __name__ == "__main__":
    main()
