#!/usr/bin/env python3
"""
Phase 6: 404 URLへの内部リンクを検出するスクリプト。
AtomPub APIで全記事を取得し、404 URLが内部リンクとして参照されていないか確認する。
"""

import sys
import re
import csv
import time
import base64
import html as html_mod
import urllib.request
import urllib.error
from pathlib import Path

sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)

ENV_FILE = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task"
    r"\youtube-slides\食事指導シリーズ\_shared\.env"
)
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
OUTPUT_CSV = Path(__file__).parent / "broken_links_found.csv"

# 検索対象の404 URL（entry系のみ — pagination/rss/feedは内部リンク対象外）
TARGET_404_URLS = [
    "https://hinyan1016.hatenablog.com/entry/2026/04/09/154436",
    "https://hinyan1016.hatenablog.com/entry/2026/03/24/152604",
    "https://hinyan1016.hatenablog.com/entry/2026/04/04/073318",
    "https://hinyan1016.hatenablog.com/entry/2025/10/03/195935",
    "https://hinyan1016.hatenablog.com/entry/2025/10/21/190655",
    "https://hinyan1016.hatenablog.com/entry/2025/10/16/185512",
    "https://hinyan1016.hatenablog.com/entry/2026/03/11/145004",
    "https://hinyan1016.hatenablog.com/entry/2025/10/08/215946",
    "https://hinyan1016.hatenablog.com/entry/2026/03/27/213946",
    "https://hinyan1016.hatenablog.com/entry/2026/03/29/184905",
    "https://hinyan1016.hatenablog.com/entry/2026/02/10/181109",
    "https://hinyan1016.hatenablog.com/entry/2026/03/29/191433",
    "https://hinyan1016.hatenablog.com/entry/2026/03/09/040507",
    "https://hinyan1016.hatenablog.com/entry/2026/03/17/062532",
    "https://hinyan1016.hatenablog.com/entry/2026/02/12/145825",
    "https://hinyan1016.hatenablog.com/entry/2026/03/24/160851",
    "https://hinyan1016.hatenablog.com/entry/2026/03/12/101552",
    "https://hinyan1016.hatenablog.com/entry/2026/02/26/160217",
    "https://hinyan1016.hatenablog.com/entry/2026/02/26/160118",
    "https://hinyan1016.hatenablog.com/entry/2025/12/16/115230",
    "https://hinyan1016.hatenablog.com/entry/2026/03/05/external_seizure_prophylaxis",
    "https://hinyan1016.hatenablog.com/entry/2026/03/05/arachnoid_cyst",
    "https://hinyan1016.hatenablog.com/entry/2026/03/05/vestibular_rehab",
    "https://hinyan1016.hatenablog.com/entry/2026/03/05/transfer_family",
    # .html系の不正URLも検索
    "01_forgetfulness_vs_dementia_blog.html",
    "02_migraine_5_things_blog.html",
    "03_sleep_disorders_blog.html",
    "alzheimer-disease",
]

# 検索用パターン（URLの末尾部分で部分一致検索）
SEARCH_PATTERNS = []
for url in TARGET_404_URLS:
    # entryパスの部分一致用
    if "/entry/" in url:
        path = url.split("hatenablog.com")[1]
        SEARCH_PATTERNS.append(path)
    else:
        SEARCH_PATTERNS.append(url)


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def get_auth(env):
    auth_str = "{}:{}".format(HATENA_ID, env["HATENA_API_KEY"])
    return "Basic {}".format(base64.b64encode(auth_str.encode()).decode())


def fetch_feed_page(auth, page_url):
    """AtomPubフィードの1ページを取得"""
    req = urllib.request.Request(page_url, headers={"Authorization": auth})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def extract_entries(feed_xml):
    """フィードXMLからエントリのURL・タイトル・コンテンツを抽出"""
    entries = []
    # エントリごとに分割
    entry_blocks = re.findall(r"<entry>(.*?)</entry>", feed_xml, re.DOTALL)
    for block in entry_blocks:
        # URL (alternate link)
        url_m = re.search(r'<link[^>]+rel="alternate"[^>]+href="([^"]+)"', block)
        url = url_m.group(1) if url_m else ""

        # タイトル
        title_m = re.search(r"<title[^>]*>([^<]+)</title>", block)
        title = html_mod.unescape(title_m.group(1)) if title_m else ""

        # コンテンツ
        content_m = re.search(r"<content[^>]*>(.*?)</content>", block, re.DOTALL)
        content = ""
        if content_m:
            content = content_m.group(1)
            if content.startswith("<![CDATA[") and content.endswith("]]>"):
                content = content[9:-3]
            content = html_mod.unescape(content)

        # entry_id
        id_m = re.search(r"<id>tag:blog.hatena.ne.jp,2013:blog-hinyan1016-(\d+)-(\d+)</id>", block)
        entry_id = id_m.group(2) if id_m else ""

        # draft status
        is_draft = "<app:draft>yes</app:draft>" in block

        entries.append({
            "url": url,
            "title": title,
            "content": content,
            "entry_id": entry_id,
            "is_draft": is_draft,
        })
    return entries


def get_next_page(feed_xml):
    """次ページのURLを取得（なければNone）"""
    m = re.search(r'<link[^>]+rel="next"[^>]+href="([^"]+)"', feed_xml)
    return html_mod.unescape(m.group(1)) if m else None


def search_content(content, url):
    """コンテンツ内で404 URLへのリンクを検索"""
    found = []
    for pattern in SEARCH_PATTERNS:
        if pattern in content:
            # href属性内での出現を確認
            href_pattern = 'href="[^"]*{}[^"]*"'.format(re.escape(pattern))
            matches = re.findall(href_pattern, content)
            if matches:
                found.append((pattern, matches))
    return found


def main():
    env = load_env()
    auth = get_auth(env)

    base_url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(HATENA_ID, BLOG_DOMAIN)
    page_url = base_url
    page_num = 0
    total_entries = 0
    broken_links = []

    print("=== 404 URLへの内部リンク検索 ===\n")

    while page_url:
        page_num += 1
        print(f"Page {page_num}: {page_url[:80]}...", flush=True)

        try:
            feed_xml = fetch_feed_page(auth, page_url)
        except Exception as e:
            print(f"  ERROR: {e}")
            break

        entries = extract_entries(feed_xml)
        total_entries += len(entries)
        print(f"  → {len(entries)}件取得（累計{total_entries}件）")

        for entry in entries:
            if entry["is_draft"]:
                continue  # 下書き記事はスキップ
            found = search_content(entry["content"], entry["url"])
            if found:
                for pattern, matches in found:
                    print(f"  *** FOUND: {entry['title'][:40]} → {pattern}")
                    broken_links.append({
                        "source_url": entry["url"],
                        "source_title": entry["title"],
                        "source_entry_id": entry["entry_id"],
                        "broken_link_pattern": pattern,
                        "match_count": len(matches),
                        "matches": "; ".join(matches[:3]),
                    })

        page_url = get_next_page(feed_xml)
        if page_url:
            time.sleep(0.5)

    # 結果出力
    print(f"\n=== 検索完了 ===")
    print(f"スキャン記事数: {total_entries}")
    print(f"壊れた内部リンク: {len(broken_links)}件")

    if broken_links:
        with open(OUTPUT_CSV, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=[
                "source_url", "source_title", "source_entry_id",
                "broken_link_pattern", "match_count", "matches",
            ])
            writer.writeheader()
            writer.writerows(broken_links)
        print(f"結果: {OUTPUT_CSV}")
    else:
        print("壊れた内部リンクは見つかりませんでした。")


if __name__ == "__main__":
    main()
