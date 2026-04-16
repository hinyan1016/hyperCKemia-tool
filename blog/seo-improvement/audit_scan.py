#!/usr/bin/env python3
"""
SEO改善 Phase 1: 全記事スキャン・データ収集
AtomPub APIで全記事を取得し、以下を分析してCSV出力:
- 記事URL、タイトル、公開日、カテゴリ
- 本文文字数（HTMLタグ除去後）
- YouTube埋め込み有無・動画ID
- 内部リンク数（hinyan1016.hatenablog.com宛）
- 下書きかどうか
"""

import sys
import re
import csv
import time
import base64
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from pathlib import Path
from html.parser import HTMLParser

sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
ATOM_NS = "http://www.w3.org/2005/Atom"
APP_NS = "http://www.w3.org/2007/app"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
OUTPUT_DIR = Path(__file__).parent
OUTPUT_CSV = OUTPUT_DIR / "audit_report.csv"


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


class HTMLTextExtractor(HTMLParser):
    """HTMLからテキストのみ抽出"""
    def __init__(self):
        super().__init__()
        self.text_parts = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._skip = True

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self._skip = False

    def handle_data(self, data):
        if not self._skip:
            self.text_parts.append(data)

    def get_text(self):
        return "".join(self.text_parts).strip()


def strip_html(html_content):
    """HTMLタグを除去してテキスト文字数を返す"""
    if not html_content:
        return 0, ""
    extractor = HTMLTextExtractor()
    try:
        extractor.feed(html_content)
        text = extractor.get_text()
        return len(text), text
    except Exception:
        # フォールバック: タグ除去
        text = re.sub(r'<[^>]+>', '', html_content)
        text = text.strip()
        return len(text), text


def extract_youtube_ids(html_content):
    """YouTube動画IDを抽出"""
    if not html_content:
        return []
    ids = set()
    # iframe src
    for m in re.finditer(r'youtube\.com/embed/([a-zA-Z0-9_-]{11})', html_content):
        ids.add(m.group(1))
    # youtu.be リンク
    for m in re.finditer(r'youtu\.be/([a-zA-Z0-9_-]{11})', html_content):
        ids.add(m.group(1))
    # youtube.com/watch?v=
    for m in re.finditer(r'youtube\.com/watch\?v=([a-zA-Z0-9_-]{11})', html_content):
        ids.add(m.group(1))
    return list(ids)


def count_internal_links(html_content):
    """内部リンク（自ブログ宛）の数を数える"""
    if not html_content:
        return 0
    pattern = r'href="https?://{}'.format(re.escape(BLOG_DOMAIN))
    return len(re.findall(pattern, html_content))


def has_json_ld_video(html_content):
    """既にVideoObject JSON-LDがあるか"""
    if not html_content:
        return False
    return '"VideoObject"' in html_content or "'VideoObject'" in html_content


def get_all_entries(env):
    """AtomPub APIで全記事データを取得"""
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    auth_str = "{}:{}".format(hatena_id, env["HATENA_API_KEY"])
    auth_b64 = base64.b64encode(auth_str.encode()).decode()
    auth = "Basic {}".format(auth_b64)

    entries = []
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
    page = 0

    while url:
        page += 1
        req = urllib.request.Request(url, headers={"Authorization": auth})
        try:
            with urllib.request.urlopen(req) as resp:
                body = resp.read().decode("utf-8")
        except urllib.error.HTTPError as e:
            print("  API エラー (page {}): HTTP {}".format(page, e.code))
            break

        root = ET.fromstring(body)

        for entry in root.findall("{%s}entry" % ATOM_NS):
            # タイトル
            title_el = entry.find("{%s}title" % ATOM_NS)
            title = title_el.text if title_el is not None and title_el.text else ""

            # URL
            link_alt = ""
            link_edit = ""
            for link in entry.findall("{%s}link" % ATOM_NS):
                rel = link.get("rel", "")
                if rel == "alternate":
                    link_alt = link.get("href", "")
                elif rel == "edit":
                    link_edit = link.get("href", "")

            # 公開日
            published_el = entry.find("{%s}published" % ATOM_NS)
            published = published_el.text[:10] if published_el is not None and published_el.text else ""

            # カテゴリ
            categories = []
            for cat in entry.findall("{%s}category" % ATOM_NS):
                term = cat.get("term", "")
                if term:
                    categories.append(term)

            # 本文
            content_el = entry.find("{%s}content" % ATOM_NS)
            content_html = ""
            if content_el is not None and content_el.text:
                content_html = content_el.text

            # 下書き判定
            control = entry.find("{%s}control" % APP_NS)
            is_draft = False
            if control is not None:
                draft_el = control.find("{%s}draft" % APP_NS)
                if draft_el is not None and draft_el.text == "yes":
                    is_draft = True

            # 分析
            char_count, _ = strip_html(content_html)
            youtube_ids = extract_youtube_ids(content_html)
            internal_links = count_internal_links(content_html)
            has_video_ld = has_json_ld_video(content_html)

            entries.append({
                "url": link_alt,
                "edit_url": link_edit,
                "title": title,
                "published": published,
                "categories": "|".join(categories),
                "is_draft": is_draft,
                "char_count": char_count,
                "has_youtube": len(youtube_ids) > 0,
                "youtube_ids": "|".join(youtube_ids),
                "youtube_count": len(youtube_ids),
                "internal_links": internal_links,
                "has_video_jsonld": has_video_ld,
            })

        # 次ページ
        next_url = None
        for link in root.findall("{%s}link" % ATOM_NS):
            if link.get("rel") == "next":
                next_url = link.get("href")
        url = next_url

        count = len(entries)
        print("  page {} 完了 ... 累計 {} 件".format(page, count))
        time.sleep(0.3)

    return entries


def write_csv(entries):
    """CSV出力"""
    fields = [
        "url", "title", "published", "categories", "is_draft",
        "char_count", "has_youtube", "youtube_count", "youtube_ids",
        "internal_links", "has_video_jsonld", "edit_url",
    ]
    with open(OUTPUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for entry in entries:
            writer.writerow(entry)
    print("\nCSV出力: {}".format(OUTPUT_CSV))


def print_summary(entries):
    """サマリーレポート表示"""
    total = len(entries)
    published = [e for e in entries if not e["is_draft"]]
    drafts = [e for e in entries if e["is_draft"]]

    with_youtube = [e for e in published if e["has_youtube"]]
    without_video_ld = [e for e in with_youtube if not e["has_video_jsonld"]]

    thin_500 = [e for e in published if e["char_count"] < 500]
    thin_1000 = [e for e in published if e["char_count"] < 1000]

    no_internal = [e for e in published if e["internal_links"] == 0]

    # 文字数分布
    bins = [0, 500, 1000, 2000, 3000, 5000, 10000, 999999]
    bin_labels = ["0-499", "500-999", "1000-1999", "2000-2999", "3000-4999", "5000-9999", "10000+"]
    bin_counts = [0] * len(bin_labels)
    for e in published:
        for i in range(len(bins) - 1):
            if bins[i] <= e["char_count"] < bins[i + 1]:
                bin_counts[i] += 1
                break

    print("\n" + "=" * 60)
    print("SEO監査レポート")
    print("=" * 60)
    print("\n--- 基本統計 ---")
    print("全記事数: {}".format(total))
    print("  公開: {}".format(len(published)))
    print("  下書き: {}".format(len(drafts)))

    print("\n--- YouTube埋め込み (Phase 2対象) ---")
    print("YouTube埋め込みあり: {} 件".format(len(with_youtube)))
    print("  うちVideoObject JSON-LD未設定: {} 件 ← Phase 2で対応".format(len(without_video_ld)))

    print("\n--- コンテンツ量 (Phase 3対象) ---")
    print("文字数分布 (公開記事):")
    for label, count in zip(bin_labels, bin_counts):
        bar = "#" * min(count, 50)
        print("  {:>12s}: {:4d} {}".format(label, count, bar))
    print("500文字未満: {} 件 ← 薄いコンテンツ候補".format(len(thin_500)))
    print("1000文字未満: {} 件".format(len(thin_1000)))

    print("\n--- 内部リンク (Phase 4対象) ---")
    print("内部リンクゼロの孤立記事: {} 件".format(len(no_internal)))

    # 薄いコンテンツ上位表示
    if thin_500:
        print("\n--- 500文字未満の記事 (上位20件) ---")
        sorted_thin = sorted(thin_500, key=lambda e: e["char_count"])
        for i, e in enumerate(sorted_thin[:20], 1):
            print("{:2d}. [{}文字] {}".format(i, e["char_count"], e["title"][:50]))
            print("    {}".format(e["url"]))

    # 孤立記事上位表示
    if no_internal:
        print("\n--- 内部リンクゼロ記事 (上位20件, 文字数順) ---")
        sorted_orphan = sorted(no_internal, key=lambda e: e["char_count"])
        for i, e in enumerate(sorted_orphan[:20], 1):
            print("{:2d}. [{}文字] {}".format(i, e["char_count"], e["title"][:50]))

    print("\n" + "=" * 60)
    print("詳細データ: {}".format(OUTPUT_CSV))
    print("=" * 60)


def main():
    print("SEO監査スキャン開始")
    print("=" * 40)

    env = load_env()
    print("\n[1/3] 全記事データ取得中...")
    entries = get_all_entries(env)
    print("取得完了: {} 件".format(len(entries)))

    print("\n[2/3] CSV出力中...")
    write_csv(entries)

    print("\n[3/3] サマリーレポート生成中...")
    print_summary(entries)


if __name__ == "__main__":
    main()
