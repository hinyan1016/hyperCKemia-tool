#!/usr/bin/env python3
"""
インフォグラフィック未対応記事の棚卸しスキャン。

- AtomPub APIで全エントリを取得（ローカルキャッシュ）
- 各エントリを判定: has_infographic / grey_zone / no_infographic
- 除外基準: 日記/雑記カテゴリ・500字未満・インデックスページ
- 出力: CSVとコンソールサマリ
"""

from __future__ import annotations

import base64
import hashlib
import random
import re
import string
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

ROOT = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task")
ENV_FILE = ROOT / "youtube-slides" / "食事指導シリーズ" / "_shared" / ".env"
PROJECT_DIR = ROOT / "blog" / "infographic-project"
ENTRIES_DIR = PROJECT_DIR / "data" / "entries"
OUTPUT_CSV = PROJECT_DIR / "data" / "inventory.csv"
SUMMARY_FILE = PROJECT_DIR / "data" / "inventory_summary.txt"

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"

EXCLUDED_CATEGORIES = {"日記", "雑記", "お知らせ", "ブログ運営", "プライベート"}
SHORT_CONTENT_THRESHOLD = 500


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def basic_auth(username: str, api_key: str) -> str:
    token = base64.b64encode(f"{username}:{api_key}".encode()).decode()
    return f"Basic {token}"


def fetch_page(url: str, auth: str) -> bytes:
    req = urllib.request.Request(url, method="GET", headers={"Authorization": auth})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


@dataclass(frozen=True)
class Entry:
    entry_id: str
    title: str
    categories: tuple[str, ...]
    published: str
    updated: str
    draft: bool
    body_html: str
    alternate_url: str


def parse_entry_xml(entry_xml: str) -> Optional[Entry]:
    """AtomPubのentry XMLをEntryに変換。"""
    id_match = re.search(r"<id>tag:[^,]+,\d+-\d+-\d+:[^/]+/(\d+)</id>", entry_xml)
    if not id_match:
        id_match = re.search(r"/atom/entry/(\d+)", entry_xml)
    if not id_match:
        return None
    entry_id = id_match.group(1)

    title_match = re.search(r"<title[^>]*>([^<]*)</title>", entry_xml)
    title = title_match.group(1) if title_match else ""

    cats = re.findall(r'<category[^>]*term="([^"]+)"', entry_xml)
    published_match = re.search(r"<published>([^<]+)</published>", entry_xml)
    published = published_match.group(1) if published_match else ""
    updated_match = re.search(r"<updated>([^<]+)</updated>", entry_xml)
    updated = updated_match.group(1) if updated_match else ""

    draft = False
    draft_match = re.search(r"<app:draft>([^<]+)</app:draft>", entry_xml)
    if draft_match and draft_match.group(1).strip().lower() == "yes":
        draft = True

    content_match = re.search(
        r'<content[^>]*type="text/html"[^>]*>(.*?)</content>', entry_xml, re.DOTALL
    )
    body_html = ""
    if content_match:
        body_html = content_match.group(1)
        body_html = body_html.strip()
        if body_html.startswith("<![CDATA["):
            body_html = body_html[len("<![CDATA[") :]
            if body_html.endswith("]]>"):
                body_html = body_html[:-3]
        body_html = (
            body_html.replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&amp;", "&")
            .replace("&quot;", '"')
            .replace("&#39;", "'")
        )

    alt_match = re.search(
        r'<link[^>]+rel="alternate"[^>]+href="([^"]+)"', entry_xml
    )
    alternate_url = alt_match.group(1) if alt_match else ""

    return Entry(
        entry_id=entry_id,
        title=title,
        categories=tuple(cats),
        published=published,
        updated=updated,
        draft=draft,
        body_html=body_html,
        alternate_url=alternate_url,
    )


def fetch_all_entries(env: dict[str, str]) -> list[Entry]:
    """全エントリをAtomPubで取得。ローカルキャッシュを活用。"""
    api_key = env["HATENA_API_KEY"]
    auth = basic_auth(HATENA_ID, api_key)

    ENTRIES_DIR.mkdir(parents=True, exist_ok=True)

    collection_url = (
        f"https://blog.hatena.ne.jp/{HATENA_ID}/{BLOG_DOMAIN}/atom/entry"
    )

    all_entries: list[Entry] = []
    seen_ids: set[str] = set()
    next_url: Optional[str] = collection_url
    page_num = 0

    while next_url:
        page_num += 1
        print(f"[fetch] page {page_num}: {next_url[:80]}...", flush=True)
        try:
            data = fetch_page(next_url, auth)
        except urllib.error.HTTPError as e:
            print(f"[error] HTTP {e.code}: {e.read()[:200]}")
            break

        xml = data.decode("utf-8")
        entry_blocks = re.findall(r"<entry>.*?</entry>", xml, re.DOTALL)

        for block in entry_blocks:
            entry = parse_entry_xml(block)
            if entry is None:
                continue
            if entry.draft:
                continue
            if entry.entry_id in seen_ids:
                continue
            seen_ids.add(entry.entry_id)
            all_entries.append(entry)

            cache_path = ENTRIES_DIR / f"{entry.entry_id}.xml"
            if not cache_path.exists():
                cache_path.write_text(block, encoding="utf-8")

        next_match = re.search(
            r'<link[^>]+rel="next"[^>]+href="([^"]+)"', xml
        )
        if next_match:
            next_url = next_match.group(1).replace("&amp;", "&")
        else:
            next_url = None

        time.sleep(0.5)

        if page_num > 500:
            print("[warn] page limit reached, stopping")
            break

    return all_entries


def strip_tags(html: str) -> str:
    text = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.DOTALL | re.I)
    text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.DOTALL | re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def detect_infographic(body_html: str) -> tuple[int, str]:
    """0=なし, 1=あり, 2=要手動確認。理由コードも返す。"""
    if "<!-- インフォグラフィック -->" in body_html:
        return 1, "comment_marker"
    if re.search(r'<img[^>]+alt="[^"]*インフォグラフィック', body_html):
        return 1, "alt_marker"
    if re.search(
        r'<img[^>]+src="[^"]*infographic[^"]*"', body_html, flags=re.I
    ):
        return 1, "filename_marker"
    if re.search(
        r'<img[^>]+src="[^"]*[_-]?info(graph|graphic)[^"]*"',
        body_html,
        flags=re.I,
    ):
        return 1, "filename_variant"
    head = body_html[:800]
    big_img = re.search(
        r'<img[^>]+(?:width="(?P<w>\d+)")', head
    )
    if big_img:
        try:
            w = int(big_img.group("w"))
            if w >= 600:
                return 2, "big_image_in_head"
        except (TypeError, ValueError):
            pass
    fotolife_imgs = re.findall(
        r'<img[^>]+src="(https?://[^"]*f\.(?:st-)?hatena[^"]*)"', head
    )
    if fotolife_imgs:
        return 2, "fotolife_image_in_head"
    return 0, "no_marker"


def count_internal_hatena_links(html: str) -> int:
    return len(
        re.findall(
            r'<a[^>]+href="https?://hinyan1016\.hatenablog\.com/', html
        )
    )


def count_all_links(html: str) -> int:
    return len(re.findall(r"<a\s[^>]*href=", html))


def should_exclude(entry: Entry, body_text: str) -> tuple[bool, str]:
    for cat in entry.categories:
        if cat in EXCLUDED_CATEGORIES:
            return True, f"category:{cat}"
    word_count = len(body_text)
    if word_count < SHORT_CONTENT_THRESHOLD:
        return True, f"short:{word_count}chars"
    total_links = count_all_links(entry.body_html)
    internal_links = count_internal_hatena_links(entry.body_html)
    if total_links >= 5 and internal_links / total_links >= 0.8:
        return True, f"index_page:{internal_links}/{total_links}"
    return False, ""


def year_of(iso_datetime: str) -> str:
    m = re.match(r"(\d{4})", iso_datetime)
    return m.group(1) if m else "unknown"


def main() -> int:
    env = load_env()
    print("[step 1] AtomPub取得中...", flush=True)
    entries = fetch_all_entries(env)
    print(f"[step 1] {len(entries)}件取得完了", flush=True)

    print("[step 2] 判定と集計中...", flush=True)
    rows: list[dict[str, object]] = []
    counts = {
        "total": 0,
        "has_infographic": 0,
        "grey_zone": 0,
        "no_infographic": 0,
        "excluded": 0,
    }
    by_year_need: dict[str, int] = {}
    by_detection: dict[str, int] = {}
    by_exclusion: dict[str, int] = {}

    for entry in entries:
        counts["total"] += 1
        body_text = strip_tags(entry.body_html)
        word_count = len(body_text)

        excluded, excl_reason = should_exclude(entry, body_text)
        if excluded:
            counts["excluded"] += 1
            by_exclusion[excl_reason.split(":")[0]] = (
                by_exclusion.get(excl_reason.split(":")[0], 0) + 1
            )
            status = "excluded"
            ig_flag = -1
            ig_reason = excl_reason
        else:
            ig_flag, ig_reason = detect_infographic(entry.body_html)
            if ig_flag == 1:
                counts["has_infographic"] += 1
                status = "has_infographic"
            elif ig_flag == 2:
                counts["grey_zone"] += 1
                status = "grey_zone"
            else:
                counts["no_infographic"] += 1
                status = "no_infographic"
                y = year_of(entry.published)
                by_year_need[y] = by_year_need.get(y, 0) + 1
            by_detection[ig_reason] = by_detection.get(ig_reason, 0) + 1

        rows.append(
            {
                "entry_id": entry.entry_id,
                "title": entry.title.replace(",", "、"),
                "category": "|".join(entry.categories),
                "published": entry.published,
                "word_count": word_count,
                "status": status,
                "detection_reason": ig_reason,
                "url": entry.alternate_url,
            }
        )

    print("[step 3] CSV出力中...", flush=True)
    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        f.write(
            "entry_id,title,category,published,word_count,status,detection_reason,url\n"
        )
        for r in rows:
            f.write(
                "{entry_id},{title},{category},{published},{word_count},{status},{detection_reason},{url}\n".format(
                    **r
                )
            )

    lines = [
        "=" * 60,
        "インフォグラフィック未対応記事 棚卸しサマリ",
        f"生成日時: {datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')}",
        "=" * 60,
        f"総記事数（下書き除く）: {counts['total']}",
        f"  インフォグラフィックあり: {counts['has_infographic']}",
        f"  要手動確認（グレーゾーン）: {counts['grey_zone']}",
        f"  インフォグラフィックなし: {counts['no_infographic']}",
        f"  除外（日記/短文/インデックス）: {counts['excluded']}",
        "",
        "--- 判定理由別内訳 ---",
    ]
    for k, v in sorted(by_detection.items(), key=lambda x: -x[1]):
        lines.append(f"  {k}: {v}")
    lines.append("")
    lines.append("--- 除外理由別内訳 ---")
    for k, v in sorted(by_exclusion.items(), key=lambda x: -x[1]):
        lines.append(f"  {k}: {v}")
    lines.append("")
    lines.append("--- 年別・インフォグラフィック必要件数 ---")
    for y, v in sorted(by_year_need.items()):
        lines.append(f"  {y}: {v}件")
    lines.append("")
    lines.append(f"CSV出力: {OUTPUT_CSV}")

    summary = "\n".join(lines)
    SUMMARY_FILE.parent.mkdir(parents=True, exist_ok=True)
    SUMMARY_FILE.write_text(summary, encoding="utf-8")
    print(summary)
    return 0


if __name__ == "__main__":
    sys.exit(main())
