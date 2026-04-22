#!/usr/bin/env python3
"""
インベントリ再判定 v2 — 既に取得済みのキャッシュXMLを使い、
より緩和した検出ロジックで判定し直す。

v2の変更点:
1. <iframe> を除去してから判定（YouTube埋め込みで判定がズレる問題）
2. Hatenaの <figure class="figure-image-fotolife"> パターンを検出
3. Fotolife由来の <img> もインフォグラフィック扱い
4. 判定結果を4段階に: has_ig_strong / has_ig_likely / needs_ig / excluded
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

ROOT = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task")
PROJECT_DIR = ROOT / "blog" / "infographic-project"
ENTRIES_DIR = PROJECT_DIR / "data" / "entries"
OUTPUT_CSV = PROJECT_DIR / "data" / "inventory_v2.csv"
SUMMARY_FILE = PROJECT_DIR / "data" / "inventory_summary_v2.txt"
NEEDS_CSV = PROJECT_DIR / "data" / "targets_v2.csv"

EXCLUDED_CATEGORIES = {"日記", "雑記", "お知らせ", "ブログ運営", "プライベート"}
SHORT_CONTENT_THRESHOLD = 500


@dataclass(frozen=True)
class Entry:
    entry_id: str
    title: str
    categories: tuple[str, ...]
    published: str
    body_html: str
    alternate_url: str


def parse_entry_xml(entry_xml: str) -> Optional[Entry]:
    id_match = re.search(r"/atom/entry/(\d+)", entry_xml)
    if not id_match:
        return None
    entry_id = id_match.group(1)

    title_match = re.search(r"<title[^>]*>([^<]*)</title>", entry_xml)
    title = title_match.group(1) if title_match else ""

    cats = re.findall(r'<category[^>]*term="([^"]+)"', entry_xml)
    published_match = re.search(r"<published>([^<]+)</published>", entry_xml)
    published = published_match.group(1) if published_match else ""

    content_match = re.search(
        r'<content[^>]*type="text/html"[^>]*>(.*?)</content>',
        entry_xml,
        re.DOTALL,
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
        body_html=body_html,
        alternate_url=alternate_url,
    )


def strip_tags(html: str) -> str:
    text = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.DOTALL | re.I)
    text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.DOTALL | re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def detect_v2(body_html: str) -> tuple[str, str]:
    """
    Returns (status, reason):
      status in {'has_ig_strong', 'has_ig_likely', 'needs_ig'}
    """
    # iframeと script を除去して判定を安定化
    cleaned = re.sub(
        r"<iframe[^>]*>.*?</iframe>", "", body_html, flags=re.DOTALL | re.I
    )
    cleaned = re.sub(
        r"<script[^>]*>.*?</script>", "", cleaned, flags=re.DOTALL | re.I
    )

    # Strong signal 1: 明示マーカー
    if "<!-- インフォグラフィック -->" in cleaned:
        return "has_ig_strong", "comment_marker"
    if re.search(r'<img[^>]+alt="[^"]*インフォグラフィック', cleaned):
        return "has_ig_strong", "alt_marker_infographic"
    if re.search(r'<img[^>]+alt="[^"]*(まとめ|要約|インフォ)', cleaned):
        return "has_ig_strong", "alt_marker_summary"

    # Strong signal 2: インフォ/infographic をURLに含む
    if re.search(r'<img[^>]+src="[^"]*infograph', cleaned, flags=re.I):
        return "has_ig_strong", "filename_infographic"

    # 本文冒頭エリア（最初の4000文字、iframe除去後）
    head = cleaned[:4000]

    # Likely signal: Hatenaの<figure class="figure-image-fotolife">
    if re.search(r'<figure[^>]*figure-image-fotolife', head, flags=re.I):
        return "has_ig_likely", "fotolife_figure_head"

    # Likely signal: Fotolife由来の<img>が冒頭にある
    if re.search(
        r'<img[^>]+src="https?://cdn-ak\.f\.st-hatena\.com/images/fotolife',
        head,
    ):
        return "has_ig_likely", "fotolife_img_head"

    # Likely signal: 幅600px以上の画像
    for m in re.finditer(r'<img[^>]+width="(\d+)"', head):
        try:
            w = int(m.group(1))
            if w >= 600:
                return "has_ig_likely", f"big_img_head_w{w}"
        except ValueError:
            continue

    # Likely signal: <figure>内の<img>
    if re.search(r"<figure[^>]*>.*?<img[^>]+>.*?</figure>", head, re.DOTALL):
        return "has_ig_likely", "figure_img_head"

    # 冒頭に画像が一切ない → インフォグラフィック要
    return "needs_ig", "no_image_in_head"


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
    wc = len(body_text)
    if wc < SHORT_CONTENT_THRESHOLD:
        return True, f"short:{wc}chars"
    total = count_all_links(entry.body_html)
    internal = count_internal_hatena_links(entry.body_html)
    if total >= 5 and internal / total >= 0.8:
        return True, f"index_page:{internal}/{total}"
    return False, ""


def year_of(iso: str) -> str:
    m = re.match(r"(\d{4})", iso)
    return m.group(1) if m else "unknown"


def main() -> int:
    xml_files = sorted(ENTRIES_DIR.glob("*.xml"))
    if not xml_files:
        print(f"[error] No cached XML in {ENTRIES_DIR}")
        return 1

    print(f"[step 1] {len(xml_files)}件のキャッシュを再判定中...", flush=True)

    rows: list[dict[str, object]] = []
    counts = {
        "total": 0,
        "has_ig_strong": 0,
        "has_ig_likely": 0,
        "needs_ig": 0,
        "excluded": 0,
    }
    by_year_need: dict[str, int] = {}
    by_reason: dict[str, int] = {}
    by_exclusion: dict[str, int] = {}

    for xml_path in xml_files:
        xml = xml_path.read_text(encoding="utf-8")
        entry = parse_entry_xml(xml)
        if entry is None:
            continue

        counts["total"] += 1
        body_text = strip_tags(entry.body_html)

        excluded, excl_reason = should_exclude(entry, body_text)
        if excluded:
            counts["excluded"] += 1
            ekey = excl_reason.split(":")[0]
            by_exclusion[ekey] = by_exclusion.get(ekey, 0) + 1
            status = "excluded"
            reason = excl_reason
        else:
            status, reason = detect_v2(entry.body_html)
            counts[status] = counts.get(status, 0) + 1
            by_reason[reason] = by_reason.get(reason, 0) + 1
            if status == "needs_ig":
                y = year_of(entry.published)
                by_year_need[y] = by_year_need.get(y, 0) + 1

        rows.append(
            {
                "entry_id": entry.entry_id,
                "title": entry.title.replace(",", "、"),
                "category": "|".join(entry.categories),
                "published": entry.published,
                "word_count": len(body_text),
                "status": status,
                "reason": reason,
                "url": entry.alternate_url,
            }
        )

    # Full inventory CSV
    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        f.write(
            "entry_id,title,category,published,word_count,status,reason,url\n"
        )
        for r in rows:
            f.write(
                "{entry_id},{title},{category},{published},{word_count},{status},{reason},{url}\n".format(
                    **r
                )
            )

    # Needs-only CSV (sorted by published ASC = 古い順)
    needs_rows = [r for r in rows if r["status"] == "needs_ig"]
    needs_rows.sort(key=lambda r: str(r["published"]))
    with NEEDS_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        f.write(
            "entry_id,title,category,published,word_count,reason,url\n"
        )
        for r in needs_rows:
            f.write(
                "{entry_id},{title},{category},{published},{word_count},{reason},{url}\n".format(
                    **r
                )
            )

    lines = [
        "=" * 60,
        "インフォグラフィック未対応記事 棚卸しサマリ v2",
        "=" * 60,
        f"総記事数（下書き除く）: {counts['total']}",
        f"  インフォグラフィックあり(強): {counts['has_ig_strong']}",
        f"  インフォグラフィックあり(推定): {counts['has_ig_likely']}",
        f"  インフォグラフィック要: {counts['needs_ig']}",
        f"  除外（日記/短文/インデックス）: {counts['excluded']}",
        "",
        "--- 判定理由別内訳 ---",
    ]
    for k, v in sorted(by_reason.items(), key=lambda x: -x[1]):
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
    lines.append(f"全件CSV: {OUTPUT_CSV}")
    lines.append(f"対象CSV: {NEEDS_CSV}")

    summary = "\n".join(lines)
    SUMMARY_FILE.write_text(summary, encoding="utf-8")
    try:
        print(summary)
    except UnicodeEncodeError:
        print(summary.encode("utf-8", errors="replace").decode("utf-8"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
