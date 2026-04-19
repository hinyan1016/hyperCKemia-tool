#!/usr/bin/env python3
"""Phase A: YouTube iframe形式の診断スキャン.

GSC動画未登録492件（「動画再生ページに動画がありません」）の原因を特定する。
全記事のYouTube埋め込み形式を4パターンに分類し、修正対象を抽出する。

パターン:
    P1 bare_only        : テキストリンク/bare URLのみ、iframe無し
    P2 iframe_plain     : iframeあり、?feature=oembed なし、<cite>無し
    P3 iframe_oembed    : iframeに ?feature=oembed あるが <cite hatena-citation>無し
    P4 iframe_hatena    : oembed+<cite hatena-citation>完全形（正常）
    P5 jsonld_only      : JSON-LDだけあってiframe/URL両方無し

出力:
    video_iframe_audit.csv — 記事ごとの分類結果
    video_iframe_audit_summary.txt — パターン別集計サマリー

使い方:
    python video_iframe_audit.py              # 全記事スキャン
    python video_iframe_audit.py --limit N    # 先頭N件のみ（テスト用）
"""
from __future__ import annotations

import argparse
import base64
import csv
import html as html_mod
import logging
import re
import sys
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger(__name__)

ATOM_NS = "http://www.w3.org/2005/Atom"
APP_NS = "http://www.w3.org/2007/app"
ENV_FILE = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task"
    r"\youtube-slides\食事指導シリーズ\_shared\.env"
)
OUTPUT_CSV = Path(__file__).parent / "video_iframe_audit.csv"
OUTPUT_SUMMARY = Path(__file__).parent / "video_iframe_audit_summary.txt"

# パターン検出用正規表現
RE_IFRAME_YT = re.compile(
    r'<iframe[^>]*src="[^"]*(?:youtube\.com/embed|youtu\.be)[^"]*"[^>]*>.*?</iframe>',
    re.IGNORECASE | re.DOTALL,
)
RE_OEMBED_PARAM = re.compile(r"feature=oembed", re.IGNORECASE)
RE_HATENA_CITATION = re.compile(
    r'<cite[^>]*class="[^"]*hatena-citation[^"]*"',
    re.IGNORECASE,
)
RE_BARE_YT_URL = re.compile(
    r"https?://(?:www\.)?(?:youtube\.com/watch\?v=|youtu\.be/)[A-Za-z0-9_-]{11}",
    re.IGNORECASE,
)
RE_JSONLD_VIDEO = re.compile(r'"VideoObject"', re.IGNORECASE)
RE_YT_ID = re.compile(r"(?:v=|youtu\.be/|embed/)([A-Za-z0-9_-]{11})")


@dataclass(frozen=True)
class Classification:
    """1記事分の分類結果."""

    url: str
    title: str
    published: str
    is_draft: bool
    has_bare_url: bool
    iframe_count: int
    oembed_count: int
    hatena_citation_count: int
    has_jsonld: bool
    pattern: str  # P1-P5 or N_none
    youtube_ids: tuple[str, ...]
    edit_url: str


@dataclass
class Counters:
    total: int = 0
    published: int = 0
    drafts: int = 0
    has_any_yt: int = 0
    patterns: dict[str, int] = field(default_factory=dict)


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def make_auth_header(env: dict[str, str]) -> str:
    hatena_id = env["HATENA_ID"]
    api_key = env["HATENA_API_KEY"]
    token = base64.b64encode(f"{hatena_id}:{api_key}".encode()).decode()
    return f"Basic {token}"


def classify_content(raw_content: str) -> tuple[str, dict[str, int | bool]]:
    """HTMLを分類し、パターン名と指標を返す.

    AtomPub GETはHTMLを XMLエスケープして返すため、
    html.unescape() してから判定する。
    """
    content = html_mod.unescape(raw_content or "")

    iframes = RE_IFRAME_YT.findall(content)
    iframe_count = len(iframes)
    oembed_count = sum(1 for f in iframes if RE_OEMBED_PARAM.search(f))
    hatena_citation_count = len(RE_HATENA_CITATION.findall(content))
    has_bare_url = bool(RE_BARE_YT_URL.search(content))
    has_jsonld = bool(RE_JSONLD_VIDEO.search(content))

    metrics = {
        "has_bare_url": has_bare_url,
        "iframe_count": iframe_count,
        "oembed_count": oembed_count,
        "hatena_citation_count": hatena_citation_count,
        "has_jsonld": has_jsonld,
    }

    # 判定ツリー（優先度順）
    if iframe_count == 0 and not has_bare_url and not has_jsonld:
        return "N_none", metrics
    if iframe_count == 0 and not has_bare_url and has_jsonld:
        return "P5_jsonld_only", metrics
    if iframe_count == 0 and has_bare_url:
        return "P1_bare_only", metrics
    # iframe あり
    if iframe_count >= 1 and oembed_count >= 1 and hatena_citation_count >= 1:
        return "P4_iframe_hatena", metrics
    if iframe_count >= 1 and oembed_count >= 1 and hatena_citation_count == 0:
        return "P3_iframe_oembed", metrics
    return "P2_iframe_plain", metrics


def fetch_entries(
    env: dict[str, str],
    limit: int | None = None,
) -> list[Classification]:
    auth = make_auth_header(env)
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    url: str | None = (
        f"https://blog.hatena.ne.jp/{hatena_id}/{blog_domain}/atom/entry"
    )

    results: list[Classification] = []
    page = 0

    while url:
        page += 1
        req = urllib.request.Request(url, headers={"Authorization": auth})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                body = resp.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            logger.error("page %d: HTTP %d", page, exc.code)
            break

        root = ET.fromstring(body)
        for entry in root.findall(f"{{{ATOM_NS}}}entry"):
            title_el = entry.find(f"{{{ATOM_NS}}}title")
            title = (title_el.text or "") if title_el is not None else ""

            link_alt = ""
            link_edit = ""
            for link in entry.findall(f"{{{ATOM_NS}}}link"):
                rel = link.get("rel", "")
                if rel == "alternate":
                    link_alt = link.get("href", "")
                elif rel == "edit":
                    link_edit = link.get("href", "")

            pub_el = entry.find(f"{{{ATOM_NS}}}published")
            published = ""
            if pub_el is not None and pub_el.text:
                published = pub_el.text[:10]

            control = entry.find(f"{{{APP_NS}}}control")
            is_draft = False
            if control is not None:
                draft_el = control.find(f"{{{APP_NS}}}draft")
                if draft_el is not None and draft_el.text == "yes":
                    is_draft = True

            content_el = entry.find(f"{{{ATOM_NS}}}content")
            raw = ""
            if content_el is not None and content_el.text:
                raw = content_el.text

            pattern, metrics = classify_content(raw)
            yt_ids = tuple(sorted(set(RE_YT_ID.findall(html_mod.unescape(raw)))))

            results.append(
                Classification(
                    url=link_alt,
                    title=title,
                    published=published,
                    is_draft=is_draft,
                    has_bare_url=bool(metrics["has_bare_url"]),
                    iframe_count=int(metrics["iframe_count"]),
                    oembed_count=int(metrics["oembed_count"]),
                    hatena_citation_count=int(metrics["hatena_citation_count"]),
                    has_jsonld=bool(metrics["has_jsonld"]),
                    pattern=pattern,
                    youtube_ids=yt_ids,
                    edit_url=link_edit,
                )
            )

            if limit is not None and len(results) >= limit:
                logger.info("limit %d reached", limit)
                return results

        next_url: str | None = None
        for link in root.findall(f"{{{ATOM_NS}}}link"):
            if link.get("rel") == "next":
                next_url = link.get("href")
        url = next_url
        logger.info("page %d done (cumulative %d)", page, len(results))
        time.sleep(0.3)

    return results


def write_csv(results: list[Classification]) -> None:
    fields = [
        "url",
        "title",
        "published",
        "is_draft",
        "pattern",
        "has_bare_url",
        "iframe_count",
        "oembed_count",
        "hatena_citation_count",
        "has_jsonld",
        "youtube_ids",
        "edit_url",
    ]
    with open(OUTPUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for r in results:
            writer.writerow(
                {
                    "url": r.url,
                    "title": r.title,
                    "published": r.published,
                    "is_draft": r.is_draft,
                    "pattern": r.pattern,
                    "has_bare_url": r.has_bare_url,
                    "iframe_count": r.iframe_count,
                    "oembed_count": r.oembed_count,
                    "hatena_citation_count": r.hatena_citation_count,
                    "has_jsonld": r.has_jsonld,
                    "youtube_ids": "|".join(r.youtube_ids),
                    "edit_url": r.edit_url,
                }
            )
    logger.info("CSV written: %s", OUTPUT_CSV)


def summarize(results: list[Classification]) -> str:
    counters = Counters()
    counters.total = len(results)
    for r in results:
        if r.is_draft:
            counters.drafts += 1
        else:
            counters.published += 1
        if r.pattern != "N_none":
            counters.has_any_yt += 1
        counters.patterns[r.pattern] = counters.patterns.get(r.pattern, 0) + 1

    pattern_labels = {
        "P1_bare_only": "P1 bare URLのみ / iframe無し（最優先修正）",
        "P2_iframe_plain": "P2 古いiframe（?feature=oembed無し）",
        "P3_iframe_oembed": "P3 oembed iframeだが <cite>欠落",
        "P4_iframe_hatena": "P4 完全な oembed + <cite>（正常／Google待ち）",
        "P5_jsonld_only": "P5 JSON-LDのみ（動画本体無し）",
        "N_none": "N YouTube関連なし",
    }

    lines = [
        "=== YouTube iframe 診断結果 ===",
        f"総記事数: {counters.total}（公開 {counters.published} / 下書き {counters.drafts}）",
        f"YouTube関連を含む記事: {counters.has_any_yt}",
        "",
        "--- パターン別件数 ---",
    ]
    # 公開のみで集計
    pub_patterns: dict[str, int] = {}
    for r in results:
        if r.is_draft:
            continue
        pub_patterns[r.pattern] = pub_patterns.get(r.pattern, 0) + 1

    for key in ["P1_bare_only", "P2_iframe_plain", "P3_iframe_oembed", "P4_iframe_hatena", "P5_jsonld_only", "N_none"]:
        label = pattern_labels[key]
        count_all = counters.patterns.get(key, 0)
        count_pub = pub_patterns.get(key, 0)
        lines.append(f"  {label}")
        lines.append(f"    全体 {count_all:4d} / 公開 {count_pub:4d}")
    lines.append("")
    # 修正対象
    fixable = sum(pub_patterns.get(k, 0) for k in ["P1_bare_only", "P2_iframe_plain", "P3_iframe_oembed"])
    lines.append(f"修正対象候補（P1+P2+P3 公開のみ）: {fixable} 件")
    lines.append(
        f"うちP4（完了形でGoogle待ち）: {pub_patterns.get('P4_iframe_hatena', 0)} 件"
    )
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="YouTube iframe diagnostic scan")
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Limit number of entries (for testing)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    env = load_env()
    for required in ("HATENA_ID", "HATENA_API_KEY", "HATENA_BLOG_DOMAIN"):
        if required not in env:
            logger.error("missing env var: %s", required)
            return 1

    logger.info("Fetching entries from AtomPub API...")
    results = fetch_entries(env, limit=args.limit)
    write_csv(results)
    summary = summarize(results)
    print()
    print(summary)
    OUTPUT_SUMMARY.write_text(summary, encoding="utf-8")
    logger.info("Summary written: %s", OUTPUT_SUMMARY)
    return 0


if __name__ == "__main__":
    sys.exit(main())
