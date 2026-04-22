"""インフォグラフィック有無の判定ロジック（v2）。

inventory_rescan_v2.py を整理してモジュール化したもの。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional


EXCLUDED_CATEGORIES = {"日記", "雑記", "お知らせ", "ブログ運営", "プライベート"}
SHORT_CONTENT_THRESHOLD = 500
HEAD_WINDOW = 4000


@dataclass(frozen=True)
class Entry:
    entry_id: str
    title: str
    categories: tuple[str, ...]
    published: str
    body_html: str
    alternate_url: str


def parse_entry_xml(entry_xml: str) -> Optional[Entry]:
    """AtomPubのentry XMLをEntryに変換。"""
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
        raw = content_match.group(1).strip()
        if raw.startswith("<![CDATA["):
            raw = raw[len("<![CDATA[") :]
            if raw.endswith("]]>"):
                raw = raw[:-3]
        body_html = (
            raw.replace("&lt;", "<")
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


def detect_v2(body_html: str) -> tuple[str, str]:
    """インフォグラフィックの有無を判定。

    Returns:
        (status, reason) where status in {'has_ig_strong', 'has_ig_likely', 'needs_ig'}
    """
    cleaned = re.sub(
        r"<iframe[^>]*>.*?</iframe>", "", body_html, flags=re.DOTALL | re.I
    )
    cleaned = re.sub(
        r"<script[^>]*>.*?</script>", "", cleaned, flags=re.DOTALL | re.I
    )

    # Strong signals
    if "<!-- インフォグラフィック -->" in cleaned:
        return "has_ig_strong", "comment_marker"
    if re.search(r'<img[^>]+alt="[^"]*インフォグラフィック', cleaned):
        return "has_ig_strong", "alt_marker_infographic"
    if re.search(r'<img[^>]+alt="[^"]*(まとめ|要約)', cleaned):
        return "has_ig_strong", "alt_marker_summary"
    if re.search(r'<img[^>]+src="[^"]*infograph', cleaned, flags=re.I):
        return "has_ig_strong", "filename_infographic"

    head = cleaned[:HEAD_WINDOW]

    if re.search(r"<figure[^>]*figure-image-fotolife", head, flags=re.I):
        return "has_ig_likely", "fotolife_figure_head"
    if re.search(
        r'<img[^>]+src="https?://cdn-ak\.f\.st-hatena\.com/images/fotolife',
        head,
    ):
        return "has_ig_likely", "fotolife_img_head"
    for m in re.finditer(r'<img[^>]+width="(\d+)"', head):
        try:
            if int(m.group(1)) >= 600:
                return "has_ig_likely", f"big_img_head_w{m.group(1)}"
        except ValueError:
            continue
    if re.search(r"<figure[^>]*>.*?<img[^>]+>.*?</figure>", head, re.DOTALL):
        return "has_ig_likely", "figure_img_head"

    return "needs_ig", "no_image_in_head"
