"""インフォグラフィック<img>タグを記事HTMLの適切な位置に挿入する。"""

from __future__ import annotations

import re


INFOGRAPHIC_TEMPLATE = (
    "\n<!-- インフォグラフィック -->\n"
    '<div style="text-align:center; margin:20px 0;">\n'
    "{img_tag}\n"
    "</div>\n\n"
)


def insert_infographic(body_html: str, img_tag: str) -> str:
    """3段階フォールバックで img_tag を body_html に挿入。

    優先順位:
      Level 1: <!-- 目次 --> コメントの直前
      Level 2: 最初の <h2> の直前
      Level 3: 最初の </p> の直後
      Level 4 (最終): 本文先頭
    """
    block = INFOGRAPHIC_TEMPLATE.format(img_tag=img_tag)

    if "<!-- 目次 -->" in body_html:
        return body_html.replace("<!-- 目次 -->", block + "<!-- 目次 -->", 1)

    h2_match = re.search(r"<h2[^>]*>", body_html)
    if h2_match:
        pos = h2_match.start()
        return body_html[:pos] + block + body_html[pos:]

    p_match = re.search(r"</p>", body_html)
    if p_match:
        pos = p_match.end()
        return body_html[:pos] + block + body_html[pos:]

    return block + body_html
