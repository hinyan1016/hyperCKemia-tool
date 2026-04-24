#!/usr/bin/env python3
"""その症状大丈夫シリーズ #02-#13 の目次リンク不具合修正スクリプト.

post_drafts.py の convert_to_hatena() が h2 の id を日本語化する一方
TOC の href (英語スラッグ) を更新し忘れていたため TOC と見出しが不一致になり
目次クリックで飛べなくなっていた問題を修正する.

位置ベースで TOC の N 番目アンカーと本文の N 番目 h2 id を対応付け,
TOC 側の href を正しい id に書き換える.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

BASE = Path(__file__).resolve().parent

TARGETS: list[str] = [
    "02_slurred_speech_hatena.html",
    "03_weakness_hatena.html",
    "04_facial_twitch_hatena.html",
    "05_leg_cramps_hatena.html",
    "06_autonomic_nerves_hatena.html",
    "07_diplopia_hatena.html",
    "08_dysphagia_hatena.html",
    "09_neck_pain_hatena.html",
    "10_numbness_hatena.html",
    "11_gait_hatena.html",
    "12_head_injury_hatena.html",
    "13_weather_headache_hatena.html",
]

TOC_BLOCK_RE = re.compile(
    r'(<ul class="table-of-contents">)(.*?)(</ul>)', re.DOTALL
)


def fix_toc(html: str) -> tuple[str, dict[str, str]]:
    """HTML内のTOC hrefをh2 idに追従させる.

    Returns:
        (修正後HTML, 変更マッピング {旧アンカー: 新アンカー})
    Raises:
        ValueError: TOC未検出 / 件数不一致の場合
    """
    toc_match = TOC_BLOCK_RE.search(html)
    if not toc_match:
        raise ValueError('<ul class="table-of-contents"> が見つかりません')
    toc_open, toc_inner, toc_close = toc_match.groups()

    toc_ids = re.findall(r'href="#([^"]+)"', toc_inner)
    h2_ids = re.findall(r'<h2\s+id="([^"]+)"', html)

    if len(toc_ids) != len(h2_ids):
        raise ValueError(
            "件数不一致: TOC={} / h2={}".format(len(toc_ids), len(h2_ids))
        )

    mapping: dict[str, str] = {
        old: new for old, new in zip(toc_ids, h2_ids) if old != new
    }
    if not mapping:
        return html, mapping

    new_inner = toc_inner
    for old, new in mapping.items():
        new_inner = new_inner.replace(
            'href="#{}"'.format(old),
            'href="#{}"'.format(new),
        )
    new_html = html.replace(
        toc_open + toc_inner + toc_close,
        toc_open + new_inner + toc_close,
        1,
    )
    return new_html, mapping


def main() -> int:
    fixed = 0
    skipped = 0
    errors = 0
    for name in TARGETS:
        path = BASE / name
        if not path.exists():
            print("[MISS] {} 存在しません".format(name))
            errors += 1
            continue

        html = path.read_text(encoding="utf-8")
        try:
            new_html, mapping = fix_toc(html)
        except ValueError as exc:
            print("[ERR ] {}: {}".format(name, exc))
            errors += 1
            continue

        if not mapping:
            print("[SKIP] {} (既に一致済み)".format(name))
            skipped += 1
            continue

        path.write_text(new_html, encoding="utf-8")
        print("[FIX ] {} ({} 件)".format(name, len(mapping)))
        for old, new in mapping.items():
            print("       #{} -> #{}".format(old, new))
        fixed += 1

    print("\n集計: 修正={} / スキップ={} / エラー={}".format(fixed, skipped, errors))
    return 0 if errors == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
