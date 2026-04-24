#!/usr/bin/env python3
"""その症状大丈夫シリーズのはてな公開/下書き記事のTOCリンク不具合を一括修正.

現状のバグ: TOC の href (英語スラッグ) と h2 の id (日本語) が不一致で目次が飛ばない.
対応: カテゴリ "その症状大丈夫" を持つ全エントリに対し AtomPub GET -> TOC 修正 -> PUT.

- 位置ベースで TOC N 番目アンカーと h2 N 番目 id を対応付ける
- draft 状態は現行値を保持
- インフォグラフィック/YouTube埋め込み等の他要素には一切触れない
- dry-run モード (--dry-run) で変更内容を確認可能
"""

from __future__ import annotations

import argparse
import base64
import re
import sys
import time
import urllib.error
import urllib.request
from typing import List, Tuple

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"
TARGET_CATEGORY = "その症状大丈夫"
EXCLUDE_TITLE_KEYWORDS = ("手のふるえ", "手のふるえ",)  # #01 はTOC不具合なしなのでスキップ

TOC_BLOCK_RE = re.compile(
    r'(<ul class="table-of-contents">)(.*?)(</ul>)', re.DOTALL
)


def auth_header() -> str:
    b64 = base64.b64encode(f"{HATENA_ID}:{API_KEY}".encode()).decode()
    return f"Basic {b64}"


def http_get(url: str) -> str:
    req = urllib.request.Request(url, headers={"Authorization": auth_header()})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def http_put(url: str, xml: str) -> str:
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
        return resp.read().decode("utf-8")


def list_entries() -> List[dict]:
    """AtomPubコレクションから全エントリを取得."""
    entries: List[dict] = []
    url = f"https://blog.hatena.ne.jp/{HATENA_ID}/{BLOG_DOMAIN}/atom/entry"
    page = 0
    while url and page < 50:
        body = http_get(url)
        for e in re.findall(r"<entry>(.*?)</entry>", body, re.DOTALL):
            title_m = re.search(r"<title[^>]*>([^<]*)</title>", e)
            id_m = re.search(r"<id>([^<]+)</id>", e)
            if not (title_m and id_m):
                continue
            title = title_m.group(1)
            entry_id_full = id_m.group(1)
            numeric_id = entry_id_full.rsplit("-", 1)[-1]
            categories = re.findall(r'<category[^>]*term="([^"]+)"', e)
            draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", e)
            is_draft = (draft_m.group(1) == "yes") if draft_m else False
            entries.append({
                "id": numeric_id,
                "title": title,
                "categories": categories,
                "is_draft": is_draft,
            })
        next_m = re.search(r'<link rel="next" href="([^"]+)"', body)
        url = next_m.group(1) if next_m else None
        page += 1
    return entries


def get_entry_content(entry_id: str) -> Tuple[str, str, List[str], bool]:
    """エントリ詳細からタイトル・content・カテゴリ・draft状態を取得."""
    url = f"https://blog.hatena.ne.jp/{HATENA_ID}/{BLOG_DOMAIN}/atom/entry/{entry_id}"
    body = http_get(url)
    title_m = re.search(r"<title[^>]*>([^<]*)</title>", body)
    title = title_m.group(1) if title_m else ""
    content_m = re.search(
        r'<content[^>]*type="text/html"[^>]*>(.*?)</content>', body, re.DOTALL
    )
    raw_content = content_m.group(1) if content_m else ""
    if raw_content.startswith("<![CDATA["):
        raw_content = raw_content[len("<![CDATA["):]
        if raw_content.endswith("]]>"):
            raw_content = raw_content[:-3]
    else:
        # XMLエスケープの解除
        raw_content = (
            raw_content.replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&quot;", '"')
            .replace("&#39;", "'")
            .replace("&amp;", "&")
        )
    categories = re.findall(r'<category[^>]*term="([^"]+)"', body)
    draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", body)
    is_draft = (draft_m.group(1) == "yes") if draft_m else False
    return title, raw_content, categories, is_draft


def fix_toc_in_content(content: str) -> Tuple[str, dict]:
    """TOCのhrefをh2のidに一致させる."""
    toc_match = TOC_BLOCK_RE.search(content)
    if not toc_match:
        return content, {}
    toc_open, toc_inner, toc_close = toc_match.groups()
    toc_ids = re.findall(r'href="#([^"]+)"', toc_inner)
    h2_ids = re.findall(r'<h2\s+id="([^"]+)"', content)
    if len(toc_ids) != len(h2_ids) or not toc_ids:
        return content, {}
    mapping = {old: new for old, new in zip(toc_ids, h2_ids) if old != new}
    if not mapping:
        return content, {}
    new_inner = toc_inner
    for old, new in mapping.items():
        new_inner = new_inner.replace(
            f'href="#{old}"', f'href="#{new}"'
        )
    new_content = content.replace(
        toc_open + toc_inner + toc_close,
        toc_open + new_inner + toc_close,
        1,
    )
    return new_content, mapping


def put_entry(entry_id: str, title: str, content: str, categories: List[str], is_draft: bool) -> str:
    cat_xml = "\n".join(f'  <category term="{c}" />' for c in categories)
    xml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<entry xmlns="http://www.w3.org/2005/Atom" '
        'xmlns:app="http://www.w3.org/2007/app">\n'
        f"  <title>{title}</title>\n"
        f"  <author><name>{HATENA_ID}</name></author>\n"
        f'  <content type="text/html"><![CDATA[{content}]]></content>\n'
        f"{cat_xml}\n"
        "  <app:control>\n"
        f"    <app:draft>{'yes' if is_draft else 'no'}</app:draft>\n"
        "  </app:control>\n"
        "</entry>"
    )
    url = f"https://blog.hatena.ne.jp/{HATENA_ID}/{BLOG_DOMAIN}/atom/entry/{entry_id}"
    body = http_put(url, xml)
    url_m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
    return url_m.group(1) if url_m else "(URL不明)"


def should_skip_title(title: str) -> bool:
    return any(kw in title for kw in EXCLUDE_TITLE_KEYWORDS)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="PUTせず差分のみ表示")
    parser.add_argument("--limit", type=int, default=None, help="先頭N件だけ処理")
    args = parser.parse_args()

    print("[1] エントリ一覧取得中...")
    all_entries = list_entries()
    print(f"    取得: {len(all_entries)}件")

    targets = [
        e for e in all_entries
        if TARGET_CATEGORY in e["categories"] and not should_skip_title(e["title"])
    ]
    print(f"[2] 対象: カテゴリ「{TARGET_CATEGORY}」= {len(targets)}件")
    for t in targets:
        state = "draft" if t["is_draft"] else "public"
        print(f"    - [{state}] {t['title']} (id={t['id']})")

    if args.limit:
        targets = targets[: args.limit]
        print(f"    limit適用: {len(targets)}件に絞り込み")

    if not targets:
        print("対象なし。終了。")
        return 0

    fixed = 0
    skipped = 0
    errors = 0
    for t in targets:
        eid = t["id"]
        print(f"\n[3] 処理中: {t['title']} (id={eid})")
        try:
            title, content, categories, is_draft = get_entry_content(eid)
        except urllib.error.HTTPError as e:
            print(f"    [ERR] GET失敗: HTTP {e.code}")
            errors += 1
            continue

        new_content, mapping = fix_toc_in_content(content)
        if not mapping:
            print("    [SKIP] 修正不要（TOC/h2が既に一致 または TOC未検出）")
            skipped += 1
            continue

        print(f"    [FIX] {len(mapping)}件のアンカーを置換予定")
        for old, new in mapping.items():
            print(f"         #{old} -> #{new}")

        if args.dry_run:
            print("    [DRY] PUTはスキップ")
            fixed += 1
            continue

        try:
            result_url = put_entry(eid, title, new_content, categories, is_draft)
            print(f"    [PUT] 成功: {result_url} (draft={'yes' if is_draft else 'no'})")
            fixed += 1
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="replace")[:300]
            print(f"    [ERR] PUT失敗: HTTP {e.code} / {err_body}")
            errors += 1
        time.sleep(1)  # レート制限回避

    print(f"\n集計: 修正={fixed} / スキップ={skipped} / エラー={errors}")
    return 0 if errors == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
