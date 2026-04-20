"""
ブログ記事の参考文献にリンクを自動付与するスクリプト。

処理フロー:
  1. 対象記事のAtomPub edit_urlを特定（GET from entry list）
  2. edit_url からGETして現在のcontent HTMLを取得
  3. ref-item ブロックごとに <a href> の有無を判定
  4. 無い場合は本文から PMID / doi / URL を抽出して <a> で包む
  5. 改変 content を PUT し記事を更新

実行:
  python fix_reflinks.py --url <blog_entry_url> [--dry-run]

注意:
  - 既存 link がある ref-item は変更しない
  - PMID/DOI/URLが本文に見つからなければスキップ（後続のPubMed検索フェーズで対処）
  - dry-run で差分だけ確認可能
"""

from __future__ import annotations

import argparse
import base64
import logging
import re
import sys
import time
import urllib.error
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Final

ENV_FILE: Final = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env"
)

ATOM_NS = {"atom": "http://www.w3.org/2005/Atom", "app": "http://www.w3.org/2007/app"}

log = logging.getLogger("fix_reflinks")


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def auth_header(env: dict[str, str]) -> str:
    auth_str = f"{env['HATENA_ID']}:{env['HATENA_API_KEY']}"
    return "Basic " + base64.b64encode(auth_str.encode()).decode()


@dataclass(frozen=True)
class EntryInfo:
    edit_url: str
    title: str
    categories: tuple[str, ...]
    content: str
    draft: bool


def find_entry(env: dict[str, str], target_url: str) -> EntryInfo | None:
    """alternate link == target_url のエントリを AtomPub 一覧から検索"""
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_url: str | None = f"https://blog.hatena.ne.jp/{hatena_id}/{blog_domain}/atom/entry"

    page = 0
    while api_url and page < 200:
        page += 1
        req = urllib.request.Request(
            api_url, headers={"Authorization": auth_header(env)}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read().decode("utf-8")
        root = ET.fromstring(body)
        for entry in root.findall("atom:entry", ATOM_NS):
            alt_link = None
            edit_link = None
            for link in entry.findall("atom:link", ATOM_NS):
                if link.get("rel") == "alternate":
                    alt_link = link.get("href")
                elif link.get("rel") == "edit":
                    edit_link = link.get("href")
            if alt_link == target_url and edit_link:
                return _parse_entry(entry, edit_link)

        next_url: str | None = None
        for link in root.findall("atom:link", ATOM_NS):
            if link.get("rel") == "next":
                next_url = link.get("href")
                break
        api_url = next_url
    return None


def get_entry_by_edit_url(env: dict[str, str], edit_url: str) -> EntryInfo:
    req = urllib.request.Request(edit_url, headers={"Authorization": auth_header(env)})
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = resp.read().decode("utf-8")
    entry = ET.fromstring(body)
    return _parse_entry(entry, edit_url)


def _parse_entry(entry: ET.Element, edit_url: str) -> EntryInfo:
    title_elem = entry.find("atom:title", ATOM_NS)
    content_elem = entry.find("atom:content", ATOM_NS)
    cats = tuple(
        c.get("term", "")
        for c in entry.findall("atom:category", ATOM_NS)
        if c.get("term")
    )
    draft_elem = entry.find("app:control/app:draft", ATOM_NS)
    draft = (draft_elem is not None) and (draft_elem.text == "yes")
    return EntryInfo(
        edit_url=edit_url,
        title=title_elem.text if title_elem is not None and title_elem.text else "",
        categories=cats,
        content=content_elem.text if content_elem is not None and content_elem.text else "",
        draft=draft,
    )


def put_entry(env: dict[str, str], info: EntryInfo, new_content: str) -> bool:
    draft_val = "yes" if info.draft else "no"
    cat_xml = "\n".join(f'  <category term="{_xml_escape(c)}" />' for c in info.categories)
    xml = f"""<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{_xml_escape(info.title)}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{new_content}]]></content>
{cat_xml}
  <app:control>
    <app:draft>{draft_val}</app:draft>
  </app:control>
</entry>"""
    data = xml.encode("utf-8")
    req = urllib.request.Request(
        info.edit_url,
        data=data,
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth_header(env),
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            log.info(f"  PUT成功: HTTP {resp.status}")
            return True
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        log.error(f"  PUT失敗: HTTP {e.code}: {body[:400]}")
        return False


def _xml_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


# --- リンク判定＆置換ロジック -----------------------------------------------

PMID_PATTERN = re.compile(r"PMID[:：]\s*(\d{5,9})", re.IGNORECASE)
DOI_PATTERN = re.compile(r"\b(?:doi[:：]\s*)?(10\.\d{4,9}/[\w\-\.:;()/]+)", re.IGNORECASE)
URL_PATTERN = re.compile(r"https?://[^\s<>\"']+[^\s<>\"'.,)]", re.IGNORECASE)


def find_ref_items_without_link(content: str) -> list[tuple[int, int]]:
    """ref-item ブロックの (start, end) リスト。リンクを持たないもののみ。"""
    starts = [m.start() for m in re.finditer(
        r'<[a-z]+[^>]*class="[^"]*ref-item[^"]*"', content, re.IGNORECASE
    )]
    if not starts:
        return []
    boundaries = starts + [len(content)]
    ranges: list[tuple[int, int]] = []
    for i in range(len(starts)):
        block_start = boundaries[i]
        block_end = boundaries[i + 1]
        block = content[block_start:block_end]
        if re.search(r'<a\s[^>]*href=', block, re.IGNORECASE):
            continue
        ranges.append((block_start, block_end))
    return ranges


def find_ref_content_span(block: str) -> tuple[int, int] | None:
    """block内の <span class="ref-content"> の内側 (start, end)。見つからなければNone。"""
    m = re.search(r'<(span|div)[^>]*class="[^"]*ref-content[^"]*"[^>]*>', block, re.IGNORECASE)
    if not m:
        return None
    tag = m.group(1)
    inner_start = m.end()
    # 対応する閉じタグを探す（ネスト無視で最初の同種閉じタグ）
    close_match = re.search(rf'</{tag}>', block[inner_start:], re.IGNORECASE)
    if not close_match:
        return None
    return inner_start, inner_start + close_match.start()


def extract_link_target(text: str) -> str | None:
    """ref-content テキストから PMID / DOI / URL を順に探してリンクURLを返す。"""
    plain = re.sub(r'<[^>]+>', '', text)
    # 1. PMID
    m = PMID_PATTERN.search(plain)
    if m:
        return f"https://pubmed.ncbi.nlm.nih.gov/{m.group(1)}/"
    # 2. URL（平文中）
    m = URL_PATTERN.search(plain)
    if m:
        return m.group(0)
    # 3. DOI
    m = DOI_PATTERN.search(plain)
    if m:
        return f"https://doi.org/{m.group(1)}"
    return None


def wrap_with_link(ref_content_inner: str, url: str) -> str:
    """ref-content の内側テキスト全体を <a href="url" target="_blank"> で包む。"""
    safe_url = _xml_escape(url)
    return f'<a href="{safe_url}" target="_blank" rel="noopener noreferrer">{ref_content_inner}</a>'


def process_content(content: str) -> tuple[str, int, int]:
    """
    戻り値: (new_content, linked_count, skipped_count)
    """
    ranges = find_ref_items_without_link(content)
    if not ranges:
        return content, 0, 0

    # 後ろから処理して位置ずれを回避
    linked = 0
    skipped = 0
    new_content = content
    for start, end in reversed(ranges):
        block = new_content[start:end]
        span = find_ref_content_span(block)
        if not span:
            skipped += 1
            continue
        inner_start, inner_end = span
        inner = block[inner_start:inner_end]
        url = extract_link_target(inner)
        if not url:
            skipped += 1
            continue
        new_inner = wrap_with_link(inner, url)
        new_block = block[:inner_start] + new_inner + block[inner_end:]
        new_content = new_content[:start] + new_block + new_content[end:]
        linked += 1
    return new_content, linked, skipped


# --- メイン処理 --------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", required=True, help="対象ブログ記事URL (alternate link)")
    parser.add_argument("--edit-url", help="AtomPub edit URL（既知の場合は--urlの検索を省略）")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--log-level", default="INFO")
    args = parser.parse_args()

    logging.basicConfig(level=args.log_level, format="%(message)s")

    env = load_env()
    log.info(f"[1] エントリ特定: {args.url}")
    if args.edit_url:
        info = get_entry_by_edit_url(env, args.edit_url)
    else:
        info = find_entry(env, args.url)
        if info is None:
            log.error("  エントリが見つかりませんでした")
            sys.exit(1)
    log.info(f"  edit_url = {info.edit_url}")
    log.info(f"  title    = {info.title[:60]}")
    log.info(f"  content  = {len(info.content)} chars")

    log.info("[2] ref-item分析＆リンク付与")
    new_content, linked, skipped = process_content(info.content)
    log.info(f"  リンク付与: {linked}件 / スキップ: {skipped}件")

    if linked == 0:
        log.info("  付与対象なし。終了。")
        return

    if args.dry_run:
        log.info("[3] dry-run: 更新せず終了")
        return

    log.info("[3] PUT (記事更新)")
    put_entry(env, info, new_content)


if __name__ == "__main__":
    main()
