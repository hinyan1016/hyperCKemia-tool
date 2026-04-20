"""
ref-contentテキストからPubMed検索で PMID を特定し、リンクを付与するスクリプト。

特徴:
- 本文にPMID/URL/DOIがない「著者+誌名+年」の具体文献が対象
- PubMed E-utilities API (esearch) で検索
- タイトルの類似度照合（最初30文字の完全一致）で誤マッチを回避
- 結果をCSVにログし、確定分のみAtomPubで記事更新
"""

from __future__ import annotations

import csv
import json
import logging
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Final

sys.path.insert(0, str(Path(__file__).parent))
import fix_reflinks  # type: ignore

ESEARCH: Final = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
ESUMMARY: Final = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
LOG: Final = logging.getLogger("fix_reflinks_pubmed")

# NCBI利用規約: 3req/sec以下。安全マージンを取って 0.5 sec 間隔
API_DELAY = 0.5


def strip_html(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", s)).strip()


def looks_specific(text: str) -> bool:
    """PubMed検索対象となる具体的文献かどうかを粗く判定"""
    plain = strip_html(text)
    if len(plain) < 30:
        return False
    has_author = bool(re.search(r"[A-Z][a-z]+\s+[A-Z][A-Z]?", plain))
    has_year = bool(re.search(r"\b(19|20)\d{2}[;,]\s*\d", plain)) or bool(
        re.search(r"\((19|20)\d{2}\)", plain)
    )
    return has_author and has_year


def build_query(text: str) -> str | None:
    """文献文字列からPubMed用クエリを構築。
    - 最初の著者の姓 + title の最初の数語 + 年を組み合わせる
    """
    plain = strip_html(text)
    # 著者の1人目を抜く（姓+イニシャル：例 "Porter RJ,"）
    m = re.match(r"([A-Z][a-z\-]+)\s+[A-Z][A-Z]?", plain)
    if not m:
        return None
    author = m.group(1)
    # 年 (1990-2029想定)
    year_match = re.search(r"\b(19|20)\d{2}\b", plain)
    year = year_match.group(0) if year_match else ""
    # タイトル候補: "著者群. タイトル. 誌名" のタイトル部分
    # 最初のピリオド後〜次のピリオドまで
    after_first = plain[m.end():]
    # Skip until we hit ". " (end of author list) then take next sentence
    title_match = re.search(r"\.\s+([A-Z][^.]{20,140})\.", after_first)
    title = title_match.group(1) if title_match else ""
    if not title:
        return None

    parts = [f"{author}[Author]"]
    if year:
        parts.append(f"{year}[PDAT]")
    # タイトルは最初の3-5単語を使う
    words = title.split()
    title_chunk = " ".join(words[:6])
    parts.append(f"{title_chunk}[Title]")
    return " AND ".join(parts)


def pubmed_search(query: str) -> list[str]:
    params = urllib.parse.urlencode(
        {"db": "pubmed", "term": query, "retmode": "json", "retmax": "3"}
    )
    url = f"{ESEARCH}?{params}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data.get("esearchresult", {}).get("idlist", [])
    except (urllib.error.URLError, json.JSONDecodeError, TimeoutError):
        return []


def pubmed_summary(pmid: str) -> dict:
    params = urllib.parse.urlencode(
        {"db": "pubmed", "id": pmid, "retmode": "json"}
    )
    url = f"{ESUMMARY}?{params}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data.get("result", {}).get(pmid, {})
    except (urllib.error.URLError, json.JSONDecodeError, TimeoutError):
        return {}


def verify_match(ref_text: str, summary: dict) -> bool:
    """PubMed summary titleが参考文献テキストと十分一致するか"""
    pm_title = summary.get("title", "")
    if not pm_title:
        return False
    # 両方を正規化（小文字/記号除去）して先頭20文字比較
    def norm(s: str) -> str:
        return re.sub(r"[^a-z0-9]", "", s.lower())
    pm_norm = norm(pm_title)
    ref_norm = norm(strip_html(ref_text))
    if len(pm_norm) < 20:
        return False
    return pm_norm[:30] in ref_norm  # PubMedタイトルの先頭30文字が参考文献内に含まれる


def process_content_pubmed(content: str, delay: float = API_DELAY) -> tuple[str, int, int, list[dict]]:
    """content内の<a>なしref-itemに対しPubMed検索で可能な限りリンク付与"""
    ranges = fix_reflinks.find_ref_items_without_link(content)
    if not ranges:
        return content, 0, 0, []
    linked = 0
    skipped = 0
    records: list[dict] = []
    new_content = content

    for start, end in reversed(ranges):
        block = new_content[start:end]
        span = fix_reflinks.find_ref_content_span(block)
        if not span:
            skipped += 1
            continue
        inner_start, inner_end = span
        inner = block[inner_start:inner_end]
        inner_plain = strip_html(inner)
        if not looks_specific(inner):
            skipped += 1
            continue
        query = build_query(inner)
        if not query:
            skipped += 1
            records.append({"text": inner_plain[:100], "reason": "no_query"})
            continue
        pmids = pubmed_search(query)
        time.sleep(delay)
        if not pmids:
            skipped += 1
            records.append({"text": inner_plain[:100], "reason": "no_hit", "query": query[:80]})
            continue
        # 先頭のPMIDで検証
        pmid = pmids[0]
        summary = pubmed_summary(pmid)
        time.sleep(delay)
        if not verify_match(inner, summary):
            skipped += 1
            records.append({"text": inner_plain[:100], "reason": "mismatch", "pmid": pmid, "pm_title": summary.get("title", "")[:80]})
            continue
        url = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"
        new_inner = fix_reflinks.wrap_with_link(inner, url)
        new_block = block[:inner_start] + new_inner + block[inner_end:]
        new_content = new_content[:start] + new_block + new_content[end:]
        linked += 1
        records.append({"text": inner_plain[:100], "reason": "linked", "pmid": pmid, "url": url})

    return new_content, linked, skipped, records


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-log", default=str(Path(__file__).parent / "audit_reflinks_v2.log"))
    parser.add_argument("--result-csv", default=str(Path(__file__).parent / "pubmed_fix_log.csv"))
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--limit", type=int, default=0, help="テスト用: 最初のN記事のみ")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    # URL抽出
    text = Path(args.audit_log).read_text(encoding="utf-8")
    urls = list(dict.fromkeys(re.findall(r"https://hinyan1016\.hatenablog\.com/entry/\S+", text)))
    if args.limit > 0:
        urls = urls[: args.limit]
    LOG.info(f"対象 {len(urls)} 記事、{'[dry-run]' if args.dry_run else '[live update]'}")

    env = fix_reflinks.load_env()
    csv_path = Path(args.result_csv)
    with csv_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["ts", "url", "status", "linked", "skipped", "hits_sample"])
        for i, url in enumerate(urls, 1):
            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            LOG.info(f"--- [{i}/{len(urls)}] {url}")
            try:
                info = fix_reflinks.find_entry(env, url)
                if info is None:
                    writer.writerow([ts, url, "not_found", 0, 0, ""])
                    continue
                new_content, linked, skipped, records = process_content_pubmed(info.content)
                hit_sample = ";".join(
                    [r.get("pmid", "") for r in records if r.get("reason") == "linked"]
                )[:200]
                if linked == 0:
                    LOG.info(f"  リンク0 / スキップ {skipped}")
                    writer.writerow([ts, url, "no_link", 0, skipped, hit_sample])
                    continue
                if args.dry_run:
                    LOG.info(f"  [dry-run] 付与可 {linked} / スキップ {skipped}")
                    writer.writerow([ts, url, "dry_run", linked, skipped, hit_sample])
                    continue
                ok = fix_reflinks.put_entry(env, info, new_content)
                status = "updated" if ok else "put_failed"
                LOG.info(f"  付与 {linked}件 / スキップ {skipped}件 / {status}")
                writer.writerow([ts, url, status, linked, skipped, hit_sample])
            except Exception as e:  # noqa: BLE001
                LOG.error(f"  ERROR: {e}")
                writer.writerow([ts, url, "error", 0, 0, str(e)[:150]])
            time.sleep(1.0)


if __name__ == "__main__":
    main()
