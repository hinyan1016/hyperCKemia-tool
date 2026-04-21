"""
fix_reflinks_pubmed.py の緩和版 v2。

変更点:
1. verify_match: 先頭30→先頭18 + 単語集合の40%以上一致を許容
2. build_query: 3段階でトライ (A: 著者+タイトル / B: 著者+年+誌名の一部 / C: タイトル長め)
3. 複数候補(最大5件)をsummaryで検証し最良マッチを採用
4. --limit / --min-missing / --urls で対象を絞れる
5. CSVにmatch_score / strategyを記録
"""

from __future__ import annotations

import argparse
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
LOG: Final = logging.getLogger("fix_reflinks_pubmed_v2")

API_DELAY = 0.4


def strip_html(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", s)).strip()


def looks_specific(text: str) -> bool:
    plain = strip_html(text)
    if len(plain) < 25:
        return False
    has_author = bool(re.search(r"[A-Z][a-z]+\s+[A-Z][A-Z]?", plain))
    has_year = bool(re.search(r"\b(19|20)\d{2}\b", plain))
    return has_author and has_year


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


def tokens(s: str) -> set[str]:
    words = re.findall(r"[A-Za-z]{4,}", s.lower())
    STOP = {"with", "from", "this", "that", "have", "been", "were", "also",
            "after", "about", "their", "other", "such", "than", "into",
            "study", "review", "trial", "patient", "patients"}
    return {w for w in words if w not in STOP}


def extract_author(plain: str) -> str | None:
    m = re.match(r"([A-Z][a-z\-]+)\s+[A-Z][A-Z]?", plain)
    return m.group(1) if m else None


def extract_year(plain: str) -> str | None:
    m = re.search(r"\b(19|20)\d{2}\b", plain)
    return m.group(0) if m else None


def extract_title(plain: str) -> str | None:
    """タイトルを抽出。複数フォーマット対応。
    1. "Last FN, Last FN. Title words here. Journal. Year." (著者→タイトル)
    2. "Title words here. Journal. Year." (著者なし)
    """
    # フォーマット1: 著者リストの後 (et al. / FN. の直後)
    m = re.search(r"(?:et\s+al\.|[A-Z][A-Z]?\.)\s+([A-Z][^.]{10,200}?)\.\s", plain)
    if m:
        return m.group(1).strip()
    # フォーマット2: 文先頭がタイトル（コロン含み可）
    parts = re.split(r"\.\s+", plain)
    if parts:
        first = parts[0].strip()
        # 著者リスト/書誌情報パターンは除外
        if 15 < len(first) < 250 and re.match(r"^[A-Z]", first):
            # ただし "Author N, Author N" 形式は除外
            if not re.match(r"^[A-Z][a-z]+\s+[A-Z][A-Z]?,", first):
                return first
    # フォールバック: 著者後の最初の長い文
    for p in parts[1:]:
        if 20 < len(p) < 200 and re.match(r"^[A-Z]", p):
            return p.strip()
    return None


def build_queries(text: str) -> list[tuple[str, str]]:
    """[(strategy_name, query)] を返す。順に試す。"""
    plain = strip_html(text)
    author = extract_author(plain)
    year = extract_year(plain)
    title = extract_title(plain)

    out: list[tuple[str, str]] = []
    # A: 著者+タイトル先頭6語+年
    if author and title:
        words = title.split()
        chunk = " ".join(words[:6])
        parts = [f"{author}[Author]", f"{chunk}[Title]"]
        if year:
            parts.append(f"{year}[PDAT]")
        out.append(("A_author_title", " AND ".join(parts)))
    # B: タイトル先頭8語+年 (著者なしOK)
    if title:
        words = title.split()
        chunk = " ".join(words[:8])
        parts = [f"{chunk}[Title]"]
        if year:
            parts.append(f"{year}[PDAT]")
        out.append(("B_title_year", " AND ".join(parts)))
    # B2: タイトル先頭4語+年 (短縮 - 雑誌情報混じり対策)
    if title:
        words = title.split()
        if len(words) >= 4:
            chunk = " ".join(words[:4])
            parts = [f"{chunk}[Title]"]
            if year:
                parts.append(f"{year}[PDAT]")
            out.append(("B2_title_short", " AND ".join(parts)))
    # C: 著者+年のみ（複数ヒットから verify で絞る）
    if author and year:
        out.append(("C_author_year", f"{author}[Author] AND {year}[PDAT]"))
    return out


def pubmed_search(query: str, retmax: int = 5) -> list[str]:
    params = urllib.parse.urlencode(
        {"db": "pubmed", "term": query, "retmode": "json", "retmax": str(retmax)}
    )
    url = f"{ESEARCH}?{params}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data.get("esearchresult", {}).get("idlist", [])
    except (urllib.error.URLError, json.JSONDecodeError, TimeoutError):
        return []


def pubmed_summary(pmid: str) -> dict:
    params = urllib.parse.urlencode({"db": "pubmed", "id": pmid, "retmode": "json"})
    url = f"{ESUMMARY}?{params}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data.get("result", {}).get(pmid, {})
    except (urllib.error.URLError, json.JSONDecodeError, TimeoutError):
        return {}


def pubmed_summary_many(pmids: list[str]) -> dict[str, dict]:
    if not pmids:
        return {}
    params = urllib.parse.urlencode(
        {"db": "pubmed", "id": ",".join(pmids), "retmode": "json"}
    )
    url = f"{ESUMMARY}?{params}"
    try:
        with urllib.request.urlopen(url, timeout=20) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        result = data.get("result", {})
        out = {pid: result.get(pid, {}) for pid in pmids if pid in result}
        return out
    except (urllib.error.URLError, json.JSONDecodeError, TimeoutError):
        return {}


def score_match(ref_text: str, summary: dict) -> float:
    """0.0〜1.0 のマッチスコア。先頭一致+単語カバレッジの両方を要求。"""
    pm_title = summary.get("title", "")
    if not pm_title:
        return 0.0
    ref_plain = strip_html(ref_text)
    ref_norm = norm(ref_plain)
    pm_norm = norm(pm_title)
    if len(pm_norm) < 15:
        return 0.0

    # 単語カバレッジ
    pm_tokens = tokens(pm_title)
    ref_tokens = tokens(ref_plain)
    if not pm_tokens:
        return 0.0
    overlap = pm_tokens & ref_tokens
    cover = len(overlap) / len(pm_tokens)

    # 著者/年のボーナス
    bonus = 0.0
    ref_author = extract_author(ref_plain)
    if ref_author:
        authors = summary.get("authors", [])
        auth_list_norm = " ".join(norm(a.get("name", "")) for a in authors)
        if norm(ref_author) in auth_list_norm:
            bonus += 0.15
    ref_year = extract_year(ref_plain)
    if ref_year:
        pm_date = summary.get("pubdate", "")
        if ref_year in pm_date:
            bonus += 0.10

    # 強い signal: 先頭40文字一致（高信頼）
    if len(pm_norm) >= 40 and pm_norm[:40] in ref_norm:
        return min(1.0, 0.95 + bonus)
    # 中 signal: 先頭25文字一致 + カバレッジ補強
    if len(pm_norm) >= 25 and pm_norm[:25] in ref_norm and cover >= 0.4:
        return min(1.0, 0.80 + bonus + (cover - 0.4) * 0.3)
    # 弱: カバレッジ単独 (高い場合のみ)
    if cover >= 0.7:
        return min(1.0, cover * 0.85 + bonus)
    # 著者+年が一致してて、カバレッジも中くらいあれば許容
    if bonus >= 0.20 and cover >= 0.5:
        return min(1.0, cover * 0.7 + bonus)
    return cover * 0.5 + bonus  # 低スコア（参考値）


THRESHOLD = 0.75


def find_best_match(ref_text: str, pmids: list[str]) -> tuple[str | None, float, dict]:
    if not pmids:
        return None, 0.0, {}
    summaries = pubmed_summary_many(pmids[:5])
    best_pmid = None
    best_score = 0.0
    best_summary = {}
    for pid, summ in summaries.items():
        s = score_match(ref_text, summ)
        if s > best_score:
            best_score = s
            best_pmid = pid
            best_summary = summ
    return best_pmid, best_score, best_summary


def process_content_pubmed(content: str, delay: float = API_DELAY) -> tuple[str, int, int, list[dict]]:
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
        queries = build_queries(inner)
        if not queries:
            skipped += 1
            records.append({"text": inner_plain[:100], "reason": "no_query"})
            continue

        matched = False
        for strategy, query in queries:
            pmids = pubmed_search(query, retmax=5)
            time.sleep(delay)
            if not pmids:
                continue
            pmid, score, summary = find_best_match(inner, pmids)
            time.sleep(delay)
            if pmid and score >= THRESHOLD:
                url = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"
                new_inner = fix_reflinks.wrap_with_link(inner, url)
                new_block = block[:inner_start] + new_inner + block[inner_end:]
                new_content = new_content[:start] + new_block + new_content[end:]
                linked += 1
                records.append({
                    "text": inner_plain[:100],
                    "reason": "linked",
                    "pmid": pmid,
                    "url": url,
                    "score": round(score, 3),
                    "strategy": strategy,
                    "pm_title": summary.get("title", "")[:100],
                })
                matched = True
                break
            else:
                records.append({
                    "text": inner_plain[:80],
                    "reason": "low_score",
                    "pmid": pmid or "",
                    "score": round(score, 3),
                    "strategy": strategy,
                    "pm_title": summary.get("title", "")[:80] if summary else "",
                })
        if not matched:
            skipped += 1

    return new_content, linked, skipped, records


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-log", default=str(Path(__file__).parent / "audit_reflinks_v2.log"))
    parser.add_argument("--result-csv", default=str(Path(__file__).parent / "pubmed_fix_v2_log.csv"))
    parser.add_argument("--detail-csv", default=str(Path(__file__).parent / "pubmed_fix_v2_detail.csv"))
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--limit", type=int, default=0, help="最初のN記事のみ")
    parser.add_argument("--urls", nargs="*", help="明示URL群（audit-logを無視）")
    parser.add_argument("--urls-file", help="URLリストファイル（1行1URL）")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    if args.urls:
        urls = args.urls
    elif args.urls_file:
        urls = [u.strip() for u in Path(args.urls_file).read_text(encoding="utf-8").splitlines() if u.strip()]
    else:
        text = Path(args.audit_log).read_text(encoding="utf-8")
        urls = list(dict.fromkeys(re.findall(r"https://hinyan1016\.hatenablog\.com/entry/\S+", text)))
    if args.limit > 0:
        urls = urls[: args.limit]
    LOG.info(f"対象 {len(urls)} 記事、{'[dry-run]' if args.dry_run else '[live update]'} threshold={THRESHOLD}")

    env = fix_reflinks.load_env()

    detail_f = Path(args.detail_csv).open("w", encoding="utf-8", newline="")
    detail_w = csv.writer(detail_f)
    detail_w.writerow(["url", "reason", "strategy", "score", "pmid", "ref_text", "pm_title"])

    with Path(args.result_csv).open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["ts", "url", "status", "linked", "skipped"])
        for i, url in enumerate(urls, 1):
            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            LOG.info(f"--- [{i}/{len(urls)}] {url}")
            try:
                info = fix_reflinks.find_entry(env, url)
                if info is None:
                    writer.writerow([ts, url, "not_found", 0, 0])
                    continue
                new_content, linked, skipped, records = process_content_pubmed(info.content)
                for r in records:
                    detail_w.writerow([
                        url, r.get("reason", ""), r.get("strategy", ""),
                        r.get("score", ""), r.get("pmid", ""),
                        r.get("text", ""), r.get("pm_title", ""),
                    ])
                detail_f.flush()
                if linked == 0:
                    LOG.info(f"  リンク0 / スキップ {skipped}")
                    writer.writerow([ts, url, "no_link", 0, skipped])
                    continue
                if args.dry_run:
                    LOG.info(f"  [dry-run] 付与可 {linked} / スキップ {skipped}")
                    writer.writerow([ts, url, "dry_run", linked, skipped])
                    continue
                ok = fix_reflinks.put_entry(env, info, new_content)
                status = "updated" if ok else "put_failed"
                LOG.info(f"  付与 {linked}件 / スキップ {skipped}件 / {status}")
                writer.writerow([ts, url, status, linked, skipped])
            except Exception as e:  # noqa: BLE001
                LOG.error(f"  ERROR: {e}")
                writer.writerow([ts, url, "error", 0, 0])
            time.sleep(1.0)

    detail_f.close()


if __name__ == "__main__":
    main()
