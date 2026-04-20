"""
ブログ参考文献リンク監査v2。
ref-item要素ブロックを正しく抽出し、リンク不足かつ汎用記述でない
（実在する論文/書籍参照）ものだけを検出する。
"""
import re
import sys
import io
import time
from urllib.request import urlopen, Request
from urllib.error import URLError

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = "https://hinyan1016.hatenablog.com"
MONTHS = [
    "2025/07", "2025/08", "2025/09", "2025/10", "2025/11", "2025/12",
    "2026/01", "2026/02", "2026/03", "2026/04",
]

# 汎用記述判定の正規表現 - これらはリンク付与対象外
GENERIC_PATTERNS = [
    r'に関する総説$',
    r'に関する研究$',
    r'の概要$',
    r'の文献$',
    r'の研究$',
    r'に関する解説$',
    r'関連の研究論文$',
    r'有病率調査$',
    r'^疫学',
    r'関連する医学文献',
]

def is_generic(text):
    """汎用記述（リンク不要）の判定。短く&著者・PMID・年号がない場合。"""
    # HTMLタグを除去
    plain = re.sub(r'<[^>]+>', '', text).strip()
    plain = re.sub(r'\s+', ' ', plain)
    if len(plain) < 10:
        return True
    # PMID/DOI/URL/西暦年を含むなら具体的な文献
    if re.search(r'PMID[:：]\s*\d+', plain, re.IGNORECASE):
        return False
    if re.search(r'doi[:：]', plain, re.IGNORECASE):
        return False
    if re.search(r'\b(19|20)\d{2}[;,]\s*\d', plain):  # 雑誌書式 "2024;45(3)" のような
        return False
    if re.search(r'[A-Z][a-z]+\s+[A-Z][A-Z]?', plain):  # 著者名 "Patel H" のような
        return False
    # 汎用記述パターン
    for pat in GENERIC_PATTERNS:
        if re.search(pat, plain):
            return True
    # 上記のいずれにも当てはまらず短い → 汎用可能性
    if len(plain) < 40:
        return True
    return False

def fetch(url, retries=2):
    for i in range(retries + 1):
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0 audit-bot"})
            with urlopen(req, timeout=15) as resp:
                return resp.read().decode("utf-8", errors="ignore")
        except URLError:
            if i == retries:
                return ""
            time.sleep(1)

def get_archive_entries(year_month):
    urls = set()
    page = 1
    while page <= 10:
        url = f"{BASE}/archive/{year_month}?page={page}"
        html = fetch(url)
        if not html:
            break
        found = re.findall(r'entry-title-link[^>]*href="([^"]+)"', html)
        new_ones = [u for u in found if u not in urls]
        if not new_ones:
            break
        urls.update(new_ones)
        page += 1
    return sorted(urls)

def extract_title(html):
    m = re.search(r'<h1[^>]*class="entry-title"[^>]*>.*?<a[^>]*>(.*?)</a>', html, re.DOTALL)
    if m:
        return re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', '', m.group(1))).strip()
    m = re.search(r'<title>([^<]+)</title>', html)
    return m.group(1).strip() if m else ""

def audit_article(html):
    """記事HTMLを分析し、リンク不足かつ具体的文献の数を返す"""
    starts = [m.start() for m in re.finditer(r'<[a-z]+[^>]*class="[^"]*ref-item[^"]*"', html, re.IGNORECASE)]
    if not starts:
        return 0, 0, 0, []
    starts.append(len(html))

    total = len(starts) - 1
    with_link = 0
    missing_specific = 0
    samples = []

    for i in range(total):
        block = html[starts[i]:starts[i + 1]]
        # ref-content内のリンク有無
        has_link = bool(re.search(r'<a\s[^>]*href=', block, re.IGNORECASE))
        if has_link:
            with_link += 1
            continue
        # リンクなし → 具体的文献か汎用記述か判定
        content_match = re.search(r'class="[^"]*ref-content[^"]*"[^>]*>(.+)', block, re.IGNORECASE | re.DOTALL)
        if content_match:
            content_html = content_match.group(1)
            # 最初の閉じタグまで
            end = content_html.find('</span>')
            if end == -1:
                end = content_html.find('</div>')
            content_html = content_html[:end] if end != -1 else content_html[:500]
            if not is_generic(content_html):
                missing_specific += 1
                if len(samples) < 2:
                    plain = re.sub(r'<[^>]+>', '', content_html).strip()[:150]
                    samples.append(plain)
    return total, with_link, missing_specific, samples

def main():
    all_flagged = []
    stats = {}
    for ym in MONTHS:
        entries = get_archive_entries(ym)
        stats[ym] = {"entries": len(entries), "flagged": 0, "specific_missing": 0}
        print(f"--- {ym} ({len(entries)} entries) ---", flush=True)
        for entry_url in entries:
            html = fetch(entry_url)
            if not html:
                continue
            total, with_link, missing_specific, samples = audit_article(html)
            if missing_specific > 0:
                stats[ym]["flagged"] += 1
                stats[ym]["specific_missing"] += missing_specific
                all_flagged.append({
                    "ym": ym,
                    "url": entry_url,
                    "total": total,
                    "with_link": with_link,
                    "missing_specific": missing_specific,
                    "title": extract_title(html),
                    "samples": samples,
                })
            time.sleep(0.12)
    # 結果
    print("\n=== 月別サマリー ===")
    total_articles = sum(s["entries"] for s in stats.values())
    total_flagged = sum(s["flagged"] for s in stats.values())
    total_missing = sum(s["specific_missing"] for s in stats.values())
    for ym, s in stats.items():
        print(f"  {ym}: {s['entries']:3d}件中 {s['flagged']:2d}件に具体文献のリンク不足計{s['specific_missing']}個")
    print(f"\n合計: 全{total_articles}記事、{total_flagged}記事に具体的文献のリンク不足 計{total_missing}個")

    print(f"\n=== 要対応記事 一覧 (具体文献のリンク不足が多い順) ===")
    all_flagged.sort(key=lambda x: -x["missing_specific"])
    for m in all_flagged:
        print(f"[{m['missing_specific']}件不足/全{m['total']}] {m['title'][:50]}")
        print(f"  {m['url']}")
        for s in m["samples"]:
            print(f"  例: {s[:120]}")
        print()

if __name__ == "__main__":
    main()
