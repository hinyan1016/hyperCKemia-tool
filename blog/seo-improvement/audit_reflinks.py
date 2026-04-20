"""
ブログ参考文献リンク監査スクリプト。
2025/07〜2026/03の全記事をスキャンして、ref-content内に<a>がないものを検出する。
"""
import re
import sys
import time
from html.parser import HTMLParser
from urllib.request import urlopen, Request
from urllib.error import URLError

BASE = "https://hinyan1016.hatenablog.com"
MONTHS = [
    "2025/07", "2025/08", "2025/09", "2025/10", "2025/11", "2025/12",
    "2026/01", "2026/02", "2026/03", "2026/04",
]

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
    """月別アーカイブ（複数ページ）から全記事URLを収集"""
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
        return re.sub(r'\s+', ' ', m.group(1).strip())
    m = re.search(r'<title>([^<]+)</title>', html)
    return m.group(1).strip() if m else ""

def count_reflinks(html):
    """ref-content要素を抽出し、<a>の有無を分類"""
    blocks = re.findall(r'<[^>]*class="[^"]*ref-content[^"]*"[^>]*>(.*?)</[a-z]+>', html, re.DOTALL | re.IGNORECASE)
    total = len(blocks)
    with_link = sum(1 for b in blocks if re.search(r'<a\s[^>]*href=', b, re.IGNORECASE))
    return total, with_link

def audit():
    missing = []
    total_articles = 0
    for ym in MONTHS:
        print(f"--- {ym} ---", flush=True)
        entries = get_archive_entries(ym)
        total_articles += len(entries)
        print(f"  {len(entries)} entries", flush=True)
        for entry_url in entries:
            html = fetch(entry_url)
            if not html:
                continue
            total, with_link = count_reflinks(html)
            if total > 0 and with_link < total:
                missing.append({
                    "url": entry_url,
                    "total": total,
                    "with_link": with_link,
                    "missing": total - with_link,
                    "title": extract_title(html),
                })
            time.sleep(0.15)
    return missing, total_articles

if __name__ == "__main__":
    missing, total = audit()
    print(f"\n=== スキャン結果 ===", flush=True)
    print(f"総記事数: {total}", flush=True)
    print(f"リンク不足記事: {len(missing)}", flush=True)
    missing.sort(key=lambda x: -x["missing"])
    for m in missing:
        print(f"  [{m['missing']}/{m['total']}] {m['title'][:60]} {m['url']}", flush=True)
