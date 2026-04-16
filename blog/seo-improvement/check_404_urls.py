#!/usr/bin/env python3
"""404 URLの生存確認スクリプト。各URLにHTTP HEADリクエストを送り、ステータスコードを記録する。"""

import csv
import urllib.request
import urllib.error
import time
import ssl

INPUT_CSV = "blog/seo-improvement/404_urls.csv"
OUTPUT_CSV = "blog/seo-improvement/404_check_result.csv"

# SSL検証を無効化（社内ネットワーク等の問題回避）
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

def check_url(url):
    """URLにGETリクエストを送り、ステータスコードを返す"""
    try:
        req = urllib.request.Request(url, method="GET")
        req.add_header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
            return resp.status, ""
    except urllib.error.HTTPError as e:
        return e.code, str(e.reason)
    except urllib.error.URLError as e:
        return 0, str(e.reason)
    except Exception as e:
        return 0, str(e)

def main():
    with open(INPUT_CSV, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    results = []
    for i, row in enumerate(rows):
        url = row["url"]
        print(f"[{i+1}/{len(rows)}] {url[:80]}...", end=" ", flush=True)
        status, error = check_url(url)
        print(f"→ {status} {error}")
        results.append({
            "url": url,
            "crawl_date": row["crawl_date"],
            "category": row["category"],
            "http_status": status,
            "error": error,
        })
        time.sleep(0.3)  # はてなへの負荷軽減

    # 結果をCSV出力
    with open(OUTPUT_CSV, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["url", "crawl_date", "category", "http_status", "error"])
        writer.writeheader()
        writer.writerows(results)

    # サマリー
    status_counts = {}
    for r in results:
        s = r["http_status"]
        status_counts[s] = status_counts.get(s, 0) + 1

    print("\n=== サマリー ===")
    for s, c in sorted(status_counts.items()):
        print(f"  HTTP {s}: {c}件")
    print(f"  合計: {len(results)}件")
    print(f"\n結果: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
