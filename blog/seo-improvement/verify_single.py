"""単一記事でref-item/ref-contentの正しい検出を検証"""
import re
import sys
import io
from urllib.request import urlopen, Request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

url = sys.argv[1] if len(sys.argv) > 1 else "https://hinyan1016.hatenablog.com/entry/2025/10/19/050931"
req = Request(url, headers={"User-Agent": "Mozilla/5.0 audit"})
html = urlopen(req, timeout=20).read().decode("utf-8", errors="ignore")

# ref-item要素ごとにブロックを取得する（非貪欲ではなく次のref-item or 終端まで）
starts = [m.start() for m in re.finditer(r'<[a-z]+[^>]*class="[^"]*ref-item[^"]*"', html, re.IGNORECASE)]
starts.append(len(html))

print(f"URL: {url}")
print(f"ref-item数: {len(starts) - 1}")

with_link = 0
without_link = 0
samples_no_link = []
for i in range(len(starts) - 1):
    block = html[starts[i]:starts[i + 1]]
    # ref-contentの範囲を抽出してその中で<a>を探すのが正しい
    # ただしblockはref-item全体なので、その中のref-content内の<a>を判定
    content_match = re.search(r'class="[^"]*ref-content[^"]*"[^>]*>(.+?)(?=<[a-z]+[^>]*class="[^"]*ref-item|\Z)', block, re.IGNORECASE | re.DOTALL)
    if content_match:
        content_html = content_match.group(1)
        # ref-content要素の閉じタグまでで切る: 次のdiv/spanの閉じタグを探す
        # 簡単に: content_html全体から<a href=を検索
        if re.search(r'<a\s[^>]*href=', content_html[:2000], re.IGNORECASE):
            with_link += 1
        else:
            without_link += 1
            if len(samples_no_link) < 3:
                samples_no_link.append(content_html[:200])
    else:
        # ref-contentが見つからない場合は block全体で判定
        if re.search(r'<a\s[^>]*href=', block, re.IGNORECASE):
            with_link += 1
        else:
            without_link += 1

print(f"<a>あり: {with_link}")
print(f"<a>なし: {without_link}")
print(f"\n--- <a>なしサンプル ---")
for s in samples_no_link:
    print(repr(s[:180]))
    print("---")
