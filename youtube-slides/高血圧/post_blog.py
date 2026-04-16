#!/usr/bin/env python3
"""高血圧ブログ記事をはてなブログに下書き投稿する"""

import re
import base64
import urllib.request
import urllib.error
from pathlib import Path

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"

BLOG_HTML = Path(__file__).parent / "高血圧_blog.html"
CATEGORIES = ["高血圧", "生活習慣病", "医学教育"]


def main():
    content = BLOG_HTML.read_text(encoding="utf-8")

    # タイトル抽出
    m = re.search(r"<h1>(.*?)</h1>", content)
    title = m.group(1) if m else "高血圧の新基準と減塩対策"

    # YouTube冒頭の前に導入文を挿入（meta description対策）
    intro = '<p style="font-size:1.1em;margin-bottom:1em;"><strong>{}</strong></p>\n'.format(
        title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )
    yt_match = re.search(r'(<div[^>]*>\s*<iframe[^>]*youtube\.com)', content, re.IGNORECASE)
    if yt_match:
        pos = yt_match.start()
        content = content[:pos] + intro + content[pos:]

    # AtomPub XML構築
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in CATEGORIES)
    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>{author}</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>yes</app:draft>
  </app:control>
</entry>""".format(title=title, author=HATENA_ID, content=content, categories=cat_xml)

    # 投稿
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(HATENA_ID, BLOG_DOMAIN)
    auth = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()

    req = urllib.request.Request(
        url,
        data=xml.encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth),
        },
    )

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            entry_url = url_match.group(1) if url_match else "(URL取得失敗)"
            print("[成功] 下書き投稿完了")
            print("URL: {}".format(entry_url))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        print("[失敗] HTTP {}: {}".format(e.code, error_body[:300]))


if __name__ == "__main__":
    main()
