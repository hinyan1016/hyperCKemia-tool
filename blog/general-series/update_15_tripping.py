#!/usr/bin/env python3
"""
からだの不思議 #15 ブログ更新スクリプト
- ローカルHTML（インラインスタイル済み）をそのままPUT
- VideoObject + FAQ JSON-LD + YouTube動画リンク + インフォグラフィック含む
"""

import sys
import re
import urllib.request
import urllib.error
import base64
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")

BLOG_TITLE = "「最近つまずきやすい」は脳の問題？ ― 歩行とつまずきの原因を解剖学レベルで解説【からだの不思議 #15】"
ENTRY_ID = "17179246901372000531"


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def main():
    categories = ["医学教育", "からだの不思議", "歩行障害", "脳神経内科"]

    print("=== からだの不思議 #15 ブログ更新 ===\n")

    env = load_env()

    # 1. ローカルHTML読み込み
    html_path = Path(__file__).parent / "15_frequent_tripping_blog.html"
    print("[1/2] ローカルHTML読み込み...")
    with open(html_path, encoding="utf-8") as f:
        content = f.read().strip()
    print("  読み込み完了（{}文字）".format(len(content)))

    # 残存classチェック
    classes = re.findall(r'class="[^"]*"', content)
    if classes:
        print("  WARNING: remaining classes: {}".format(classes))
    else:
        print("  classチェック: OK（残存なし）")

    # JSON-LDチェック
    if "VideoObject" in content:
        print("  VideoObject JSON-LD: OK")
    else:
        print("  WARNING: VideoObject JSON-LD なし")

    if "FAQPage" in content:
        print("  FAQ JSON-LD: OK")
    else:
        print("  WARNING: FAQ JSON-LD なし")

    # 2. PUT更新
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]

    auth_str = "{}:{}".format(hatena_id, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    entry_id = ENTRY_ID
    print("\n[2/2] はてなブログ更新中...")

    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)

    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>no</app:draft>
  </app:control>
</entry>""".format(title=BLOG_TITLE, content=content, categories=cat_xml)

    put_url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        hatena_id, blog_domain, entry_id
    )
    data = xml.encode("utf-8")
    req_put = urllib.request.Request(put_url, data=data, method="PUT", headers={
        "Content-Type": "application/xml; charset=utf-8",
        "Authorization": "Basic {}".format(auth_b64),
    })

    try:
        with urllib.request.urlopen(req_put) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            entry_url = url_match.group(1) if url_match else "(URL取得失敗)"
            print("\n[成功] ブログ更新完了")
            print("  URL: {}".format(entry_url))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        print("\n[失敗] HTTP {}: {}".format(e.code, error_body[:500]))
        sys.exit(1)


if __name__ == "__main__":
    main()
