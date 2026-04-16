#!/usr/bin/env python3
"""令和8年度診療報酬改定ブログ記事にYouTube動画リンクを追加して更新"""

import re
import base64
import urllib.request
import urllib.error
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\r8_shinryouhoushu_kaitei_blog_hatena.html")
ENTRY_URL = "https://hinyan1016.hatenablog.com/entry/2026/04/10/141419"


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def get_entry_id(env):
    """一覧からエントリIDを取得"""
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
    auth_str = "{}:{}".format(hatena_id, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    req = urllib.request.Request(url, headers={
        "Authorization": "Basic {}".format(auth_b64),
    })

    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")

    # エントリURLに一致するエントリのedit URLを探す
    entries = body.split("<entry")
    for entry in entries[1:]:
        if "2026/04/10/141419" in entry:
            m = re.search(r'<link rel="edit"[^>]*href="([^"]+)"', entry)
            if m:
                return m.group(1)
    return None


def update_entry(env, edit_url, content):
    hatena_id = env["HATENA_ID"]
    api_key = env["HATENA_API_KEY"]

    auth_str = "{}:{}".format(hatena_id, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    title = "令和8年度診療報酬改定の全体像 ― +3.09%改定の内訳と医科の主要変更点を徹底解説"

    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
  <category term="診療報酬改定" />
  <category term="医療経営" />
  <category term="医学教育" />
  <app:control>
    <app:draft>no</app:draft>
  </app:control>
</entry>""".format(title=title, content=content)

    data = xml.encode("utf-8")
    req = urllib.request.Request(
        edit_url,
        data=data,
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth_b64),
        },
    )

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            entry_url = url_match.group(1) if url_match else "(URL取得失敗)"
            return True, entry_url
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        return False, "HTTP {}: {}".format(e.code, error_body[:500])


def main():
    env = load_env()

    print("エントリID取得中...")
    edit_url = get_entry_id(env)
    if not edit_url:
        print("[失敗] エントリが見つかりません")
        return

    print("edit URL: {}".format(edit_url))

    content = HTML_FILE.read_text(encoding="utf-8")
    print("HTML読込完了 ({} bytes)".format(len(content)))

    print("更新中...")
    ok, result = update_entry(env, edit_url, content)

    if ok:
        print("[成功] {}".format(result))
    else:
        print("[失敗] {}".format(result))


if __name__ == "__main__":
    main()
