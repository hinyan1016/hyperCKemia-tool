#!/usr/bin/env python3
"""低活動性せん妄ブログ記事を下書き再投稿"""
import re
import urllib.request
import urllib.error
import base64
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HATENA_HTML = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\hypoactive_delirium_blog_hatena.html")
TITLE = "低活動性せん妄 ― 「静かなせん妄」を見逃さないための実践ガイド"
CATEGORIES = ["医学教育", "脳神経内科", "せん妄"]

def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env

env = load_env()
with open(HATENA_HTML, encoding="utf-8") as f:
    content = f.read()

cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in CATEGORIES)
xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>yes</app:draft>
  </app:control>
</entry>""".format(title=TITLE, content=content, categories=cat_xml)

hatena_id = env["HATENA_ID"]
blog_domain = env["HATENA_BLOG_DOMAIN"]
api_key = env["HATENA_API_KEY"]
url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()

req = urllib.request.Request(url, data=xml.encode("utf-8"), method="POST", headers={
    "Content-Type": "application/xml; charset=utf-8",
    "Authorization": "Basic {}".format(auth_b64),
})

try:
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")
        url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
        entry_url = url_match.group(1) if url_match else "(URL取得失敗)"
        id_match = re.search(r'<id>tag:[^,]+,\d+:entry-(\d+)</id>', body)
        entry_id = id_match.group(1) if id_match else "(ID取得失敗)"
        print("投稿成功!")
        print("URL: {}".format(entry_url))
        print("Entry ID: {}".format(entry_id))
except urllib.error.HTTPError as e:
    print("失敗: {} {}".format(e.code, e.read().decode("utf-8", errors="replace")[:500]))
