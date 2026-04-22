#!/usr/bin/env python3
"""1000記事突破記事をはてなブログに下書き投稿/更新する

使い方:
  python post_1000_milestone.py            # 新規下書き投稿（POST）
  python post_1000_milestone.py --update   # 既存下書きを更新（PUT）
"""
import base64
import sys
import urllib.request
import urllib.error
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
HTML_FILE = SCRIPT_DIR / "1000_articles_milestone_hatena.html"
ENV_FILE = SCRIPT_DIR.parent / "youtube-slides" / "食事指導シリーズ" / "_shared" / ".env"

TITLE = "ブログ開始から1年余りで1000記事を突破。AIとともに続けてきた発信の現在地"
CATEGORIES = ["雑記", "AI活用", "ブログ運営"]
EDIT_URL = "https://blog.hatena.ne.jp/hinyan1016/hinyan1016.hatenablog.com/atom/entry/17179246901379254551"


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def build_atom(title, content, cats, draft=True):
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in cats)
    return """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{cats}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(
        title=title,
        content=content,
        cats=cat_xml,
        draft="yes" if draft else "no",
    )


def main():
    update = "--update" in sys.argv
    env = load_env()
    with open(HTML_FILE, encoding="utf-8") as f:
        body = f.read()

    if update:
        url = EDIT_URL
        method = "PUT"
    else:
        url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(
            env["HATENA_ID"], env["HATENA_BLOG_DOMAIN"]
        )
        method = "POST"

    auth = base64.b64encode(
        "{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"]).encode()
    ).decode()

    xml = build_atom(TITLE, body, CATEGORIES, draft=True)
    req = urllib.request.Request(
        url,
        data=xml.encode("utf-8"),
        method=method,
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic " + auth,
        },
    )

    print("title: {}".format(TITLE))
    print("mode : draft ({})".format(method))
    print("size : {:,} bytes".format(len(body.encode("utf-8"))))
    print("url  : {}".format(url))
    print("sending...")

    try:
        with urllib.request.urlopen(req) as resp:
            body_resp = resp.read().decode("utf-8")
            import re
            m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body_resp)
            edit_m = re.search(r'<link rel="edit"[^>]*href="([^"]+)"', body_resp)
            print("[OK] HTTP {}".format(resp.status))
            print("alternate: {}".format(m.group(1) if m else "(n/a)"))
            print("edit     : {}".format(edit_m.group(1) if edit_m else "(n/a)"))
    except urllib.error.HTTPError as e:
        print("[FAIL] HTTP {}".format(e.code))
        print(e.read().decode("utf-8", errors="replace")[:500])


if __name__ == "__main__":
    main()
