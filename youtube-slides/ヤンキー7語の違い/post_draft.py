#!/usr/bin/env python3
"""ヤンキー7語の違い 記事を下書き投稿する"""
import base64
import re
import urllib.request
import urllib.error
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
HTML_FILE = SCRIPT_DIR / "blog_hatena.html"
ENV_FILE = Path("C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/食事指導シリーズ/_shared/.env")

TITLE = "ヤンキー・ツッパリ・不良・やんちゃ・チンピラ・暴走族・半グレの違いとは"
CATEGORIES = ["日本語雑学", "意味の違い", "言葉の歴史"]


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def build_entry(title, content, categories, draft=True):
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)
    return """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(title=title, content=content, categories=cat_xml, draft=draft_val)


def post(env, xml):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(
        env["HATENA_ID"], env["HATENA_BLOG_DOMAIN"]
    )
    auth = base64.b64encode("{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"]).encode()).decode()
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
            m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            edit_m = re.search(r'<link rel="edit"[^>]*href="([^"]+)"', body)
            return True, (m.group(1) if m else "(url unknown)"), (edit_m.group(1) if edit_m else "")
    except urllib.error.HTTPError as e:
        return False, "HTTP {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:300]), ""


def main():
    env = load_env()
    with open(HTML_FILE, encoding="utf-8") as f:
        content = f.read()

    print("タイトル: {}".format(TITLE))
    print("カテゴリ: {}".format(", ".join(CATEGORIES)))
    print("文字数: {} (HTML含む)".format(len(content)))
    print("モード: 下書き")
    print("投稿中...")

    xml = build_entry(TITLE, content, CATEGORIES, draft=True)
    ok, url, edit_url = post(env, xml)
    if ok:
        print("[成功]")
        print("  記事URL: {}".format(url))
        print("  編集URL: {}".format(edit_url))
    else:
        print("[失敗] {}".format(url))


if __name__ == "__main__":
    main()
