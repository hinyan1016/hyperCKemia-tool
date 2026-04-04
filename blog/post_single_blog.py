#!/usr/bin/env python3
"""
はてなブログ 単発記事投稿スクリプト
AtomPub API を使用して下書き投稿する

使い方:
  python post_single_blog.py <HTMLファイル> <タイトル> [--publish] [--categories cat1,cat2]
"""

import os
import sys
import re
import urllib.request
import urllib.error
import base64
from pathlib import Path
from html.parser import HTMLParser

# .env は食事指導シリーズの共有フォルダから読み込む
ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def fix_youtube_first(content, title):
    """YouTube埋め込みが本文冒頭にある場合、導入文を前に挿入する（meta description対策）"""
    if not content:
        return content
    # divタグを剥がして最初の実質コンテンツを確認
    inner = re.sub(r'^(\s*<div[^>]*>\s*)+', '', content.strip(), flags=re.IGNORECASE)
    patterns = [
        r'^\s*<p>\s*<iframe[^>]*youtube\.com',
        r'^\s*<p>\s*<a\s[^>]*href="https?://(youtu\.be|www\.youtube\.com)',
        r'^\s*<iframe[^>]*youtube\.com',
        r'^\s*<p>\s*youtu\.be',
    ]
    is_yt_first = any(re.match(pat, inner, re.IGNORECASE) for pat in patterns)
    if not is_yt_first:
        return content

    intro = '<p style="font-size:1.1em;margin-bottom:1em;"><strong>{}</strong></p>\n'.format(
        title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )
    # YouTube iframeを含む<p>の前に挿入
    m = re.search(r'(<p>\s*<iframe[^>]*youtube\.com)', content, re.IGNORECASE)
    if m:
        pos = m.start()
        print("[自動修正] YouTube埋め込みの前に導入文を挿入しました（meta description対策）")
        return content[:pos] + intro + content[pos:]
    # <p><a ...youtu.be の前に挿入
    m = re.search(r'(<p>\s*<a\s[^>]*href="https?://youtu\.be)', content, re.IGNORECASE)
    if m:
        pos = m.start()
        print("[自動修正] YouTubeリンクの前に導入文を挿入しました（meta description対策）")
        return content[:pos] + intro + content[pos:]
    return content


def build_atom_entry(title, content, categories, draft=True):
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join(
        '  <category term="{}" />'.format(c) for c in categories
    )
    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(
        title=title, content=content, categories=cat_xml, draft=draft_val
    )
    return xml


def post_entry(env, xml_body):
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
    auth_str = "{}:{}".format(hatena_id, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    data = xml_body.encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        method="POST",
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
    if len(sys.argv) < 3:
        print("使い方: python post_single_blog.py <HTMLファイル> <タイトル> [--publish] [--categories cat1,cat2]")
        sys.exit(1)

    html_file = sys.argv[1]
    title = sys.argv[2]
    publish = "--publish" in sys.argv

    # カテゴリ
    categories = ["医学教育", "ビタミンB12"]
    for arg in sys.argv[3:]:
        if arg.startswith("--categories"):
            idx = sys.argv.index(arg)
            if idx + 1 < len(sys.argv):
                categories = sys.argv[idx + 1].split(",")

    # HTML読み込み
    with open(html_file, encoding="utf-8") as f:
        content = f.read()

    # YouTube冒頭の自動修正
    content = fix_youtube_first(content, title)

    mode = "公開" if publish else "下書き"
    print("タイトル: {}".format(title))
    print("モード: {}".format(mode))
    print("カテゴリ: {}".format(", ".join(categories)))
    print("投稿中...")

    env = load_env()
    xml = build_atom_entry(title, content, categories, draft=not publish)
    ok, result = post_entry(env, xml)

    if ok:
        print("[成功] {}".format(result))
    else:
        print("[失敗] {}".format(result))
        sys.exit(1)


if __name__ == "__main__":
    main()
