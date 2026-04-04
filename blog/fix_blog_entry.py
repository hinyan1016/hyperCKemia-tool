#!/usr/bin/env python3
"""
壊れたブログ記事をローカルHTMLから修復するスクリプト
AtomPub API の PUT で既存エントリを更新する
"""

import re
import xml.etree.ElementTree as ET
import urllib.request
import urllib.error
import base64
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\general-series\02_migraine_5_things_blog.html")

ENTRY_ID = "17179246901371583907"
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"

NS = {
    'atom': 'http://www.w3.org/2005/Atom',
    'app': 'http://www.w3.org/2007/app',
}


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def get_entry(env, entry_id):
    api_key = env["HATENA_API_KEY"]
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    auth_b64 = base64.b64encode("{}:{}".format(HATENA_ID, api_key).encode()).decode()

    req = urllib.request.Request(url, method="GET",
        headers={"Authorization": "Basic {}".format(auth_b64)})

    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def update_entry(env, entry_id, xml_body):
    api_key = env["HATENA_API_KEY"]
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    auth_b64 = base64.b64encode("{}:{}".format(HATENA_ID, api_key).encode()).decode()

    data = xml_body.encode("utf-8")
    req = urllib.request.Request(url, data=data, method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth_b64),
        })

    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def main():
    env = load_env()

    # 1. ローカルHTMLを読み込む
    html_content = HTML_FILE.read_text(encoding="utf-8")
    # <body>内のコンテンツのみ抽出
    body_match = re.search(r'<body>(.*)</body>', html_content, re.DOTALL)
    if body_match:
        content = body_match.group(1).strip()
    else:
        content = html_content
    print("ローカルHTML読み込み完了 ({}文字)".format(len(content)))

    # 2. 現在のエントリを取得してメタ情報を保持
    print("エントリ取得中...")
    entry_xml = get_entry(env, ENTRY_ID)

    # タイトル抽出
    title_match = re.search(r'<title[^>]*>([^<]+)</title>', entry_xml)
    title = title_match.group(1) if title_match else "タイトル不明"
    print("タイトル: {}".format(title))

    # カテゴリ抽出
    categories = re.findall(r'<category[^>]*term="([^"]+)"', entry_xml)
    print("カテゴリ: {}".format(", ".join(categories)))

    # 3. 新しいXMLを構築
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)

    new_xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    new_xml += '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">\n'
    new_xml += '  <title>{}</title>\n'.format(title)
    new_xml += '  <author><name>{}</name></author>\n'.format(HATENA_ID)
    new_xml += '  <content type="text/html"><![CDATA[{}]]></content>\n'.format(content)
    new_xml += '{}\n'.format(cat_xml)
    new_xml += '  <app:control>\n'
    new_xml += '    <app:draft>no</app:draft>\n'
    new_xml += '  </app:control>\n'
    new_xml += '</entry>'

    # 4. エントリを更新
    print("エントリ更新中...")
    result = update_entry(env, ENTRY_ID, new_xml)

    url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
    if url_match:
        print("更新成功: {}".format(url_match.group(1)))
    else:
        print("更新完了")


if __name__ == "__main__":
    main()
