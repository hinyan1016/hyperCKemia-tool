#!/usr/bin/env python3
"""
低活動性せん妄ブログ記事を更新（PUT）し、重複記事を削除する
"""
import os
import re
import urllib.request
import urllib.error
import base64
import xml.etree.ElementTree as ET
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HATENA_HTML = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\hypoactive_delirium_blog_hatena.html")
TARGET_URL = "https://hinyan1016.hatenablog.com/entry/2026/04/09/160554"
DUPLICATE_URL = "https://hinyan1016.hatenablog.com/entry/2026/04/09/161339"
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


def get_auth_header(env):
    auth_str = "{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"])
    return "Basic " + base64.b64encode(auth_str.encode()).decode()


def find_entry_id(env, target_url):
    """AtomPub APIからエントリ一覧を取得し、URLでエントリIDを特定"""
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)

    req = urllib.request.Request(api_url, headers={
        "Authorization": get_auth_header(env),
    })
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")

    # Parse XML to find entry with matching alternate link
    ns = {"atom": "http://www.w3.org/2005/Atom", "app": "http://www.w3.org/2007/app"}
    root = ET.fromstring(body)
    for entry in root.findall("atom:entry", ns):
        for link in entry.findall("atom:link", ns):
            if link.get("rel") == "alternate" and link.get("href") == target_url:
                # Get edit link
                for elink in entry.findall("atom:link", ns):
                    if elink.get("rel") == "edit":
                        return elink.get("href")
    return None


def update_entry(env, edit_url, title, content, categories, draft=True):
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)
    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(title=title, content=content, categories=cat_xml, draft=draft_val)

    data = xml.encode("utf-8")
    req = urllib.request.Request(edit_url, data=data, method="PUT", headers={
        "Content-Type": "application/xml; charset=utf-8",
        "Authorization": get_auth_header(env),
    })
    try:
        with urllib.request.urlopen(req) as resp:
            print("  更新成功: {}".format(resp.status))
            return True
    except urllib.error.HTTPError as e:
        print("  更新失敗: {} {}".format(e.code, e.read().decode("utf-8", errors="replace")[:300]))
        return False


def delete_entry(env, edit_url):
    req = urllib.request.Request(edit_url, method="DELETE", headers={
        "Authorization": get_auth_header(env),
    })
    try:
        with urllib.request.urlopen(req) as resp:
            print("  削除成功: {}".format(resp.status))
            return True
    except urllib.error.HTTPError as e:
        print("  削除失敗: {} {}".format(e.code, e.read().decode("utf-8", errors="replace")[:300]))
        return False


def main():
    env = load_env()

    # Read updated HTML
    with open(HATENA_HTML, encoding="utf-8") as f:
        content = f.read()

    # 1. Find and update original entry
    print("[1/3] 元記事のエントリIDを検索中...")
    edit_url = find_entry_id(env, TARGET_URL)
    if edit_url:
        print("  見つかりました: {}".format(edit_url))
        print("[2/3] 元記事を更新中...")
        update_entry(env, edit_url, TITLE, content, CATEGORIES, draft=True)
    else:
        print("  元記事が見つかりません")

    # 2. Find and delete duplicate
    print("[3/3] 重複記事を削除中...")
    dup_url = find_entry_id(env, DUPLICATE_URL)
    if dup_url:
        print("  重複記事を検出: {}".format(dup_url))
        delete_entry(env, dup_url)
    else:
        print("  重複記事は見つかりませんでした")

    print("完了!")


if __name__ == "__main__":
    main()
