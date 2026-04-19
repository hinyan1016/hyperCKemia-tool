#!/usr/bin/env python3
"""
はてなブログ既存記事更新スクリプト
AtomPub APIで「若年ミオクロニー」記事を検索し、HTMLを更新する
"""

import os
import re
import base64
import urllib.request
import urllib.error
from pathlib import Path
from xml.etree import ElementTree as ET

# パス設定
ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\posts\jme_blog.html")

# カテゴリ
CATEGORIES = ["医学教育", "てんかん", "神経内科", "鑑別診断"]

# タイトル
TITLE = "若年ミオクロニーてんかん（JME）― 見逃されやすい思春期てんかんの診断と最新治療戦略"


def load_env():
    """.envファイルから設定を読み込む"""
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def make_auth_header(env):
    """Basic認証ヘッダーを作成"""
    auth_str = "{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"])
    auth_b64 = base64.b64encode(auth_str.encode()).decode()
    return "Basic {}".format(auth_b64)


def find_entry_edit_url(env):
    """エントリ一覧から「若年ミオクロニー」を含むエントリのedit URLを取得"""
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    base_url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
    auth = make_auth_header(env)

    url = base_url
    while url:
        req = urllib.request.Request(url, headers={"Authorization": auth})
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")

        # XML名前空間
        ns = {
            "atom": "http://www.w3.org/2005/Atom",
            "app": "http://www.w3.org/2007/app",
        }
        root = ET.fromstring(body)

        for entry in root.findall("atom:entry", ns):
            title_el = entry.find("atom:title", ns)
            if title_el is not None and "若年ミオクロニー" in (title_el.text or ""):
                # edit URLを取得
                for link in entry.findall("atom:link", ns):
                    if link.get("rel") == "edit":
                        edit_url = link.get("href")
                        print("記事発見: {}".format(title_el.text))
                        print("Edit URL: {}".format(edit_url))
                        return edit_url

        # 次ページ
        next_url = None
        for link in root.findall("atom:link", ns):
            if link.get("rel") == "next":
                next_url = link.get("href")
                break
        url = next_url

    return None


def build_update_xml(title, content, categories, draft=True):
    """更新用AtomPub XMLを構築"""
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join(
        '  <category term="{}" />'.format(c) for c in categories
    )
    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>{author}</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(
        title=title, author="hinyan1016", content=content,
        categories=cat_xml, draft=draft_val
    )
    return xml


def update_entry(env, edit_url, xml_body):
    """PUTリクエストで記事を更新"""
    auth = make_auth_header(env)
    data = xml_body.encode("utf-8")
    req = urllib.request.Request(
        edit_url,
        data=data,
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth,
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
    print("=" * 60)
    print("はてなブログ記事更新: JME")
    print("=" * 60)

    # 環境変数読み込み
    env = load_env()
    print("認証情報読み込み完了: HATENA_ID={}".format(env["HATENA_ID"]))

    # HTML読み込み
    with open(HTML_FILE, encoding="utf-8") as f:
        html_content = f.read()
    print("HTML読み込み完了: {} bytes".format(len(html_content)))

    # エントリ検索
    print("\nエントリ一覧を検索中...")
    edit_url = find_entry_edit_url(env)
    if not edit_url:
        print("[エラー] 「若年ミオクロニー」を含む記事が見つかりませんでした")
        return

    # 更新XML構築
    xml = build_update_xml(TITLE, html_content, CATEGORIES, draft=True)
    print("\n更新XMLを構築しました ({} bytes)".format(len(xml)))

    # 更新実行
    print("PUTリクエスト送信中...")
    ok, result = update_entry(env, edit_url, xml)

    if ok:
        print("\n[成功] 記事を更新しました")
        print("URL: {}".format(result))
    else:
        print("\n[失敗] {}".format(result))


if __name__ == "__main__":
    main()
