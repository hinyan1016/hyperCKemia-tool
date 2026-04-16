#!/usr/bin/env python3
"""
はてなブログ記事にYouTube動画リンクを追加するスクリプト
AtomPub API を使用して既存記事を更新する
"""

import re
import urllib.request
import urllib.error
import base64
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")

ENTRY_ID = "17179246901371583907"
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"


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
    auth_str = "{}:{}".format(HATENA_ID, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    req = urllib.request.Request(
        url,
        method="GET",
        headers={
            "Authorization": "Basic {}".format(auth_b64),
        },
    )

    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def update_entry(env, entry_id, xml_body):
    api_key = env["HATENA_API_KEY"]
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    auth_str = "{}:{}".format(HATENA_ID, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    data = xml_body.encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth_b64),
        },
    )

    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def main():
    env = load_env()

    # 1. 現在のエントリを取得
    print("エントリ取得中...")
    entry_xml = get_entry(env, ENTRY_ID)

    # 2. コンテンツ部分を抽出
    content_match = re.search(r'<content[^>]*>(.*?)</content>', entry_xml, re.DOTALL)
    if not content_match:
        print("ERROR: contentが見つかりません")
        return

    content = content_match.group(1)
    # CDATA除去
    if content.startswith("<![CDATA["):
        content = content[9:]
    if content.endswith("]]>"):
        content = content[:-3]

    # 3. YouTube動画リンクが既にあるか確認
    if "youtu.be/UJrNkbENiVw" in content or "youtube.com/watch?v=UJrNkbENiVw" in content:
        print("YouTube動画リンクは既に存在します。スキップ。")
        return

    # 4. 関連リンクセクションの前にYouTubeリンクを挿入
    youtube_block = (
        '<div class="important-box">\n'
        '<strong>&#9654; この記事の動画版はこちら</strong><br />\n'
        '<a href="https://youtu.be/UJrNkbENiVw" target="_blank" rel="noopener">'
        '片頭痛持ちが知っておくべき5つのこと｜脳神経内科医が解説【からだの不思議 #02】（YouTube）</a>\n'
        '</div>\n\n'
    )

    target = '<!-- 関連リンク -->'
    if target in content:
        new_content = content.replace(target, youtube_block + target)
        print("YouTube動画リンクを関連リンクセクションの前に挿入しました")
    else:
        # 末尾に追加
        new_content = content + '\n' + youtube_block
        print("YouTube動画リンクを末尾に追加しました")

    # 5. XMLを更新
    new_xml = entry_xml.replace(content_match.group(0),
        '<content type="text/html"><![CDATA[{}]]></content>'.format(new_content))

    # 6. エントリを更新
    print("エントリ更新中...")
    result = update_entry(env, ENTRY_ID, new_xml)

    url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
    if url_match:
        print("更新成功: {}".format(url_match.group(1)))
    else:
        print("更新完了（URL取得失敗）")


if __name__ == "__main__":
    main()
