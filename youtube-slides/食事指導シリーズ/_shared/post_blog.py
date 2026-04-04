#!/usr/bin/env python3
"""
食事指導シリーズ ブログ一括投稿スクリプト
はてなブログ AtomPub API を使用して下書き投稿する

使い方:
  python post_blog.py                  # 未投稿の記事を全て下書き投稿
  python post_blog.py 03               # 03_脂質異常症 のみ投稿
  python post_blog.py 03 04 05         # 複数指定
  python post_blog.py --publish 03     # 下書きではなく公開投稿
  python post_blog.py --list           # 投稿状況を確認
"""

import os
import sys
import re
import json
import urllib.request
import urllib.error
import base64
from pathlib import Path
from html.parser import HTMLParser

# パス設定
SCRIPT_DIR = Path(__file__).parent
SERIES_DIR = SCRIPT_DIR.parent
ENV_FILE = SCRIPT_DIR / ".env"
STATUS_FILE = SCRIPT_DIR / "post_status.json"

# トピックフォルダのマッピング
TOPICS = {
    "01": "01_総論",
    "02": "02_高血圧",
    "03": "03_脂質異常症",
    "04": "04_糖尿病",
    "05": "05_肥満",
    "06": "06_CKD",
    "07": "07_肝疾患",
    "08": "08_嚥下障害",
    "09": "09_術後消化管",
    "10": "10_フレイル",
}

# カテゴリタグ
CATEGORIES = ["食事指導シリーズ", "医学教育", "栄養"]


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


def load_status():
    """投稿状況を読み込む"""
    if STATUS_FILE.exists():
        with open(STATUS_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_status(status):
    """投稿状況を保存する"""
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(status, f, ensure_ascii=False, indent=2)


def find_blog_html(topic_num):
    """トピック番号からブログHTMLファイルを探す"""
    folder_name = TOPICS.get(topic_num)
    if not folder_name:
        return None, None
    blog_dir = SERIES_DIR / folder_name / "blog"
    if not blog_dir.exists():
        return folder_name, None
    html_files = list(blog_dir.glob("*.html"))
    if not html_files:
        return folder_name, None
    return folder_name, html_files[0]


def extract_title(html_content):
    """HTMLからh1タイトルを抽出する"""
    match = re.search(r"<h1>(.*?)</h1>", html_content)
    if match:
        return match.group(1)
    return None


def fix_youtube_first(content, title):
    """YouTube埋め込みが本文冒頭にある場合、導入文を前に挿入する（meta description対策）"""
    if not content:
        return content
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
    m = re.search(r'(<p>\s*<iframe[^>]*youtube\.com)', content, re.IGNORECASE)
    if m:
        pos = m.start()
        print("  [自動修正] YouTube埋め込みの前に導入文を挿入（meta description対策）")
        return content[:pos] + intro + content[pos:]
    m = re.search(r'(<p>\s*<a\s[^>]*href="https?://youtu\.be)', content, re.IGNORECASE)
    if m:
        pos = m.start()
        print("  [自動修正] YouTubeリンクの前に導入文を挿入（meta description対策）")
        return content[:pos] + intro + content[pos:]
    return content


def build_atom_entry(title, content, categories, draft=True):
    """AtomPub形式のXMLエントリを構築する"""
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
    """はてなブログ AtomPub APIに投稿する"""
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
            # レスポンスからURLを抽出
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            entry_url = url_match.group(1) if url_match else "(URL取得失敗)"
            return True, entry_url
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        return False, "HTTP {}: {}".format(e.code, error_body[:200])


def post_topic(topic_num, env, draft=True):
    """指定トピックを投稿する"""
    folder_name, html_path = find_blog_html(topic_num)
    if not folder_name:
        print("  [スキップ] {} は不明なトピック番号です".format(topic_num))
        return False
    if not html_path:
        print("  [スキップ] {} にブログHTMLがありません".format(folder_name))
        return False

    with open(html_path, encoding="utf-8") as f:
        content = f.read()

    title = extract_title(content)

    # YouTube冒頭の自動修正
    content = fix_youtube_first(content, title) if title else content
    if not title:
        print("  [エラー] {} のタイトル(h1)が見つかりません".format(folder_name))
        return False

    mode = "公開" if not draft else "下書き"
    print("  投稿中: {} ({})...".format(title, mode))

    xml = build_atom_entry(title, content, CATEGORIES, draft=draft)
    ok, result = post_entry(env, xml)

    if ok:
        print("  [成功] {}".format(result))
        # ステータス更新
        status = load_status()
        status[topic_num] = {
            "title": title,
            "url": result,
            "draft": draft,
            "folder": folder_name,
        }
        save_status(status)
        return True
    else:
        print("  [失敗] {}".format(result))
        return False


def show_list():
    """投稿状況を一覧表示する"""
    status = load_status()
    print("\n食事指導シリーズ 投稿状況")
    print("=" * 60)
    for num in sorted(TOPICS.keys()):
        folder = TOPICS[num]
        if num in status:
            s = status[num]
            state = "下書き" if s.get("draft") else "公開済"
            print("  {} [{}] {} → {}".format(num, state, s["title"], s["url"]))
        else:
            _, html_path = find_blog_html(num)
            has_html = "HTML有" if html_path else "HTML無"
            print("  {} [未投稿] {} ({})".format(num, folder, has_html))
    print()


def main():
    args = sys.argv[1:]

    if "--list" in args:
        show_list()
        return

    publish = "--publish" in args
    args = [a for a in args if not a.startswith("--")]

    env = load_env()

    # 投稿対象を決定
    if args:
        targets = [a.zfill(2) for a in args]
    else:
        # 未投稿のものを全て対象にする
        status = load_status()
        targets = [num for num in sorted(TOPICS.keys()) if num not in status]
        if not targets:
            print("全記事が投稿済みです。--list で状況を確認できます。")
            return

    print("\n食事指導シリーズ ブログ投稿")
    print("=" * 60)
    print("モード: {}".format("公開" if publish else "下書き"))
    print("対象: {} 件\n".format(len(targets)))

    success = 0
    for num in targets:
        if post_topic(num, env, draft=not publish):
            success += 1
        print()

    print("完了: {}/{} 件成功".format(success, len(targets)))


if __name__ == "__main__":
    main()
