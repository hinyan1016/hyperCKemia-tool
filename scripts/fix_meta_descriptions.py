#!/usr/bin/env python3
"""
はてなブログ meta description 修正スクリプト
記事冒頭の YouTube iframe/リンクの前に導入文を追加して、
meta description が「youtu.be ...」にならないようにする。

使い方:
  python fix_meta_descriptions.py --scan          # 対象記事をスキャン
  python fix_meta_descriptions.py --preview       # 修正プレビュー（変更なし）
  python fix_meta_descriptions.py --fix           # 実際に修正を適用
  python fix_meta_descriptions.py --fix --entry URL  # 特定記事のみ修正
"""

import os
import sys
import re
import urllib.request
import urllib.error
import base64
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from io import StringIO

# stdout を UTF-8 に強制
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

# .env パス
ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")

ATOM_NS = "http://www.w3.org/2005/Atom"
APP_NS = "http://www.w3.org/2007/app"


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def make_auth_header(env):
    auth_str = "{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"])
    auth_b64 = base64.b64encode(auth_str.encode()).decode()
    return "Basic {}".format(auth_b64)


def get_entries(env, max_pages=30):
    """全記事をAtomPub APIで取得"""
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    auth = make_auth_header(env)

    entries = []
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)

    for page in range(max_pages):
        req = urllib.request.Request(url, headers={"Authorization": auth})
        try:
            with urllib.request.urlopen(req) as resp:
                body = resp.read().decode("utf-8")
        except urllib.error.HTTPError as e:
            print("  API エラー: HTTP {}".format(e.code))
            break

        root = ET.fromstring(body)

        for entry in root.findall("{%s}entry" % ATOM_NS):
            title_el = entry.find("{%s}title" % ATOM_NS)
            content_el = entry.find("{%s}content" % ATOM_NS)
            link_alt = None
            link_edit = None
            for link in entry.findall("{%s}link" % ATOM_NS):
                if link.get("rel") == "alternate":
                    link_alt = link.get("href")
                if link.get("rel") == "edit":
                    link_edit = link.get("href")

            title = title_el.text if title_el is not None else ""
            content = content_el.text if content_el is not None else ""

            entries.append({
                "title": title,
                "content": content,
                "url": link_alt,
                "edit_url": link_edit,
                "xml_entry": ET.tostring(entry, encoding="unicode"),
            })

        # 次ページリンク
        next_link = None
        for link in root.findall("{%s}link" % ATOM_NS):
            if link.get("rel") == "next":
                next_link = link.get("href")
        if not next_link:
            break
        url = next_link
        time.sleep(0.5)

    return entries


def find_first_text_content(html):
    """HTMLの最初のテキストコンテンツ（タグを除いた）を取得"""
    if not html:
        return ""
    # タグを除去してテキストを取得
    text = re.sub(r'<[^>]+>', ' ', html)
    text = re.sub(r'\s+', ' ', text).strip()
    return text[:100]


def content_starts_with_youtube(content):
    """
    記事本文の最初のテキスト/埋め込みがYouTubeかチェック。
    実際のパターン: <div class="medical-article"> ... <p><iframe ...youtube...>
    または: 冒頭テキストが youtu.be で始まる
    """
    if not content:
        return False

    stripped = content.strip()

    # divタグなどを剥がして最初の実質コンテンツを見つける
    # <div ...> タグを除去して最初の <p> や <iframe> を見る
    inner = re.sub(r'^(\s*<div[^>]*>\s*)+', '', stripped, flags=re.IGNORECASE)

    patterns = [
        # <p><iframe ...youtube...>
        r'^\s*<p>\s*<iframe[^>]*youtube\.com',
        # <p><a href="https://youtu.be/...">
        r'^\s*<p>\s*<a\s[^>]*href="https?://(youtu\.be|www\.youtube\.com)',
        # <iframe ...youtube...> 直接
        r'^\s*<iframe[^>]*youtube\.com',
        # テキストで youtu.be
        r'^\s*<p>\s*youtu\.be',
        r'^\s*youtu\.be',
        r'^\s*https?://youtu\.be',
    ]
    for pat in patterns:
        if re.match(pat, inner, re.IGNORECASE):
            return True

    # フォールバック: 最初のテキストがyoutu.beで始まるか
    first_text = find_first_text_content(inner[:500])
    if first_text.strip().startswith("youtu.be") or first_text.strip().startswith("https://youtu.be"):
        return True

    return False


def already_has_intro(content):
    """既に導入文が追加済みかチェック"""
    if not content:
        return False
    stripped = content.strip()
    inner = re.sub(r'^(\s*<div[^>]*>\s*)+', '', stripped, flags=re.IGNORECASE)
    # 最初の要素が<p><strong>で始まる場合、既に導入文がある可能性
    if re.match(r'^\s*<p[^>]*>\s*<strong>', inner, re.IGNORECASE):
        return True
    return False


def extract_youtube_id(content):
    """冒頭のYouTube動画IDを抽出"""
    match = re.search(r'youtube\.com/embed/([A-Za-z0-9_-]+)', content)
    if match:
        return match.group(1)
    match = re.search(r'youtu\.be/([A-Za-z0-9_-]+)', content)
    if match:
        return match.group(1)
    return None


def add_intro_to_content(content, title):
    """
    記事本文の最初のYouTube要素の前に導入文を挿入。
    <div class="medical-article"> の直後に挿入する。
    """
    if not content:
        return content

    # <div class="medical-article"> の直後に導入文を挿入
    # パターン: <div class="medical-article"> <div class="medical-article"> <p><iframe...
    # 最も内側の<div>の直後、<p><iframe>の前に挿入

    intro = '<p style="font-size:1.1em;margin-bottom:1em;"><strong>{}</strong></p>\n'.format(title)

    # 最初のYouTube iframe/リンクを含む<p>タグの前に挿入
    # パターン1: <p><iframe ...youtube...
    match = re.search(r'(<p>\s*<iframe[^>]*youtube\.com)', content, re.IGNORECASE)
    if match:
        pos = match.start()
        return content[:pos] + intro + content[pos:]

    # パターン2: <p><a ...youtu.be...
    match = re.search(r'(<p>\s*<a\s[^>]*href="https?://youtu\.be)', content, re.IGNORECASE)
    if match:
        pos = match.start()
        return content[:pos] + intro + content[pos:]

    # フォールバック: 最初の<p>の前に挿入
    match = re.search(r'(<p[^>]*>)', content, re.IGNORECASE)
    if match:
        pos = match.start()
        return content[:pos] + intro + content[pos:]

    # それでもダメなら先頭に追加
    return intro + content


def update_entry(env, edit_url, title, new_content, original_xml):
    """AtomPub APIで記事を更新"""
    auth = make_auth_header(env)

    # 元のXMLからカテゴリなどを保持
    root = ET.fromstring(original_xml)
    categories = []
    for cat in root.findall("{%s}category" % ATOM_NS):
        term = cat.get("term")
        if term:
            categories.append(term)

    # draft状態を保持
    draft_el = root.find(".//{%s}draft" % APP_NS)
    draft_val = draft_el.text if draft_el is not None else "no"

    # カテゴリのXML（特殊文字をエスケープ）
    cat_xml = "\n".join('  <category term="{}" />'.format(
        c.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")
    ) for c in categories)

    # タイトルの特殊文字もエスケープ
    safe_title = title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

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
        title=safe_title, content=new_content, categories=cat_xml, draft=draft_val
    )

    data = xml.encode("utf-8")
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
            return True, "OK"
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        return False, "HTTP {}: {}".format(e.code, error_body[:300])


def main():
    args = sys.argv[1:]
    if not args or "--help" in args:
        print(__doc__)
        return

    mode = args[0]  # --scan, --preview, --fix, --debug
    target_url = None
    if "--entry" in args:
        idx = args.index("--entry")
        if idx + 1 < len(args):
            target_url = args[idx + 1]

    env = load_env()
    print("記事を取得中...")
    entries = get_entries(env)
    print("取得完了: {} 件\n".format(len(entries)))

    if mode == "--debug":
        print("=== 全記事の冒頭を表示（最初の15件） ===\n")
        for i, entry in enumerate(entries[:15], 1):
            content = entry["content"] or ""
            stripped = content.strip()
            # divを剥がして内容を見る
            inner = re.sub(r'^(\s*<div[^>]*>\s*)+', '', stripped, flags=re.IGNORECASE)
            preview = inner[:200].replace("\n", " ")
            is_yt = content_starts_with_youtube(content)
            marker = " [YouTube冒頭]" if is_yt else ""
            print("{}. {}{}".format(i, entry["title"], marker))
            print("   冒頭: {}".format(preview))
            print()
        return

    # 対象記事を抽出
    affected = []
    for entry in entries:
        if content_starts_with_youtube(entry["content"]) and not already_has_intro(entry["content"]):
            affected.append(entry)

    if mode == "--scan":
        print("=== YouTube冒頭の記事: {} 件 ===\n".format(len(affected)))
        for i, entry in enumerate(affected, 1):
            yt_id = extract_youtube_id(entry["content"] or "")
            print("{}. {}".format(i, entry["title"]))
            print("   URL: {}".format(entry["url"]))
            if yt_id:
                print("   YouTube: https://youtube.com/watch?v={}".format(yt_id))
            print()
        print("修正対象: {} 件".format(len(affected)))
        print("次のステップ: python fix_meta_descriptions.py --preview")

    elif mode == "--preview":
        print("=== 修正プレビュー ===\n")
        for i, entry in enumerate(affected, 1):
            if target_url and entry["url"] != target_url:
                continue
            content = entry["content"] or ""
            new_content = add_intro_to_content(content, entry["title"])
            # 変更部分を表示
            inner_old = re.sub(r'^(\s*<div[^>]*>\s*)+', '', content.strip(), flags=re.IGNORECASE)
            inner_new = re.sub(r'^(\s*<div[^>]*>\s*)+', '', new_content.strip(), flags=re.IGNORECASE)
            print("{}. {}".format(i, entry["title"]))
            print("   変更前: {}".format(inner_old[:100].replace("\n", " ")))
            print("   変更後: {}".format(inner_new[:160].replace("\n", " ")))
            print()

    elif mode == "--fix":
        print("=== 修正適用 ===\n")
        success = 0
        total = 0
        for i, entry in enumerate(affected, 1):
            if target_url and entry["url"] != target_url:
                continue
            total += 1
            content = entry["content"] or ""
            new_content = add_intro_to_content(content, entry["title"])
            print("{}. {} ...".format(i, entry["title"]))
            ok, msg = update_entry(
                env, entry["edit_url"], entry["title"],
                new_content, entry["xml_entry"]
            )
            if ok:
                print("   [成功]")
                success += 1
            else:
                print("   [失敗] {}".format(msg))
            time.sleep(1)  # API制限対策
        print("\n完了: {}/{} 件修正".format(success, total))

    else:
        print("不明なオプション: {}".format(mode))
        print("--scan / --preview / --fix / --debug を指定してください")


if __name__ == "__main__":
    main()
