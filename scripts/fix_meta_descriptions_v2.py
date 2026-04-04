#!/usr/bin/env python3
"""
はてなブログ meta description 修正スクリプト v2
公開ページの meta description を直接チェックし、
「youtu.be」が先頭にある記事に導入文を追加する。

使い方:
  python fix_meta_descriptions_v2.py --scan        # 問題記事を検出
  python fix_meta_descriptions_v2.py --preview      # 修正プレビュー
  python fix_meta_descriptions_v2.py --fix          # 修正適用
  python fix_meta_descriptions_v2.py --fix-one URL  # 1件だけ修正（テスト用）
"""

import sys
import re
import urllib.request
import urllib.error
import base64
import time
import xml.etree.ElementTree as ET
from pathlib import Path

sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

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


def get_all_entries(env, max_pages=30):
    """全記事をAtomPub APIで取得（コンテンツ含む）"""
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
            entries.append({
                "title": title_el.text if title_el is not None else "",
                "content": content_el.text if content_el is not None else "",
                "url": link_alt,
                "edit_url": link_edit,
                "xml_entry": ET.tostring(entry, encoding="unicode"),
            })

        next_link = None
        for link in root.findall("{%s}link" % ATOM_NS):
            if link.get("rel") == "next":
                next_link = link.get("href")
        if not next_link:
            break
        url = next_link
        time.sleep(0.3)
    return entries


def get_meta_description(page_url):
    """公開ページの meta description を取得"""
    try:
        req = urllib.request.Request(page_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode('utf-8')
        m = re.search(r'<meta name="description" content="([^"]*)"', html)
        return m.group(1) if m else ""
    except Exception:
        return ""


def desc_starts_with_youtube(desc):
    """meta descriptionがyoutu.beで始まるか"""
    d = desc.strip().lower()
    return d.startswith("youtu.be") or d.startswith("https://youtu.be") or d.startswith("http://youtu.be")


def find_first_youtube_in_content(content):
    """
    コンテンツ内の最初のYouTube要素（iframe or リンク）の位置を返す。
    YouTube要素を含む最初の<p>タグの開始位置を返す。
    """
    if not content:
        return -1

    # パターン1: <p><iframe ...youtube...> or <p> <iframe ...youtube...>
    m = re.search(r'<p>\s*<iframe[^>]*youtube\.com', content, re.IGNORECASE)
    if m:
        return m.start()

    # パターン2: <p><a ...youtu.be...>
    m = re.search(r'<p>\s*<a\s[^>]*href="https?://youtu\.be', content, re.IGNORECASE)
    if m:
        return m.start()

    # パターン3: <p>youtu.be/... (テキストリンク)
    m = re.search(r'<p>\s*youtu\.be/', content, re.IGNORECASE)
    if m:
        return m.start()

    # パターン4: <p>https://youtu.be/...
    m = re.search(r'<p>\s*https?://youtu\.be/', content, re.IGNORECASE)
    if m:
        return m.start()

    # パターン5: <iframe ...youtube...> (pタグなし)
    m = re.search(r'<iframe[^>]*youtube\.com', content, re.IGNORECASE)
    if m:
        return m.start()

    # パターン6: youtu.be テキスト直接
    m = re.search(r'youtu\.be/[A-Za-z0-9_-]', content, re.IGNORECASE)
    if m:
        return m.start()

    return -1


def add_intro_before_youtube(content, title):
    """YouTube要素の前に導入文を挿入"""
    if not content:
        return content

    intro = '<p style="font-size:1.1em;margin-bottom:1em;"><strong>{}</strong></p>\n'.format(
        title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )

    pos = find_first_youtube_in_content(content)
    if pos >= 0:
        return content[:pos] + intro + content[pos:]

    # YouTube要素が見つからない場合、コンテンツ先頭に追加
    # ただし<div>タグの直後に挿入する
    m = re.search(r'((?:<div[^>]*>\s*)+)', content, re.IGNORECASE)
    if m:
        end = m.end()
        return content[:end] + intro + content[end:]

    return intro + content


def update_entry(env, edit_url, title, new_content, original_xml):
    """AtomPub APIで記事を更新"""
    auth = make_auth_header(env)

    root = ET.fromstring(original_xml)
    categories = []
    for cat in root.findall("{%s}category" % ATOM_NS):
        term = cat.get("term")
        if term:
            categories.append(term)

    draft_el = root.find(".//{%s}draft" % APP_NS)
    draft_val = draft_el.text if draft_el is not None else "no"

    cat_xml = "\n".join('  <category term="{}" />'.format(
        c.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")
    ) for c in categories)

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
        edit_url, data=data, method="PUT",
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

    mode = args[0]

    env = load_env()

    # --fix-one: 1件だけテスト
    if mode == "--fix-one" and len(args) > 1:
        target_url = args[1]
        print("記事を取得中...")
        entries = get_all_entries(env)
        print("取得完了: {} 件\n".format(len(entries)))
        for entry in entries:
            if entry["url"] == target_url:
                desc = get_meta_description(entry["url"])
                if not desc_starts_with_youtube(desc):
                    print("この記事はyoutu.be問題なし: {}".format(desc[:80]))
                    return
                print("修正中: {}".format(entry["title"]))
                new_content = add_intro_before_youtube(entry["content"], entry["title"])
                ok, msg = update_entry(env, entry["edit_url"], entry["title"],
                                       new_content, entry["xml_entry"])
                print("[成功]" if ok else "[失敗] {}".format(msg))
                if ok:
                    time.sleep(2)
                    new_desc = get_meta_description(entry["url"])
                    print("修正後DESC: {}".format(new_desc[:100]))
                return
        print("記事が見つかりません: {}".format(target_url))
        return

    print("記事を取得中...")
    entries = get_all_entries(env)
    print("取得完了: {} 件".format(len(entries)))

    # 公開ページのmeta descriptionをチェック
    print("meta descriptionをチェック中...\n")
    problems = []
    for i, entry in enumerate(entries):
        if not entry["url"]:
            continue
        desc = get_meta_description(entry["url"])
        if desc_starts_with_youtube(desc):
            problems.append({
                "entry": entry,
                "desc": desc,
            })
        if (i + 1) % 50 == 0:
            print("  ... {}/{} 件チェック済み（問題: {} 件）".format(
                i + 1, len(entries), len(problems)))
        time.sleep(0.2)

    print("\n問題記事: {} 件\n".format(len(problems)))

    if mode == "--scan":
        for i, p in enumerate(problems, 1):
            print("{}. {}".format(i, p["entry"]["title"]))
            print("   URL: {}".format(p["entry"]["url"]))
            print("   DESC: {}".format(p["desc"][:100]))
            print()

    elif mode == "--preview":
        for i, p in enumerate(problems, 1):
            entry = p["entry"]
            content = entry["content"] or ""
            new_content = add_intro_before_youtube(content, entry["title"])
            # 挿入位置の前後を表示
            pos = find_first_youtube_in_content(content)
            print("{}. {}".format(i, entry["title"]))
            if pos >= 0:
                before = content[max(0,pos-30):pos+50].replace("\n"," ")
                print("   挿入位置付近: ...{}...".format(before))
            else:
                print("   YouTube要素未検出 → 先頭に挿入")
            print()

    elif mode == "--fix":
        print("=== 修正適用 ===\n")
        success = 0
        for i, p in enumerate(problems, 1):
            entry = p["entry"]
            content = entry["content"] or ""
            new_content = add_intro_before_youtube(content, entry["title"])

            if new_content == content:
                print("{}. {} ... [変更なし（スキップ）]".format(i, entry["title"]))
                continue

            print("{}. {} ...".format(i, entry["title"]))
            ok, msg = update_entry(env, entry["edit_url"], entry["title"],
                                   new_content, entry["xml_entry"])
            if ok:
                print("   [成功]")
                success += 1
            else:
                print("   [失敗] {}".format(msg))
            time.sleep(1)

        print("\n完了: {}/{} 件修正".format(success, len(problems)))


if __name__ == "__main__":
    main()
