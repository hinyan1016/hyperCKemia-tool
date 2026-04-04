#!/usr/bin/env python3
"""
はてなブログ全記事のmeta descriptionを直接チェック。
「youtu.be」が先頭付近にある記事を検出する。
APIからの本文チェックではなく、実際の公開ページのmeta descriptionを確認。
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


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def get_all_entry_urls(env, max_pages=30):
    """AtomPub APIで全記事のURL一覧を取得"""
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    auth_str = "{}:{}".format(hatena_id, env["HATENA_API_KEY"])
    auth_b64 = base64.b64encode(auth_str.encode()).decode()
    auth = "Basic {}".format(auth_b64)

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
            title = title_el.text if title_el is not None else ""
            link_alt = None
            link_edit = None
            for link in entry.findall("{%s}link" % ATOM_NS):
                if link.get("rel") == "alternate":
                    link_alt = link.get("href")
                if link.get("rel") == "edit":
                    link_edit = link.get("href")
            entries.append({
                "title": title,
                "url": link_alt,
                "edit_url": link_edit,
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


def check_meta_description(page_url):
    """公開ページのmeta descriptionを取得"""
    try:
        req = urllib.request.Request(page_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode('utf-8')
        m = re.search(r'<meta name="description" content="([^"]*)"', html)
        if m:
            return m.group(1)
    except Exception as e:
        return "ERROR: {}".format(str(e)[:50])
    return ""


def main():
    env = load_env()
    print("記事URL一覧を取得中...")
    entries = get_all_entry_urls(env)
    print("取得完了: {} 件\n".format(len(entries)))

    print("各記事のmeta descriptionをチェック中...")
    print("（youtu.be / youtube.com が先頭50文字以内にある記事を検出）\n")

    problems = []
    for i, entry in enumerate(entries):
        if not entry["url"]:
            continue
        desc = check_meta_description(entry["url"])
        # 先頭50文字以内にyoutu.beやyoutube.comがあるか
        head = desc[:50].lower() if desc else ""
        has_yt = "youtu.be" in head or "youtube.com" in head
        if has_yt:
            problems.append({
                "title": entry["title"],
                "url": entry["url"],
                "desc_head": desc[:100],
            })

        # 進捗表示（50件ごと）
        if (i + 1) % 50 == 0:
            print("  ... {}/{} 件チェック済み（問題: {} 件）".format(
                i + 1, len(entries), len(problems)))

        time.sleep(0.3)  # サーバー負荷対策

    print("\n=== 結果 ===\n")
    print("チェック: {} 件".format(len(entries)))
    print("問題あり: {} 件\n".format(len(problems)))

    if problems:
        for i, p in enumerate(problems, 1):
            print("{}. {}".format(i, p["title"]))
            print("   URL: {}".format(p["url"]))
            print("   DESC: {}".format(p["desc_head"]))
            print()

    # 結果をファイルに保存
    output_file = Path(__file__).parent / "meta_desc_problems.txt"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("meta description問題記事一覧（{}件）\n\n".format(len(problems)))
        for i, p in enumerate(problems, 1):
            f.write("{}. {}\n".format(i, p["title"]))
            f.write("   URL: {}\n".format(p["url"]))
            f.write("   DESC: {}\n\n".format(p["desc_head"]))
    print("結果を {} に保存しました".format(output_file))


if __name__ == "__main__":
    main()
