#!/usr/bin/env python3
"""
その症状大丈夫#05 寝ている時に足がつる
- ローカル 05_leg_cramps_hatena.html (YouTube埋め込み挿入済み) を AtomPub PUT
- ★ 安全パターン: PUT前に本番GETして「インフォグラ/画像/ライブのみ要素」を検出
- draft状態を保持（公開はユーザーに委ねる）
"""

import re
import sys
import base64
import urllib.request
import urllib.error
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"

HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\symptom-concern-series\05_leg_cramps_hatena.html")
TITLE_KEYWORD = "寝ている時に足がつる"
TARGET_URL_FRAGMENT = "2026/04/25/081330"
VIDEO_ID = "KYNd5Z6e8sU"
PRESERVE_DRAFT_STATE = True  # 本番のdraft状態を維持
FORCE_PROCEED = False  # 警告があってもCIで通したい時にTrue


def auth_header():
    b64 = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()
    return "Basic " + b64


def find_entry_id():
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(HATENA_ID, BLOG_DOMAIN)
    page = 0
    while url and page < 30:
        req = urllib.request.Request(url, headers={"Authorization": auth_header()})
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
        entries = re.findall(r"<entry>(.*?)</entry>", body, re.DOTALL)
        for e in entries:
            title_m = re.search(r"<title[^>]*>([^<]*)</title>", e)
            if not title_m or TITLE_KEYWORD not in title_m.group(1):
                continue
            alt_m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', e)
            alt_url = alt_m.group(1) if alt_m else ""
            if TARGET_URL_FRAGMENT and TARGET_URL_FRAGMENT not in alt_url:
                continue
            id_m = re.search(r"<id>tag:blog\.hatena\.ne\.jp[^:]*:[^:]*-(\d+)</id>", e)
            if id_m:
                return id_m.group(1), title_m.group(1), alt_url
            id_m = re.search(r"<id>([^<]+)</id>", e)
            if id_m:
                return id_m.group(1).rsplit("-", 1)[-1], title_m.group(1), alt_url
        next_m = re.search(r'<link rel="next" href="([^"]+)"', body)
        url = next_m.group(1) if next_m else None
        page += 1
    raise RuntimeError("対象記事が見つかりません")


def get_entry(entry_id):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, headers={"Authorization": auth_header()})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def extract_content(entry_xml):
    m = re.search(r"<content[^>]*>(.*?)</content>", entry_xml, re.DOTALL)
    if not m:
        return ""
    body = m.group(1)
    cdata = re.match(r"\s*<!\[CDATA\[(.*?)\]\]>\s*", body, re.DOTALL)
    return cdata.group(1) if cdata else body


def safety_diff(live_html, local_html):
    """本番にあってlocalに無い要素を検出。空配列なら安全。"""
    warnings = []
    if "<!-- infographic -->" in live_html and "<!-- infographic -->" not in local_html:
        warnings.append("[CRITICAL] インフォグラフィックブロック (<!-- infographic -->) が本番にありlocalに無い")
    live_imgs = set(re.findall(r'src="([^"]+fotolife[^"]+)"', live_html))
    local_imgs = set(re.findall(r'src="([^"]+fotolife[^"]+)"', local_html))
    for img in live_imgs - local_imgs:
        warnings.append("[CRITICAL] 本番Fotolife画像がlocalに無い: " + img)
    live_iframes = set(re.findall(r'src="(https?://[^"]+)"', live_html))
    live_iframes = {u for u in live_iframes if "youtube" not in u and "fotolife" not in u}
    local_iframes = set(re.findall(r'src="(https?://[^"]+)"', local_html))
    local_iframes = {u for u in local_iframes if "youtube" not in u and "fotolife" not in u}
    for u in live_iframes - local_iframes:
        warnings.append("[WARN] 本番に他iframe/外部リソースがlocalに無い: " + u)
    return warnings


def put_entry(entry_id, title, content, categories, is_draft):
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)
    xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml += '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">\n'
    xml += "  <title>{}</title>\n".format(title)
    xml += "  <author><name>{}</name></author>\n".format(HATENA_ID)
    xml += '  <content type="text/html"><![CDATA[{}]]></content>\n'.format(content)
    xml += "{}\n".format(cat_xml)
    xml += "  <app:control>\n    <app:draft>{}</app:draft>\n  </app:control>\n".format(
        "yes" if is_draft else "no"
    )
    xml += "</entry>"

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(
        url,
        data=xml.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth_header(),
        },
    )
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")
    url_m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
    return url_m.group(1) if url_m else "(URL不明)"


def main():
    print("[1/5] AtomPubで対象エントリを検索中...")
    entry_id, title, alt_url = find_entry_id()
    print("      entry_id={} / title={} / url={}".format(entry_id, title, alt_url))

    print("[2/5] 既存エントリ取得（本番content + draft/category取得）...")
    current = get_entry(entry_id)
    categories = re.findall(r'<category[^>]*term="([^"]+)"', current)
    draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", current)
    is_draft = (draft_m.group(1) == "yes") if draft_m else True
    DEFAULT_CATS = ["その症状大丈夫", "脳神経内科", "セルフチェック"]
    if not categories:
        categories = DEFAULT_CATS
    print("      categories={} / draft={}".format(categories, is_draft))
    live_html = extract_content(current)

    print("[3/5] ローカルHTMLとの安全差分チェック...")
    local_html = HTML_FILE.read_text(encoding="utf-8")
    if "<!-- youtube embed -->" not in local_html or VIDEO_ID not in local_html:
        raise RuntimeError("ローカルHTMLにYouTube埋め込み（VIDEO_ID={}）が見つかりません。先に挿入してください。".format(VIDEO_ID))
    warnings = safety_diff(live_html, local_html)
    if warnings:
        print("      ⚠️ 本番にlocal未反映の要素を検出:")
        for w in warnings:
            print("        - " + w)
        if not FORCE_PROCEED:
            print("\n      [STOP] PUTを中止します。FORCE_PROCEED=Trueにするか、要素をlocalに反映してください。")
            sys.exit(2)
    else:
        print("      ✅ 差分なし。安全にPUT可能。")

    print("[4/5] PUT 準備（draft={}維持）...".format("yes" if is_draft else "no"))
    if not PRESERVE_DRAFT_STATE:
        is_draft = False

    print("[5/5] AtomPub PUT中...")
    result = put_entry(entry_id, title, local_html, categories, is_draft)
    print("      更新成功: {}".format(result))


if __name__ == "__main__":
    try:
        main()
    except urllib.error.HTTPError as e:
        print("HTTP Error {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:500]))
        sys.exit(1)
