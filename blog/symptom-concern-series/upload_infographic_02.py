#!/usr/bin/env python3
"""
その症状大丈夫#02 ろれつが回らない
- はてなフォトライフにインフォグラフィックをアップロード
- 02_slurred_speech_hatena.htmlに画像タグを挿入（リード直後、YouTube埋め込みの直前）
- AtomPub APIで既存の下書きを更新（draft=yes維持）
"""

from __future__ import annotations

import re
import sys
import hashlib
import base64
import random
import string
import urllib.request
import urllib.error
from datetime import datetime, timezone
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"

IMAGE_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\その症状大丈夫02_ろれつ\infographic.png")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\symptom-concern-series\02_slurred_speech_hatena.html")

IMAGE_ALT = "ろれつが回りにくい 緊急性の見分け方 インフォグラフィック — 構音障害と失語症の違い・突然vs徐々に・FASTチェック・TIA・受診の目安"
TITLE_KEYWORD = "ろれつ"
REQUIRE_DRAFT = True


def auth_header() -> str:
    b64 = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()
    return "Basic {}".format(b64)


def wsse_header() -> str:
    nonce_raw = "".join(random.choices(string.ascii_letters + string.digits, k=40))
    nonce = nonce_raw.encode("utf-8")
    created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    digest = hashlib.sha1(nonce + created.encode() + API_KEY.encode()).digest()
    pwd = base64.b64encode(digest).decode()
    nonce_b64 = base64.b64encode(nonce).decode()
    return 'UsernameToken Username="{}", PasswordDigest="{}", Nonce="{}", Created="{}"'.format(
        HATENA_ID, pwd, nonce_b64, created
    )


def upload_fotolife() -> str:
    with open(IMAGE_FILE, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml += '<entry xmlns="http://purl.org/atom/ns#">\n'
    xml += "  <title>slurred_speech_infographic</title>\n"
    xml += '  <content mode="base64" type="image/png">{}</content>\n'.format(b64)
    xml += "</entry>"

    req = urllib.request.Request(
        "https://f.hatena.ne.jp/atom/post",
        data=xml.encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "X-WSSE": wsse_header(),
            "Authorization": 'WSSE profile="UsernameToken"',
        },
    )
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")
    m = re.search(r"<hatena:imageurl>([^<]+)</hatena:imageurl>", body)
    if m:
        return m.group(1)
    m = re.search(r'https://cdn-ak\.f\.st-hatena\.com/images/fotolife/[^<"\s]+', body)
    if m:
        return m.group(0)
    raise RuntimeError("imageurl not found: " + body[:500])


def find_entry_id() -> tuple[str, str]:
    """AtomPubコレクションから対象記事のentry idを検索。
    REQUIRE_DRAFT=True の場合は draft=yes の下書きのみ対象。"""
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
            draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", e)
            is_draft = (draft_m.group(1) == "yes") if draft_m else False
            if REQUIRE_DRAFT and not is_draft:
                print("      候補スキップ（draft=no）: {}".format(title_m.group(1)))
                continue
            id_m = re.search(r"<id>tag:blog\.hatena\.ne\.jp[^:]*:[^:]*-(\d+)</id>", e)
            if id_m:
                return id_m.group(1), title_m.group(1)
            id_m = re.search(r"<id>([^<]+)</id>", e)
            if id_m:
                return id_m.group(1).rsplit("-", 1)[-1], title_m.group(1)
        next_m = re.search(r'<link rel="next" href="([^"]+)"', body)
        url = next_m.group(1) if next_m else None
        page += 1
    raise RuntimeError(
        "対象記事が見つかりません（TITLE_KEYWORD={}, REQUIRE_DRAFT={}）".format(TITLE_KEYWORD, REQUIRE_DRAFT)
    )


def get_entry(entry_id: str) -> str:
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, headers={"Authorization": auth_header()})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def put_entry(entry_id: str, title: str, content: str, categories: list[str], is_draft: bool) -> str:
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


def main() -> None:
    # 再実行防止：引数で既存URLを指定可能
    if len(sys.argv) > 1 and sys.argv[1].startswith("http"):
        image_url = sys.argv[1]
        print("[1/4] 既存URL使用: {}".format(image_url))
    else:
        print("[1/4] インフォグラフィックをフォトライフへアップロード中...")
        image_url = upload_fotolife()
        print("      画像URL: {}".format(image_url))

    print("[2/4] 下書き記事のentry id検索中...")
    entry_id, title = find_entry_id()
    print("      entry_id={} / title={}".format(entry_id, title))

    print("[3/4] 既存エントリからcategoryとdraft状態を取得...")
    current = get_entry(entry_id)
    categories = re.findall(r'<category[^>]*term="([^"]+)"', current)
    draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", current)
    is_draft = (draft_m.group(1) == "yes") if draft_m else True
    # シリーズ既定カテゴリ（空なら補完）
    DEFAULT_CATS = ["その症状大丈夫", "脳神経内科", "セルフチェック"]
    if not categories:
        categories = DEFAULT_CATS
        print("      categories空→既定値補完: {}".format(categories))
    print("      categories={} / draft={}".format(categories, is_draft))

    # 4. HTMLに画像を挿入（すでに挿入済みならスキップ）
    html = HTML_FILE.read_text(encoding="utf-8")
    marker = "<!-- infographic -->"
    if marker in html:
        print("[4/4] 画像は既にHTMLに挿入済み。そのままPUTします")
    else:
        img_block = (
            "\n<!-- infographic -->\n"
            '<div style="text-align:center; margin:24px 0;">\n'
            '<img src="{url}" alt="{alt}" '
            'style="max-width:100%; height:auto; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.1);" '
            'loading="lazy">\n'
            "</div>\n\n"
        ).format(url=image_url, alt=IMAGE_ALT)

        # YouTube埋め込みの直前に挿入（リード文のあと）
        target = "<!-- youtube embed -->"
        if target not in html:
            raise RuntimeError("挿入先（YouTube embedマーカー）が見つかりません")
        html = html.replace(target, img_block + target, 1)
        HTML_FILE.write_text(html, encoding="utf-8")
        print("[4/4] ローカルHTMLに挿入完了（{}バイト）".format(len(html.encode("utf-8"))))

    # 5. AtomPub PUT
    print("      AtomPub PUT中（draft={}）...".format("yes" if is_draft else "no"))
    result = put_entry(entry_id, title, html, categories, is_draft)
    print("      更新成功: {}".format(result))


if __name__ == "__main__":
    try:
        main()
    except urllib.error.HTTPError as e:
        print("HTTP Error {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:500]))
        sys.exit(1)
