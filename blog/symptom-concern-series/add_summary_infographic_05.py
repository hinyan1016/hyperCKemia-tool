#!/usr/bin/env python3
"""
その症状大丈夫#05 - まとめインフォグラフィック追加
- ユーザー画像を Fotolife にアップロード
- <h2 id="まとめ">まとめ</h2> 直後に <figure>+<img>+<figcaption> を挿入
- 安全パターン: PUT前に本番GET→差分検出
- 公開記事のため draft=no を維持
"""

import re
import sys
import io
import hashlib
import base64
import random
import string
import urllib.request
import urllib.error
from datetime import datetime, timezone
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"

IMAGE_FILE = Path(r"C:\Users\jsber\Downloads\ChatGPT Image 2026年4月26日 16_25_21.png")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\symptom-concern-series\05_leg_cramps_hatena.html")

IMAGE_TITLE = "leg_cramps_summary_infographic"
IMAGE_ALT = (
    "寝ている時に足がつる 危険なサインと予防法 まとめインフォグラフィック — "
    "①有痛性筋けいれんとは／②よくある原因（脱水・ミネラル不足・筋疲労・冷え・加齢・妊娠後期）／"
    "③薬の副作用（利尿薬・スタチン・ARB/ACE阻害薬・β刺激薬）／"
    "④隠れた病気（末梢神経障害・脊柱管狭窄症・ALS・甲状腺機能低下症・下肢動脈疾患）／"
    "⑤要注意の「つり方」5つ／"
    "⑥A受診の目安／Bつった時の対処／C予防する5つの習慣／3つのポイント"
)
FIGURE_CAPTION = "図：「寝ている時に足がつる」記事の全体まとめ。原因→危険サイン→受診の目安→予防までを1枚で整理"

TITLE_KEYWORD = "寝ている時に足がつる"
TARGET_URL_FRAGMENT = "2026/04/26/105809"  # 公開後の確定URL
PRESERVE_DRAFT_STATE = True  # 本番のdraft状態（=no）を維持
FORCE_PROCEED = False


def auth_header():
    b64 = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()
    return "Basic " + b64


def wsse_header():
    nonce_raw = "".join(random.choices(string.ascii_letters + string.digits, k=40))
    nonce = nonce_raw.encode("utf-8")
    created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    digest = hashlib.sha1(nonce + created.encode() + API_KEY.encode()).digest()
    pwd = base64.b64encode(digest).decode()
    nonce_b64 = base64.b64encode(nonce).decode()
    return 'UsernameToken Username="{}", PasswordDigest="{}", Nonce="{}", Created="{}"'.format(
        HATENA_ID, pwd, nonce_b64, created
    )


def upload_fotolife():
    if not IMAGE_FILE.exists():
        raise FileNotFoundError("画像が見つかりません: {}".format(IMAGE_FILE))
    with open(IMAGE_FILE, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml += '<entry xmlns="http://purl.org/atom/ns#">\n'
    xml += "  <title>{}</title>\n".format(IMAGE_TITLE)
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
    m = re.search(r"https://cdn-ak\.f\.st-hatena\.com/images/fotolife/[^<\"\s]+", body)
    if m:
        return m.group(0)
    raise RuntimeError("imageurl not found: " + body[:500])


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
    """本番にあって local に無い要素を検出"""
    warnings = []
    if "<!-- infographic -->" in live_html and "<!-- infographic -->" not in local_html:
        warnings.append("[CRITICAL] インフォグラ既に本番にありlocalに無い（重複の可能性）")
    live_imgs = set(re.findall(r'src="([^"]+fotolife[^"]+)"', live_html))
    local_imgs = set(re.findall(r'src="([^"]+fotolife[^"]+)"', local_html))
    for img in live_imgs - local_imgs:
        warnings.append("[CRITICAL] 本番Fotolife画像がlocalに無い: " + img)
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
    if len(sys.argv) > 1 and sys.argv[1].startswith("http"):
        image_url = sys.argv[1]
        print("[1/6] 既存URL使用: {}".format(image_url))
    else:
        print("[1/6] まとめインフォグラフィックを Fotolife へアップロード中...")
        image_url = upload_fotolife()
        print("      画像URL: {}".format(image_url))

    print("[2/6] 対象エントリ検索中...")
    entry_id, title, alt_url = find_entry_id()
    print("      entry_id={} / title={} / url={}".format(entry_id, title, alt_url))

    print("[3/6] 既存エントリ取得（content + draft/category）...")
    current = get_entry(entry_id)
    categories = re.findall(r'<category[^>]*term="([^"]+)"', current)
    draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", current)
    is_draft = (draft_m.group(1) == "yes") if draft_m else False
    if not categories:
        categories = ["その症状大丈夫", "脳神経内科", "セルフチェック"]
    print("      categories={} / draft={}".format(categories, is_draft))
    live_html = extract_content(current)

    print("[4/6] ローカルHTMLにインフォグラフィック挿入...")
    html = HTML_FILE.read_text(encoding="utf-8")

    figure_block = (
        "\n<!-- infographic -->\n"
        '<figure style="text-align:center; margin:24px auto;">\n'
        '<img src="{url}" alt="{alt}" '
        'style="max-width:100%; height:auto; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.1);" '
        'loading="lazy">\n'
        '<figcaption style="font-size:0.9em; color:#666; margin-top:8px;">{caption}</figcaption>\n'
        "</figure>\n\n"
    ).format(url=image_url, alt=IMAGE_ALT, caption=FIGURE_CAPTION)

    summary_marker = '<h2 id="まとめ">まとめ</h2>'
    if summary_marker not in html:
        raise RuntimeError("挿入先（まとめh2）が見つかりません")
    if "<!-- infographic -->" in html:
        print("      [warn] 既に <!-- infographic --> がローカルに存在。重複防止のためスキップ")
    else:
        html = html.replace(summary_marker, summary_marker + figure_block, 1)
        HTML_FILE.write_text(html, encoding="utf-8")
        print("      ローカルHTML更新（{}バイト）".format(len(html.encode("utf-8"))))

    print("[5/6] 本番との安全差分チェック...")
    warnings = safety_diff(live_html, html)
    if warnings:
        print("      ⚠️ 警告:")
        for w in warnings:
            print("        - " + w)
        if not FORCE_PROCEED:
            print("      [STOP] PUT中止")
            sys.exit(2)
    else:
        print("      ✅ 差分なし。安全にPUT可能。")

    print("[6/6] AtomPub PUT中（draft={}）...".format("yes" if is_draft else "no"))
    if not PRESERVE_DRAFT_STATE:
        is_draft = False
    result = put_entry(entry_id, title, html, categories, is_draft)
    print("      更新成功: {}".format(result))


if __name__ == "__main__":
    try:
        main()
    except urllib.error.HTTPError as e:
        print("HTTP Error {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:500]))
        sys.exit(1)
