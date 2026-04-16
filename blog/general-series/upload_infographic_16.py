#!/usr/bin/env python3
"""
からだの不思議 #16 インフォグラフィックアップロード＆ブログ挿入
1. はてなフォトライフにPNG画像をアップロード（WSSE認証）
2. ブログ記事にインフォグラフィックを挿入（AtomPub PUT）
"""

import sys
import re
import html as html_mod
import hashlib
import base64
import random
import string
import urllib.request
import urllib.error
from datetime import datetime, timezone
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
IMAGE_FILE = Path(r"C:\Users\jsber\Downloads\Gemini_Generated_Image_1c8ufw1c8ufw1c8u.png")

ENTRY_ID = "17179246901372000539"
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
BLOG_TITLE = "けいれんを目撃したときの正しい対応 ― やってはいけないことと救急車を呼ぶ判断【からだの不思議 #16】"


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def make_wsse_header(username, api_key):
    nonce_raw = ''.join(random.choices(string.ascii_letters + string.digits, k=40))
    nonce = nonce_raw.encode('utf-8')
    created = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    digest_raw = hashlib.sha1(nonce + created.encode('utf-8') + api_key.encode('utf-8')).digest()
    password_digest = base64.b64encode(digest_raw).decode('utf-8')
    nonce_b64 = base64.b64encode(nonce).decode('utf-8')
    wsse = (
        'UsernameToken Username="{}", PasswordDigest="{}", Nonce="{}", Created="{}"'
        .format(username, password_digest, nonce_b64, created)
    )
    return wsse


def upload_to_fotolife(env, image_path):
    """はてなフォトライフにPNG画像をアップロード"""
    api_key = env["HATENA_API_KEY"]

    with open(image_path, 'rb') as f:
        image_data = f.read()

    image_b64 = base64.b64encode(image_data).decode('utf-8')
    title = "seizure_first_aid_infographic_16"

    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://purl.org/atom/ns#">
  <title>{title}</title>
  <content mode="base64" type="image/png">{data}</content>
</entry>""".format(title=title, data=image_b64)

    url = "https://f.hatena.ne.jp/atom/post"
    wsse = make_wsse_header(HATENA_ID, api_key)

    req = urllib.request.Request(
        url,
        data=xml.encode('utf-8'),
        method="POST",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "X-WSSE": wsse,
            "Authorization": 'WSSE profile="UsernameToken"',
        },
    )

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode('utf-8')
            url_match = re.search(r'<hatena:imageurl>([^<]+)</hatena:imageurl>', body)
            if url_match:
                return url_match.group(1)
            url_match = re.search(r'https://cdn-ak\.f\.st-hatena\.com/images/fotolife/[^<"]+', body)
            if url_match:
                return url_match.group(0)
            return "URL not found in response: " + body[:500]
    except urllib.error.HTTPError as e:
        error_body = e.read().decode('utf-8', errors='replace')
        return "ERROR HTTP {}: {}".format(e.code, error_body[:500])


def get_entry_content(env):
    """既存エントリーの本文を取得"""
    api_key = env["HATENA_API_KEY"]
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, ENTRY_ID)
    auth_b64 = base64.b64encode("{}:{}".format(HATENA_ID, api_key).encode()).decode()

    req = urllib.request.Request(url, headers={"Authorization": "Basic {}".format(auth_b64)})
    with urllib.request.urlopen(req) as resp:
        entry_xml = resp.read().decode("utf-8")

    # content抽出（CDATA or HTML-escaped）
    m = re.search(r'<content[^>]*>\s*<!\[CDATA\[(.*?)\]\]>\s*</content>', entry_xml, re.DOTALL)
    if m:
        return m.group(1), entry_xml
    m = re.search(r'<content[^>]*>(.*?)</content>', entry_xml, re.DOTALL)
    if m:
        return html_mod.unescape(m.group(1)), entry_xml
    return None, entry_xml


def put_entry(env, content):
    """エントリーを更新（下書き維持）"""
    api_key = env["HATENA_API_KEY"]
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, ENTRY_ID)
    auth_b64 = base64.b64encode("{}:{}".format(HATENA_ID, api_key).encode()).decode()

    categories = ["医学教育", "からだの不思議", "てんかん", "救急"]
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)

    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>{author}</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>yes</app:draft>
  </app:control>
</entry>""".format(title=BLOG_TITLE, author=HATENA_ID, content=content, categories=cat_xml)

    req = urllib.request.Request(url, data=xml.encode("utf-8"), method="PUT", headers={
        "Content-Type": "application/xml; charset=utf-8",
        "Authorization": "Basic {}".format(auth_b64),
    })

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            return True, url_match.group(1) if url_match else "(URL取得失敗)"
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        return False, "HTTP {}: {}".format(e.code, error_body[:500])


def main():
    env = load_env()

    # 1. 画像アップロード
    print("=== からだの不思議 #16 インフォグラフィック挿入 ===\n")
    print("[1/3] はてなフォトライフにアップロード中...")
    image_url = upload_to_fotolife(env, IMAGE_FILE)
    print("  画像URL: {}".format(image_url))

    if image_url.startswith("ERROR") or image_url.startswith("URL not found"):
        print("アップロード失敗。中断します。")
        sys.exit(1)

    # 2. 既存記事を取得し、インフォグラフィックを挿入
    print("\n[2/3] 既存記事を取得・インフォグラフィック挿入...")
    content, _ = get_entry_content(env)
    if not content:
        print("ERROR: 記事本文を取得できません")
        sys.exit(1)

    infographic_html = (
        '<div style="text-align:center;margin:20px 0;">'
        '<img src="{}" alt="けいれんを目撃したときの正しい対応 インフォグラフィック — 発作中にやるべきこと・やってはいけないこと・救急車を呼ぶ基準・よくある誤解" '
        'style="max-width:100%;height:auto;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1);" '
        'loading="lazy">'
        '</div>'
    ).format(image_url)

    # h1とYouTubeボタンの後、本文の前に挿入
    # YouTubeボタンの</a></p>の後に挿入
    yt_button_end = 'YouTube動画で見る</a></p>'
    if yt_button_end in content:
        content = content.replace(
            yt_button_end,
            yt_button_end + '\n\n' + infographic_html,
            1
        )
        print("  YouTubeボタンの後にインフォグラフィック挿入完了")
    else:
        # フォールバック: 最初の<p>の前に挿入
        content = infographic_html + '\n\n' + content
        print("  記事先頭にインフォグラフィック挿入（フォールバック）")

    # 3. 記事更新
    print("\n[3/3] ブログ記事更新中...")
    ok, result = put_entry(env, content)
    if ok:
        print("\n[成功] インフォグラフィック挿入完了")
        print("  URL: {}".format(result))
    else:
        print("\n[失敗] {}".format(result))
        sys.exit(1)


if __name__ == "__main__":
    main()
