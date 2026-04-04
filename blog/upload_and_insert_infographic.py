#!/usr/bin/env python3
"""
1. はてなフォトライフに画像をアップロード（WSSE認証）
2. ブログ記事HTMLにインフォグラフィックを挿入
3. AtomPub APIでブログ記事を更新
"""

import os
import re
import hashlib
import base64
import random
import string
import urllib.request
import urllib.error
from datetime import datetime, timezone
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
IMAGE_FILE = Path(r"C:\Users\jsber\Downloads\Gemini_Generated_Image_u7bnuru7bnuru7bn.png")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\general-series\02_migraine_5_things_blog.html")

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
    title = "migraine_infographic"

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
            # 画像URLを抽出
            url_match = re.search(r'<hatena:imageurl>([^<]+)</hatena:imageurl>', body)
            if url_match:
                return url_match.group(1)
            # fallback: syntax urlを探す
            url_match = re.search(r'https://cdn-ak\.f\.st-hatena\.com/images/fotolife/[^<"]+', body)
            if url_match:
                return url_match.group(0)
            # body全体を返してデバッグ
            return "URL not found in response: " + body[:500]
    except urllib.error.HTTPError as e:
        error_body = e.read().decode('utf-8', errors='replace')
        return "ERROR HTTP {}: {}".format(e.code, error_body[:500])


def update_blog(env, html_content):
    """AtomPub APIでブログ記事を更新"""
    api_key = env["HATENA_API_KEY"]

    # 現在のエントリを取得
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, ENTRY_ID)
    auth_b64 = base64.b64encode("{}:{}".format(HATENA_ID, api_key).encode()).decode()

    req = urllib.request.Request(url, method="GET",
        headers={"Authorization": "Basic {}".format(auth_b64)})
    with urllib.request.urlopen(req) as resp:
        entry_xml = resp.read().decode("utf-8")

    # タイトルとカテゴリを保持
    title_match = re.search(r'<title[^>]*>([^<]+)</title>', entry_xml)
    title = title_match.group(1) if title_match else ""
    categories = re.findall(r'<category[^>]*term="([^"]+)"', entry_xml)
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)

    # body部分を抽出
    body_match = re.search(r'<body>(.*)</body>', html_content, re.DOTALL)
    content = body_match.group(1).strip() if body_match else html_content

    new_xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    new_xml += '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">\n'
    new_xml += '  <title>{}</title>\n'.format(title)
    new_xml += '  <author><name>{}</name></author>\n'.format(HATENA_ID)
    new_xml += '  <content type="text/html"><![CDATA[{}]]></content>\n'.format(content)
    new_xml += '{}\n'.format(cat_xml)
    new_xml += '  <app:control>\n    <app:draft>no</app:draft>\n  </app:control>\n'
    new_xml += '</entry>'

    req = urllib.request.Request(url, data=new_xml.encode("utf-8"), method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth_b64),
        })
    with urllib.request.urlopen(req) as resp:
        result = resp.read().decode("utf-8")

    url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
    return url_match.group(1) if url_match else "URL不明"


def main():
    env = load_env()

    # 1. 画像をアップロード
    print("画像アップロード中...")
    image_url = upload_to_fotolife(env, IMAGE_FILE)
    print("画像URL: {}".format(image_url))

    if image_url.startswith("ERROR") or image_url.startswith("URL not found"):
        print("アップロード失敗。中断します。")
        return

    # 2. ローカルHTMLに画像を挿入
    html = HTML_FILE.read_text(encoding="utf-8")

    # リード文の後、目次の前に挿入（SEO最適位置）
    infographic_html = (
        '\n<!-- インフォグラフィック -->\n'
        '<div style="text-align:center; margin:20px 0;">\n'
        '<img src="{}" alt="片頭痛持ちが知っておくべき5つのこと インフォグラフィック — 脳の過敏体質・薬物乱用頭痛・予防薬・トリガー管理・危険な頭痛の見分け方" '
        'style="max-width:100%; height:auto; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.1);" '
        'loading="lazy">\n'
        '</div>\n\n'
    ).format(image_url)

    target = '<!-- 目次 -->'
    if target in html:
        html = html.replace(target, infographic_html + target)
        print("インフォグラフィックを目次の前に挿入しました")
    else:
        print("目次コメントが見つかりません")
        return

    # 3. ローカルファイルも更新
    HTML_FILE.write_text(html, encoding="utf-8")
    print("ローカルHTML更新完了")

    # 4. ブログ記事を更新
    print("ブログ記事更新中...")
    result_url = update_blog(env, html)
    print("更新成功: {}".format(result_url))


if __name__ == "__main__":
    main()
