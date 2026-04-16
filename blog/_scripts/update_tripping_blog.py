#!/usr/bin/env python3
"""
からだの不思議 #15 つまずきブログ更新スクリプト
1. はてなフォトライフにインフォグラフィックをアップロード
2. ローカルHTMLにインフォグラフィックURL挿入 + YouTube iframe埋め込み追加
3. AtomPub APIで下書き記事を更新（下書きのまま維持）
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
IMAGE_FILE = Path(r"C:\Users\jsber\Downloads\Gemini_Generated_Image_rtbmijrtbmijrtbm.png")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\general-series\15_frequent_tripping_blog.html")

ENTRY_ID = "17179246901372000531"
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
YOUTUBE_ID = "UX_sXj_27Ts"


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
    title = "tripping_infographic_15"

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


def update_blog(env, html_content):
    """AtomPub APIでブログ記事を更新（下書き維持）"""
    api_key = env["HATENA_API_KEY"]

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, ENTRY_ID)
    auth_b64 = base64.b64encode("{}:{}".format(HATENA_ID, api_key).encode()).decode()

    # 現在のエントリを取得
    req = urllib.request.Request(url, method="GET",
        headers={"Authorization": "Basic {}".format(auth_b64)})
    with urllib.request.urlopen(req) as resp:
        entry_xml = resp.read().decode("utf-8")

    # タイトルとカテゴリを保持
    title_match = re.search(r'<title[^>]*>([^<]+)</title>', entry_xml)
    title = title_match.group(1) if title_match else ""
    categories = re.findall(r'<category[^>]*term="([^"]+)"', entry_xml)
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)

    # draft状態を保持
    draft_match = re.search(r'<app:draft>([^<]+)</app:draft>', entry_xml)
    draft_status = draft_match.group(1) if draft_match else "yes"

    content = html_content

    new_xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    new_xml += '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">\n'
    new_xml += '  <title>{}</title>\n'.format(title)
    new_xml += '  <author><name>{}</name></author>\n'.format(HATENA_ID)
    new_xml += '  <content type="text/html"><![CDATA[{}]]></content>\n'.format(content)
    new_xml += '{}\n'.format(cat_xml)
    new_xml += '  <app:control>\n    <app:draft>{}</app:draft>\n  </app:control>\n'.format(draft_status)
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

    # 2. ローカルHTMLの画像プレースホルダーを差し替え
    html = HTML_FILE.read_text(encoding="utf-8")

    html = html.replace("INFOGRAPHIC_URL_PLACEHOLDER", image_url)
    print("インフォグラフィックURLを挿入しました: {}".format(image_url))

    # 3. YouTube動画リンクをiframe埋め込みに差し替え
    youtube_link = '<p><strong style="font-weight:bold;color:#2c5aa0;">▶ YouTube動画：</strong><a href="https://youtu.be/{}">この記事の内容を動画でも解説しています</a></p>'.format(YOUTUBE_ID)
    youtube_embed = (
        '<div style="text-align:center;margin:20px 0;">\n'
        '<iframe width="560" height="315" src="https://www.youtube.com/embed/{}" '
        'frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" '
        'allowfullscreen style="max-width:100%;"></iframe>\n'
        '</div>'
    ).format(YOUTUBE_ID)

    if youtube_link in html:
        html = html.replace(youtube_link, youtube_embed)
        print("YouTube iframe埋め込みに変換しました")
    else:
        print("WARNING: YouTubeリンクが見つかりません。手動で確認してください。")

    # 4. ローカルファイルも更新
    HTML_FILE.write_text(html, encoding="utf-8")
    print("ローカルHTML更新完了")

    # 5. ブログ記事を更新（下書き維持）
    print("ブログ記事更新中...")
    result_url = update_blog(env, html)
    print("更新成功: {}".format(result_url))


if __name__ == "__main__":
    main()
