#!/usr/bin/env python3
"""プライベートクレジット問題ブログ記事をはてなブログに下書き投稿する"""

import re
import base64
import urllib.request
from pathlib import Path

# .envファイルから認証情報を読み込み
ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")

def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env

env = load_env()
HATENA_ID = env["HATENA_ID"]
BLOG_DOMAIN = env["HATENA_BLOG_DOMAIN"]
API_KEY = env["HATENA_API_KEY"]

TITLE = "プライベートクレジット問題とは？ 3兆ドル市場の構造リスクと今後のシナリオをわかりやすく解説"
CATEGORIES = ["金融リスク", "プライベートクレジット", "投資"]

# --- インラインスタイル変換 ---
def convert_to_inline(html):
    """CSSクラスをインラインスタイルに変換し、<style>ブロック等を除去"""

    # </style>以降を本文として抽出
    m = re.search(r'</style>\s*', html, re.DOTALL)
    if m:
        body = html[m.end():]
    else:
        body = html

    # 末尾のwrapper/body/html除去
    body = re.sub(r'\s*</div>\s*</body>\s*</html>\s*$', '', body)

    # h1
    body = body.replace('<h1>', '<h1 style="font-size:1.6rem;color:#1a3c5e;border-bottom:2px solid #1a3c5e;padding-bottom:8px;margin-top:0;">')

    # h2 (with and without id)
    body = re.sub(
        r'<h2( id="[^"]*")?>',
        lambda m: '<h2%s style="font-size:1.3rem;color:#1a3c5e;border-left:4px solid #c0392b;padding-left:12px;margin-top:2rem;">' % (m.group(1) or ''),
        body
    )

    # h3 — add ▶ prefix via inline style
    body = body.replace('<h3>', '<h3 style="font-size:1.1rem;color:#444;margin-top:1.2rem;">')

    # strong
    body = re.sub(r'<strong>', '<strong style="font-weight:bold;color:#1a3c5e;">', body)

    # sup
    body = re.sub(r'<sup>', '<sup style="color:#c0392b;font-weight:bold;font-size:0.75em;">', body)

    # topic-label
    body = body.replace(
        '<span class="topic-label">',
        '<span style="display:inline-block;background:#c0392b;color:#fff;font-size:0.85rem;padding:2px 12px;border-radius:4px;margin-bottom:8px;">'
    )

    # toc-box
    body = body.replace(
        '<div class="toc-box">',
        '<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;">'
    )

    # important-box
    body = body.replace(
        '<div class="important-box">',
        '<div style="background:#e8f0f8;border-left:4px solid #1a3c5e;border-radius:4px;padding:12px 16px;margin:16px 0;">'
    )

    # warning-box
    body = body.replace(
        '<div class="warning-box">',
        '<div style="background:#fff3cd;border-left:4px solid #ffc107;border-radius:4px;padding:12px 16px;margin:16px 0;">'
    )

    # alert-box
    body = body.replace(
        '<div class="alert-box">',
        '<div style="background:#f8d7da;border-left:4px solid #c0392b;border-radius:4px;padding:12px 16px;margin:16px 0;">'
    )

    # neutral-box
    body = body.replace(
        '<div class="neutral-box">',
        '<div style="background:#e8f5e9;border-left:4px solid #27ae60;border-radius:4px;padding:12px 16px;margin:16px 0;">'
    )

    # related-box
    body = body.replace(
        '<div class="related-box">',
        '<div style="background:#f0f4f8;border-left:4px solid #1a3c5e;border-radius:4px;padding:12px 16px;margin:20px 0;">'
    )

    # figure-box
    body = body.replace(
        '<div class="figure-box">',
        '<div style="background:#f0f4f8;border:1px solid #cdd5de;border-radius:8px;padding:20px;margin:20px 0;text-align:center;">'
    )

    # fig-title
    body = body.replace(
        '<div class="fig-title">',
        '<div style="font-weight:bold;color:#1a3c5e;margin-bottom:12px;font-size:1rem;">'
    )

    # flow-diagram
    body = body.replace(
        '<div class="flow-diagram">',
        '<div style="display:flex;flex-wrap:wrap;justify-content:center;align-items:center;gap:8px;margin:12px 0;">'
    )

    # flow-node variants
    body = body.replace(
        '<div class="flow-node risk">',
        '<div style="background:#c0392b;color:#fff;padding:8px 16px;border-radius:6px;font-size:0.9rem;text-align:center;">'
    )
    body = body.replace(
        '<div class="flow-node warning">',
        '<div style="background:#e67e22;color:#fff;padding:8px 16px;border-radius:6px;font-size:0.9rem;text-align:center;">'
    )
    body = body.replace(
        '<div class="flow-node">',
        '<div style="background:#1a3c5e;color:#fff;padding:8px 16px;border-radius:6px;font-size:0.9rem;text-align:center;">'
    )

    # flow-arrow
    body = body.replace(
        '<div class="flow-arrow">',
        '<div style="font-size:1.2rem;color:#888;">'
    )

    # data-table
    body = body.replace(
        '<table class="data-table">',
        '<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:0.95rem;">'
    )

    # th — inline styles (重要: はてなブログではインライン必須)
    # 既にstyle付きのthはそのまま、class無しの素のthにデフォルトスタイル追加
    body = re.sub(
        r'<th>',
        '<th style="background:#1a3c5e;color:#fff;font-weight:bold;padding:10px 12px;text-align:left;border:1px solid #ddd;">',
        body
    )
    # style="background:#c0392b;" 付きのthにも残りのスタイルを追加
    body = re.sub(
        r'<th style="background:#c0392b;">',
        '<th style="background:#c0392b;color:#fff;font-weight:bold;padding:10px 12px;text-align:left;border:1px solid #ddd;">',
        body
    )
    body = re.sub(
        r'<th style="background:#1a3c5e;">',
        '<th style="background:#1a3c5e;color:#fff;font-weight:bold;padding:10px 12px;text-align:left;border:1px solid #ddd;">',
        body
    )

    # td
    body = re.sub(
        r'<td>',
        '<td style="padding:10px 12px;border-bottom:1px solid #dee2e6;vertical-align:top;border:1px solid #ddd;">',
        body
    )
    # td with existing style
    body = re.sub(
        r'<td style="',
        '<td style="padding:10px 12px;border-bottom:1px solid #dee2e6;vertical-align:top;border:1px solid #ddd;',
        body
    )

    # timeline
    body = body.replace(
        '<div class="timeline">',
        '<div style="border-left:3px solid #c0392b;padding-left:20px;margin:16px 0;">'
    )
    body = body.replace(
        '<div class="timeline-item">',
        '<div style="margin-bottom:16px;">'
    )
    body = body.replace(
        '<div class="timeline-date">',
        '<div style="font-weight:bold;color:#c0392b;font-size:0.9rem;">'
    )

    # references
    body = body.replace(
        '<div class="references">',
        '<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">'
    )
    body = body.replace(
        '<div class="ref-item">',
        '<div style="margin-bottom:12px;padding:10px;background:#f8f9fa;border-radius:4px;font-size:0.9rem;">'
    )
    body = body.replace(
        '<span class="ref-numbers">',
        '<span style="color:#c0392b;font-weight:bold;margin-right:8px;font-size:0.85em;">'
    )
    body = body.replace(
        '<span class="ref-content">',
        '<span style="color:#555;">'
    )

    # HTMLコメント除去
    body = re.sub(r'<!--.*?-->', '', body, flags=re.DOTALL)

    return body.strip()


# --- Basic認証 ---
def make_auth_header():
    auth_str = "{}:{}".format(HATENA_ID, API_KEY)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()
    return "Basic {}".format(auth_b64)


# --- 記事投稿 ---
def post_draft(title, body_html, categories):
    cat_xml = ''.join('<category term="%s" />' % c for c in categories)

    entry_xml = '''<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>%s</title>
  <author><name>%s</name></author>
  %s
  <content type="text/html"><![CDATA[%s]]></content>
  <app:control>
    <app:draft>yes</app:draft>
  </app:control>
</entry>''' % (title, HATENA_ID, cat_xml, body_html)

    url = "https://blog.hatena.ne.jp/%s/%s/atom/entry" % (HATENA_ID, BLOG_DOMAIN)

    req = urllib.request.Request(url, data=entry_xml.encode('utf-8'), method='POST')
    req.add_header('Content-Type', 'application/atom+xml; charset=utf-8')
    req.add_header('Authorization', make_auth_header())

    try:
        with urllib.request.urlopen(req) as resp:
            result = resp.read().decode('utf-8')
            # entry IDを抽出
            m = re.search(r'<id>tag:blog.hatena.ne.jp,\d+:blog-[^-]+-(\d+)</id>', result)
            entry_id = m.group(1) if m else "unknown"
            # URLを抽出
            m2 = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
            entry_url = m2.group(1) if m2 else "unknown"
            print("SUCCESS: Draft posted!")
            print("Entry ID: %s" % entry_id)
            print("URL: %s" % entry_url)
            return result
    except urllib.error.HTTPError as e:
        print("ERROR %d: %s" % (e.code, e.read().decode('utf-8')[:500]))
        raise


if __name__ == "__main__":
    # HTMLファイル読み込み
    html_path = r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\posts\private_credit_crisis_blog.html"
    with open(html_path, "r", encoding="utf-8") as f:
        raw_html = f.read()

    # インラインスタイル変換
    body = convert_to_inline(raw_html)

    # 投稿
    post_draft(TITLE, body, CATEGORIES)
