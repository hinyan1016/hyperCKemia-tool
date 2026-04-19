#!/usr/bin/env python3
"""プライベートクレジット記事にYouTubeリンクを追加してPUT更新"""

import re
import base64
import urllib.request
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\posts\private_credit_crisis_blog.html")
ENTRY_ID = "17179246901374206438"
TITLE = "プライベートクレジット問題とは？ 3兆ドル市場の構造リスクと今後のシナリオをわかりやすく解説"
CATEGORIES = ["金融リスク", "プライベートクレジット", "投資"]


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def convert_to_inline(html):
    """CSSクラスをインラインスタイルに変換し、<style>ブロック等を除去"""
    m = re.search(r'</style>\s*', html, re.DOTALL)
    if m:
        body = html[m.end():]
    else:
        body = html

    body = re.sub(r'\s*</div>\s*</body>\s*</html>\s*$', '', body)

    body = body.replace('<h1>', '<h1 style="font-size:1.6rem;color:#1a3c5e;border-bottom:2px solid #1a3c5e;padding-bottom:8px;margin-top:0;">')
    body = re.sub(
        r'<h2( id="[^"]*")?>',
        lambda m: '<h2%s style="font-size:1.3rem;color:#1a3c5e;border-left:4px solid #c0392b;padding-left:12px;margin-top:2rem;">' % (m.group(1) or ''),
        body
    )
    body = body.replace('<h3>', '<h3 style="font-size:1.1rem;color:#444;margin-top:1.2rem;">')
    body = re.sub(r'<strong>', '<strong style="font-weight:bold;color:#1a3c5e;">', body)
    body = re.sub(r'<sup>', '<sup style="color:#c0392b;font-weight:bold;font-size:0.75em;">', body)

    body = body.replace('<span class="topic-label">', '<span style="display:inline-block;background:#c0392b;color:#fff;font-size:0.85rem;padding:2px 12px;border-radius:4px;margin-bottom:8px;">')
    body = body.replace('<div class="toc-box">', '<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;">')
    body = body.replace('<div class="important-box">', '<div style="background:#e8f0f8;border-left:4px solid #1a3c5e;border-radius:4px;padding:12px 16px;margin:16px 0;">')
    body = body.replace('<div class="warning-box">', '<div style="background:#fff3cd;border-left:4px solid #ffc107;border-radius:4px;padding:12px 16px;margin:16px 0;">')
    body = body.replace('<div class="alert-box">', '<div style="background:#f8d7da;border-left:4px solid #c0392b;border-radius:4px;padding:12px 16px;margin:16px 0;">')
    body = body.replace('<div class="neutral-box">', '<div style="background:#e8f5e9;border-left:4px solid #27ae60;border-radius:4px;padding:12px 16px;margin:16px 0;">')
    body = body.replace('<div class="related-box">', '<div style="background:#f0f4f8;border-left:4px solid #1a3c5e;border-radius:4px;padding:12px 16px;margin:20px 0;">')
    body = body.replace('<div class="figure-box">', '<div style="background:#f0f4f8;border:1px solid #cdd5de;border-radius:8px;padding:20px;margin:20px 0;text-align:center;">')
    body = body.replace('<div class="fig-title">', '<div style="font-weight:bold;color:#1a3c5e;margin-bottom:12px;font-size:1rem;">')
    body = body.replace('<div class="flow-diagram">', '<div style="display:flex;flex-wrap:wrap;justify-content:center;align-items:center;gap:8px;margin:12px 0;">')
    body = body.replace('<div class="flow-node risk">', '<div style="background:#c0392b;color:#fff;padding:8px 16px;border-radius:6px;font-size:0.9rem;text-align:center;">')
    body = body.replace('<div class="flow-node warning">', '<div style="background:#e67e22;color:#fff;padding:8px 16px;border-radius:6px;font-size:0.9rem;text-align:center;">')
    body = body.replace('<div class="flow-node">', '<div style="background:#1a3c5e;color:#fff;padding:8px 16px;border-radius:6px;font-size:0.9rem;text-align:center;">')
    body = body.replace('<div class="flow-arrow">', '<div style="font-size:1.2rem;color:#888;">')

    body = body.replace('<table class="data-table">', '<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:0.95rem;">')
    body = re.sub(r'<th>', '<th style="background:#1a3c5e;color:#fff;font-weight:bold;padding:10px 12px;text-align:left;border:1px solid #ddd;">', body)
    body = re.sub(r'<th style="background:#c0392b;">', '<th style="background:#c0392b;color:#fff;font-weight:bold;padding:10px 12px;text-align:left;border:1px solid #ddd;">', body)
    body = re.sub(r'<th style="background:#1a3c5e;">', '<th style="background:#1a3c5e;color:#fff;font-weight:bold;padding:10px 12px;text-align:left;border:1px solid #ddd;">', body)
    body = re.sub(r'<td>', '<td style="padding:10px 12px;border-bottom:1px solid #dee2e6;vertical-align:top;border:1px solid #ddd;">', body)
    body = re.sub(r'<td style="', '<td style="padding:10px 12px;border-bottom:1px solid #dee2e6;vertical-align:top;border:1px solid #ddd;', body)

    body = body.replace('<div class="timeline">', '<div style="border-left:3px solid #c0392b;padding-left:20px;margin:16px 0;">')
    body = body.replace('<div class="timeline-item">', '<div style="margin-bottom:16px;">')
    body = body.replace('<div class="timeline-date">', '<div style="font-weight:bold;color:#c0392b;font-size:0.9rem;">')

    body = body.replace('<div class="references">', '<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">')
    body = body.replace('<div class="ref-item">', '<div style="margin-bottom:12px;padding:10px;background:#f8f9fa;border-radius:4px;font-size:0.9rem;">')
    body = body.replace('<span class="ref-numbers">', '<span style="color:#c0392b;font-weight:bold;margin-right:8px;font-size:0.85em;">')
    body = body.replace('<span class="ref-content">', '<span style="color:#555;">')

    body = re.sub(r'<!--.*?-->', '', body, flags=re.DOTALL)
    return body.strip()


def update_entry(env, body_html):
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]

    cat_xml = ''.join('<category term="%s" />' % c for c in CATEGORIES)

    entry_xml = '''<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>%s</title>
  <author><name>%s</name></author>
  %s
  <content type="text/html"><![CDATA[%s]]></content>
  <app:control>
    <app:draft>no</app:draft>
  </app:control>
</entry>''' % (TITLE, hatena_id, cat_xml, body_html)

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(hatena_id, blog_domain, ENTRY_ID)
    auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()

    req = urllib.request.Request(url, data=entry_xml.encode('utf-8'), method='PUT')
    req.add_header('Content-Type', 'application/atom+xml; charset=utf-8')
    req.add_header('Authorization', 'Basic {}'.format(auth_b64))

    try:
        with urllib.request.urlopen(req) as resp:
            print("SUCCESS: Entry updated!")
            print("Status: {}".format(resp.status))
    except urllib.error.HTTPError as e:
        print("ERROR {}: {}".format(e.code, e.read().decode('utf-8')[:500]))
        raise


if __name__ == "__main__":
    env = load_env()
    raw_html = HTML_FILE.read_text(encoding="utf-8")
    body = convert_to_inline(raw_html)
    update_entry(env, body)
