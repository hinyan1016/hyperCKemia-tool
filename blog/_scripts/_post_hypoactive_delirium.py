#!/usr/bin/env python3
"""
低活動性せん妄ブログ記事: <style>除去 → インラインスタイル変換 → はてなブログ下書き投稿
"""
import os
import sys
import re
import urllib.request
import urllib.error
import base64
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\posts\hypoactive_delirium_blog.html")
TITLE = "低活動性せん妄 ― 「静かなせん妄」を見逃さないための実践ガイド"
CATEGORIES = ["医学教育", "脳神経内科", "せん妄"]


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


# CSS class → inline style mapping
INLINE_STYLES = {
    "medical-article": "max-width:100%;padding:16px;line-height:1.7;color:#333;",
    "toc-box": "background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;",
    "important-box": "background:#e8f4fd;border-left:4px solid #2c5aa0;border-radius:4px;padding:12px 16px;margin:16px 0;",
    "warning-box": "background:#fff3cd;border-left:4px solid #ffc107;border-radius:4px;padding:12px 16px;margin:16px 0;",
    "emergency-box": "background:#f8d7da;border-left:4px solid #dc3545;border-radius:4px;padding:12px 16px;margin:16px 0;",
    "clinical-tip": "background:#d4edda;border-left:4px solid #28a745;border-radius:4px;padding:12px 16px;margin:16px 0;",
    "drug-table": "width:100%;border-collapse:collapse;margin:16px 0;font-size:0.95rem;",
    "compare-table": "width:100%;border-collapse:collapse;margin:16px 0;font-size:0.95rem;",
    "references": "margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;",
    "ref-item": "margin-bottom:12px;padding:10px;background:#f8f9fa;border-radius:4px;font-size:0.9rem;",
    "ref-numbers": "color:#dc3545;font-weight:bold;margin-right:8px;font-size:0.85em;",
    "ref-content": "color:#555;",
    "medical-details": "background:#f8f9fa;margin:16px 0;border-radius:4px;border:1px solid #dee2e6;",
    "medical-details-content": "padding:16px;background:white;border-radius:0 0 4px 4px;",
}


def convert_to_inline(html):
    # 1. Remove DOCTYPE, <html>, <head>...</head>, <body> wrappers
    html = re.sub(r'<!DOCTYPE[^>]*>', '', html, flags=re.IGNORECASE)
    html = re.sub(r'<html[^>]*>', '', html, flags=re.IGNORECASE)
    html = re.sub(r'</html>', '', html, flags=re.IGNORECASE)
    html = re.sub(r'<head>.*?</head>', '', html, flags=re.IGNORECASE | re.DOTALL)
    html = re.sub(r'</?body>', '', html, flags=re.IGNORECASE)

    # 2. Remove <style>...</style> block
    html = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.IGNORECASE | re.DOTALL)

    # 3. Convert class="xxx" to style="..."
    def replace_class(m):
        tag_start = m.group(1)
        classes = m.group(2).split()
        existing_style = ""
        # Check if there's already a style attribute
        style_match = re.search(r'style="([^"]*)"', tag_start)
        if style_match:
            existing_style = style_match.group(1).rstrip(";") + ";"
            tag_start = re.sub(r'\s*style="[^"]*"', '', tag_start)

        inline = existing_style
        for cls in classes:
            if cls in INLINE_STYLES:
                inline += INLINE_STYLES[cls]

        if inline:
            return '{} style="{}"'.format(tag_start, inline)
        else:
            return tag_start

    html = re.sub(r'(<\w+[^>]*?)\s+class="([^"]*)"', replace_class, html)

    # 4. Add inline styles to h1, h2, h3, strong, sup, th/td
    html = re.sub(
        r'<h1(?![^>]*style)([^>]*)>',
        r'<h1 style="font-size:1.6rem;color:#2c5aa0;border-bottom:2px solid #2c5aa0;padding-bottom:8px;margin-top:0;"\1>',
        html
    )
    html = re.sub(
        r'<h2(?![^>]*style)([^>]*)>',
        r'<h2 style="font-size:1.3rem;color:#2c5aa0;border-left:4px solid #2c5aa0;padding-left:12px;margin-top:1.5rem;"\1>',
        html
    )
    html = re.sub(
        r'<h3(?![^>]*style)([^>]*)>',
        lambda m: '<h3 style="font-size:1.1rem;color:#444;margin-top:1rem;"{}>{}'.format(
            m.group(1), '&#9654; '
        ) if '::before' not in m.group(0) else m.group(0),
        html
    )
    # Actually, h3::before adds ▶, so we need to prepend it to h3 content
    # Replace <h3 ...>text with <h3 ...>▶ text (for h3 without ref class)
    html = re.sub(
        r'(<h3[^>]*>)(?!参考文献)',
        lambda m: m.group(1) + '▶ ' if 'references' not in m.group(0) else m.group(1),
        html
    )
    # Remove duplicate ▶ if already present
    html = html.replace('▶ ▶', '▶')

    # th inline styles - ensure they have the right background
    html = re.sub(
        r'<th(?![^>]*style)([^>]*)>',
        r'<th style="background:#2c5aa0;color:white;font-weight:bold;border:1px solid #ddd;padding:12px;text-align:left;"\1>',
        html
    )
    # td inline styles
    html = re.sub(
        r'<td(?![^>]*style)([^>]*)>',
        r'<td style="border:1px solid #ddd;padding:12px;text-align:left;"\1>',
        html
    )

    # strong color
    html = re.sub(
        r'<strong(?![^>]*style)>',
        r'<strong style="font-weight:bold;color:#2c5aa0;">',
        html
    )

    # sup color
    html = re.sub(
        r'<sup(?![^>]*style)>',
        r'<sup style="color:#dc3545;font-weight:bold;font-size:0.75em;">',
        html
    )

    # 5. Remove outer wrapper div if it just has the medical-article style
    # Keep content inside
    html = re.sub(r'^\s*<div[^>]*style="[^"]*max-width:100%[^"]*"[^>]*>\s*', '', html.strip())
    # Remove matching closing </div> at the very end
    html = re.sub(r'\s*</div>\s*$', '', html.strip())

    # 6. Clean up whitespace
    html = re.sub(r'\n{3,}', '\n\n', html)

    return html.strip()


def build_atom_entry(title, content, categories, draft=True):
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join(
        '  <category term="{}" />'.format(c) for c in categories
    )
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
        title=title, content=content, categories=cat_xml, draft=draft_val
    )
    return xml


def post_entry(env, xml_body):
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
    auth_str = "{}:{}".format(hatena_id, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    data = xml_body.encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        method="POST",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth_b64),
        },
    )

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            entry_url = url_match.group(1) if url_match else "(URL取得失敗)"
            entry_id_match = re.search(r'<id>tag:[^,]+,\d+:entry-(\d+)</id>', body)
            entry_id = entry_id_match.group(1) if entry_id_match else "(ID取得失敗)"
            return True, entry_url, entry_id
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        return False, "HTTP {}: {}".format(e.code, error_body[:500]), None


def main():
    print("=" * 60)
    print("低活動性せん妄 ブログ記事 下書き投稿")
    print("=" * 60)

    # Read HTML
    with open(HTML_FILE, encoding="utf-8") as f:
        raw_html = f.read()

    # Convert to inline styles
    print("[1/3] <style>除去・インラインスタイル変換...")
    content = convert_to_inline(raw_html)
    print("  変換後サイズ: {} bytes".format(len(content)))

    # Save converted version for reference
    out_path = HTML_FILE.parent / "hypoactive_delirium_blog_hatena.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(content)
    print("  変換済みHTML保存: {}".format(out_path.name))

    # Post to Hatena Blog
    print("[2/3] はてなブログに下書き投稿中...")
    print("  タイトル: {}".format(TITLE))
    print("  カテゴリ: {}".format(", ".join(CATEGORIES)))

    env = load_env()
    xml = build_atom_entry(TITLE, content, CATEGORIES, draft=True)
    ok, url_or_err, entry_id = post_entry(env, xml)

    if ok:
        print("[3/3] 投稿成功!")
        print("  URL: {}".format(url_or_err))
        print("  Entry ID: {}".format(entry_id))
    else:
        print("[失敗] {}".format(url_or_err))
        sys.exit(1)


if __name__ == "__main__":
    main()
