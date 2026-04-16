#!/usr/bin/env python3
"""
rt-PA血栓溶解療法ブログ記事: CSSクラス→インラインスタイル変換 + はてなブログ下書き投稿
"""
import re
import sys
import os
import urllib.request
import urllib.error
import base64
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\rtpa_thrombolysis_mnemonic_blog.html")
TITLE = "rt-PA静注療法の適応と禁忌を覚えるコツ ― 「時・血・画・出・偽」で整理する"
CATEGORIES = ["医学教育", "脳神経内科", "脳卒中"]


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
    """CSSクラスをインラインスタイルに変換し、<style>ブロックを除去"""

    # 1. <div class="medical-article">の中身だけ取得
    m = re.search(r'<div class="medical-article">\s*<style>.*?</style>\s*(.*?)</div>\s*<!--\s*/\.medical-article\s*-->', html, re.DOTALL)
    if not m:
        # fallback: styleブロック除去だけ
        content = re.sub(r'<style>.*?</style>', '', html, flags=re.DOTALL)
        content = re.sub(r'<!DOCTYPE.*?<body>\s*', '', content, flags=re.DOTALL | re.IGNORECASE)
        content = re.sub(r'</body>\s*</html>\s*$', '', content, flags=re.IGNORECASE)
        return content.strip()

    content = m.group(1).strip()

    # 2. h1 → inline style
    content = re.sub(
        r'<h1([^>]*)>',
        r'<h1\1 style="font-size:1.6rem;color:#2c5aa0;border-bottom:2px solid #2c5aa0;padding-bottom:8px;margin-top:0;">',
        content
    )

    # 3. h2 → inline style
    content = re.sub(
        r'<h2([^>]*)>',
        r'<h2\1 style="font-size:1.3rem;color:#2c5aa0;border-left:4px solid #2c5aa0;padding-left:12px;margin-top:1.5rem;">',
        content
    )

    # 4. h3 → inline style (add ▶ prefix manually since ::before won't work)
    def replace_h3(m):
        attrs = m.group(1)
        return '<h3{} style="font-size:1.1rem;color:#444;margin-top:1rem;">▶ '.format(attrs)
    content = re.sub(r'<h3([^>]*)>', replace_h3, content)

    # 5. strong → inline
    content = re.sub(
        r'<strong>',
        '<strong style="font-weight:bold;color:#2c5aa0;">',
        content
    )

    # 6. sup → inline
    content = re.sub(
        r'<sup>',
        '<sup style="color:#dc3545;font-weight:bold;font-size:0.75em;">',
        content
    )

    # 7. div class replacements
    class_to_style = {
        'toc-box': 'background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;',
        'important-box': 'background:#e8f4fd;border-left:4px solid #2c5aa0;border-radius:4px;padding:12px 16px;margin:16px 0;',
        'warning-box': 'background:#fff3cd;border-left:4px solid #ffc107;border-radius:4px;padding:12px 16px;margin:16px 0;',
        'emergency-box': 'background:#f8d7da;border-left:4px solid #dc3545;border-radius:4px;padding:12px 16px;margin:16px 0;',
        'references': 'margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;',
        'ref-item': 'margin-bottom:12px;padding:10px;background:#f8f9fa;border-radius:4px;font-size:0.9rem;',
        'flow-step': 'display:flex;align-items:flex-start;margin:8px 0;padding:10px 14px;background:#f8f9fa;border-radius:4px;border:1px solid #dee2e6;',
        'mnemonic-grid': 'display:flex;flex-wrap:wrap;gap:12px;margin:16px 0;',
        'mnemonic-card': 'flex:1 1 140px;background:#2c5aa0;color:white;border-radius:8px;padding:16px 12px;text-align:center;min-width:120px;',
    }
    for cls, style in class_to_style.items():
        content = re.sub(
            r'<div class="{}">'.format(re.escape(cls)),
            '<div style="{}">'.format(style),
            content
        )

    # 8. span class replacements
    span_styles = {
        'ref-numbers': 'color:#dc3545;font-weight:bold;margin-right:8px;font-size:0.85em;',
        'ref-content': 'color:#555;',
        'kanji': 'font-size:2.2rem;font-weight:bold;display:block;',
        'label': 'font-size:0.85rem;margin-top:4px;opacity:0.9;',
        'number-highlight-red': 'display:inline-block;background:#dc3545;color:white;border-radius:4px;padding:1px 8px;font-weight:bold;font-size:0.95em;',
        'number-highlight': 'display:inline-block;background:#2c5aa0;color:white;border-radius:4px;padding:1px 8px;font-weight:bold;font-size:0.95em;',
    }
    for cls, style in span_styles.items():
        content = re.sub(
            r'<span class="{}">'.format(re.escape(cls)),
            '<span style="{}">'.format(style),
            content
        )

    # 9. flow-num and flow-text divs
    content = re.sub(
        r'<div class="flow-num">',
        '<div style="background:#2c5aa0;color:white;border-radius:50%;width:28px;height:28px;display:inline-flex;align-items:center;justify-content:center;font-weight:bold;font-size:0.9rem;flex-shrink:0;margin-right:12px;">',
        content
    )
    content = re.sub(
        r'<div class="flow-text">',
        '<div style="color:#333;flex:1;">',
        content
    )

    # 10. table class
    content = re.sub(
        r'<table class="drug-table">',
        '<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:0.95rem;">',
        content
    )

    # 11. details/summary (medical-details)
    content = re.sub(
        r'<details class="medical-details">',
        '<details style="background:#f8f9fa;margin:16px 0;border-radius:4px;border:1px solid #dee2e6;">',
        content
    )
    content = re.sub(
        r'<summary>',
        '<summary style="background:#2c5aa0;color:white;padding:12px;cursor:pointer;font-weight:bold;border-radius:4px 4px 0 0;">',
        content
    )
    content = re.sub(
        r'<div class="medical-details-content">',
        '<div style="padding:16px;background:white;border-radius:0 0 4px 4px;">',
        content
    )

    # 12. td/th にborder追加（thは既にinline style済み）
    content = re.sub(
        r'<td>',
        '<td style="border:1px solid #ddd;padding:12px;text-align:left;">',
        content
    )

    # 13. 残っているclass属性を除去（安全のため）
    # content = re.sub(r'\s+class="[^"]*"', '', content)

    # 14. 全体をdivで囲む（line-height等の基本スタイル）
    content = '<div style="max-width:100%;padding:16px;line-height:1.7;color:#333;">\n' + content + '\n</div>'

    return content


def build_atom_entry(title, content, categories, draft=True):
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)
    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(title=title, content=content, categories=cat_xml, draft=draft_val)
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
        url, data=data, method="POST",
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
            return True, entry_url
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        return False, "HTTP {}: {}".format(e.code, error_body[:500])


def main():
    # HTML読み込み
    with open(HTML_FILE, encoding="utf-8") as f:
        raw_html = f.read()

    # インラインスタイル変換
    content = convert_to_inline(raw_html)

    print("タイトル: {}".format(TITLE))
    print("カテゴリ: {}".format(", ".join(CATEGORIES)))
    print("モード: 下書き")
    print("変換後サイズ: {} bytes".format(len(content.encode("utf-8"))))
    print("投稿中...")

    env = load_env()
    xml = build_atom_entry(TITLE, content, CATEGORIES, draft=True)
    ok, result = post_entry(env, xml)

    if ok:
        print("[成功] {}".format(result))
    else:
        print("[失敗] {}".format(result))
        sys.exit(1)


if __name__ == "__main__":
    main()
