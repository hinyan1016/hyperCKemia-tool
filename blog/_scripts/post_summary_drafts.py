#!/usr/bin/env python3
"""
まとめ記事 下書き投稿スクリプト
CSS変数・クラスをインラインスタイルに変換して、はてなブログにAtomPub APIで下書き投稿する

使い方:
  python post_summary_drafts.py                  # 両方を下書き投稿
  python post_summary_drafts.py blog              # ブログまとめのみ
  python post_summary_drafts.py youtube           # YouTubeまとめのみ
"""

import os
import sys
import re
import base64
import urllib.request
import urllib.error
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")

# CSS変数の解決マップ
CSS_VARS = {
    "--c-navy": "#0f2b44",
    "--c-blue": "#1e6ba5",
    "--c-blue-light": "#e9f2fb",
    "--c-accent": "#e8850c",
    "--c-accent-light": "#fff6eb",
    "--c-green": "#0d7c5f",
    "--c-green-light": "#e6f7f2",
    "--c-red": "#c0392b",
    "--c-red-light": "#fdecea",
    "--c-purple": "#6c3483",
    "--c-purple-light": "#f3e8fa",
    "--c-gray": "#6b7280",
    "--c-bg": "#f7f8fa",
    "--c-card": "#ffffff",
    "--c-border": "#e5e7eb",
    "--radius": "12px",
    "--shadow-sm": "0 1px 3px rgba(0,0,0,0.06)",
    "--shadow-md": "0 4px 12px rgba(0,0,0,0.08)",
    # YouTube版
    "--yt-red": "#c4302b",
    "--yt-red-dark": "#9a1c17",
    "--yt-red-light": "#fef2f1",
    "--yt-navy": "#0f2b44",
    "--yt-blue": "#1e6ba5",
    "--yt-blue-light": "#e9f2fb",
    "--yt-green": "#0d7c5f",
    "--yt-green-light": "#e6f7f2",
    "--yt-gray": "#6b7280",
    "--yt-bg": "#f7f8fa",
    "--yt-card": "#ffffff",
    "--yt-border": "#e5e7eb",
}

def resolve_var(val):
    """CSS変数 var(--xxx) を実際の値に置換"""
    def replacer(m):
        var_name = m.group(1).strip()
        return CSS_VARS.get(var_name, m.group(0))
    return re.sub(r'var\(([^)]+)\)', replacer, val)


# ========== ブログまとめ記事のCSSクラス→インラインスタイル ==========
BLOG_CLASS_MAP = {
    "hub": "max-width:100%;margin:0 auto;font-family:-apple-system,BlinkMacSystemFont,'Hiragino Kaku Gothic ProN','Hiragino Sans',Meiryo,sans-serif;line-height:1.8;color:#1f2937;background:#f7f8fa;",
    "hub-hero": "background:linear-gradient(135deg,#0f2b44 0%,#1e6ba5 100%);color:#fff;padding:40px 24px 32px;border-radius:0 0 24px 24px;text-align:center;position:relative;overflow:hidden;",
    "hub-hero-date": "display:inline-block;background:rgba(255,255,255,0.15);padding:4px 16px;border-radius:20px;font-size:0.82em;margin-bottom:12px;letter-spacing:0.02em;",
    "hub-hero-sub": "font-size:0.92em;opacity:0.85;margin:0 0 16px;",
    "hub-hero-stats": "display:flex;justify-content:center;gap:24px;margin-top:16px;",
    "hub-hero-stat-num": "display:block;font-size:1.6em;font-weight:800;line-height:1.2;",
    "hub-hero-stat-label": "font-size:0.75em;opacity:0.7;",
    "hub-author": "display:flex;align-items:center;gap:14px;background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:16px 20px;margin:-20px 16px 24px;position:relative;z-index:1;box-shadow:0 4px 12px rgba(0,0,0,0.08);",
    "hub-author-icon": "font-size:2.2em;flex-shrink:0;",
    "hub-author-text": "flex:1;font-size:0.88em;color:#6b7280;line-height:1.6;",
    "hub-toc": "background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:20px 24px;margin:0 16px 28px;box-shadow:0 1px 3px rgba(0,0,0,0.06);",
    "hub-toc-title": "font-size:0.9em;font-weight:700;color:#0f2b44;margin:0 0 12px;display:flex;align-items:center;gap:6px;",
    "hub-toc-list": "display:flex;flex-wrap:wrap;gap:8px;list-style:none;padding:0;margin:0;",
    "hub-toc-count": "background:#1e6ba5;color:#fff;font-size:0.72em;padding:1px 7px;border-radius:10px;font-weight:700;",
    "hub-lead": "background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:20px 24px;margin:0 16px 24px;font-size:0.92em;color:#374151;line-height:1.9;box-shadow:0 1px 3px rgba(0,0,0,0.06);",
    "hub-section": "margin:0 16px 32px;",
    "hub-section-header": "display:flex;align-items:center;gap:10px;margin-bottom:6px;",
    "hub-section-icon": "width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.1em;flex-shrink:0;",
    "hub-section-desc": "font-size:0.88em;color:#6b7280;margin:0 0 16px;padding-left:46px;",
    "hub-card": "display:flex;gap:14px;padding:14px 16px;background:#fff;border:1px solid #e5e7eb;border-radius:12px;margin-bottom:10px;text-decoration:none;color:inherit;",
    "hub-card-icon": "font-size:1.4em;flex-shrink:0;width:36px;text-align:center;padding-top:2px;",
    "hub-card-body": "flex:1;",
    "hub-card-title": "font-size:0.95em;font-weight:700;color:#0f2b44;margin:0 0 4px;line-height:1.5;",
    "hub-card-excerpt": "font-size:0.84em;color:#6b7280;margin:0;line-height:1.6;",
    "hub-badges": "display:flex;flex-wrap:wrap;gap:5px;margin-bottom:5px;",
    "hub-badge": "display:inline-block;font-size:0.68em;font-weight:700;padding:2px 9px;border-radius:6px;letter-spacing:0.02em;",
    "hub-badge-new": "background:#fdecea;color:#c0392b;",
    "hub-badge-ddx": "background:#e9f2fb;color:#1e6ba5;",
    "hub-badge-neuro": "background:#eef2f7;color:#3b6b9e;",
    "hub-badge-general": "background:#e6f7f2;color:#0d7c5f;",
    "hub-badge-treatment": "background:#f3e8fa;color:#6c3483;",
    "hub-badge-series": "background:#fff6eb;color:#e8850c;",
    "hub-badge-tool": "background:#fce4ec;color:#c62828;",
    "hub-badge-date": "background:transparent;color:#6b7280;font-weight:500;padding-left:0;",
    "hub-badge-pro": "background:#fef9e7;color:#b7950b;",
    "hub-badge-public": "background:#e6f7f2;color:#0d7c5f;",
    "hub-series-banner": "border-radius:12px;padding:20px;margin-bottom:16px;display:flex;align-items:center;gap:16px;",
    "hub-series-banner-icon": "font-size:2.4em;flex-shrink:0;",
    "hub-progress": "display:flex;align-items:center;gap:8px;font-size:0.78em;color:#6b7280;margin-bottom:12px;",
    "hub-progress-bar": "flex:1;height:6px;background:#e5e7eb;border-radius:3px;overflow:hidden;",
    "hub-progress-fill": "height:100%;border-radius:3px;",
    "hub-grid": "display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px;margin-bottom:16px;",
    "hub-grid-card": "background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:16px;",
    "hub-grid-card-num": "display:inline-block;width:26px;height:26px;line-height:26px;text-align:center;border-radius:8px;font-size:0.78em;font-weight:800;margin-bottom:8px;",
    "hub-btn": "display:inline-block;padding:7px 20px;border-radius:8px;font-size:0.84em;font-weight:700;text-decoration:none;color:#fff;",
    "hub-link": "font-size:0.82em;font-weight:700;text-decoration:none;",
    "hub-sep": "border:none;border-top:1px solid #e5e7eb;margin:32px 16px;",
    "hub-faq": "background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:24px;margin:0 16px 24px;",
    "hub-faq-q": "font-weight:700;color:#0f2b44;font-size:0.95em;margin:0 0 4px;",
    "hub-faq-a": "font-size:0.88em;color:#6b7280;margin:0;padding-left:22px;line-height:1.7;",
    "hub-footer": "background:#fff;border-top:1px solid #e5e7eb;padding:20px 24px;margin-top:16px;border-radius:0 0 12px 12px;font-size:0.85em;color:#6b7280;line-height:1.8;",
}

# ========== YouTubeまとめ記事のCSSクラス→インラインスタイル ==========
YT_CLASS_MAP = {
    "yt": "max-width:100%;margin:0 auto;font-family:-apple-system,BlinkMacSystemFont,'Hiragino Kaku Gothic ProN','Hiragino Sans',Meiryo,sans-serif;line-height:1.8;color:#1f2937;background:#f7f8fa;",
    "yt-hero": "background:linear-gradient(135deg,#9a1c17 0%,#c4302b 60%,#e74c3c 100%);color:#fff;padding:40px 24px 32px;border-radius:0 0 24px 24px;text-align:center;position:relative;overflow:hidden;",
    "yt-hero-date": "display:inline-block;background:rgba(255,255,255,0.15);padding:4px 16px;border-radius:20px;font-size:0.82em;margin-bottom:12px;",
    "yt-hero-sub": "font-size:0.92em;opacity:0.85;margin:0 0 16px;",
    "yt-hero-stats": "display:flex;justify-content:center;gap:24px;margin-top:16px;",
    "yt-hero-stat-num": "display:block;font-size:1.6em;font-weight:800;line-height:1.2;",
    "yt-hero-stat-label": "font-size:0.75em;opacity:0.7;",
    "yt-cta": "display:flex;justify-content:center;margin:-18px 16px 24px;position:relative;z-index:1;",
    "yt-author": "display:flex;align-items:center;gap:14px;background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:16px 20px;margin:0 16px 24px;box-shadow:0 1px 3px rgba(0,0,0,0.06);",
    "yt-author-icon": "font-size:2.2em;flex-shrink:0;",
    "yt-author-text": "flex:1;font-size:0.88em;color:#6b7280;line-height:1.6;",
    "yt-toc": "background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:20px 24px;margin:0 16px 28px;box-shadow:0 1px 3px rgba(0,0,0,0.06);",
    "yt-toc-title": "font-size:0.9em;font-weight:700;color:#0f2b44;margin:0 0 12px;",
    "yt-toc-list": "display:flex;flex-wrap:wrap;gap:8px;",
    "yt-toc-count": "background:#c4302b;color:#fff;font-size:0.72em;padding:1px 7px;border-radius:10px;font-weight:700;",
    "yt-section": "margin:0 16px 32px;",
    "yt-section-header": "display:flex;align-items:center;gap:10px;margin-bottom:6px;",
    "yt-section-icon": "width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.1em;flex-shrink:0;",
    "yt-section-desc": "font-size:0.88em;color:#6b7280;margin:0 0 16px;padding-left:46px;",
    "yt-vcard": "display:flex;gap:14px;padding:12px;background:#fff;border:1px solid #e5e7eb;border-radius:12px;margin-bottom:10px;",
    "yt-vcard-thumb": "flex-shrink:0;width:180px;position:relative;border-radius:8px;overflow:hidden;",
    "yt-vcard-info": "flex:1;",
    "yt-vcard-meta": "display:flex;flex-wrap:wrap;gap:5px;margin-bottom:5px;",
    "yt-vcard-desc": "font-size:0.84em;color:#6b7280;margin:0;line-height:1.6;",
    "yt-badge": "display:inline-block;font-size:0.68em;font-weight:700;padding:2px 9px;border-radius:6px;",
    "yt-badge-new": "background:#fef2f1;color:#c4302b;",
    "yt-badge-series": "background:#e6f7f2;color:#0d7c5f;",
    "yt-badge-ddx": "background:#e9f2fb;color:#1e6ba5;",
    "yt-badge-neuro": "background:#eef2f7;color:#3b6b9e;",
    "yt-badge-general": "background:#e6f7f2;color:#0d7c5f;",
    "yt-badge-finance": "background:#fff3e0;color:#e65100;",
    "yt-progress": "display:flex;align-items:center;gap:8px;font-size:0.78em;color:#6b7280;margin-bottom:12px;",
    "yt-progress-bar": "flex:1;height:6px;background:#e5e7eb;border-radius:3px;overflow:hidden;",
    "yt-progress-fill": "height:100%;border-radius:3px;",
    "yt-sep": "border:none;border-top:1px solid #e5e7eb;margin:32px 16px;",
    "yt-footer": "background:#fff;border-top:1px solid #e5e7eb;padding:20px 24px;margin-top:16px;font-size:0.85em;color:#6b7280;line-height:1.8;",
}


def convert_to_inline(html_content, class_map):
    """CSSクラスをインラインスタイルに変換する"""

    # <style>ブロックを除去
    html_content = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL)

    # <!DOCTYPE>, <html>, <head>, <body> タグを除去
    html_content = re.sub(r'<!DOCTYPE[^>]*>', '', html_content)
    html_content = re.sub(r'<html[^>]*>', '', html_content)
    html_content = re.sub(r'</html>', '', html_content)
    html_content = re.sub(r'<head>.*?</head>', '', html_content, flags=re.DOTALL)
    html_content = re.sub(r'<body>', '', html_content)
    html_content = re.sub(r'</body>', '', html_content)
    # JSON-LDスクリプトも除去（はてなブログではサポートされない）
    html_content = re.sub(r'<script type="application/ld\+json">.*?</script>', '', html_content, flags=re.DOTALL)
    # トップへ戻るボタンのscriptも除去
    html_content = re.sub(r'<button[^>]*id="(hubToTop|ytToTop)"[^>]*>.*?</button>', '', html_content, flags=re.DOTALL)
    html_content = re.sub(r'<script>\s*\(function\(\)\{.*?\}\)\(\);\s*</script>', '', html_content, flags=re.DOTALL)

    def replace_classes(m):
        full_tag = m.group(0)
        tag_name = m.group(1)
        attrs = m.group(2)

        # class属性を抽出
        class_match = re.search(r'class="([^"]*)"', attrs)
        if not class_match:
            return full_tag

        classes = class_match.group(1).split()
        existing_style = ""
        style_match = re.search(r'style="([^"]*)"', attrs)
        if style_match:
            existing_style = style_match.group(1).rstrip(";") + ";"

        # クラスに対応するインラインスタイルを組み立て
        inline_styles = existing_style
        for cls in classes:
            if cls in class_map:
                inline_styles += class_map[cls]

        # CSS変数を解決
        inline_styles = resolve_var(inline_styles)

        # 既存の style="" と class="" を除去して、新しい style="" に置換
        new_attrs = re.sub(r'\s*class="[^"]*"', '', attrs)
        new_attrs = re.sub(r'\s*style="[^"]*"', '', new_attrs)
        if inline_styles:
            new_attrs = new_attrs.rstrip() + ' style="' + inline_styles + '"'

        return "<{}{}>".format(tag_name, new_attrs)

    # 全タグの class 属性をインラインスタイルに変換
    html_content = re.sub(r'<(\w+)((?:\s+[^>]*)?)>', replace_classes, html_content)

    # h1タグのスタイル追加（hero内のh1）
    # h2/h3/h4 のデフォルトスタイル
    html_content = html_content.strip()

    return html_content


def extract_title_from_meta(html_content):
    """<title>タグからタイトルを抽出"""
    m = re.search(r'<title>(.*?)</title>', html_content)
    if m:
        return m.group(1)
    return None


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


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


def post_article(html_path, class_map, categories, env):
    """HTMLファイルを読み込み、変換して下書き投稿"""
    with open(html_path, encoding="utf-8") as f:
        raw_html = f.read()

    title = extract_title_from_meta(raw_html)
    if not title:
        print("  [エラー] タイトルが見つかりません: {}".format(html_path))
        return False

    print("  タイトル: {}".format(title))
    print("  変換中...")

    content = convert_to_inline(raw_html, class_map)

    # <a> タグの中のスタイル（色・太字）を追加
    content = re.sub(
        r'(<a\s)',
        r'\1style="color:#1e6ba5;font-weight:600;text-decoration:none;" ',
        content
    )

    print("  変換後サイズ: {:.1f} KB".format(len(content.encode("utf-8")) / 1024))
    print("  下書き投稿中...")

    xml = build_atom_entry(title, content, categories, draft=True)
    ok, result = post_entry(env, xml)

    if ok:
        print("  [成功] {}".format(result))
        return True
    else:
        print("  [失敗] {}".format(result))
        return False


def main():
    args = sys.argv[1:]
    targets = args if args else ["blog", "youtube"]

    env = load_env()
    success = 0

    for target in targets:
        if target == "blog":
            print("\n===== ブログまとめ記事 =====")
            html_path = SCRIPT_DIR / "featured_articles_april9.html"
            categories = ["まとめ記事", "脳神経内科", "鑑別診断", "食事指導"]
            if post_article(html_path, BLOG_CLASS_MAP, categories, env):
                success += 1

        elif target == "youtube":
            print("\n===== YouTubeまとめ記事 =====")
            html_path = SCRIPT_DIR / "youtube_featured_april9.html"
            categories = ["まとめ記事", "YouTube", "脳神経内科"]
            if post_article(html_path, YT_CLASS_MAP, categories, env):
                success += 1

    print("\n完了: {}/{} 件成功".format(success, len(targets)))


if __name__ == "__main__":
    main()
