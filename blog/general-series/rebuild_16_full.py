#!/usr/bin/env python3
"""
からだの不思議 #16 ブログ記事 完全再構築
ローカルHTMLから全要素を一括構築してPUT
- インラインスタイル変換
- YouTube動画リンク
- インフォグラフィック
- 関連記事
- FAQ + VideoObject JSON-LD
"""

import sys
import re
import base64
import urllib.request
import urllib.error
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\general-series\16_seizure_first_aid_blog.html")

ENTRY_ID = "17179246901372000539"
VIDEO_ID = "E75RZdF4pd4"
INFOGRAPHIC_URL = "https://cdn-ak.f.st-hatena.com/images/fotolife/h/hinyan1016/20260415/20260415133838.png"
BLOG_TITLE = "けいれんを目撃したときの正しい対応 ― やってはいけないことと救急車を呼ぶ判断【からだの不思議 #16】"
VIDEO_TITLE = "けいれんを目撃したときの正しい対応は？｜脳神経内科医がやってはいけないことと救急車を呼ぶ判断を解説【からだの不思議 #16】"

sys.stdout.reconfigure(encoding='utf-8')


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def build_content():
    """ローカルHTMLから完全なブログ本文を構築"""
    raw = HTML_FILE.read_text(encoding="utf-8")

    # <body>内の<div class="medical-article">...</div>を抽出
    m = re.search(r'<div class="medical-article">(.*?)</div>\s*</body>', raw, re.DOTALL)
    if not m:
        print("ERROR: medical-article div not found")
        sys.exit(1)
    c = m.group(1)

    # <style>除去
    c = re.sub(r'<style>.*?</style>', '', c, flags=re.DOTALL)

    # === インラインスタイル変換 ===
    c = c.replace('<span class="series-label">',
        '<span style="display:inline-block;background:#2c5aa0;color:#fff;font-size:0.85rem;padding:2px 12px;border-radius:4px;margin-bottom:8px;">')

    c = re.sub(r'<h2( id="[^"]*")?>', lambda m:
        '<h2{} style="font-size:1.3rem;color:#2c5aa0;border-left:4px solid #2c5aa0;padding-left:12px;margin-top:2rem;">'.format(m.group(1) or ''), c)

    # h3 → テキスト▶挿入
    c = re.sub(r'<h3>', '<h3>▶ ', c)

    # toc-box
    c = c.replace('<div class="toc-box">',
        '<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;">')
    c = re.sub(r'(<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;">\s*)<h3>▶ ',
        r'\1<h3>', c)

    c = c.replace('<div class="important-box">',
        '<div style="background:#e8f4fd;border-left:4px solid #2c5aa0;border-radius:4px;padding:12px 16px;margin:16px 0;">')
    c = c.replace('<div class="warning-box">',
        '<div style="background:#fff3cd;border-left:4px solid #ffc107;border-radius:4px;padding:12px 16px;margin:16px 0;">')
    c = c.replace('<div class="emergency-box">',
        '<div style="background:#f8d7da;border-left:4px solid #dc3545;border-radius:4px;padding:12px 16px;margin:16px 0;">')

    c = c.replace('<table class="comparison-table">',
        '<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:0.95rem;">')
    c = re.sub(r'<th style="([^"]*)">',
        lambda m: '<th style="background:#2c5aa0;color:#fff;padding:10px 12px;text-align:left;font-weight:bold;{}">'.format(m.group(1)), c)
    c = re.sub(r'<td>', '<td style="padding:10px 12px;border-bottom:1px solid #dee2e6;vertical-align:top;">', c)

    c = c.replace('<div class="checklist-ok">',
        '<div style="background:#eaf6ea;border:1px solid #b2dfdb;border-radius:4px;padding:16px;margin:16px 0;">')
    c = c.replace('<div class="references">',
        '<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">')
    c = c.replace('<div class="ref-item">',
        '<div style="margin-bottom:12px;padding:10px;background:#f8f9fa;border-radius:4px;font-size:0.9rem;">')
    c = c.replace('<span class="ref-numbers">',
        '<span style="color:#dc3545;font-weight:bold;margin-right:8px;font-size:0.85em;">')
    c = c.replace('<span class="ref-content">', '<span style="color:#555;">')
    c = re.sub(r'<sup>', '<sup style="color:#dc3545;font-weight:bold;font-size:0.75em;">', c)
    c = re.sub(r'<strong>(?!<)', '<strong style="color:#2c5aa0;">', c)

    c = c.strip()

    # === YouTube動画ボタン（h1直後） ===
    yt_btn = ('<p style="text-align:center;margin:20px 0;">'
              '<a href="https://youtu.be/{}" target="_blank" rel="noopener" '
              'style="display:inline-block;background:#c0392b;color:#fff;padding:12px 24px;'
              'border-radius:6px;text-decoration:none;font-weight:bold;font-size:1.05rem;">'
              '&#9654; YouTube動画で見る</a></p>').format(VIDEO_ID)

    c = re.sub(r'(</h1>)', r'\1\n\n' + yt_btn, c, count=1)

    # === インフォグラフィック（YouTubeボタン直後） ===
    infographic = ('<div style="text-align:center;margin:20px 0;">'
                   '<img src="{}" '
                   'alt="けいれんを目撃したときの正しい対応 インフォグラフィック" '
                   'style="max-width:100%;height:auto;border-radius:8px;'
                   'box-shadow:0 2px 8px rgba(0,0,0,0.1);" loading="lazy">'
                   '</div>').format(INFOGRAPHIC_URL)

    c = c.replace(yt_btn, yt_btn + '\n\n' + infographic)

    # === 関連記事（参考文献の前） ===
    related = ('<div style="background:#eaf6ea;border-left:4px solid #27ae60;'
               'border-radius:4px;padding:12px 16px;margin:20px 0;">'
               '<p style="font-size:1.1rem;color:#444;margin-top:0;">'
               '<strong style="color:#27ae60;">&#128279; 関連記事</strong></p>'
               '<ul>'
               '<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/03/214301">'
               '手がしびれて目が覚めた ― 6つの原因と対処法【からだの不思議 #03】</a></li>'
               '<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/05/081707">'
               '機能性神経障害（FND）の真実 ― けいれんに似た発作の正体【からだの不思議 #05】</a></li>'
               '<li><a href="https://hinyan1016.hatenablog.com/entry/2026/03/23/120000">'
               'しびれと受診科 ― どの診療科を受診すべき？【からだの不思議 #14】</a></li>'
               '</ul></div>')

    ref_marker = '<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">'
    c = c.replace(ref_marker, related + '\n\n' + ref_marker, 1)

    # === FAQ構造化データ ===
    faq = ('<script type="application/ld+json">'
           '{"@context":"https://schema.org","@type":"FAQPage","mainEntity":['
           '{"@type":"Question","name":"けいれん発作を目撃したら最初に何をすべきですか？",'
           '"acceptedAnswer":{"@type":"Answer","text":"最も重要なのは時間を確認することです。'
           '発作が始まった時刻を記録し、周囲の危険物を除き、頭を守り、横向き（回復体位）にしてください。'
           '5分以上続く場合は119番通報してください。"}},'
           '{"@type":"Question","name":"けいれん中に口にタオルや指を入れてもいいですか？",'
           '"acceptedAnswer":{"@type":"Answer","text":"絶対にやめてください。'
           '舌を飲み込むことは医学的にありえません。口に物を入れると窒息・歯の骨折・'
           '介助者の咬傷などを引き起こす最も危険な行為です。"}},'
           '{"@type":"Question","name":"けいれんが何分続いたら救急車を呼ぶべきですか？",'
           '"acceptedAnswer":{"@type":"Answer","text":"5分以上続く場合は119番通報してください。'
           '国際てんかん連盟（ILAE）のガイドラインでは5分以上の発作をてんかん重積状態として'
           '早急な治療が必要としています。"}}]}'
           '</script>')

    # === VideoObject JSON-LD ===
    video_ld = ('<script type="application/ld+json">'
                '{{"@context":"https://schema.org","@type":"VideoObject",'
                '"name":"{}",'
                '"description":"けいれん発作を目撃したときの正しい対応・やってはいけないこと・'
                '救急車を呼ぶ判断を脳神経内科医が解説",'
                '"thumbnailUrl":"https://img.youtube.com/vi/{}/maxresdefault.jpg",'
                '"uploadDate":"2026-04-15",'
                '"contentUrl":"https://www.youtube.com/watch?v={}",'
                '"embedUrl":"https://www.youtube.com/embed/{}"}}'
                '</script>').format(VIDEO_TITLE, VIDEO_ID, VIDEO_ID, VIDEO_ID)

    c = c + '\n\n' + faq + '\n' + video_ld

    # 残存classチェック
    classes = re.findall(r'class="[^"]*"', c)
    if classes:
        print("  WARNING: remaining classes: {}".format(classes))

    return c


def main():
    env = load_env()

    print("=== からだの不思議 #16 ブログ完全再構築 ===\n")

    # 1. コンテンツ構築
    print("[1/2] ローカルHTMLから全コンテンツ構築...")
    content = build_content()
    print("  構築完了（{}文字）".format(len(content)))

    # 要素確認
    checks = [
        ('YouTube E75RZdF4pd4', VIDEO_ID in content),
        ('Infographic img', INFOGRAPHIC_URL in content),
        ('VideoObject', 'VideoObject' in content),
        ('FAQPage', 'FAQPage' in content),
        ('Related #03', 'entry/2026/04/03' in content),
        ('Related #05', 'entry/2026/04/05' in content),
        ('Related #14', 'entry/2026/03/23' in content),
        ('h2 style', 'border-left:4px solid #2c5aa0' in content),
        ('emergency-box', 'background:#f8d7da' in content),
        ('DOI links', 'doi.org' in content),
    ]
    all_ok = True
    for name, ok in checks:
        status = 'OK' if ok else 'NG'
        print("  {} {}".format(status, name))
        if not ok:
            all_ok = False

    if not all_ok:
        print("\nERROR: 一部要素が欠落。中断します。")
        sys.exit(1)

    # 2. PUT更新
    print("\n[2/2] はてなブログにPUT更新...")

    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]

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
</entry>""".format(title=BLOG_TITLE, author=hatena_id, content=content, categories=cat_xml)

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(hatena_id, blog_domain, ENTRY_ID)
    auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()

    req = urllib.request.Request(url, data=xml.encode("utf-8"), method="PUT", headers={
        "Content-Type": "application/xml; charset=utf-8",
        "Authorization": "Basic {}".format(auth_b64),
    })

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            entry_url = url_match.group(1) if url_match else "(URL取得失敗)"
            print("\n[成功] 完全再構築＆更新完了")
            print("  URL: {}".format(entry_url))
            print("  編集: https://blog.hatena.ne.jp/hinyan1016/hinyan1016.hatenablog.com/edit?entry={}".format(ENTRY_ID))
            print("\n※ 編集画面で「HTML編集」タブに切り替えると全コンテンツが確認できます")
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        print("\n[失敗] HTTP {}: {}".format(e.code, error_body[:500]))
        sys.exit(1)


if __name__ == "__main__":
    main()
