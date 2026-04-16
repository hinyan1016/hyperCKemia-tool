#!/usr/bin/env python3
"""
からだの不思議 #16 ブログ更新スクリプト（全内容再構築版）
- ローカルHTMLからインラインスタイル変換
- YouTube動画リンク + VideoObject JSON-LD追加
- FAQ構造化データ + 関連記事追加
- はてなAtomPub APIでPUT更新
"""

import sys
import re
import urllib.request
import urllib.error
import base64
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")

ENTRY_ID = "17179246901372000539"
VIDEO_ID = "E75RZdF4pd4"
VIDEO_TITLE = "けいれんを目撃したときの正しい対応は？｜脳神経内科医がやってはいけないことと救急車を呼ぶ判断を解説【からだの不思議 #16】"
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


def convert_to_inline():
    """ローカルHTMLからインラインスタイルに変換"""
    html_path = Path(__file__).parent / "16_seizure_first_aid_blog.html"
    with open(html_path, encoding="utf-8") as f:
        raw = f.read()

    # <body>内の<div class="medical-article">...</div>を抽出
    m = re.search(r'<div class="medical-article">(.*?)</div>\s*</body>', raw, re.DOTALL)
    if not m:
        print("ERROR: medical-article div not found")
        sys.exit(1)
    content = m.group(1)

    # <style>...</style>ブロック除去
    content = re.sub(r'<style>.*?</style>', '', content, flags=re.DOTALL)

    # --- インラインスタイル変換 ---
    # series-label
    content = content.replace(
        '<span class="series-label">',
        '<span style="display:inline-block;background:#2c5aa0;color:#fff;font-size:0.85rem;padding:2px 12px;border-radius:4px;margin-bottom:8px;">'
    )

    # h2 (with optional id) — はてなはh2のstyleを保持する
    content = re.sub(
        r'<h2( id="[^"]*")?>',
        lambda m: '<h2{} style="font-size:1.3rem;color:#2c5aa0;border-left:4px solid #2c5aa0;padding-left:12px;margin-top:2rem;">'.format(m.group(1) or ''),
        content
    )

    # h3 — ▶ をテキストとして挿入（はてながstyleを除去しても▶は残る）
    content = re.sub(
        r'<h3>',
        '<h3>▶ ',
        content
    )

    # toc-box
    content = content.replace(
        '<div class="toc-box">',
        '<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;">'
    )
    # toc-box内のh3は▶不要
    content = re.sub(
        r'(<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;">\s*)<h3>▶ ',
        r'\1<h3>',
        content
    )

    # important-box
    content = content.replace(
        '<div class="important-box">',
        '<div style="background:#e8f4fd;border-left:4px solid #2c5aa0;border-radius:4px;padding:12px 16px;margin:16px 0;">'
    )

    # warning-box
    content = content.replace(
        '<div class="warning-box">',
        '<div style="background:#fff3cd;border-left:4px solid #ffc107;border-radius:4px;padding:12px 16px;margin:16px 0;">'
    )

    # emergency-box
    content = content.replace(
        '<div class="emergency-box">',
        '<div style="background:#f8d7da;border-left:4px solid #dc3545;border-radius:4px;padding:12px 16px;margin:16px 0;">'
    )

    # comparison-table
    content = content.replace(
        '<table class="comparison-table">',
        '<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:0.95rem;">'
    )
    # th — インラインスタイル（既存のstyle属性とマージ）
    content = re.sub(
        r'<th style="([^"]*)">',
        lambda m: '<th style="background:#2c5aa0;color:#fff;padding:10px 12px;text-align:left;font-weight:bold;{}">'.format(m.group(1)),
        content
    )
    # td
    content = re.sub(
        r'<td>',
        '<td style="padding:10px 12px;border-bottom:1px solid #dee2e6;vertical-align:top;">',
        content
    )

    # checklist-ok
    content = content.replace(
        '<div class="checklist-ok">',
        '<div style="background:#eaf6ea;border:1px solid #b2dfdb;border-radius:4px;padding:16px;margin:16px 0;">'
    )

    # references
    content = content.replace(
        '<div class="references">',
        '<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">'
    )

    # ref-item
    content = content.replace(
        '<div class="ref-item">',
        '<div style="margin-bottom:12px;padding:10px;background:#f8f9fa;border-radius:4px;font-size:0.9rem;">'
    )

    # ref-numbers
    content = content.replace(
        '<span class="ref-numbers">',
        '<span style="color:#dc3545;font-weight:bold;margin-right:8px;font-size:0.85em;">'
    )

    # ref-content
    content = content.replace(
        '<span class="ref-content">',
        '<span style="color:#555;">'
    )

    # sup
    content = re.sub(
        r'<sup>',
        '<sup style="color:#dc3545;font-weight:bold;font-size:0.75em;">',
        content
    )

    # strong — color追加（既にstyle付きは除外）
    content = re.sub(
        r'<strong>(?!<)',
        '<strong style="color:#2c5aa0;">',
        content
    )

    content = content.strip()
    return content


def main():
    categories = ["医学教育", "からだの不思議", "てんかん", "救急"]

    print("=== からだの不思議 #16 ブログ全内容再構築 ===\n")

    env = load_env()

    # 1. ローカルHTMLからインラインスタイル変換
    print("[1/3] ローカルHTMLからインラインスタイル変換...")
    content = convert_to_inline()
    print("  変換完了（{}文字）".format(len(content)))

    # 2. YouTube動画リンクをh1直後に挿入
    print("[2/3] YouTube・関連記事・構造化データ追加...")

    youtube_embed = '<p style="text-align:center;margin:20px 0;"><a href="https://youtu.be/{vid}" target="_blank" rel="noopener" style="display:inline-block;background:#c0392b;color:#fff;padding:12px 24px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:1.05rem;">&#9654; YouTube動画で見る</a></p>'.format(vid=VIDEO_ID)

    # h1の後に挿入（はてながh1のstyleを除去するのでstyle無しで出す）
    content = re.sub(r'(</h1>)', r'\1\n\n' + youtube_embed, content, count=1)

    # 関連記事（参考文献divの前に挿入）
    related_html = """
<div style="background:#eaf6ea;border-left:4px solid #27ae60;border-radius:4px;padding:12px 16px;margin:20px 0;">
<p style="font-size:1.1rem;color:#444;margin-top:0;"><strong style="color:#27ae60;">&#128279; 関連記事</strong></p>
<ul>
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/03/214301">手がしびれて目が覚めた ― 6つの原因と対処法【からだの不思議 #03】</a></li>
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/05/081707">機能性神経障害（FND）の真実 ― けいれんに似た発作の正体【からだの不思議 #05】</a></li>
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/03/23/120000">しびれと受診科 ― どの診療科を受診すべき？【からだの不思議 #14】</a></li>
</ul>
</div>"""

    content = content.replace(
        '<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">',
        related_html + '\n\n<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">',
        1
    )

    # FAQ構造化データ（末尾）
    faq_jsonld = '<script type="application/ld+json">{"@context":"https://schema.org","@type":"FAQPage","mainEntity":[{"@type":"Question","name":"けいれん発作を目撃したら最初に何をすべきですか？","acceptedAnswer":{"@type":"Answer","text":"最も重要なのは時間を確認することです。発作が始まった時刻を記録し、周囲の危険物を除き、頭を守り、横向き（回復体位）にしてください。5分以上続く場合は119番通報してください。"}},{"@type":"Question","name":"けいれん中に口にタオルや指を入れてもいいですか？","acceptedAnswer":{"@type":"Answer","text":"絶対にやめてください。舌を飲み込むことは医学的にありえません。口に物を入れると窒息・歯の骨折・介助者の咬傷などを引き起こす最も危険な行為です。"}},{"@type":"Question","name":"けいれんが何分続いたら救急車を呼ぶべきですか？","acceptedAnswer":{"@type":"Answer","text":"5分以上続く場合は119番通報してください。国際てんかん連盟（ILAE）のガイドラインでは5分以上の発作をてんかん重積状態として早急な治療が必要としています。"}}]}</script>'

    # VideoObject JSON-LD
    video_jsonld = '<script type="application/ld+json">{{"@context":"https://schema.org","@type":"VideoObject","name":"{}","description":"けいれん発作を目撃したときの正しい対応・やってはいけないこと・救急車を呼ぶ判断を脳神経内科医が解説","thumbnailUrl":"https://img.youtube.com/vi/{}/maxresdefault.jpg","uploadDate":"2026-04-15","contentUrl":"https://www.youtube.com/watch?v={}","embedUrl":"https://www.youtube.com/embed/{}"}}</script>'.format(
        VIDEO_TITLE.replace('"', '\\"'), VIDEO_ID, VIDEO_ID, VIDEO_ID
    )

    content = content + "\n\n" + faq_jsonld + "\n" + video_jsonld

    print("  YouTube動画リンク: OK")
    print("  関連記事3件: OK")
    print("  FAQ JSON-LD: OK")
    print("  VideoObject JSON-LD: OK")

    # 残存classチェック
    classes = re.findall(r'class="[^"]*"', content)
    if classes:
        print("  WARNING: remaining classes: {}".format(classes))
    else:
        print("  classチェック: OK（残存なし）")

    # 3. PUT更新
    print("\n[3/3] はてなブログ更新中...")

    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)

    xml = """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>yes</app:draft>
  </app:control>
</entry>""".format(title=BLOG_TITLE, content=content, categories=cat_xml)

    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(hatena_id, blog_domain, ENTRY_ID)
    auth_str = "{}:{}".format(hatena_id, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    data = xml.encode("utf-8")
    req = urllib.request.Request(url, data=data, method="PUT", headers={
        "Content-Type": "application/xml; charset=utf-8",
        "Authorization": "Basic {}".format(auth_b64),
    })

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            entry_url = url_match.group(1) if url_match else "(URL取得失敗)"
            print("\n[成功] ブログ更新完了")
            print("  URL: {}".format(entry_url))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        print("\n[失敗] HTTP {}: {}".format(e.code, error_body[:500]))
        sys.exit(1)


if __name__ == "__main__":
    main()
