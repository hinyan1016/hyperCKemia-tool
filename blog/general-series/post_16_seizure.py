#!/usr/bin/env python3
"""
からだの不思議 #16 けいれん対応 ブログ下書き投稿スクリプト
- <style>ブロック除去・インラインスタイル変換
- 内部リンク・FAQ構造化データ追加
- はてなAtomPub APIで下書き投稿
"""

import os
import sys
import re
import urllib.request
import urllib.error
import base64
from pathlib import Path

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


def convert_to_inline():
    """ブログHTMLをインラインスタイルに変換"""

    html_path = Path(__file__).parent / "16_seizure_first_aid_blog.html"
    with open(html_path, encoding="utf-8") as f:
        raw = f.read()

    # <body>内の<div class="medical-article">...</div>を抽出
    m = re.search(r'<div class="medical-article">(.*?)</div>\s*</body>', raw, re.DOTALL)
    if not m:
        print("ERROR: medical-article div not found")
        sys.exit(1)
    content = m.group(1)

    # <style>...</style>ブロックを除去
    content = re.sub(r'<style>.*?</style>', '', content, flags=re.DOTALL)

    # --- インラインスタイル変換 ---

    # series-label
    content = content.replace(
        '<span class="series-label">',
        '<span style="display:inline-block;background:#2c5aa0;color:#fff;font-size:0.85rem;padding:2px 12px;border-radius:4px;margin-bottom:8px;">'
    )

    # h1
    content = re.sub(
        r'<h1>',
        '<h1 style="font-size:1.6rem;color:#2c5aa0;border-bottom:2px solid #2c5aa0;padding-bottom:8px;margin-top:0;">',
        content
    )

    # h2 (with id)
    content = re.sub(
        r'<h2( id="[^"]*")?>',
        lambda m: '<h2{} style="font-size:1.3rem;color:#2c5aa0;border-left:4px solid #2c5aa0;padding-left:12px;margin-top:2rem;">'.format(m.group(1) or ''),
        content
    )

    # h3 — add ▶ prefix inline
    content = re.sub(
        r'<h3>',
        '<h3 style="font-size:1.1rem;color:#444;margin-top:1.2rem;"><span style="color:#2c5aa0;margin-right:8px;">&#9654;</span>',
        content
    )

    # toc-box
    content = content.replace(
        '<div class="toc-box">',
        '<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;">'
    )
    # toc-box内のh3は▶不要なので修正
    content = re.sub(
        r'(<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;padding:16px;margin:16px 0;">\s*)<h3 style="[^"]*"><span style="color:#2c5aa0;margin-right:8px;">&#9654;</span>',
        r'\1<h3 style="font-size:1.1rem;color:#444;margin-top:0;">',
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
    # th にインラインスタイル追加（既存のstyle属性とマージ）
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

    # sup (footnote numbers)
    content = re.sub(
        r'<sup>',
        '<sup style="color:#dc3545;font-weight:bold;font-size:0.75em;">',
        content
    )

    # strong
    # Note: はてなブログではstrongのスタイルはデフォルトで十分なので色だけ追加
    # ただし既にstyle付きのstrongは除外
    content = re.sub(
        r'<strong>(?!<)',
        '<strong style="color:#2c5aa0;">',
        content
    )

    # 末尾の空行を整理
    content = content.strip()

    # --- 内部リンク（関連記事）追加 ---
    related_html = """
<div style="background:#eaf6ea;border-left:4px solid #27ae60;border-radius:4px;padding:12px 16px;margin:20px 0;">
<h3 style="font-size:1.1rem;color:#444;margin-top:0;"><span style="color:#27ae60;margin-right:8px;">&#128279;</span>関連記事</h3>
<ul>
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/03/214301">手がしびれて目が覚めた ― 6つの原因と対処法【からだの不思議 #03】</a></li>
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/05/081707">機能性神経障害（FND）の真実 ― けいれんに似た発作の正体【からだの不思議 #05】</a></li>
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/03/23/120000">しびれと受診科 ― どの診療科を受診すべき？【からだの不思議 #14】</a></li>
</ul>
</div>"""

    # --- FAQ構造化データ ---
    faq_jsonld = """<script type="application/ld+json">
{"@context":"https://schema.org","@type":"FAQPage","mainEntity":[{"@type":"Question","name":"けいれん発作を目撃したら最初に何をすべきですか？","acceptedAnswer":{"@type":"Answer","text":"最も重要なのは時間を確認することです。発作が始まった時刻を記録し、周囲の危険物を除き、頭を守り、横向き（回復体位）にしてください。5分以上続く場合は119番通報してください。"}},{"@type":"Question","name":"けいれん中に口にタオルや指を入れてもいいですか？","acceptedAnswer":{"@type":"Answer","text":"絶対にやめてください。舌を飲み込むことは医学的にありえません。口に物を入れると窒息・歯の骨折・介助者の咬傷などを引き起こす最も危険な行為です。"}},{"@type":"Question","name":"けいれんが何分続いたら救急車を呼ぶべきですか？","acceptedAnswer":{"@type":"Answer","text":"5分以上続く場合は119番通報してください。国際てんかん連盟（ILAE）のガイドラインでは5分以上の発作を「てんかん重積状態」として早急な治療が必要としています。"}}]}
</script>"""

    # 関連記事を参考文献divの前に挿入
    content = content.replace(
        '<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">',
        related_html + '\n\n<div style="margin-top:2rem;border-top:2px solid #dee2e6;padding-top:1rem;">',
        1
    )

    # FAQ JSON-LDを末尾に追加
    content = content + "\n\n" + faq_jsonld

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
            # entry IDも取得（後で更新時に使用）
            id_match = re.search(r'<id>tag:[^,]*,\d+:entry-(\d+)</id>', body)
            entry_id = id_match.group(1) if id_match else "(ID取得失敗)"
            return True, entry_url, entry_id
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        return False, "HTTP {}: {}".format(e.code, error_body[:500]), None


def main():
    title = "けいれんを目撃したときの正しい対応 ― やってはいけないことと救急車を呼ぶ判断【からだの不思議 #16】"
    categories = ["医学教育", "からだの不思議", "てんかん", "救急"]

    print("=== からだの不思議 #16 ブログ下書き投稿 ===")
    print("タイトル: {}".format(title))
    print("カテゴリ: {}".format(", ".join(categories)))

    # インラインスタイル変換
    print("\n[1/2] HTML変換中...")
    content = convert_to_inline()
    print("  変換完了（{}文字）".format(len(content)))

    # 投稿
    print("\n[2/2] はてなブログに下書き投稿中...")
    env = load_env()
    xml = build_atom_entry(title, content, categories, draft=True)
    ok, url_or_err, entry_id = post_entry(env, xml)

    if ok:
        print("\n[成功] 下書き投稿完了")
        print("  URL: {}".format(url_or_err))
        print("  Entry ID: {}".format(entry_id))
    else:
        print("\n[失敗] {}".format(url_or_err))
        sys.exit(1)


if __name__ == "__main__":
    main()
