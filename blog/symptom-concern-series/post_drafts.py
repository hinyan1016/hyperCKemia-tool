#!/usr/bin/env python3
"""
その症状大丈夫シリーズ ブログ下書き投稿スクリプト
ブログHTMLをはてな形式に変換し、AtomPub APIで下書き投稿する
"""

import re
import sys
import urllib.request
import urllib.error
import base64

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"
CATEGORIES = ["その症状大丈夫", "脳神経内科", "セルフチェック"]


def convert_to_hatena(html):
    """ブログHTMLをはてな形式に変換する"""
    # body内のdiv.medical-articleの中身だけ取得
    m = re.search(r'<div class="medical-article">\s*<style>.*?</style>\s*(.*?)\s*</div>\s*</body>', html, re.DOTALL)
    if not m:
        raise ValueError("medical-article div not found")
    content = m.group(1)

    # series-label除去
    content = re.sub(r'<span class="series-label">.*?</span>\s*', '', content)

    # h1除去（タイトルはAtomPubで指定）
    content = re.sub(r'<h1>.*?</h1>\s*', '', content)

    # toc-box → table-of-contents
    def convert_toc(m):
        inner = m.group(1)
        # h3タイトル除去、ul以下のみ残す
        ul_m = re.search(r'(<ul>.*?</ul>)', inner, re.DOTALL)
        if ul_m:
            ul = ul_m.group(1)
            # id付きアンカーを日本語テキストベースのアンカーに変換
            # はてなブログは[:contents]で自動目次を生成するため、シンプルなリストで十分
            return '<ul class="table-of-contents">\n' + re.sub(r'<ul>|</ul>', '', ul).strip() + '\n</ul>'
        return ''
    content = re.sub(r'<div class="toc-box">(.*?)</div>', convert_toc, content, flags=re.DOTALL)

    # h2のidを日本語テキストに変換（はてなブログ用）
    # 同時にTOCのhref（英語スラッグ）も日本語idへ追従させる
    id_mapping = {}

    def make_clean_id(text):
        clean = re.sub(r'[―–—]', '--', text)
        clean = re.sub(r'[^\w\s\u3000-\u9fff\u30a0-\u30ff\u3040-\u309f-]', '', clean)
        clean = re.sub(r'\s+', '', clean)
        return clean

    def convert_h2_id(m):
        old_id = m.group(1)
        text = m.group(2)
        new_id = make_clean_id(text)
        if old_id != new_id:
            id_mapping[old_id] = new_id
        return '<h2 id="' + new_id + '">' + text + '</h2>'
    content = re.sub(r'<h2 id="([^"]+)">(.*?)</h2>', convert_h2_id, content)

    # TOCのhref="#旧ID"を新IDに追従（見出しと一致させる）
    for old_id, new_id in id_mapping.items():
        content = content.replace(
            'href="#' + old_id + '"',
            'href="#' + new_id + '"',
        )

    # important-box → inline style
    content = re.sub(
        r'<div class="important-box">\s*',
        '<p style="background:#e8f4fd; border-left:4px solid #2c5aa0; padding:12px 16px; border-radius:4px;">',
        content
    )
    # warning-box → inline style
    content = re.sub(
        r'<div class="warning-box">\s*',
        '<p style="background:#fff3cd; border-left:4px solid #ffc107; padding:12px 16px; border-radius:4px;">',
        content
    )
    # emergency-box → inline style
    content = re.sub(
        r'<div class="emergency-box">\s*',
        '<p style="background:#f8d7da; border-left:4px solid #dc3545; padding:12px 16px; border-radius:4px;">',
        content
    )
    # summary-box → inline style
    content = re.sub(
        r'<div class="summary-box">\s*',
        '<p style="background:#e8f4fd; border:2px solid #2c5aa0; padding:16px; border-radius:8px;">',
        content
    )
    # 上記divの閉じタグをpの閉じタグに（最初のマッチングする</div>）
    # Note: これらのboxの中に<ol>や<ul>がある場合はそのまま残す
    # 閉じ</div>を</p>に変換（boxクラスに対応するもの）
    # シンプルに: </div>が残っている場合、boxの閉じを処理

    # comparison-table → inline styled table
    def convert_table(m):
        table_content = m.group(1)
        rows = re.findall(r'<tr>(.*?)</tr>', table_content, re.DOTALL)
        result = '<table style="width:100%; border-collapse:collapse; margin:20px 0; font-size:95%;">\n<tbody>\n'
        for i, row in enumerate(rows):
            cells = re.findall(r'<td(?:\s[^>]*)?>(.*?)</td>', row, re.DOTALL)
            if i == 0:
                # ヘッダー行
                result += '<tr>\n'
                for cell in cells:
                    result += '<td style="padding:10px 12px; background:#2c5aa0; color:#fff; font-weight:bold;">' + cell.strip() + '</td>\n'
                result += '</tr>\n'
            else:
                bg = ' style="background:#f9f9f9;"' if i % 2 == 1 else ''
                result += '<tr' + bg + '>\n'
                for cell in cells:
                    # 既存のstyle属性を保持しつつpadding追加
                    cell_str = cell.strip()
                    result += '<td style="padding:10px 12px;">' + cell_str + '</td>\n'
                result += '</tr>\n'
        result += '</tbody>\n</table>'
        return result
    content = re.sub(r'<table class="comparison-table">\s*(.*?)\s*</table>', convert_table, content, flags=re.DOTALL)

    # sup内の参考文献番号変換 [1] → 1)
    content = re.sub(r'<sup>\[(\d+)\]</sup>', r'<sup>\1)</sup>', content)

    # セルフチェックツールリンク変換
    content = re.sub(
        r'<p><a href="(https://hinyan1016\.github\.io/symptom-checker/[^"]+)">&#x1F50D; (.*?)</a></p>',
        r'<p><strong>▶ セルフチェック：</strong><a href="\1">\2</a></p>',
        content
    )

    # faq-section/faq-item/faq-q → シンプル形式
    content = re.sub(r'<div class="faq-section">\s*', '', content)
    content = re.sub(r'<div class="faq-item">\s*', '', content)
    content = re.sub(r'<p class="faq-q">(.*?)</p>', r'<p><strong>\1</strong></p>', content)

    # references → シンプル形式
    content = re.sub(r'<div class="references">\s*', '', content)
    content = re.sub(r'<h3>参考文献</h3>', '<h2>参考文献</h2>', content)

    def convert_ref(m):
        num = m.group(1)
        ref_content = m.group(2).strip()
        return '<p>' + num + ') ' + ref_content + '</p>'
    content = re.sub(
        r'<div class="ref-item">\s*<span class="ref-numbers">\[(\d+)\]</span>\s*<span class="ref-content">(.*?)</span>\s*</div>',
        convert_ref, content, flags=re.DOTALL
    )

    # 残りの</div>を除去（box系・faq系・references系の閉じタグ）
    # ただし構造的な</div>は残さない
    content = content.replace('</div>', '</p>')

    # </p></p>の重複を修正
    content = re.sub(r'</p>\s*</p>', '</p>', content)

    # <ol>や<ul>を含むbox内の</p>タグ位置を修正
    # box内で<ol>...</ol>の後に</p>が来るようにする

    # 空行の整理
    content = re.sub(r'\n{3,}', '\n\n', content)
    content = content.strip()

    return content


def extract_title(html):
    """HTMLからh1タイトルを抽出"""
    m = re.search(r'<h1>(.*?)</h1>', html)
    return m.group(1) if m else None


def post_draft(title, content, categories):
    """はてなブログにdraft投稿"""
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
</entry>""".format(title=title, author=HATENA_ID, content=content, categories=cat_xml)

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(HATENA_ID, BLOG_DOMAIN)
    auth_b64 = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()

    req = urllib.request.Request(
        url,
        data=xml.encode("utf-8"),
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
            return True, entry_url
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        return False, "HTTP {}: {}".format(e.code, error_body[:300])


def main():
    import os
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 投稿対象ファイル
    files = {
        "02": "02_slurred_speech_blog.html",
        "03": "03_weakness_blog.html",
        "04": "04_facial_twitch_blog.html",
        "05": "05_leg_cramps_blog.html",
        "06": "06_autonomic_nerves_blog.html",
        "07": "07_diplopia_blog.html",
        "08": "08_dysphagia_blog.html",
        "09": "09_neck_pain_blog.html",
        "10": "10_numbness_blog.html",
    }

    # 引数でフィルタ
    targets = sys.argv[1:] if len(sys.argv) > 1 else sorted(files.keys())

    print("\nその症状大丈夫シリーズ ブログ下書き投稿")
    print("=" * 60)
    print("対象: {} 件\n".format(len(targets)))

    success = 0
    for num in targets:
        if num not in files:
            print("  [スキップ] #{} は対象外です".format(num))
            continue

        filepath = os.path.join(script_dir, files[num])
        if not os.path.exists(filepath):
            print("  [スキップ] {} が見つかりません".format(files[num]))
            continue

        with open(filepath, encoding="utf-8") as f:
            html = f.read()

        title = extract_title(html)
        if not title:
            print("  [エラー] #{} のタイトルが見つかりません".format(num))
            continue

        print("  #{} 変換中: {}".format(num, title))
        hatena_content = convert_to_hatena(html)

        # hatena版を保存
        hatena_path = os.path.join(script_dir, files[num].replace("_blog.html", "_hatena.html"))
        with open(hatena_path, "w", encoding="utf-8") as f:
            f.write(hatena_content)
        print("    → hatena版保存: {}".format(os.path.basename(hatena_path)))

        # 下書き投稿
        print("    → 下書き投稿中...")
        ok, result = post_draft(title, hatena_content, CATEGORIES)
        if ok:
            print("    → [成功] {}".format(result))
            success += 1
        else:
            print("    → [失敗] {}".format(result))
        print()

    print("完了: {}/{} 件成功".format(success, len(targets)))


if __name__ == "__main__":
    main()
