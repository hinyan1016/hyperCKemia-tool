#!/usr/bin/env python3
"""
からだの不思議 #17 ブログSEO強化スクリプト
- 既存コンテンツ（YouTube iframe・インフォグラフィック）を保持
- VideoObject JSON-LD追加
- FAQ構造化データ追加
- 内部リンク修正・追加
"""

import sys
import re
import urllib.request
import urllib.error
import base64
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")

ENTRY_ID = "17179246901372000545"
VIDEO_ID = "O_bmlvUY4UE"
BLOG_TITLE = "認知症の親を病院に連れて行くには ― 受診を嫌がる場合の対処法【からだの不思議 #17】"


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def get_entry_content(env):
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]
    auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        hatena_id, blog_domain, ENTRY_ID
    )
    req = urllib.request.Request(url, headers={"Authorization": "Basic {}".format(auth_b64)})
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")

    m = re.search(r'<content[^>]*>\s*<!\[CDATA\[(.*?)\]\]>\s*</content>', body, re.DOTALL)
    if m:
        return m.group(1)
    m = re.search(r'<content[^>]*>(.*?)</content>', body, re.DOTALL)
    if m:
        import html as html_mod
        return html_mod.unescape(m.group(1))
    return None


def verify_content(content, label):
    """YouTube iframe・インフォグラフィックの存在を検証"""
    ok = True
    if VIDEO_ID not in content:
        print("  [FAIL] {} - YouTube embed ({}) が見つかりません！".format(label, VIDEO_ID))
        ok = False
    else:
        print("  [OK] {} - YouTube embed".format(label))

    if "fotolife" not in content:
        print("  [FAIL] {} - インフォグラフィック (fotolife) が見つかりません！".format(label, ))
        ok = False
    else:
        print("  [OK] {} - インフォグラフィック".format(label))

    return ok


def main():
    categories = ["医学教育", "からだの不思議", "認知症", "受診"]

    print("=== からだの不思議 #17 SEO強化 ===\n")

    env = load_env()

    # 1. 既存コンテンツ取得
    print("[1/4] 既存コンテンツ取得...")
    content = get_entry_content(env)
    if not content:
        print("  ERROR: コンテンツ取得失敗")
        sys.exit(1)
    print("  取得完了（{}文字）".format(len(content)))

    # 2. 既存要素の検証
    print("\n[2/4] 既存要素の検証...")
    if not verify_content(content, "取得時"):
        print("  ERROR: 既存要素が欠損しています。中断します。")
        sys.exit(1)

    # 3. SEO要素追加
    print("\n[3/4] SEO要素追加...")

    # 3a. 内部リンク修正（壊れたファイル名リンクを正規URLに置換＋追加）
    old_related = """<div style="background: #eaf6ea; border-left: 4px solid #27ae60; border-radius: 4px; padding: 12px 16px; margin: 20px 0;"><strong>関連記事・ツール</strong><br />
<ul style="margin-bottom: 0;">
<li><a href="01_forgetfulness_vs_dementia_blog.html">【#01】「あの人の名前が出てこない」は認知症のサインか？ ― "ふつうの物忘れ"との違い</a></li>
</ul>
</div>"""

    new_related = """<div style="background: #eaf6ea; border-left: 4px solid #27ae60; border-radius: 4px; padding: 12px 16px; margin: 20px 0;"><strong>関連記事</strong>
<ul style="margin-bottom: 0;">
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/01/230038">「あの人の名前が出てこない」は認知症のサイン？ ― ふつうの物忘れとの違い【からだの不思議 #01】</a></li>
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/09/135347">デジャヴの正体を神経科学で解説【からだの不思議 #09】</a></li>
<li><a href="https://hinyan1016.hatenablog.com/entry/2026/04/15/104135">「最近つまずきやすい」は脳の問題？ ― 歩行とつまずきの原因【からだの不思議 #15】</a></li>
</ul>
</div>"""

    if old_related in content:
        content = content.replace(old_related, new_related)
        print("  内部リンク修正: OK（壊れたリンク修正＋3件に拡充）")
    else:
        print("  WARNING: 旧関連記事ブロックが見つかりません。末尾に追加します")
        content = content + "\n\n" + new_related

    # 3b. VideoObject JSON-LD
    video_jsonld = '<script type="application/ld+json">{"@context":"https://schema.org","@type":"VideoObject","name":"認知症の親を病院に連れて行くには ― 受診を嫌がる場合の対処法【からだの不思議 #17】","description":"認知症が疑われる親が受診を拒否する理由（病識低下・スティグマ）と、実際に外来で伝えている具体的な対処法を脳神経内科医が解説","thumbnailUrl":"https://img.youtube.com/vi/O_bmlvUY4UE/maxresdefault.jpg","uploadDate":"2026-04-16","contentUrl":"https://www.youtube.com/watch?v=O_bmlvUY4UE","embedUrl":"https://www.youtube.com/embed/O_bmlvUY4UE"}</script>'

    # 3c. FAQ JSON-LD
    faq_jsonld = '<script type="application/ld+json">{"@context":"https://schema.org","@type":"FAQPage","mainEntity":[{"@type":"Question","name":"認知症の親が病院に行きたがらないのはなぜですか？","acceptedAnswer":{"@type":"Answer","text":"主な理由は「病識低下（アノソグノシア）」という脳の変化です。自分の認知機能の低下を自覚できなくなるため、本人は本当に「自分はおかしくない」と感じています。アルツハイマー型認知症の40〜81%にみられます。"}},{"@type":"Question","name":"受診を拒否する親をどうやって病院に連れて行けばいいですか？","acceptedAnswer":{"@type":"Answer","text":"3つのアプローチが有効です。(1)「健康診断」「脳ドック」など別の名目を使う、(2)かかりつけ医に相談して紹介してもらう、(3)まず家族だけでもの忘れ外来や地域包括支援センターに相談する。「認知症の検査」とは言わないのがポイントです。"}},{"@type":"Question","name":"もの忘れ外来では何をしますか？","acceptedAnswer":{"@type":"Answer","text":"一般的に問診・神経心理検査（MMSE・MoCAなど）・血液検査・頭部MRIを行います。1回の受診で診断がつかないこともありますが、ベースラインの記録を残すこと自体が将来の比較に重要です。"}}]}</script>'

    # 末尾に構造化データ追加
    content = content + "\n\n" + video_jsonld + "\n" + faq_jsonld
    print("  VideoObject JSON-LD: 追加")
    print("  FAQ JSON-LD: 追加")

    # 4. 更新前の最終検証
    print("\n[4/4] 更新前の最終検証...")
    if not verify_content(content, "更新前"):
        print("  ERROR: 更新処理で既存要素が消失しました！中断します。")
        sys.exit(1)

    # VideoObject/FAQ確認
    if "VideoObject" not in content:
        print("  ERROR: VideoObject追加に失敗")
        sys.exit(1)
    if "FAQPage" not in content:
        print("  ERROR: FAQ追加に失敗")
        sys.exit(1)
    print("  [OK] VideoObject JSON-LD")
    print("  [OK] FAQ JSON-LD")
    print("  [OK] すべての要素が揃っています")

    # PUT更新
    print("\nはてなブログ更新中...")
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]
    auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()

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

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        hatena_id, blog_domain, ENTRY_ID
    )
    req = urllib.request.Request(url, data=xml.encode("utf-8"), method="PUT", headers={
        "Content-Type": "application/xml; charset=utf-8",
        "Authorization": "Basic {}".format(auth_b64),
    })

    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            entry_url = url_match.group(1) if url_match else "(URL取得失敗)"

            # 更新後の検証
            m2 = re.search(r'<content[^>]*>\s*<!\[CDATA\[(.*?)\]\]>\s*</content>', body, re.DOTALL)
            if m2:
                updated = m2.group(1)
                print("\n更新後の検証...")
                verify_content(updated, "更新後")
                print("  VideoObject:", "OK" if "VideoObject" in updated else "MISSING!")
                print("  FAQPage:", "OK" if "FAQPage" in updated else "MISSING!")
                print("  内部リンク:", len(re.findall(r'hinyan1016\.hatenablog\.com/entry/', updated)), "件")

            print("\n[成功] ブログ更新完了")
            print("  URL: {}".format(entry_url))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        print("\n[失敗] HTTP {}: {}".format(e.code, error_body[:500]))
        sys.exit(1)


if __name__ == "__main__":
    main()
