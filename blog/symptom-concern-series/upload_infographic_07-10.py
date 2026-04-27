#!/usr/bin/env python3
"""その症状大丈夫 #07-#10 インフォグラフィック反映バッチ。

各話の処理:
1. はてなフォトライフへインフォグラフィックをアップロード
2. ローカル `<NN>_<key>_hatena.html` のリード文直後に <img> ブロックを挿入
   （`<!-- youtube embed -->` があればその直前、なければ `<!-- 目次 -->` の直前）
3. AtomPub PUT で既存の下書きを更新（draft=yes 維持）

引数なしで全話処理。`07 09` のように番号指定で部分実行可。
既存の Fotolife URL を再利用したい場合は `07=<url>` 形式で指定。
"""

from __future__ import annotations

import base64
import hashlib
import random
import re
import string
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"

ROOT = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\symptom-concern-series")
INFOGRAPHIC_DIR = ROOT / "infographics"


@dataclass(frozen=True)
class Episode:
    num: str
    title_keyword: str
    image_filename: str
    html_filename: str
    fotolife_title: str
    alt: str


EPISODES: list[Episode] = [
    Episode(
        num="07",
        title_keyword="物が二重",
        image_filename="07_diplopia.png",
        html_filename="07_diplopia_hatena.html",
        fotolife_title="diplopia_infographic",
        alt=(
            "複視 単眼性vs両眼性 インフォグラフィック ― 6つの原因疾患"
            "（疲労・糖尿病・甲状腺眼症・重症筋無力症・脳動脈瘤・脳卒中）・"
            "自分でできる片目テスト5ステップ・緊急サイン・症状で選ぶ受診先"
        ),
    ),
    Episode(
        num="08",
        title_keyword="むせやすくなった",
        image_filename="08_dysphagia.png",
        html_filename="08_dysphagia_hatena.html",
        fotolife_title="dysphagia_infographic",
        alt=(
            "嚥下障害「むせやすくなった」インフォグラフィック ― "
            "加齢性vs病的・主な原因6つ・見逃しやすい脳神経の病気6つ（脳卒中・"
            "パーキンソン・重症筋無力症・ALS・不顕性誤嚥・誤嚥性肺炎）・"
            "セルフチェック5項目・誤嚥予防の5習慣"
        ),
    ),
    Episode(
        num="09",
        title_keyword="首が痛い",
        image_filename="09_neck_pain.png",
        html_filename="09_neck_pain_hatena.html",
        fotolife_title="neck_pain_infographic",
        alt=(
            "首の痛み 肩こりvs神経系のサイン インフォグラフィック ― "
            "6つの原因疾患（肩こり・頚椎症・椎間板ヘルニア・頚椎症性脊髄症・"
            "椎骨動脈解離・髄膜炎）・受けるべき検査5項目・危険サイン・"
            "受診先早見表・セルフケア5習慣"
        ),
    ),
    Episode(
        num="10",
        title_keyword="手足がしびれる",
        image_filename="10_numbness.png",
        html_filename="10_numbness_hatena.html",
        fotolife_title="numbness_infographic",
        alt=(
            "手足のしびれ 感覚低下型vs異常感覚型 インフォグラフィック ― "
            "分布パターン6・主な原因疾患6つ（脳卒中・手根管症候群・"
            "糖尿病性神経障害・ギラン・バレー症候群・頚椎症性脊髄症・"
            "ビタミンB12欠乏）・緊急サイン・予防の5習慣"
        ),
    ),
]


def auth_header() -> str:
    b64 = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()
    return "Basic {}".format(b64)


def wsse_header() -> str:
    nonce_raw = "".join(random.choices(string.ascii_letters + string.digits, k=40))
    nonce = nonce_raw.encode("utf-8")
    created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    digest = hashlib.sha1(nonce + created.encode() + API_KEY.encode()).digest()
    pwd = base64.b64encode(digest).decode()
    nonce_b64 = base64.b64encode(nonce).decode()
    return 'UsernameToken Username="{}", PasswordDigest="{}", Nonce="{}", Created="{}"'.format(
        HATENA_ID, pwd, nonce_b64, created
    )


def upload_fotolife(image_path: Path, fotolife_title: str) -> str:
    with open(image_path, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml += '<entry xmlns="http://purl.org/atom/ns#">\n'
    xml += "  <title>{}</title>\n".format(fotolife_title)
    xml += '  <content mode="base64" type="image/png">{}</content>\n'.format(b64)
    xml += "</entry>"

    req = urllib.request.Request(
        "https://f.hatena.ne.jp/atom/post",
        data=xml.encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "X-WSSE": wsse_header(),
            "Authorization": 'WSSE profile="UsernameToken"',
        },
    )
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")
    m = re.search(r"<hatena:imageurl>([^<]+)</hatena:imageurl>", body)
    if m:
        return m.group(1)
    m = re.search(r'https://cdn-ak\.f\.st-hatena\.com/images/fotolife/[^<"\s]+', body)
    if m:
        return m.group(0)
    raise RuntimeError("imageurl not found: " + body[:500])


def find_draft_entry(title_keyword: str) -> tuple[str, str]:
    """draft=yes の下書きから title_keyword を含むエントリを検索。"""
    url: str | None = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(HATENA_ID, BLOG_DOMAIN)
    page = 0
    while url and page < 30:
        req = urllib.request.Request(url, headers={"Authorization": auth_header()})
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
        entries = re.findall(r"<entry>(.*?)</entry>", body, re.DOTALL)
        for e in entries:
            title_m = re.search(r"<title[^>]*>([^<]*)</title>", e)
            if not title_m or title_keyword not in title_m.group(1):
                continue
            draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", e)
            is_draft = (draft_m.group(1) == "yes") if draft_m else False
            if not is_draft:
                print("        候補スキップ（draft=no）: {}".format(title_m.group(1)))
                continue
            id_m = re.search(r"<id>tag:blog\.hatena\.ne\.jp[^:]*:[^:]*-(\d+)</id>", e)
            if id_m:
                return id_m.group(1), title_m.group(1)
            id_m = re.search(r"<id>([^<]+)</id>", e)
            if id_m:
                return id_m.group(1).rsplit("-", 1)[-1], title_m.group(1)
        next_m = re.search(r'<link rel="next" href="([^"]+)"', body)
        url = next_m.group(1) if next_m else None
        page += 1
    raise RuntimeError("draft entry not found for keyword={}".format(title_keyword))


def get_entry(entry_id: str) -> str:
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, headers={"Authorization": auth_header()})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def put_entry(
    entry_id: str,
    title: str,
    content: str,
    categories: list[str],
    is_draft: bool,
) -> str:
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)
    xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml += '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">\n'
    xml += "  <title>{}</title>\n".format(title)
    xml += "  <author><name>{}</name></author>\n".format(HATENA_ID)
    xml += '  <content type="text/html"><![CDATA[{}]]></content>\n'.format(content)
    xml += "{}\n".format(cat_xml)
    xml += "  <app:control>\n    <app:draft>{}</app:draft>\n  </app:control>\n".format(
        "yes" if is_draft else "no"
    )
    xml += "</entry>"

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(
        url,
        data=xml.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth_header(),
        },
    )
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")
    url_m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
    return url_m.group(1) if url_m else "(URL不明)"


def insert_infographic_block(html: str, image_url: str, alt: str) -> str:
    """インフォグラ block を `<!-- youtube embed -->` の直前に、なければ `<!-- 目次 -->` の直前に挿入。"""
    if "<!-- infographic -->" in html:
        return html

    img_block = (
        "\n<!-- infographic -->\n"
        '<div style="text-align:center; margin:24px 0;">\n'
        '<img src="{url}" alt="{alt}" '
        'style="max-width:100%; height:auto; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.1);" '
        'loading="lazy">\n'
        "</div>\n\n"
    ).format(url=image_url, alt=alt)

    if "<!-- youtube embed -->" in html:
        return html.replace("<!-- youtube embed -->", img_block + "<!-- youtube embed -->", 1)
    if "<!-- 目次 -->" in html:
        return html.replace("<!-- 目次 -->", img_block + "<!-- 目次 -->", 1)
    # フォールバック: TOC マーカーがない場合は ul.table-of-contents の直前に挿入
    toc_marker = '<ul class="table-of-contents">'
    if toc_marker in html:
        return html.replace(toc_marker, img_block + toc_marker, 1)
    raise RuntimeError(
        "挿入先マーカーが見つかりません（<!-- youtube embed --> / <!-- 目次 --> / <ul class=\"table-of-contents\">）"
    )


def process_episode(ep: Episode, override_url: str | None = None) -> tuple[bool, str]:
    print("\n[#{}] {} の処理を開始".format(ep.num, ep.title_keyword))
    image_path = INFOGRAPHIC_DIR / ep.image_filename
    html_path = ROOT / ep.html_filename

    if not image_path.exists():
        return False, "image file not found: {}".format(image_path)
    if not html_path.exists():
        return False, "html file not found: {}".format(html_path)

    # 1. Fotolife アップロード（または既存URL使用）
    if override_url:
        image_url = override_url
        print("  [1/4] 既存URL使用: {}".format(image_url))
    else:
        print("  [1/4] Fotolife アップロード中...")
        image_url = upload_fotolife(image_path, ep.fotolife_title)
        print("        画像URL: {}".format(image_url))

    # 2. 下書き entry id 検索
    print("  [2/4] 下書き entry id 検索中...")
    entry_id, title = find_draft_entry(ep.title_keyword)
    print("        entry_id={} / title={}".format(entry_id, title))

    # 3. 既存エントリから category と draft 状態を取得
    print("  [3/4] 既存エントリから category/draft 状態を取得...")
    current = get_entry(entry_id)
    categories = re.findall(r'<category[^>]*term="([^"]+)"', current)
    draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", current)
    is_draft = (draft_m.group(1) == "yes") if draft_m else True

    DEFAULT_CATS = ["その症状大丈夫", "脳神経内科", "セルフチェック"]
    if not categories:
        categories = DEFAULT_CATS
        print("        categories 空 → 既定値補完: {}".format(categories))
    print("        categories={} / draft={}".format(categories, is_draft))

    # 4. ローカル HTML に画像を挿入し、AtomPub PUT
    html = html_path.read_text(encoding="utf-8")
    if "<!-- infographic -->" in html:
        print("  [4/4] 画像は既に HTML に挿入済み。そのまま PUT します")
    else:
        html = insert_infographic_block(html, image_url, ep.alt)
        html_path.write_text(html, encoding="utf-8")
        print("  [4/4] ローカル HTML に挿入完了（{} bytes）".format(len(html.encode("utf-8"))))

    print("        AtomPub PUT 中（draft={}）...".format("yes" if is_draft else "no"))
    result = put_entry(entry_id, title, html, categories, is_draft)
    print("        更新成功: {}".format(result))
    return True, result


def parse_args(argv: list[str]) -> tuple[list[str], dict[str, str]]:
    """引数を「処理対象番号リスト」と「番号→既存URL map」に分解。"""
    nums: list[str] = []
    overrides: dict[str, str] = {}
    for a in argv:
        if "=" in a:
            num, url = a.split("=", 1)
            overrides[num] = url
            nums.append(num)
        else:
            nums.append(a)
    return nums, overrides


def main() -> None:
    nums, overrides = parse_args(sys.argv[1:])
    targets = nums if nums else [ep.num for ep in EPISODES]

    print("その症状大丈夫 #07-#10 インフォグラフィック反映バッチ")
    print("=" * 60)
    print("対象: {}".format(", ".join(targets)))
    if overrides:
        print("既存URL: {}".format(overrides))

    success = 0
    failures: list[tuple[str, str]] = []
    for ep in EPISODES:
        if ep.num not in targets:
            continue
        try:
            ok, info = process_episode(ep, overrides.get(ep.num))
            if ok:
                success += 1
            else:
                failures.append((ep.num, info))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")[:300]
            failures.append((ep.num, "HTTP {}: {}".format(e.code, body)))
            print("        [HTTP Error {}] {}".format(e.code, body))
        except Exception as e:  # noqa: BLE001
            failures.append((ep.num, str(e)))
            print("        [Error] {}".format(e))

    print()
    print("=" * 60)
    print("完了: {} / {} 件成功".format(success, len(targets)))
    if failures:
        print("失敗:")
        for num, info in failures:
            print("  #{}: {}".format(num, info))
    sys.exit(0 if not failures else 1)


if __name__ == "__main__":
    main()
