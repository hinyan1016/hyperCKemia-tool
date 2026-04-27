#!/usr/bin/env python3
"""その症状大丈夫 #04-#10 のインフォグラフィック画像を ``<a>`` ラッパーで包む。

クリック時に画像が新タブで原寸表示されるよう、本番AtomPubで以下を一括適用:
1. 各話の AtomPub GET で本番 HTML を取得
2. ``<img src="...">`` を ``<a href="同URL" target="_blank" rel="noopener noreferrer"><img ...></a>`` に置換
3. ``<img>`` の style に ``cursor:zoom-in`` を追加
4. 既に ``<a>`` で wrap 済みなら冪等的に SKIP
5. AtomPub PUT で本番更新（draft 状態を維持）
6. ローカル ``*_hatena.html`` も同期更新

引数で番号指定（例: 07 09）。引数なしで全7件処理。
"""

from __future__ import annotations

import base64
import html
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"

ROOT = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\symptom-concern-series")


@dataclass(frozen=True)
class Episode:
    num: str
    title_keyword: str
    image_pattern: str  # Fotolife URLの一意部分（YYYYMMDDHHMMSS）
    html_filename: str


EPISODES: list[Episode] = [
    Episode("04", "顔がピクピク", "20260426042347", "04_facial_twitch_hatena.html"),
    Episode("05", "足がつる", "20260426163057", "05_leg_cramps_hatena.html"),
    Episode("06", "自律神経失調症", "20260427214026", "06_autonomic_nerves_hatena.html"),
    Episode("07", "物が二重", "20260428071849", "07_diplopia_hatena.html"),
    Episode("08", "むせやすくなった", "20260428071900", "08_dysphagia_hatena.html"),
    Episode("09", "首が痛い", "20260428071911", "09_neck_pain_hatena.html"),
    Episode("10", "手足がしびれる", "20260428071921", "10_numbness_hatena.html"),
]


def auth_header() -> str:
    b64 = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()
    return "Basic {}".format(b64)


def find_entry(title_keyword: str) -> tuple[str, str, bool]:
    """draft 状態に関わらず title_keyword を含むエントリを検索。
    最新（先頭）のものを返す。戻り値: (entry_id, title, is_draft)
    """
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
            id_m = re.search(r"<id>tag:blog\.hatena\.ne\.jp[^:]*:[^:]*-(\d+)</id>", e)
            if id_m:
                return id_m.group(1), title_m.group(1), is_draft
            id_m = re.search(r"<id>([^<]+)</id>", e)
            if id_m:
                return id_m.group(1).rsplit("-", 1)[-1], title_m.group(1), is_draft
        next_m = re.search(r'<link rel="next" href="([^"]+)"', body)
        url = next_m.group(1) if next_m else None
        page += 1
    raise RuntimeError("entry not found for keyword={}".format(title_keyword))


def get_entry_xml(entry_id: str) -> str:
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, headers={"Authorization": auth_header()})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def extract_content(entry_xml: str) -> str:
    m = re.search(r'<content[^>]*type="text/html"[^>]*>(.*?)</content>', entry_xml, re.DOTALL)
    if not m:
        raise RuntimeError("content tag not found in entry XML")
    body = m.group(1)
    cdata = re.match(r"\s*<!\[CDATA\[(.*?)\]\]>\s*$", body, re.DOTALL)
    if cdata:
        return cdata.group(1)
    # AtomPubは CDATA を使わず HTML エンティティ符号化で返すことがある
    return html.unescape(body)


def add_anchor_wrapper(content: str, image_pattern: str) -> tuple[str, bool, str]:
    """img タグを ``<a>`` でラップ。既にラップ済みなら no-op。
    戻り値: (new_content, did_modify, reason)
    """
    # Find <img ... src="...image_pattern..." ... > tag
    img_re = re.compile(
        r'<img\s+([^>]*?src="[^"]*' + re.escape(image_pattern) + r'[^"]*"[^>]*?)\s*/?>',
        re.DOTALL,
    )
    img_match = img_re.search(content)
    if not img_match:
        return content, False, "img tag not found for pattern={}".format(image_pattern)

    img_full = img_match.group(0)
    img_attrs = img_match.group(1)

    # Check if already wrapped: look ~150 chars before for <a href=
    start = img_match.start()
    preceding = content[max(0, start - 200) : start]
    if re.search(r'<a[^>]*href="[^"]*' + re.escape(image_pattern) + r'[^"]*"[^>]*>\s*$', preceding):
        return content, False, "already wrapped"

    # Extract image URL
    src_m = re.search(r'src="([^"]+)"', img_attrs)
    if not src_m:
        return content, False, "src attribute not found"
    img_url = src_m.group(1)

    # Add cursor:zoom-in to style attribute
    if 'style="' in img_attrs:
        if "cursor:zoom-in" in img_attrs:
            new_attrs = img_attrs
        else:
            new_attrs = re.sub(
                r'style="([^"]*)"',
                lambda m: 'style="{}; cursor:zoom-in;"'.format(m.group(1).rstrip("; ")),
                img_attrs,
                count=1,
            )
    else:
        new_attrs = img_attrs + ' style="cursor:zoom-in;"'

    new_img = "<img " + new_attrs + ">"
    wrapped = '<a href="{}" target="_blank" rel="noopener noreferrer">{}</a>'.format(
        img_url, new_img
    )
    new_content = content[: img_match.start()] + wrapped + content[img_match.end() :]

    # Sanity check: only wrapping was added (no other content lost)
    expected_diff = len(wrapped) - len(img_full)
    actual_diff = len(new_content) - len(content)
    if actual_diff != expected_diff:
        return content, False, "diff size mismatch ({} vs expected {})".format(
            actual_diff, expected_diff
        )

    return new_content, True, "wrapped"


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


def process_episode(ep: Episode) -> tuple[bool, str]:
    print("\n[#{}] {} の処理を開始".format(ep.num, ep.title_keyword))

    print("  [1/5] エントリ検索...")
    entry_id, title, is_draft = find_entry(ep.title_keyword)
    print("        entry_id={} / title={} / draft={}".format(entry_id, title, is_draft))

    print("  [2/5] 本番 HTML を GET...")
    entry_xml = get_entry_xml(entry_id)
    content = extract_content(entry_xml)
    categories = re.findall(r'<category[^>]*term="([^"]+)"', entry_xml)
    print("        本番 content {} bytes / categories={}".format(len(content), categories))

    print("  [3/5] img タグを <a> でラップ（差分検証）...")
    new_content, modified, reason = add_anchor_wrapper(content, ep.image_pattern)
    if not modified:
        print("        [SKIP] {}".format(reason))
        return True, "skip: " + reason
    print("        [OK] {} → 差分 +{} bytes（<a>ラッパー追加のみ）".format(
        reason, len(new_content) - len(content)
    ))

    print("  [4/5] AtomPub PUT で本番を更新（draft={}）...".format("yes" if is_draft else "no"))
    result = put_entry(entry_id, title, new_content, categories, is_draft)
    print("        更新成功: {}".format(result))

    print("  [5/5] ローカル {} を同期...".format(ep.html_filename))
    html_path = ROOT / ep.html_filename
    if html_path.exists():
        local = html_path.read_text(encoding="utf-8")
        local_new, local_mod, local_reason = add_anchor_wrapper(local, ep.image_pattern)
        if local_mod:
            html_path.write_text(local_new, encoding="utf-8")
            print("        ローカル更新完了（+{} bytes）".format(len(local_new) - len(local)))
        else:
            print("        ローカルは {}".format(local_reason))
    else:
        print("        ローカル html ファイルなし（スキップ）")

    return True, result


def main() -> None:
    targets = sys.argv[1:] if len(sys.argv) > 1 else [ep.num for ep in EPISODES]
    print("その症状大丈夫 #04-#10 インフォグラフィック <a> ラッパー追加バッチ")
    print("=" * 60)
    print("対象: {}".format(", ".join(targets)))

    success = 0
    skipped = 0
    failures: list[tuple[str, str]] = []
    for ep in EPISODES:
        if ep.num not in targets:
            continue
        try:
            ok, info = process_episode(ep)
            if ok:
                if info.startswith("skip:"):
                    skipped += 1
                else:
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
    print("完了: 更新 {} / SKIP（既にwrap済） {} / 失敗 {}".format(
        success, skipped, len(failures)
    ))
    if failures:
        print("失敗:")
        for num, info in failures:
            print("  #{}: {}".format(num, info))
    sys.exit(0 if not failures else 1)


if __name__ == "__main__":
    main()
