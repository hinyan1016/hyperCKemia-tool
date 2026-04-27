#!/usr/bin/env python3
"""#05 足がつる ― AtomPub recon: ドラフトの存在確認＋本番content差分プレビュー"""
import re
import sys
import base64
import urllib.request
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"

LOCAL_HTML = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\symptom-concern-series\05_leg_cramps_hatena.html")
TITLE_KEYWORDS = ["寝ている時に足がつる", "足がつる", "こむら返り"]


def auth_header():
    b64 = base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()
    return "Basic " + b64


def find_entry():
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(HATENA_ID, BLOG_DOMAIN)
    page = 0
    while url and page < 30:
        req = urllib.request.Request(url, headers={"Authorization": auth_header()})
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
        entries = re.findall(r"<entry>(.*?)</entry>", body, re.DOTALL)
        for e in entries:
            title_m = re.search(r"<title[^>]*>([^<]*)</title>", e)
            if not title_m:
                continue
            title = title_m.group(1)
            if not any(kw in title for kw in TITLE_KEYWORDS):
                continue
            draft_m = re.search(r"<app:draft>(yes|no)</app:draft>", e)
            is_draft = (draft_m.group(1) == "yes") if draft_m else False
            alt_m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', e)
            alt_url = alt_m.group(1) if alt_m else ""
            id_m = re.search(r"<id>tag:blog\.hatena\.ne\.jp[^:]*:[^:]*-(\d+)</id>", e)
            entry_id = id_m.group(1) if id_m else None
            yield {"entry_id": entry_id, "title": title, "is_draft": is_draft, "url": alt_url}
        next_m = re.search(r'<link rel="next" href="([^"]+)"', body)
        url = next_m.group(1) if next_m else None
        page += 1


def get_entry_content(entry_id):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, headers={"Authorization": auth_header()})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def main():
    print("[recon 1/3] AtomPubで#05候補を検索...")
    candidates = list(find_entry())
    if not candidates:
        print("  該当なし")
        return
    for c in candidates:
        print("  - id={entry_id} draft={is_draft} title={title!r} url={url}".format(**c))

    target = candidates[0]
    print("\n[recon 2/3] 本命候補のcontent取得 (entry_id={})...".format(target["entry_id"]))
    xml = get_entry_content(target["entry_id"])
    content_m = re.search(r"<content[^>]*>(.*?)</content>", xml, re.DOTALL)
    if not content_m:
        print("  content取れず")
        return
    live_html = content_m.group(1)
    # CDATA解除
    cdata = re.match(r"\s*<!\[CDATA\[(.*?)\]\]>\s*", live_html, re.DOTALL)
    if cdata:
        live_html = cdata.group(1)

    print("\n[recon 3/3] 本番 vs ローカル 差分検出（インフォグラ/画像/動画）...")
    local_html = LOCAL_HTML.read_text(encoding="utf-8")

    def find_blocks(html):
        return {
            "infographic_marker": "<!-- infographic -->" in html,
            "youtube_marker": "<!-- youtube embed -->" in html,
            "json_ld_marker": "VideoObject" in html,
            "img_count": len(re.findall(r"<img[^>]+>", html)),
            "iframe_count": len(re.findall(r"<iframe[^>]+>", html)),
            "fotolife_imgs": list(set(re.findall(r'src="([^"]+fotolife[^"]+)"', html))),
            "youtube_iframes": list(set(re.findall(r'src="(https?://[^"]*youtube[^"]+)"', html))),
            "size_bytes": len(html.encode("utf-8")),
        }

    live_info = find_blocks(live_html)
    local_info = find_blocks(local_html)

    print("\n   |     項目          |  本番(live)  |  ローカル  |  差異")
    print("   |-------------------|--------------|------------|------")
    for k in ["infographic_marker", "youtube_marker", "json_ld_marker", "img_count", "iframe_count", "size_bytes"]:
        live_v = live_info[k]
        local_v = local_info[k]
        diff = "⚠️ 不一致" if live_v != local_v else "OK"
        print("   | {:18}| {:12} | {:10} | {}".format(k, str(live_v), str(local_v), diff))

    if live_info["fotolife_imgs"]:
        print("\n   本番Fotolife画像（要保護）:")
        for img in live_info["fotolife_imgs"]:
            print("     - " + img)
    else:
        print("\n   本番にFotolife画像なし")

    if live_info["youtube_iframes"]:
        print("\n   本番YouTube埋め込み:")
        for u in live_info["youtube_iframes"]:
            print("     - " + u)


if __name__ == "__main__":
    main()
