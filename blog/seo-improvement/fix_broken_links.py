#!/usr/bin/env python3
"""
Phase 6: 壊れた内部リンクを修正するスクリプト。
7記事から404 URLへのリンクを除去・修正する。

対象:
  - .htmlファイル名リンク → 除去（<a>タグからテキストのみ残す）
  - 削除entry/custom entryリンク → 関連記事リスト内なら<li>ごと除去
  - /alzheimer-disease → リンクタグ除去（テキスト残す）
  - 新着記事内の壊れたリンク → <li>ごと除去

使い方:
  python fix_broken_links.py --test     プレビュー（更新しない）
  python fix_broken_links.py --run      実際に更新
"""

import sys
import re
import csv
import time
import base64
import html as html_mod
import urllib.request
import urllib.error
from pathlib import Path

sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)

ENV_FILE = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task"
    r"\youtube-slides\食事指導シリーズ\_shared\.env"
)
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
LOG_FILE = Path(__file__).parent / "fix_broken_links_log.csv"

# 修正対象の記事（entry_id: 壊れたリンクパターンリスト）
TARGETS = {
    # まぶたの痙攣
    "17179246901372000461": {
        "title": "目がピクピクする",
        "broken_urls": [
            "https://hinyan1016.hatenablog.com/entry/2026/04/04/073318",
            "https://hinyan1016.hatenablog.com/entry/2026/03/27/213946",
        ],
    },
    # 機能性神経障害
    "17179246901371583952": {
        "title": "ストレスで本当に体は動かなくなるのか",
        "broken_urls": [
            "01_forgetfulness_vs_dementia_blog.html",
            "02_migraine_5_things_blog.html",
        ],
    },
    # めまいの種類
    "17179246901371583942": {
        "title": "めまいの種類を脳神経内科医が解説",
        "broken_urls": [
            "01_forgetfulness_vs_dementia_blog.html",
            "02_migraine_5_things_blog.html",
            "03_sleep_disorders_blog.html",
        ],
    },
    # 朝のしびれ
    "17179246901371583924": {
        "title": "手がしびれて目が覚めた",
        "broken_urls": [
            "https://hinyan1016.hatenablog.com/entry/2026/03/27/213946",
        ],
    },
    # 片頭痛5つのこと
    "17179246901371583907": {
        "title": "片頭痛持ちが知っておくべき5つのこと",
        "broken_urls": [
            "01_forgetfulness_vs_dementia_blog.html",
        ],
    },
    # 3月23日新着記事
    "17179246901368124243": {
        "title": "3月23日 新着記事 & 特集",
        "broken_urls": [
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/external_seizure_prophylaxis",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/arachnoid_cyst",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/vestibular_rehab",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/transfer_family",
        ],
    },
    # 3月20日新着記事
    "17179246901367351026": {
        "title": "3月20日 新着記事 & 特集",
        "broken_urls": [
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/external_seizure_prophylaxis",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/arachnoid_cyst",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/vestibular_rehab",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/transfer_family",
        ],
    },
    # 3月18日新着記事
    "17179246901366142013": {
        "title": "3月18日新着記事＆特集",
        "broken_urls": [
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/external_seizure_prophylaxis",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/arachnoid_cyst",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/vestibular_rehab",
            "https://hinyan1016.hatenablog.com/entry/2026/03/05/transfer_family",
        ],
    },
    # PCA記事
    "6802418398417184425": {
        "title": "後部皮質萎縮症(PCA)とは？",
        "broken_urls": [
            "/alzheimer-disease",
        ],
    },
}


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def get_auth(env):
    auth_str = "{}:{}".format(HATENA_ID, env["HATENA_API_KEY"])
    return "Basic {}".format(base64.b64encode(auth_str.encode()).decode())


def api_get(auth, entry_id):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        HATENA_ID, BLOG_DOMAIN, entry_id
    )
    req = urllib.request.Request(url, headers={"Authorization": auth})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def api_put(auth, entry_id, xml_body):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        HATENA_ID, BLOG_DOMAIN, entry_id
    )
    req = urllib.request.Request(
        url,
        data=xml_body.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth,
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def extract_content(entry_xml):
    m = re.search(r"(<content[^>]*>)(.*?)(</content>)", entry_xml, re.DOTALL)
    if not m:
        return None, None
    raw = m.group(2)
    if raw.startswith("<![CDATA[") and raw.endswith("]]>"):
        raw = raw[9:-3]
    else:
        raw = html_mod.unescape(raw)
    return raw, entry_xml


def fix_content(content, broken_urls):
    """壊れたリンクを修正"""
    fixed = content
    changes = 0

    for url in broken_urls:
        # パターン1: <li> 内のリンク → <li>ごと除去
        # 例: <li><a href="...broken...">テキスト</a></li>
        li_pattern = r'<li>\s*<a\s+href="[^"]*{}[^"]*"[^>]*>[^<]*</a>\s*</li>\s*'.format(
            re.escape(url)
        )
        li_matches = re.findall(li_pattern, fixed, re.DOTALL)
        if li_matches:
            fixed = re.sub(li_pattern, "", fixed, flags=re.DOTALL)
            changes += len(li_matches)
            continue

        # パターン2: <a href="...broken...">テキスト</a> → テキストのみ残す
        a_pattern = r'<a\s+href="[^"]*{}[^"]*"[^>]*>([^<]*)</a>'.format(
            re.escape(url)
        )
        a_matches = re.findall(a_pattern, fixed)
        if a_matches:
            fixed = re.sub(a_pattern, r"\1", fixed)
            changes += len(a_matches)
            continue

        # パターン3: より広いリンクパターン（テキスト内にタグを含む場合）
        a_pattern2 = r'<a\s+href="[^"]*{}[^"]*"[^>]*>(.*?)</a>'.format(
            re.escape(url)
        )
        a_matches2 = re.findall(a_pattern2, fixed, re.DOTALL)
        if a_matches2:
            fixed = re.sub(a_pattern2, r"\1", fixed, flags=re.DOTALL)
            changes += len(a_matches2)

    return fixed, changes


def rebuild_xml(original_xml, new_content):
    """XMLのcontent部分を置換"""
    def replacer(m):
        tag_open = m.group(1)
        return '{}<![CDATA[{}]]></content>'.format(tag_open, new_content)

    return re.sub(
        r"(<content[^>]*>).*?(</content>)",
        replacer,
        original_xml,
        count=1,
        flags=re.DOTALL,
    )


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "--test"
    is_test = mode != "--run"

    if is_test:
        print("=== テストモード（更新しません） ===\n")
    else:
        print("=== 実行モード（実際に更新します） ===\n")

    env = load_env()
    auth = get_auth(env)
    log_rows = []

    for entry_id, info in TARGETS.items():
        print(f"--- {info['title']} (ID: {entry_id}) ---")

        try:
            entry_xml = api_get(auth, entry_id)
        except Exception as e:
            print(f"  ERROR: 取得失敗 {e}")
            log_rows.append({
                "entry_id": entry_id,
                "title": info["title"],
                "status": "error",
                "detail": str(e),
                "changes": 0,
            })
            continue

        content, _ = extract_content(entry_xml)
        if not content:
            print("  ERROR: コンテンツ抽出失敗")
            log_rows.append({
                "entry_id": entry_id,
                "title": info["title"],
                "status": "error",
                "detail": "content extraction failed",
                "changes": 0,
            })
            continue

        fixed_content, change_count = fix_content(content, info["broken_urls"])

        if change_count == 0:
            print(f"  変更なし（リンクが見つからない）")
            log_rows.append({
                "entry_id": entry_id,
                "title": info["title"],
                "status": "no_change",
                "detail": "",
                "changes": 0,
            })
            continue

        print(f"  {change_count}件のリンクを修正")

        if is_test:
            # 差分プレビュー（最初の変更箇所を表示）
            for url in info["broken_urls"]:
                if url in content and url not in fixed_content:
                    print(f"  ✓ 除去: {url}")
                elif url in fixed_content:
                    print(f"  × 残存: {url}")
            log_rows.append({
                "entry_id": entry_id,
                "title": info["title"],
                "status": "preview",
                "detail": f"{change_count} links to fix",
                "changes": change_count,
            })
        else:
            # 実際に更新
            new_xml = rebuild_xml(entry_xml, fixed_content)
            try:
                api_put(auth, entry_id, new_xml)
                print(f"  ✓ 更新成功")
                log_rows.append({
                    "entry_id": entry_id,
                    "title": info["title"],
                    "status": "ok",
                    "detail": f"{change_count} links removed",
                    "changes": change_count,
                })
            except Exception as e:
                print(f"  ERROR: 更新失敗 {e}")
                log_rows.append({
                    "entry_id": entry_id,
                    "title": info["title"],
                    "status": "error",
                    "detail": str(e),
                    "changes": 0,
                })

        time.sleep(1)

    # ログ出力
    with open(LOG_FILE, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["entry_id", "title", "status", "detail", "changes"])
        writer.writeheader()
        writer.writerows(log_rows)

    print(f"\n=== 完了 ===")
    total_changes = sum(r["changes"] for r in log_rows)
    ok_count = sum(1 for r in log_rows if r["status"] in ("ok", "preview"))
    print(f"対象: {len(TARGETS)}記事, 修正: {total_changes}件")
    print(f"ログ: {LOG_FILE}")


if __name__ == "__main__":
    main()
