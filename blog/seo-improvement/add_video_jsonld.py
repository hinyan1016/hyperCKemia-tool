#!/usr/bin/env python3
"""
SEO改善 Phase 2: VideoObject JSON-LD 構造化データを追加
YouTube埋め込みがある記事にVideoObject JSON-LDを挿入する。

使い方:
  python add_video_jsonld.py --test        テスト（3件のみ、変更せずプレビュー）
  python add_video_jsonld.py --test-update テスト（3件のみ、実際に更新）
  python add_video_jsonld.py --run         全件実行
  python add_video_jsonld.py --resume N    N件目から再開
"""

import sys
import re
import csv
import json
import time
import base64
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
AUDIT_CSV = Path(__file__).parent / "audit_report.csv"
LOG_FILE = Path(__file__).parent / "video_jsonld_log.csv"
ATOM_NS = "http://www.w3.org/2005/Atom"
APP_NS = "http://www.w3.org/2007/app"
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"


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


def load_targets():
    """CSVからYouTube埋め込みあり・JSON-LD未設定・公開記事を抽出"""
    targets = []
    with open(AUDIT_CSV, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if (row["has_youtube"] == "True"
                    and row["has_video_jsonld"] == "False"
                    and row["is_draft"] == "False"
                    and row["youtube_ids"]):
                targets.append(row)
    return targets


def get_entry_xml(auth, entry_id):
    """AtomPub APIでエントリXMLを取得"""
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, method="GET",
                                headers={"Authorization": auth})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def put_entry_xml(auth, entry_id, xml_body):
    """AtomPub APIでエントリを更新"""
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        HATENA_ID, BLOG_DOMAIN, entry_id)
    data = xml_body.encode("utf-8")
    req = urllib.request.Request(url, data=data, method="PUT",
                                headers={
                                    "Content-Type": "application/xml; charset=utf-8",
                                    "Authorization": auth,
                                })
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def extract_content(entry_xml):
    """XMLからcontent部分を抽出"""
    m = re.search(r'<content[^>]*>(.*?)</content>', entry_xml, re.DOTALL)
    if not m:
        return None, None
    raw = m.group(1)
    content = raw
    if content.startswith("<![CDATA["):
        content = content[9:]
    if content.endswith("]]>"):
        content = content[:-3]
    return content, m.group(0)


def build_video_jsonld(video_ids, title, published_date):
    """VideoObject JSON-LDスニペットを生成"""
    objects = []
    for vid in video_ids:
        obj = {
            "@context": "https://schema.org",
            "@type": "VideoObject",
            "name": title,
            "description": title,
            "thumbnailUrl": "https://img.youtube.com/vi/{}/maxresdefault.jpg".format(vid),
            "uploadDate": published_date,
            "contentUrl": "https://www.youtube.com/watch?v={}".format(vid),
            "embedUrl": "https://www.youtube.com/embed/{}".format(vid),
        }
        objects.append(obj)

    if len(objects) == 1:
        json_str = json.dumps(objects[0], ensure_ascii=False, indent=2)
    else:
        json_str = json.dumps(objects, ensure_ascii=False, indent=2)

    snippet = '\n<script type="application/ld+json">\n{}\n</script>'.format(json_str)
    return snippet


def already_has_jsonld(content):
    """既にVideoObject JSON-LDがあるか再チェック"""
    return '"VideoObject"' in content or "VideoObject" in content


def process_entry(auth, row, dry_run=False):
    """1記事を処理"""
    edit_url = row["edit_url"]
    entry_id = edit_url.split("/")[-1]
    title = row["title"]
    published = row["published"]
    video_ids = row["youtube_ids"].split("|")

    # PLACEHOLDERをスキップ
    video_ids = [v for v in video_ids if v and v != "PLACEHOLDER"]
    if not video_ids:
        return "skip", "動画IDなし"

    # エントリ取得
    try:
        entry_xml = get_entry_xml(auth, entry_id)
    except urllib.error.HTTPError as e:
        return "error", "GET HTTP {}".format(e.code)

    content, original_tag = extract_content(entry_xml)
    if content is None:
        return "error", "content抽出失敗"

    # 既にJSON-LDあり → スキップ
    if already_has_jsonld(content):
        return "skip", "JSON-LD既存"

    # JSON-LD生成
    snippet = build_video_jsonld(video_ids, title, published)

    # 挿入（本文末尾に追加）
    new_content = content + snippet

    if dry_run:
        print("    [プレビュー] 追加するJSON-LD:")
        print("    動画ID: {}".format(", ".join(video_ids)))
        return "preview", "OK"

    # XML更新
    new_tag = '<content type="text/html"><![CDATA[{}]]></content>'.format(new_content)
    new_xml = entry_xml.replace(original_tag, new_tag)

    try:
        result = put_entry_xml(auth, entry_id, new_xml)
        url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
        result_url = url_match.group(1) if url_match else ""
        return "ok", result_url
    except urllib.error.HTTPError as e:
        return "error", "PUT HTTP {}".format(e.code)


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "--test"
    resume_from = 0
    if mode == "--resume" and len(sys.argv) > 2:
        resume_from = int(sys.argv[2])

    dry_run = (mode == "--test")
    is_test = mode in ("--test", "--test-update")
    test_limit = 3 if is_test else None

    print("=" * 60)
    print("SEO Phase 2: VideoObject JSON-LD 追加")
    print("モード: {}".format(mode))
    print("=" * 60)

    env = load_env()
    auth = get_auth(env)

    # 対象記事読み込み
    targets = load_targets()
    print("対象記事: {} 件".format(len(targets)))

    if resume_from > 0:
        targets = targets[resume_from:]
        print("{}件目から再開 (残り {} 件)".format(resume_from, len(targets)))

    if test_limit:
        targets = targets[:test_limit]
        print("テストモード: {}件のみ処理".format(test_limit))

    print()

    # ログファイル準備
    log_exists = LOG_FILE.exists()
    log_f = open(LOG_FILE, "a", encoding="utf-8-sig", newline="")
    log_writer = csv.writer(log_f)
    if not log_exists:
        log_writer.writerow(["timestamp", "entry_id", "title", "video_ids", "status", "detail"])

    stats = {"ok": 0, "skip": 0, "error": 0, "preview": 0}

    for i, row in enumerate(targets):
        entry_id = row["edit_url"].split("/")[-1]
        print("[{}/{}] {}".format(i + 1, len(targets), row["title"][:50]))

        status, detail = process_entry(auth, row, dry_run=dry_run)
        stats[status] += 1
        print("  → {} : {}".format(status, detail))

        if not dry_run:
            log_writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                entry_id,
                row["title"],
                row["youtube_ids"],
                status,
                detail,
            ])
            log_f.flush()

        # API負荷対策
        if not dry_run:
            time.sleep(1.0)
        else:
            time.sleep(0.3)

    log_f.close()

    print("\n" + "=" * 60)
    print("完了サマリー")
    print("  成功: {}".format(stats["ok"]))
    print("  スキップ: {}".format(stats["skip"]))
    print("  エラー: {}".format(stats["error"]))
    print("  プレビュー: {}".format(stats["preview"]))
    if not dry_run:
        print("ログ: {}".format(LOG_FILE))
    print("=" * 60)


if __name__ == "__main__":
    main()
