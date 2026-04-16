#!/usr/bin/env python3
"""
SEO改善 Phase 2（v2）: VideoObject JSON-LD 構造化データを安全に追加

前回(v1)の失敗原因:
  AtomPub GETはHTMLをXMLエスケープして返す（&lt;p&gt;形式）。
  v1はこれをそのままCDATAで包んでPUT → 二重エスケープで記事破損。

v2の修正:
  1. GETで取得したcontentを html.unescape() で元のHTMLに復元
  2. JSON-LDスニペットを末尾に追加
  3. CDATAで包んでPUT（revert_final.pyで実証済みの方式）
  4. 処理前のcontent全文をバックアップ

使い方:
  python add_video_jsonld_v2.py --test          2件プレビュー（更新しない）
  python add_video_jsonld_v2.py --test-update   2件実際に更新
  python add_video_jsonld_v2.py --run           全件実行
  python add_video_jsonld_v2.py --resume N      N件目から再開
  python add_video_jsonld_v2.py --batch N M     N件目からM件処理
"""

import sys
import re
import csv
import json
import time
import base64
import html as html_mod
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime

sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)

ENV_FILE = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task"
    r"\youtube-slides\食事指導シリーズ\_shared\.env"
)
AUDIT_CSV = Path(__file__).parent / "audit_report.csv"
LOG_FILE = Path(__file__).parent / "video_jsonld_v2_log.csv"
BACKUP_DIR = Path(__file__).parent / "backup_v2"
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def get_auth(env: dict[str, str]) -> str:
    auth_str = "{}:{}".format(HATENA_ID, env["HATENA_API_KEY"])
    return "Basic {}".format(base64.b64encode(auth_str.encode()).decode())


def load_targets() -> list[dict[str, str]]:
    """CSVからYouTube埋め込みあり・JSON-LD未設定・公開記事を抽出"""
    targets: list[dict[str, str]] = []
    with open(AUDIT_CSV, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if (
                row["has_youtube"] == "True"
                and row["has_video_jsonld"] == "False"
                and row["is_draft"] == "False"
                and row["youtube_ids"]
                and row["youtube_ids"] != "PLACEHOLDER"
            ):
                targets.append(dict(row))
    return targets


def api_get(auth: str, entry_id: str) -> str:
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        HATENA_ID, BLOG_DOMAIN, entry_id
    )
    req = urllib.request.Request(url, headers={"Authorization": auth})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def api_put(auth: str, entry_id: str, xml_body: str) -> str:
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
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def extract_content(entry_xml: str) -> tuple[str | None, str | None, str | None]:
    """content要素を抽出。(raw_text, full_match, open_tag)を返す。
    CDATAがあれば剥がす。なければXMLエスケープ済みテキストをそのまま返す。"""
    m = re.search(r"(<content[^>]*>)(.*?)(</content>)", entry_xml, re.DOTALL)
    if not m:
        return None, None, None
    raw = m.group(2)
    if raw.startswith("<![CDATA[") and raw.endswith("]]>"):
        raw = raw[9:-3]
    return raw, m.group(0), m.group(1)


def build_video_jsonld(
    video_ids: list[str], title: str, published_date: str
) -> str:
    """VideoObject JSON-LDスニペットを生成"""
    objects = []
    for vid in video_ids:
        obj = {
            "@context": "https://schema.org",
            "@type": "VideoObject",
            "name": title,
            "description": title,
            "thumbnailUrl": "https://img.youtube.com/vi/{}/maxresdefault.jpg".format(
                vid
            ),
            "uploadDate": published_date[:10],
            "contentUrl": "https://www.youtube.com/watch?v={}".format(vid),
            "embedUrl": "https://www.youtube.com/embed/{}".format(vid),
        }
        objects.append(obj)

    if len(objects) == 1:
        json_str = json.dumps(objects[0], ensure_ascii=False, indent=2)
    else:
        json_str = json.dumps(objects, ensure_ascii=False, indent=2)

    return '\n<script type="application/ld+json">\n{}\n</script>'.format(json_str)


def save_backup(entry_id: str, content: str) -> None:
    """処理前のcontentをバックアップ"""
    BACKUP_DIR.mkdir(exist_ok=True)
    path = BACKUP_DIR / "{}.html".format(entry_id)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


def process_entry(
    auth: str, row: dict[str, str], dry_run: bool = False
) -> tuple[str, str]:
    """1記事を処理"""
    edit_url = row["edit_url"]
    entry_id = edit_url.split("/")[-1]
    title = row["title"]
    published = row["published"]
    video_ids = [v for v in row["youtube_ids"].split("|") if v and v != "PLACEHOLDER"]

    if not video_ids:
        return "skip", "動画IDなし"

    # エントリ取得
    try:
        entry_xml = api_get(auth, entry_id)
    except urllib.error.HTTPError as e:
        return "error", "GET HTTP {}".format(e.code)

    raw, full_match, open_tag = extract_content(entry_xml)
    if raw is None or full_match is None or open_tag is None:
        return "error", "content抽出失敗"

    # ★ 重要: XMLエスケープを解除して元のHTMLに復元
    restored_html = html_mod.unescape(raw)

    # 既にJSON-LDあり → スキップ
    if '"VideoObject"' in restored_html:
        return "skip", "JSON-LD既存"

    # バックアップ保存（dry_runでも保存して安全確認に使う）
    save_backup(entry_id, restored_html)

    # JSON-LDスニペット生成
    snippet = build_video_jsonld(video_ids, title, published)

    # 末尾に追加
    new_html = restored_html + snippet

    if dry_run:
        print("    [プレビュー] 動画ID: {}".format(", ".join(video_ids)))
        print("    [プレビュー] 追加サイズ: {} bytes".format(len(snippet)))
        print("    [プレビュー] バックアップ: {}".format(
            BACKUP_DIR / "{}.html".format(entry_id)
        ))
        return "preview", "OK"

    # CDATA内の ]]> を安全処理
    safe_html = new_html.replace("]]>", "]]]]><![CDATA[>")
    new_tag = "{}<![CDATA[{}]]></content>".format(open_tag, safe_html)
    new_xml = entry_xml.replace(full_match, new_tag)

    try:
        result = api_put(auth, entry_id, new_xml)
    except urllib.error.HTTPError as e:
        return "error", "PUT HTTP {}".format(e.code)

    url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
    return "ok", url_match.group(1) if url_match else "OK"


def verify_entry(auth: str, entry_id: str) -> tuple[str, str]:
    """更新後の記事を再取得して破損チェック"""
    try:
        entry_xml = api_get(auth, entry_id)
    except urllib.error.HTTPError as e:
        return "error", "検証GET HTTP {}".format(e.code)

    raw, _, _ = extract_content(entry_xml)
    if raw is None:
        return "error", "検証content抽出失敗"

    restored = html_mod.unescape(raw)

    # 二重エスケープチェック
    if "&lt;" in restored[:200]:
        return "BROKEN", "二重エスケープ検出"

    # JSON-LDが含まれているか
    if '"VideoObject"' not in restored:
        return "WARN", "JSON-LDが見つからない"

    return "ok", "正常"


def main() -> None:
    mode = sys.argv[1] if len(sys.argv) > 1 else "--test"
    resume_from = 0
    batch_limit: int | None = None
    if mode == "--resume" and len(sys.argv) > 2:
        resume_from = int(sys.argv[2])
    elif mode == "--batch" and len(sys.argv) > 3:
        resume_from = int(sys.argv[2])
        batch_limit = int(sys.argv[3])

    dry_run = mode == "--test"
    is_test = mode in ("--test", "--test-update")
    test_limit = 2 if is_test else batch_limit

    print("=" * 60)
    print("SEO Phase 2 v2: VideoObject JSON-LD 安全追加")
    print("モード: {}".format(mode))
    print("=" * 60)

    env = load_env()
    auth = get_auth(env)

    targets = load_targets()
    print("全対象記事: {} 件".format(len(targets)))

    if resume_from > 0:
        targets = targets[resume_from:]
        print("{}件目から再開 (残り {} 件)".format(resume_from, len(targets)))

    if test_limit:
        targets = targets[:test_limit]
        print("テストモード: {}件のみ処理".format(test_limit))

    print()

    # ログファイル準備
    log_new = not LOG_FILE.exists()
    log_f = open(LOG_FILE, "a", encoding="utf-8-sig", newline="")
    log_writer = csv.writer(log_f)
    if log_new:
        log_writer.writerow([
            "timestamp", "entry_id", "title", "video_ids",
            "status", "detail", "verify_status", "verify_detail",
        ])

    stats: dict[str, int] = {"ok": 0, "skip": 0, "error": 0, "preview": 0}

    for i, row in enumerate(targets):
        entry_id = row["edit_url"].split("/")[-1]
        print("[{}/{}] {}".format(i + 1, len(targets), row["title"][:60]))

        status, detail = process_entry(auth, row, dry_run=dry_run)
        stats[status] = stats.get(status, 0) + 1
        print("  -> {} : {}".format(status, detail))

        verify_status = ""
        verify_detail = ""

        # 更新成功した場合は即座に検証
        if status == "ok":
            time.sleep(1.0)
            verify_status, verify_detail = verify_entry(auth, entry_id)
            print("  -> 検証: {} : {}".format(verify_status, verify_detail))

            if verify_status == "BROKEN":
                print("  *** 破損検出！処理を中断します ***")
                log_writer.writerow([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    entry_id, row["title"], row["youtube_ids"],
                    status, detail, verify_status, verify_detail,
                ])
                log_f.flush()
                log_f.close()
                print("\nバックアップから復元してください:")
                print("  {}".format(BACKUP_DIR / "{}.html".format(entry_id)))
                sys.exit(1)

        if not dry_run:
            log_writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                entry_id, row["title"], row["youtube_ids"],
                status, detail, verify_status, verify_detail,
            ])
            log_f.flush()

        # API負荷対策
        if not dry_run and status == "ok":
            time.sleep(1.0)
        else:
            time.sleep(0.3)

    log_f.close()

    print("\n" + "=" * 60)
    print("完了サマリー")
    for k, v in sorted(stats.items()):
        print("  {}: {}".format(k, v))
    if not dry_run:
        print("ログ: {}".format(LOG_FILE))
    print("=" * 60)


if __name__ == "__main__":
    main()
