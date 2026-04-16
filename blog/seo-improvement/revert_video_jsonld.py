#!/usr/bin/env python3
"""
緊急修復: VideoObject JSON-LD追加で発生したエスケープ破損を修復

問題の原因:
  AtomPub APIのGETはCDATAなしでHTMLエスケープして返す。
  元のスクリプトがそのままCDATAで包んでPUTしたため、
  &lt;p&gt;が文字通り保存され、ブラウザに生HTMLコードが表示された。

修復方法:
  1. GETでcontent取得（APIがXMLエスケープ済み）
  2. html.unescapeを2回適用→元のraw HTMLに戻す
  3. JSON-LDスニペットを除去
  4. CDATAで包んでPUT→正しくraw HTMLが保存される

使い方:
  python revert_video_jsonld.py --test       3件プレビュー
  python revert_video_jsonld.py --test-fix   3件実際に修復
  python revert_video_jsonld.py --run        全件修復
  python revert_video_jsonld.py --resume N   N件目から再開
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
from datetime import datetime

sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
LOG_FILE = Path(__file__).parent / "video_jsonld_log.csv"
REVERT_LOG = Path(__file__).parent / "revert_log.csv"
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


def get_entry_xml(auth, entry_id):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, method="GET",
                                headers={"Authorization": auth})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def put_entry_xml(auth, entry_id, xml_body):
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


def extract_raw_content(entry_xml):
    """XMLからcontent要素のraw textを抽出（CDATAがあれば剥がす）"""
    m = re.search(r'(<content[^>]*>)(.*?)(</content>)', entry_xml, re.DOTALL)
    if not m:
        return None, None
    raw = m.group(2)
    # CDATA剥がし
    if raw.startswith("<![CDATA[") and raw.endswith("]]>"):
        raw = raw[9:-3]
    return raw, m.group(0)


def remove_video_jsonld(content):
    """VideoObject JSON-LDスニペットをraw HTML状態で除去"""
    # raw HTML形式: <script type="application/ld+json">\n{...VideoObject...}\n</script>
    # 単一オブジェクト
    fixed = re.sub(
        r'\n?<script\s+type="application/ld\+json">\s*\{[^<]*?"VideoObject"[^<]*?\}\s*</script>',
        '', content, flags=re.DOTALL)
    # 配列形式
    fixed = re.sub(
        r'\n?<script\s+type="application/ld\+json">\s*\[[^\]]*?"VideoObject"[^\]]*?\]\s*</script>',
        '', fixed, flags=re.DOTALL)
    return fixed


def load_updated_entries():
    """ログファイルから更新済みエントリIDを取得"""
    entries = []
    if not LOG_FILE.exists():
        print("ERROR: ログファイルが見つかりません: {}".format(LOG_FILE))
        sys.exit(1)
    with open(LOG_FILE, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["status"] == "ok":
                entries.append({
                    "entry_id": row["entry_id"],
                    "title": row["title"],
                })
    # 重複除去（同じentry_idが複数回ログされている場合）
    seen = set()
    unique = []
    for e in entries:
        if e["entry_id"] not in seen:
            seen.add(e["entry_id"])
            unique.append(e)
    return unique


def revert_entry(auth, entry_id, dry_run=False):
    """1記事を修復"""
    try:
        entry_xml = get_entry_xml(auth, entry_id)
    except urllib.error.HTTPError as e:
        return "error", "GET HTTP {}".format(e.code)

    raw_content, original_tag = extract_raw_content(entry_xml)
    if raw_content is None:
        return "error", "content抽出失敗"

    # 破損チェック: &amp;lt; (二重エスケープ) or VideoObject の存在
    has_double_escape = "&amp;lt;" in raw_content
    has_video_jsonld = "VideoObject" in raw_content

    if not has_double_escape and not has_video_jsonld:
        return "skip", "修復不要"

    # APIが返すXMLエスケープ済みcontentを2回unescapeして元のraw HTMLに復元
    # 1回目: XMLエスケープ解除（APIの通常処理）
    # 2回目: 当スクリプトがCDATA PUTで追加した余分なエスケープを解除
    restored = html_mod.unescape(html_mod.unescape(raw_content))

    # JSON-LD除去（raw HTML状態）
    restored = remove_video_jsonld(restored)

    if dry_run:
        changes = []
        if has_double_escape:
            changes.append("エスケープ修復")
        if has_video_jsonld:
            changes.append("JSON-LD除去")
        print("    修復前先頭: {}".format(raw_content[:80]))
        print("    修復後先頭: {}".format(restored[:80]))
        # JSON-LD残存チェック
        if "VideoObject" in restored:
            print("    [警告] JSON-LDが除去しきれていない可能性あり")
        return "preview", ", ".join(changes)

    # CDATA内の ]]> をエスケープ（稀だが安全対策）
    safe_content = restored.replace("]]>", "]]]]><![CDATA[>")

    # 新しいcontent要素を構築してPUT
    new_tag = '<content type="text/html"><![CDATA[{}]]></content>'.format(safe_content)
    new_xml = entry_xml.replace(original_tag, new_tag)

    try:
        result = put_entry_xml(auth, entry_id, new_xml)
    except urllib.error.HTTPError as e:
        return "error", "PUT HTTP {}".format(e.code)

    # 検証: ブラウザ表示を確認
    time.sleep(0.5)
    try:
        verify_xml = get_entry_xml(auth, entry_id)
        v_raw, _ = extract_raw_content(verify_xml)
        if v_raw:
            # API GETは1回エスケープして返すので、元のHTMLが正しければ
            # &lt;p&gt;が見えるはず（=正常）。&amp;lt;があれば二重で異常。
            if "&amp;lt;" in v_raw:
                return "warn", "まだ二重エスケープ残存"
            if "VideoObject" in v_raw:
                return "warn", "JSON-LD残存"
    except Exception:
        pass  # 検証失敗は致命的ではない

    url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
    result_url = url_match.group(1) if url_match else "OK"
    return "ok", result_url


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "--test"
    resume_from = 0
    if mode == "--resume" and len(sys.argv) > 2:
        resume_from = int(sys.argv[2])

    dry_run = (mode == "--test")
    is_test = mode in ("--test", "--test-fix")
    test_limit = 3 if is_test else None

    print("=" * 60)
    print("緊急修復: ブログ記事エスケープ破損 + JSON-LD除去")
    print("モード: {}".format(mode))
    print("=" * 60)

    env = load_env()
    auth = get_auth(env)

    entries = load_updated_entries()
    print("修復対象: {} 件".format(len(entries)))

    if resume_from > 0:
        entries = entries[resume_from:]
        print("{}件目から再開 (残り {} 件)".format(resume_from, len(entries)))

    if test_limit:
        entries = entries[:test_limit]
        print("テストモード: {}件のみ処理".format(test_limit))
    print()

    # ログ準備
    log_new = not REVERT_LOG.exists() or REVERT_LOG.stat().st_size == 0
    log_f = open(REVERT_LOG, "a", encoding="utf-8-sig", newline="")
    log_writer = csv.writer(log_f)
    if log_new:
        log_writer.writerow(["timestamp", "entry_id", "title", "status", "detail"])

    stats = {"ok": 0, "skip": 0, "error": 0, "preview": 0, "warn": 0}

    for i, entry in enumerate(entries):
        print("[{}/{}] {}".format(i + 1, len(entries), entry["title"][:50]))
        status, detail = revert_entry(auth, entry["entry_id"], dry_run=dry_run)
        stats[status] = stats.get(status, 0) + 1
        print("  -> {} : {}".format(status, detail))

        if not dry_run:
            log_writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                entry["entry_id"], entry["title"], status, detail,
            ])
            log_f.flush()

        # API負荷対策（修復: GET + PUT + 検証GET = 3リクエスト）
        if not dry_run:
            time.sleep(1.5)
        else:
            time.sleep(0.5)

    log_f.close()

    print("\n" + "=" * 60)
    print("修復サマリー")
    for k, v in stats.items():
        print("  {}: {}".format(k, v))
    if not dry_run:
        print("ログ: {}".format(REVERT_LOG))
    print("=" * 60)


if __name__ == "__main__":
    main()
