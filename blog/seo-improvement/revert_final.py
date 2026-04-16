#!/usr/bin/env python3
"""
最終修復スクリプト: 全対象記事を再チェックして確実に修復する。

判定ロジック:
  - APIのcontent先頭が &amp;lt; で始まる → 二重エスケープ（壊れている）
  - content内にVideoObjectがある → JSON-LD除去が必要
  - どちらでもなければ → 正常（skip）

修復:
  - html.unescape 2回で元のraw HTMLを復元
  - JSON-LD除去
  - CDATAでPUT

使い方:
  python revert_final.py --test      5件プレビュー
  python revert_final.py --test-fix  5件修復
  python revert_final.py --run       全件修復
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
VIDEO_LOG = Path(__file__).parent / "video_jsonld_log.csv"
FINAL_LOG = Path(__file__).parent / "revert_final_log.csv"
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


def api_get(auth, entry_id):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, headers={"Authorization": auth})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def api_put(auth, entry_id, xml_body):
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(
        HATENA_ID, BLOG_DOMAIN, entry_id)
    req = urllib.request.Request(url, data=xml_body.encode("utf-8"), method="PUT",
                                headers={
                                    "Content-Type": "application/xml; charset=utf-8",
                                    "Authorization": auth,
                                })
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def extract_content(entry_xml):
    """content要素のraw text（CDATA剥がし済み）と、置換用の全体マッチを返す"""
    m = re.search(r'(<content[^>]*>)(.*?)(</content>)', entry_xml, re.DOTALL)
    if not m:
        return None, None, None
    raw = m.group(2)
    if raw.startswith("<![CDATA[") and raw.endswith("]]>"):
        raw = raw[9:-3]
    return raw, m.group(0), m.group(1)


def is_double_escaped(raw):
    """本文のHTMLタグが二重エスケープされているか判定。
    正常: &lt;p&gt; (APIのXMLエスケープ)
    異常: &amp;lt;p&amp;gt; (二重エスケープ)
    先頭の空白を除いて判定。"""
    stripped = raw.lstrip()
    # 典型的なHTML開始タグの二重エスケープ形式
    if stripped.startswith("&amp;lt;"):
        return True
    # 本文中に大量の &amp;lt; がある場合も判定
    # 正常な記事でも p&lt;0.05 のような表現で少数の &amp;lt; が出る
    # 壊れた記事では数百〜数千の &amp;lt; がある
    amp_lt_count = raw.count("&amp;lt;")
    if amp_lt_count > 50:
        return True
    return False


def has_video_jsonld(raw):
    """VideoObject関連コンテンツが残っているか"""
    return "VideoObject" in raw


def remove_video_jsonld_raw(content):
    """raw HTML状態のVideoObject JSON-LDを除去"""
    fixed = re.sub(
        r'\n?<script\s+type="application/ld\+json">\s*\{[^<]*?"VideoObject"[^<]*?\}\s*</script>',
        '', content, flags=re.DOTALL)
    fixed = re.sub(
        r'\n?<script\s+type="application/ld\+json">\s*\[[^\]]*?"VideoObject"[^\]]*?\]\s*</script>',
        '', fixed, flags=re.DOTALL)
    return fixed


def load_targets():
    """video_jsonld_logからユニークなentry_idリスト取得"""
    entries = []
    seen = set()
    with open(VIDEO_LOG, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["status"] == "ok" and row["entry_id"] not in seen:
                seen.add(row["entry_id"])
                entries.append({"entry_id": row["entry_id"], "title": row["title"]})
    return entries


def fix_entry(auth, entry_id, dry_run=False):
    """1記事を検査・修復"""
    try:
        entry_xml = api_get(auth, entry_id)
    except urllib.error.HTTPError as e:
        return "error", "GET HTTP {}".format(e.code)

    raw, full_match, open_tag = extract_content(entry_xml)
    if raw is None:
        return "error", "content抽出失敗"

    need_unescape = is_double_escaped(raw)
    need_jsonld_remove = has_video_jsonld(raw)

    if not need_unescape and not need_jsonld_remove:
        return "skip", "正常"

    # 2回unescapeで元のraw HTMLに復元
    restored = html_mod.unescape(html_mod.unescape(raw))

    # JSON-LD除去
    restored = remove_video_jsonld_raw(restored)

    if dry_run:
        changes = []
        if need_unescape:
            changes.append("エスケープ修復({})".format(raw.count("&amp;lt;")))
        if need_jsonld_remove:
            changes.append("JSON-LD除去")
        still_has = "VideoObject" in restored
        print("    変更: {}".format(", ".join(changes)))
        print("    修復後先頭: {}".format(restored[:80]))
        if still_has:
            print("    [警告] JSON-LDが残存")
        return "preview", ", ".join(changes)

    # CDATA内の ]]> を安全処理
    safe = restored.replace("]]>", "]]]]><![CDATA[>")
    new_tag = '{}<![CDATA[{}]]></content>'.format(open_tag, safe)
    new_xml = entry_xml.replace(full_match, new_tag)

    try:
        result = api_put(auth, entry_id, new_xml)
    except urllib.error.HTTPError as e:
        return "error", "PUT HTTP {}".format(e.code)

    url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
    return "ok", url_match.group(1) if url_match else "OK"


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "--test"
    resume_from = 0
    if mode == "--resume" and len(sys.argv) > 2:
        resume_from = int(sys.argv[2])

    dry_run = (mode == "--test")
    is_test = mode in ("--test", "--test-fix")
    test_limit = 5 if is_test else None

    print("=" * 60)
    print("最終修復: 全対象記事の再チェック + 修復")
    print("モード: {}".format(mode))
    print("=" * 60)

    env = load_env()
    auth = get_auth(env)

    targets = load_targets()
    print("対象記事: {} 件".format(len(targets)))

    if resume_from > 0:
        targets = targets[resume_from:]
        print("{}件目から再開".format(resume_from))

    if test_limit:
        targets = targets[:test_limit]
        print("テスト: {}件のみ".format(test_limit))
    print()

    # ログ
    log_new = not FINAL_LOG.exists() or FINAL_LOG.stat().st_size == 0
    log_f = open(FINAL_LOG, "w", encoding="utf-8-sig", newline="")
    log_writer = csv.writer(log_f)
    log_writer.writerow(["timestamp", "entry_id", "title", "status", "detail"])

    stats = {"ok": 0, "skip": 0, "error": 0, "preview": 0}

    for i, entry in enumerate(targets):
        print("[{}/{}] {}".format(i + 1, len(targets), entry["title"][:50]))
        status, detail = fix_entry(auth, entry["entry_id"], dry_run=dry_run)
        stats[status] = stats.get(status, 0) + 1
        print("  -> {} : {}".format(status, detail))

        log_writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            entry["entry_id"], entry["title"], status, detail,
        ])
        log_f.flush()

        if not dry_run and status == "ok":
            time.sleep(1.0)
        else:
            time.sleep(0.3)

    log_f.close()

    print("\n" + "=" * 60)
    print("最終修復サマリー")
    for k, v in sorted(stats.items()):
        print("  {}: {}".format(k, v))
    print("ログ: {}".format(FINAL_LOG))
    print("=" * 60)


if __name__ == "__main__":
    main()
