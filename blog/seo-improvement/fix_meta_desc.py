#!/usr/bin/env python3
"""
SEO改善 Phase 5: メタディスクリプション最適化
YouTube URL が meta description に混入している記事の先頭に
リード文を挿入し、はてなの自動生成 description を改善する。

問題: 記事先頭が「タイトル → YouTube iframe」の構成のため、
はてなが自動生成する description に「youtu.be」が混入する。

対策: YouTube iframe の直前にリード文を挿入する。
リード文は記事本文の見出し（h2/h3）から自動構成する。

使い方:
  python fix_meta_desc.py --test          5件プレビュー
  python fix_meta_desc.py --test-update   5件実際に更新
  python fix_meta_desc.py --run           全件実行
  python fix_meta_desc.py --batch N M     N件目からM件処理
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

sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)

ENV_FILE = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task"
    r"\youtube-slides\食事指導シリーズ\_shared\.env"
)
PROBLEMS_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\scripts\meta_desc_problems.txt")
AUDIT_CSV = Path(__file__).parent / "audit_report.csv"
LOG_FILE = Path(__file__).parent / "meta_desc_log.csv"
BACKUP_DIR = Path(__file__).parent / "backup_meta"
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


def load_problem_urls() -> list[dict[str, str]]:
    """問題記事リストからURL・タイトルを抽出"""
    with open(PROBLEMS_FILE, encoding="utf-8") as f:
        text = f.read()
    entries = re.findall(
        r"(\d+)\.\s+(.+?)\n\s+URL:\s+(.+?)\n\s+DESC:\s+(.+?)(?=\n\n|\Z)",
        text, re.DOTALL,
    )
    results: list[dict[str, str]] = []
    for num, title, url, desc in entries:
        results.append({
            "num": num.strip(),
            "title": title.strip(),
            "url": url.strip(),
            "desc": desc.strip(),
        })
    return results


def load_url_to_entry_id() -> dict[str, str]:
    """audit CSVからURL→entry_idのマッピングを作成"""
    mapping: dict[str, str] = {}
    with open(AUDIT_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            eid = row["edit_url"].split("/")[-1]
            mapping[row["url"]] = eid
    return mapping


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
        url, data=xml_body.encode("utf-8"), method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth,
        },
    )
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def extract_content(
    entry_xml: str,
) -> tuple[str | None, str | None, str | None]:
    m = re.search(r"(<content[^>]*>)(.*?)(</content>)", entry_xml, re.DOTALL)
    if not m:
        return None, None, None
    raw = m.group(2)
    if raw.startswith("<![CDATA[") and raw.endswith("]]>"):
        raw = raw[9:-3]
    return raw, m.group(0), m.group(1)


def save_backup(entry_id: str, content: str) -> None:
    BACKUP_DIR.mkdir(exist_ok=True)
    path = BACKUP_DIR / "{}.html".format(entry_id)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


def extract_headings(html_content: str) -> list[str]:
    """h2/h3見出しテキストを抽出"""
    headings = re.findall(r"<h[23][^>]*>(.*?)</h[23]>", html_content, re.DOTALL)
    result: list[str] = []
    for h in headings:
        text = re.sub(r"<[^>]+>", "", h).strip()
        # 絵文字のUnicode除去
        text = re.sub(r"[\U0001F300-\U0001F9FF]", "", text).strip()
        if text and len(text) > 2:
            result.append(text)
    return result


def extract_first_text(html_content: str) -> str:
    """最初の意味のあるテキスト段落を抽出（iframe/img/style/scriptを除く）"""
    # pタグからテキストを抽出
    paragraphs = re.findall(r"<p[^>]*>(.*?)</p>", html_content, re.DOTALL)
    for p in paragraphs:
        # iframe, img, script, style を含む段落はスキップ
        if re.search(r"<(iframe|img|script|style)", p):
            continue
        text = re.sub(r"<[^>]+>", "", p).strip()
        # 空白や短すぎるテキストはスキップ
        if len(text) > 20 and "youtu" not in text and "docs.google" not in text:
            return text
    return ""


def generate_lead(title: str, headings: list[str], first_text: str) -> str:
    """タイトルと見出しからリード文を生成"""
    # 見出しから主要トピックを取得（番号プレフィックスを除去）
    topics: list[str] = []
    for h in headings:
        clean = re.sub(r"^\d+[\.\s、]+", "", h).strip()
        if clean and clean != title and len(clean) > 2:
            topics.append(clean)

    # first_textがタイトルとほぼ同じなら無視
    use_first_text = False
    if first_text and len(first_text) > 30:
        # タイトルとの重複チェック（先頭20文字が一致しなければOK）
        if first_text[:20] != title[:20]:
            use_first_text = True

    if use_first_text:
        lead = first_text
        if len(lead) > 150:
            cut = lead[:150]
            last_period = max(cut.rfind("。"), cut.rfind("）"), cut.rfind("す"))
            if last_period > 50:
                lead = lead[:last_period + 1]
            else:
                lead = cut + "…"
        return lead

    if len(topics) >= 2:
        topic_str = "・".join(topics[:3])
        lead = "この記事では、{}について解説します。".format(topic_str)
        return lead

    # フォールバック: タイトルベース
    return "{}について、最新のエビデンスに基づいて解説します。".format(title)


def find_youtube_iframe_pos(html_content: str) -> int:
    """最初のYouTube iframeの開始位置を返す（見つからなければ-1）"""
    m = re.search(
        r"<(p[^>]*>)?\s*<iframe[^>]*(youtube|youtu\.be)",
        html_content, re.IGNORECASE,
    )
    if m:
        return m.start()

    # パターン2: iframeがpタグ内でなく直接ある場合
    m2 = re.search(r"<iframe[^>]*(youtube|youtu\.be)", html_content, re.IGNORECASE)
    if m2:
        return m2.start()

    # パターン3: hatenablog-partsのiframe（Google Docsなど）
    m3 = re.search(r"<p>\s*<iframe[^>]*hatenablog-parts", html_content)
    if m3:
        return m3.start()

    return -1


def has_lead_marker(html_content: str) -> bool:
    """既にリード文マーカーがあるか"""
    return 'class="seo-lead"' in html_content


def process_entry(
    auth: str,
    entry: dict[str, str],
    entry_id: str,
    dry_run: bool = False,
) -> tuple[str, str]:
    # GET
    try:
        entry_xml = api_get(auth, entry_id)
    except urllib.error.HTTPError as e:
        return "error", "GET HTTP {}".format(e.code)

    raw, full_match, open_tag = extract_content(entry_xml)
    if raw is None or full_match is None or open_tag is None:
        return "error", "content抽出失敗"

    restored_html = html_mod.unescape(raw)

    # 既にリード文あり
    if has_lead_marker(restored_html):
        return "skip", "リード文既存"

    # YouTube iframeの位置を特定
    iframe_pos = find_youtube_iframe_pos(restored_html)
    if iframe_pos < 0:
        return "skip", "YouTube iframe未検出"

    # 見出しとテキストを抽出
    headings = extract_headings(restored_html)
    first_text = extract_first_text(restored_html)

    # リード文生成
    lead = generate_lead(entry["title"], headings, first_text)

    # リード文HTML
    lead_html = '<p class="seo-lead" style="font-size: 1.05em; color: #333; margin-bottom: 1em;">{}</p>\n'.format(lead)

    if dry_run:
        print("    リード文: {}".format(lead[:80]))
        print("    iframe位置: {}文字目".format(iframe_pos))
        return "preview", lead[:60]

    # バックアップ
    save_backup(entry_id, restored_html)

    # iframe前にリード文を挿入
    new_html = restored_html[:iframe_pos] + lead_html + restored_html[iframe_pos:]

    # CDATA処理
    safe_html = new_html.replace("]]>", "]]]]><![CDATA[>")
    new_tag = "{}<![CDATA[{}]]></content>".format(open_tag, safe_html)
    new_xml = entry_xml.replace(full_match, new_tag)

    try:
        result = api_put(auth, entry_id, new_xml)
    except urllib.error.HTTPError as e:
        return "error", "PUT HTTP {}".format(e.code)

    # 検証
    time.sleep(0.5)
    try:
        verify_xml = api_get(auth, entry_id)
    except urllib.error.HTTPError:
        return "ok", "更新OK（検証スキップ）"

    v_raw, _, _ = extract_content(verify_xml)
    if v_raw:
        v_restored = html_mod.unescape(v_raw)
        if "&lt;" in v_restored[:300]:
            return "BROKEN", "二重エスケープ検出"
        if 'class="seo-lead"' not in v_restored:
            return "WARN", "リード文未検出"

    url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', result)
    return "ok", url_match.group(1) if url_match else "OK"


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
    test_limit = 5 if is_test else batch_limit

    print("=" * 60)
    print("SEO Phase 5: メタディスクリプション最適化")
    print("モード: {}".format(mode))
    print("=" * 60)

    env = load_env()
    auth = get_auth(env)

    # 問題記事読み込み
    problems = load_problem_urls()
    url_to_id = load_url_to_entry_id()
    print("問題記事: {} 件".format(len(problems)))

    if resume_from > 0:
        problems = problems[resume_from:]
        print("{}件目から再開 (残り {} 件)".format(resume_from, len(problems)))

    if test_limit:
        problems = problems[:test_limit]
        print("処理件数: {} 件".format(len(problems)))

    print()

    # ログ
    log_new = not LOG_FILE.exists()
    log_f = open(LOG_FILE, "a", encoding="utf-8-sig", newline="")
    log_writer = csv.writer(log_f)
    if log_new:
        log_writer.writerow([
            "timestamp", "entry_id", "title", "url",
            "status", "detail",
        ])

    stats: dict[str, int] = {}

    for i, entry in enumerate(problems):
        entry_id = url_to_id.get(entry["url"], "")
        if not entry_id:
            print("[{}/{}] {} — ID不明、スキップ".format(
                i + 1, len(problems), entry["title"][:50]))
            stats["skip_no_id"] = stats.get("skip_no_id", 0) + 1
            continue

        print("[{}/{}] {}".format(i + 1, len(problems), entry["title"][:60]))

        status, detail = process_entry(auth, entry, entry_id, dry_run)
        stats[status] = stats.get(status, 0) + 1
        print("  -> {} : {}".format(status, detail))

        if status == "BROKEN":
            print("  *** 破損検出！中断 ***")
            log_writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                entry_id, entry["title"], entry["url"], status, detail,
            ])
            log_f.flush()
            log_f.close()
            sys.exit(1)

        if not dry_run:
            log_writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                entry_id, entry["title"], entry["url"], status, detail,
            ])
            log_f.flush()

        if not dry_run and status == "ok":
            time.sleep(1.5)
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
