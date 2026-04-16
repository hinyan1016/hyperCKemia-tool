#!/usr/bin/env python3
"""
SEO改善 Phase 4: 内部リンク強化
孤立記事（内部リンク0）の末尾に「関連記事」ブロックを自動挿入する。

マッチングロジック:
  - カテゴリ一致数でスコアリング
  - 広すぎるカテゴリは重み0.1に下げる
  - スコア上位5件を関連記事として選定

使い方:
  python add_internal_links.py --test          10件プレビュー（更新しない）
  python add_internal_links.py --test-update   10件実際に更新
  python add_internal_links.py --run           全件実行
  python add_internal_links.py --batch N M     N件目からM件処理
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
AUDIT_CSV = Path(__file__).parent / "audit_report.csv"
LOG_FILE = Path(__file__).parent / "internal_links_log.csv"
BACKUP_DIR = Path(__file__).parent / "backup_links"
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"

# 広すぎるカテゴリ（重み0.1）
BROAD_CATEGORIES = {
    "医学情報", "医師向け", "一般向け", "生活", "薬",
    "脳神経内科", "医学教育", "AI活用",
}

# 関連記事の最大数
MAX_RELATED = 5


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


def load_articles() -> list[dict[str, str]]:
    """全公開記事を読み込み"""
    articles: list[dict[str, str]] = []
    with open(AUDIT_CSV, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["is_draft"] == "False":
                articles.append(dict(row))
    return articles


def get_categories(article: dict[str, str]) -> set[str]:
    """記事のカテゴリをセットで返す"""
    return {c.strip() for c in article["categories"].split("|") if c.strip()}


def calc_relevance(
    target_cats: set[str], candidate_cats: set[str]
) -> float:
    """カテゴリ一致によるスコア計算"""
    score = 0.0
    common = target_cats & candidate_cats
    for cat in common:
        if cat in BROAD_CATEGORIES:
            score += 0.1
        else:
            score += 1.0
    return score


def find_related(
    target: dict[str, str],
    all_articles: list[dict[str, str]],
) -> list[dict[str, str]]:
    """関連記事を最大MAX_RELATED件選定"""
    target_url = target["url"]
    target_cats = get_categories(target)

    if not target_cats:
        return []

    scored: list[tuple[float, dict[str, str]]] = []
    for art in all_articles:
        if art["url"] == target_url:
            continue
        art_cats = get_categories(art)
        score = calc_relevance(target_cats, art_cats)
        if score > 0:
            scored.append((score, art))

    # スコア降順、同スコアなら新しい記事優先
    scored.sort(key=lambda x: (-x[0], x[1]["published"]), reverse=False)
    scored.sort(key=lambda x: x[0], reverse=True)

    return [art for _, art in scored[:MAX_RELATED]]


def build_related_html(related: list[dict[str, str]]) -> str:
    """関連記事HTMLブロックを生成"""
    lines = [
        '\n<div class="related-articles">',
        "<h3>関連記事</h3>",
        "<ul>",
    ]
    for art in related:
        title = art["title"]
        url = art["url"]
        lines.append('<li><a href="{}">{}</a></li>'.format(url, title))
    lines.append("</ul>")
    lines.append("</div>")
    return "\n".join(lines)


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


def process_entry(
    auth: str,
    target: dict[str, str],
    all_articles: list[dict[str, str]],
    dry_run: bool = False,
) -> tuple[str, str]:
    entry_id = target["edit_url"].split("/")[-1]

    # 関連記事を選定
    related = find_related(target, all_articles)
    if not related:
        return "skip", "関連記事なし"

    # GET
    try:
        entry_xml = api_get(auth, entry_id)
    except urllib.error.HTTPError as e:
        return "error", "GET HTTP {}".format(e.code)

    raw, full_match, open_tag = extract_content(entry_xml)
    if raw is None or full_match is None or open_tag is None:
        return "error", "content抽出失敗"

    # XMLエスケープ解除
    restored_html = html_mod.unescape(raw)

    # 既に関連記事ブロックあり → スキップ
    if "related-articles" in restored_html:
        return "skip", "関連記事ブロック既存"

    # バックアップ
    save_backup(entry_id, restored_html)

    # 関連記事HTML生成
    related_html = build_related_html(related)

    if dry_run:
        print("    関連記事:")
        for art in related:
            cats = get_categories(art)
            common = cats & get_categories(target)
            print(
                "      - {} (共通: {})".format(
                    art["title"][:45], ", ".join(common)
                )
            )
        return "preview", "{}件".format(len(related))

    # 末尾に追加
    new_html = restored_html + related_html

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
        if "related-articles" not in v_restored:
            return "WARN", "関連記事ブロック未検出"

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
    test_limit = 10 if is_test else batch_limit

    print("=" * 60)
    print("SEO Phase 4: 内部リンク強化")
    print("モード: {}".format(mode))
    print("=" * 60)

    env = load_env()
    auth = get_auth(env)

    # 全記事読み込み
    all_articles = load_articles()
    print("公開記事: {} 件".format(len(all_articles)))

    # 孤立記事を抽出
    orphans = [a for a in all_articles if int(a["internal_links"]) == 0]
    print("孤立記事: {} 件".format(len(orphans)))

    if resume_from > 0:
        orphans = orphans[resume_from:]
        print("{}件目から再開 (残り {} 件)".format(resume_from, len(orphans)))

    if test_limit:
        orphans = orphans[:test_limit]
        print("処理件数: {} 件".format(len(orphans)))

    print()

    # ログ
    log_new = not LOG_FILE.exists()
    log_f = open(LOG_FILE, "a", encoding="utf-8-sig", newline="")
    log_writer = csv.writer(log_f)
    if log_new:
        log_writer.writerow([
            "timestamp", "entry_id", "title", "status",
            "detail", "related_count",
        ])

    stats: dict[str, int] = {}

    for i, target in enumerate(orphans):
        entry_id = target["edit_url"].split("/")[-1]
        print("[{}/{}] {}".format(i + 1, len(orphans), target["title"][:60]))

        status, detail = process_entry(auth, target, all_articles, dry_run)
        stats[status] = stats.get(status, 0) + 1
        print("  -> {} : {}".format(status, detail))

        if status == "BROKEN":
            print("  *** 破損検出！処理を中断 ***")
            log_writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                entry_id, target["title"], status, detail, 0,
            ])
            log_f.flush()
            log_f.close()
            sys.exit(1)

        if not dry_run:
            related = find_related(target, all_articles)
            log_writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                entry_id, target["title"], status, detail, len(related),
            ])
            log_f.flush()

        # API負荷対策
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
