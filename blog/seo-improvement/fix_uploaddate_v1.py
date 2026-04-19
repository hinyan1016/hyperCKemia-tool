#!/usr/bin/env python3
"""Phase C'/D': VideoObject JSON-LD uploadDate を ISO 8601+JST 形式に修正.

Phase B の診断で、GSC URL検査が uploadDate に対して次の警告を出すことが判明:
    「uploadDate」の日時値が無効です（任意）
    日時プロパティ「uploadDate」にタイムゾーンがありません（任意）

原因: add_video_jsonld_v2.py 時点で uploadDate を "YYYY-MM-DD" 形式で
生成していた。Schema.org は ISO 8601 datetime+TZ を期待する。

変換例:
    "uploadDate": "2026-04-16"  →  "uploadDate": "2026-04-16T00:00:00+09:00"

処理フロー:
    1. video_iframe_audit.csv から has_jsonld=True の記事を抽出
    2. AtomPub GET で記事 XML を取得
    3. content 要素を html.unescape() してHTML復元
    4. JSON-LD 内の uploadDate を正規表現で置換
    5. 既に ISO 8601+TZ 形式 (T含む) ならスキップ
    6. CDATA で包んで PUT、GET で破損チェック

使い方:
    python fix_uploaddate_v1.py --test            # 2件dry-run（PUTなし）
    python fix_uploaddate_v1.py --test-update     # 2件実更新
    python fix_uploaddate_v1.py --target URL      # URL 指定で1件更新
    python fix_uploaddate_v1.py --batch N M       # N件目からM件処理
    python fix_uploaddate_v1.py --run             # 全件処理
"""
from __future__ import annotations

import argparse
import base64
import csv
import html as html_mod
import logging
import re
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger(__name__)

ENV_FILE = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task"
    r"\youtube-slides\食事指導シリーズ\_shared\.env"
)
AUDIT_CSV = Path(__file__).parent / "video_iframe_audit.csv"
LOG_FILE = Path(__file__).parent / "fix_uploaddate_log.csv"
BACKUP_DIR = Path(__file__).parent / "backup_uploaddate"
HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
JST_SUFFIX = "T00:00:00+09:00"

# uploadDate置換用正規表現
# 既存: "uploadDate": "2026-04-16" または "uploadDate":"2026-04-16"
# 置換先: "uploadDate": "2026-04-16T00:00:00+09:00"
# 既にT入り（ISO 8601）の場合はマッチしない
RE_UPLOADDATE_DATE_ONLY = re.compile(
    r'("uploadDate"\s*:\s*")(\d{4}-\d{2}-\d{2})(")'
)


@dataclass(frozen=True)
class Target:
    url: str
    title: str
    published: str
    edit_url: str


@dataclass(frozen=True)
class UpdateResult:
    entry_id: str
    status: str  # ok / skip_no_jsonld / skip_already_iso / error
    detail: str
    before_count: int = 0
    after_count: int = 0
    verify: str = ""


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def make_auth(env: dict[str, str]) -> str:
    token = base64.b64encode(
        f"{HATENA_ID}:{env['HATENA_API_KEY']}".encode()
    ).decode()
    return f"Basic {token}"


def load_targets(target_url: str | None = None) -> list[Target]:
    """JSON-LDを含む記事（公開のみ）を抽出."""
    targets: list[Target] = []
    with open(AUDIT_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if row["is_draft"] == "True":
                continue
            if row["has_jsonld"] != "True":
                continue
            if target_url and row["url"] != target_url:
                continue
            targets.append(
                Target(
                    url=row["url"],
                    title=row["title"],
                    published=row["published"],
                    edit_url=row["edit_url"],
                )
            )
    return targets


def api_get(auth: str, entry_id: str) -> str:
    url = (
        f"https://blog.hatena.ne.jp/{HATENA_ID}/{BLOG_DOMAIN}"
        f"/atom/entry/{entry_id}"
    )
    req = urllib.request.Request(url, headers={"Authorization": auth})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def api_put(auth: str, entry_id: str, xml_body: str) -> str:
    url = (
        f"https://blog.hatena.ne.jp/{HATENA_ID}/{BLOG_DOMAIN}"
        f"/atom/entry/{entry_id}"
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


def extract_content(
    entry_xml: str,
) -> tuple[str | None, str | None, str | None]:
    """content要素を抽出. (raw_inner, full_match, open_tag) を返す."""
    m = re.search(r"(<content[^>]*>)(.*?)(</content>)", entry_xml, re.DOTALL)
    if not m:
        return None, None, None
    raw = m.group(2)
    if raw.startswith("<![CDATA[") and raw.endswith("]]>"):
        raw = raw[9:-3]
    return raw, m.group(0), m.group(1)


def replace_uploaddate(html: str) -> tuple[str, int]:
    """uploadDate を ISO 8601+JST に置換. (新HTML, 置換件数) を返す."""
    count = [0]

    def _repl(m: re.Match[str]) -> str:
        count[0] += 1
        return f'{m.group(1)}{m.group(2)}{JST_SUFFIX}{m.group(3)}'

    new_html = RE_UPLOADDATE_DATE_ONLY.sub(_repl, html)
    return new_html, count[0]


def save_backup(entry_id: str, content: str) -> None:
    BACKUP_DIR.mkdir(exist_ok=True)
    (BACKUP_DIR / f"{entry_id}.html").write_text(content, encoding="utf-8")


def verify_entry(auth: str, entry_id: str) -> tuple[str, str]:
    try:
        entry_xml = api_get(auth, entry_id)
    except urllib.error.HTTPError as exc:
        return "error", f"GET HTTP {exc.code}"
    raw, _, _ = extract_content(entry_xml)
    if raw is None:
        return "error", "content抽出失敗"
    restored = html_mod.unescape(raw)
    if "&lt;" in restored[:200]:
        return "BROKEN", "二重エスケープ検出"
    if '"VideoObject"' not in restored:
        return "WARN", "VideoObject消失"
    # ISO 8601 形式に切り替わったか確認
    if re.search(r'"uploadDate"\s*:\s*"\d{4}-\d{2}-\d{2}T', restored):
        return "ok", "ISO 8601確認"
    return "WARN", "uploadDate変換未検出"


def process_one(
    auth: str, target: Target, dry_run: bool
) -> UpdateResult:
    entry_id = target.edit_url.rsplit("/", 1)[-1]

    try:
        entry_xml = api_get(auth, entry_id)
    except urllib.error.HTTPError as exc:
        return UpdateResult(entry_id, "error", f"GET HTTP {exc.code}")

    raw, full_match, open_tag = extract_content(entry_xml)
    if raw is None or full_match is None or open_tag is None:
        return UpdateResult(entry_id, "error", "content抽出失敗")

    restored = html_mod.unescape(raw)
    if '"VideoObject"' not in restored:
        return UpdateResult(entry_id, "skip_no_jsonld", "JSON-LD無し")

    # 既にISO 8601+T 形式なら何もしない
    before_iso = len(re.findall(
        r'"uploadDate"\s*:\s*"\d{4}-\d{2}-\d{2}T', restored
    ))
    new_html, replaced = replace_uploaddate(restored)
    if replaced == 0:
        if before_iso > 0:
            return UpdateResult(entry_id, "skip_already_iso", f"既にISO形式 ({before_iso}件)")
        return UpdateResult(entry_id, "skip_no_match", "置換対象なし")

    save_backup(entry_id, restored)

    if dry_run:
        return UpdateResult(
            entry_id,
            "preview",
            f"置換 {replaced}件",
            before_count=replaced,
            after_count=replaced,
        )

    safe_html = new_html.replace("]]>", "]]]]><![CDATA[>")
    new_tag = f"{open_tag}<![CDATA[{safe_html}]]></content>"
    new_xml = entry_xml.replace(full_match, new_tag)

    try:
        api_put(auth, entry_id, new_xml)
    except urllib.error.HTTPError as exc:
        return UpdateResult(entry_id, "error", f"PUT HTTP {exc.code}")

    time.sleep(1.0)
    verify_status, verify_detail = verify_entry(auth, entry_id)
    return UpdateResult(
        entry_id,
        "ok",
        f"置換 {replaced}件",
        before_count=replaced,
        after_count=replaced,
        verify=f"{verify_status}: {verify_detail}",
    )


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Fix VideoObject uploadDate format")
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--test", action="store_true", help="2件 dry-run")
    mode.add_argument("--test-update", action="store_true", help="2件 実更新")
    mode.add_argument("--batch", nargs=2, type=int, metavar=("START", "COUNT"))
    mode.add_argument("--run", action="store_true", help="全件更新")
    mode.add_argument("--target", type=str, help="URL指定で1件のみ実更新")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    env = load_env()
    auth = make_auth(env)

    if args.target:
        targets = load_targets(target_url=args.target)
        dry_run = False
    else:
        targets = load_targets()
        if args.test:
            targets = targets[:2]
            dry_run = True
        elif args.test_update:
            targets = targets[:2]
            dry_run = False
        elif args.batch:
            start, count = args.batch
            targets = targets[start : start + count]
            dry_run = False
        elif args.run:
            dry_run = False
        else:
            # デフォルトは --test
            targets = targets[:2]
            dry_run = True

    logger.info(
        "mode=%s dry_run=%s targets=%d",
        "target" if args.target else (
            "test" if args.test or (not any([args.test_update, args.batch, args.run, args.target])) else
            "test-update" if args.test_update else
            "batch" if args.batch else
            "run"
        ),
        dry_run,
        len(targets),
    )

    # ログファイル準備
    log_new = not LOG_FILE.exists()
    with open(LOG_FILE, "a", encoding="utf-8-sig", newline="") as log_f:
        writer = csv.writer(log_f)
        if log_new:
            writer.writerow([
                "timestamp",
                "entry_id",
                "url",
                "title",
                "status",
                "detail",
                "replaced",
                "verify",
            ])

        stats: dict[str, int] = {}
        for i, t in enumerate(targets, 1):
            logger.info("[%d/%d] %s", i, len(targets), t.title[:50])
            result = process_one(auth, t, dry_run=dry_run)
            stats[result.status] = stats.get(result.status, 0) + 1
            logger.info("  -> %s: %s %s", result.status, result.detail, result.verify)

            if result.status == "ok" and "BROKEN" in result.verify:
                logger.error("破損検出。処理中断。")
                writer.writerow([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    result.entry_id, t.url, t.title,
                    result.status, result.detail,
                    result.before_count, result.verify,
                ])
                log_f.flush()
                return 1

            if not dry_run:
                writer.writerow([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    result.entry_id, t.url, t.title,
                    result.status, result.detail,
                    result.before_count, result.verify,
                ])
                log_f.flush()

            time.sleep(1.0 if not dry_run else 0.3)

    print()
    logger.info("=== サマリー ===")
    for k, v in sorted(stats.items()):
        logger.info("  %s: %d", k, v)
    return 0


if __name__ == "__main__":
    sys.exit(main())
