"""decisions.csv を読んで OK/NG/Skip を反映する処理スクリプト。

idempotencyを processed_decisions テーブルで保証。
"""

from __future__ import annotations

import csv
import re
import shutil
import sqlite3
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from scripts.hatena_client import (
    build_entry_put_xml,
    fetch_entry,
    put_entry,
    upload_to_fotolife,
)
from scripts.html_inserter import insert_infographic


@dataclass
class ProcessResult:
    uploaded: int = 0
    ng_recorded: int = 0
    skipped: int = 0
    already_processed: int = 0
    failed: int = 0


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _build_img_tag(image_url: str, alt_text: str) -> str:
    return (
        f'<img src="{image_url}" alt="{alt_text}" '
        'style="max-width:100%; height:auto; border-radius:8px; '
        'box-shadow:0 2px 8px rgba(0,0,0,0.1);" loading="lazy">'
    )


def _extract_current_body(entry_xml: str) -> tuple[str, str, tuple[str, ...]]:
    """既存entry XMLから (title, body_html, categories) を抽出。"""
    title_match = re.search(r"<title[^>]*>([^<]*)</title>", entry_xml)
    title = title_match.group(1) if title_match else ""
    cats = tuple(re.findall(r'<category[^>]*term="([^"]+)"', entry_xml))
    content_match = re.search(
        r'<content[^>]*type="text/html"[^>]*>(.*?)</content>',
        entry_xml,
        re.DOTALL,
    )
    body = ""
    if content_match:
        raw = content_match.group(1).strip()
        if raw.startswith("<![CDATA["):
            raw = raw[len("<![CDATA[") :]
            if raw.endswith("]]>"):
                raw = raw[:-3]
        body = (
            raw.replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&amp;", "&")
            .replace("&quot;", '"')
            .replace("&#39;", "'")
        )
    return title, body, cats


def process_decisions(
    db_path: Path,
    decisions_csv: Path,
    images_in_dir: Path,
    images_archive_dir: Path,
    *,
    api_key: str,
    upload_fn: Callable = upload_to_fotolife,
    fetch_fn: Callable = fetch_entry,
    put_fn: Callable = put_entry,
    dry_run: bool = False,
) -> ProcessResult:
    result = ProcessResult()
    conn = sqlite3.connect(db_path)
    csv_id = str(decisions_csv.resolve())

    try:
        with decisions_csv.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))

        for row in rows:
            eid = row["entry_id"]
            decision = row["decision"]

            already = conn.execute(
                """SELECT 1 FROM processed_decisions
                WHERE decisions_csv_path=? AND entry_id=?""",
                (csv_id, eid),
            ).fetchone()
            if already:
                result.already_processed += 1
                continue

            try:
                if decision == "ok":
                    _handle_ok(
                        conn,
                        eid,
                        images_in_dir,
                        images_archive_dir,
                        api_key=api_key,
                        upload_fn=upload_fn,
                        fetch_fn=fetch_fn,
                        put_fn=put_fn,
                        dry_run=dry_run,
                    )
                    result.uploaded += 1
                elif decision == "ng":
                    _handle_ng(
                        conn,
                        eid,
                        row.get("ng_tags", ""),
                        row.get("ng_note", ""),
                        images_in_dir,
                        images_archive_dir,
                        dry_run=dry_run,
                    )
                    result.ng_recorded += 1
                elif decision == "skip":
                    result.skipped += 1
                else:
                    continue

                if not dry_run:
                    conn.execute(
                        """INSERT OR IGNORE INTO processed_decisions
                        (decisions_csv_path, entry_id, decision, processed_at)
                        VALUES (?, ?, ?, ?)""",
                        (csv_id, eid, decision, _now()),
                    )
                    conn.commit()
            except Exception as exc:  # noqa: BLE001
                result.failed += 1
                print(f"[ERROR] entry={eid}: {exc}", file=sys.stderr)
                conn.rollback()
    finally:
        conn.close()

    return result


def _handle_ok(
    conn: sqlite3.Connection,
    entry_id: str,
    images_in: Path,
    archive: Path,
    *,
    api_key: str,
    upload_fn: Callable,
    fetch_fn: Callable,
    put_fn: Callable,
    dry_run: bool,
) -> None:
    gen_row = conn.execute(
        """SELECT alt_text, fotolife_url FROM generations
        WHERE entry_id=? ORDER BY attempt DESC LIMIT 1""",
        (entry_id,),
    ).fetchone()
    if not gen_row:
        raise RuntimeError(f"generations に {entry_id} がない")
    alt_text, existing_url = gen_row

    image_path = images_in / f"{entry_id}.png"
    if not image_path.exists():
        raise FileNotFoundError(str(image_path))

    if dry_run:
        print(f"[DRY-RUN] would upload {image_path}")
        print(f"[DRY-RUN] would PUT entry {entry_id}")
        return

    image_url = existing_url or upload_fn(image_path, api_key)
    if str(image_url).startswith(("ERROR", "URL not found")):
        raise RuntimeError(f"Fotolife upload failed: {image_url}")

    entry_xml = fetch_fn(entry_id, api_key)
    title, body_html, categories = _extract_current_body(entry_xml)

    conn.execute(
        "UPDATE articles SET body_html_before=? WHERE entry_id=?",
        (body_html, entry_id),
    )

    img_tag = _build_img_tag(image_url, alt_text)
    new_body = insert_infographic(body_html, img_tag)

    put_xml = build_entry_put_xml(
        title=title,
        author="hinyan1016",
        body_html=new_body,
        categories=categories,
    )
    put_fn(entry_id, put_xml, api_key)

    archive.mkdir(parents=True, exist_ok=True)
    shutil.move(str(image_path), str(archive / f"{entry_id}.png"))

    conn.execute(
        """UPDATE generations
        SET fotolife_url=?, processed_at=?, decision='approved'
        WHERE entry_id=?
        AND gen_id = (
            SELECT gen_id FROM generations
            WHERE entry_id=? ORDER BY attempt DESC LIMIT 1
        )""",
        (image_url, _now(), entry_id, entry_id),
    )
    conn.execute(
        "UPDATE articles SET status='done' WHERE entry_id=?",
        (entry_id,),
    )


def _handle_ng(
    conn: sqlite3.Connection,
    entry_id: str,
    ng_tags: str,
    ng_note: str,
    images_in: Path,
    archive: Path,
    *,
    dry_run: bool,
) -> None:
    if dry_run:
        print(f"[DRY-RUN] would record NG for {entry_id}: tags={ng_tags}")
        return

    archive.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    image_path = images_in / f"{entry_id}.png"
    if image_path.exists():
        shutil.move(str(image_path), str(archive / f"{entry_id}_ng_{ts}.png"))

    conn.execute(
        """UPDATE generations
        SET decision='ng', ng_tags=?, ng_note=?, processed_at=?
        WHERE entry_id=?
        AND gen_id = (
            SELECT gen_id FROM generations
            WHERE entry_id=? ORDER BY attempt DESC LIMIT 1
        )""",
        (ng_tags, ng_note, _now(), entry_id, entry_id),
    )
    conn.execute(
        "UPDATE articles SET status='ng_pending', last_ng_date=? WHERE entry_id=?",
        (_now(), entry_id),
    )


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("decisions_csv")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--yes", action="store_true")
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    images_in = root / "images" / "in"
    archive = root / "images" / "archive"

    from scripts.hatena_client import load_env

    env_file = Path(
        r"C:\Users\jsber\OneDrive\Documents\Claude_task"
        r"\youtube-slides\食事指導シリーズ\_shared\.env"
    )
    api_key = load_env(env_file)["HATENA_API_KEY"]

    if not args.dry_run and not args.yes:
        ans = input("本番のはてなブログを更新します。続行しますか? (yes/no): ")
        if ans.strip().lower() not in ("yes", "y"):
            print("中止しました。")
            return 1

    result = process_decisions(
        db_path,
        Path(args.decisions_csv),
        images_in,
        archive,
        api_key=api_key,
        dry_run=args.dry_run,
    )
    print(
        f"完了: OK={result.uploaded} NG={result.ng_recorded} "
        f"Skip={result.skipped} 既処理={result.already_processed} "
        f"失敗={result.failed}"
    )
    return 0 if result.failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
