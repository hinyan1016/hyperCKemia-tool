"""緊急ロールバック：誤って更新した記事を body_html_before から復元する。"""

from __future__ import annotations

import argparse
import re
import sqlite3
import sys
from pathlib import Path
from typing import Callable

from scripts.hatena_client import (
    build_entry_put_xml,
    fetch_entry,
    load_env,
    put_entry,
)


def rollback_entry(
    db_path: Path,
    entry_id: str,
    *,
    api_key: str,
    fetch_fn: Callable = fetch_entry,
    put_fn: Callable = put_entry,
) -> None:
    conn = sqlite3.connect(db_path)
    try:
        row = conn.execute(
            "SELECT body_html_before FROM articles WHERE entry_id=?",
            (entry_id,),
        ).fetchone()
        if not row or not row[0]:
            raise RuntimeError(
                f"entry_id={entry_id} には body_html_before がないのでロールバック不可"
            )
        previous_body = row[0]

        entry_xml = fetch_fn(entry_id, api_key)
        title_match = re.search(r"<title[^>]*>([^<]*)</title>", entry_xml)
        title = title_match.group(1) if title_match else ""
        cats = tuple(re.findall(r'<category[^>]*term="([^"]+)"', entry_xml))

        put_xml = build_entry_put_xml(
            title=title,
            author="hinyan1016",
            body_html=previous_body,
            categories=cats,
        )
        put_fn(entry_id, put_xml, api_key)

        conn.execute(
            """UPDATE articles
            SET status='rolled_back', body_html_before=NULL
            WHERE entry_id=?""",
            (entry_id,),
        )
        conn.commit()
    finally:
        conn.close()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("entry_id")
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    env_file = Path(
        r"C:\Users\jsber\OneDrive\Documents\Claude_task"
        r"\youtube-slides\食事指導シリーズ\_shared\.env"
    )
    api_key = load_env(env_file)["HATENA_API_KEY"]

    ans = input(f"entry_id={args.entry_id} をロールバックしますか? (yes/no): ")
    if ans.strip().lower() not in ("yes", "y"):
        print("中止しました。")
        return 1

    rollback_entry(db_path, args.entry_id, api_key=api_key)
    print("ロールバック完了")
    return 0


if __name__ == "__main__":
    sys.exit(main())
