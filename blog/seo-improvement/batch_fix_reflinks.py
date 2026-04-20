"""
監査ログ audit_reflinks_v2.log から要対応記事URLを抽出し、
fix_reflinks.py のロジックを各記事に適用する。
"""

from __future__ import annotations

import csv
import logging
import re
import sys
import time
from pathlib import Path
from typing import Final

sys.path.insert(0, str(Path(__file__).parent))
import fix_reflinks  # type: ignore

AUDIT_LOG: Final = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\seo-improvement\audit_reflinks_v2.log"
)
RESULT_CSV: Final = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\seo-improvement\batch_fix_reflinks_log.csv"
)
URL_PATTERN = re.compile(r"https://hinyan1016\.hatenablog\.com/entry/\S+")

log = logging.getLogger("batch_fix_reflinks")


def extract_urls(log_path: Path) -> list[str]:
    text = log_path.read_text(encoding="utf-8")
    return list(dict.fromkeys(URL_PATTERN.findall(text)))


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    env = fix_reflinks.load_env()
    urls = extract_urls(AUDIT_LOG)
    log.info(f"対象 {len(urls)} 記事")

    with RESULT_CSV.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["timestamp", "url", "status", "linked", "skipped", "error"])
        for i, url in enumerate(urls, 1):
            log.info(f"--- [{i}/{len(urls)}] {url}")
            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            try:
                info = fix_reflinks.find_entry(env, url)
                if info is None:
                    writer.writerow([ts, url, "not_found", 0, 0, ""])
                    continue
                new_content, linked, skipped = fix_reflinks.process_content(info.content)
                if linked == 0:
                    log.info(f"  付与0件スキップ（残スキップ {skipped}）")
                    writer.writerow([ts, url, "no_target", 0, skipped, ""])
                    continue
                ok = fix_reflinks.put_entry(env, info, new_content)
                status = "updated" if ok else "put_failed"
                log.info(f"  付与 {linked}件 / スキップ {skipped}件 / {status}")
                writer.writerow([ts, url, status, linked, skipped, ""])
            except Exception as e:  # noqa: BLE001
                log.error(f"  ERROR: {e}")
                writer.writerow([ts, url, "error", 0, 0, str(e)[:200]])
            # はてなAPIレート配慮
            time.sleep(1.5)


if __name__ == "__main__":
    main()
