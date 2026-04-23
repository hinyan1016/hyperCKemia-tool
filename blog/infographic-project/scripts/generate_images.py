"""OpenAI Images API (gpt-image-1) で prompts.csv の全件画像を自動生成する。

- data/batches/<today>_prompts.csv を読み、prompt が空でない行を処理
- 既に images/in/<entry_id>.png が存在する行はスキップ（idempotency）
- 生成失敗はログに記録して次行へ（途中で止めない）
- DB generations.image_path / created_at を更新
"""

from __future__ import annotations

import base64
import csv
import sqlite3
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Callable, Protocol


DEFAULT_MODEL = "gpt-image-1"
DEFAULT_SIZE = "1024x1024"
DEFAULT_QUALITY = "medium"
DEFAULT_MAX_RETRIES = 2
DEFAULT_RETRY_DELAY_S = 4.0


class ImageGenerator(Protocol):
    """OpenAI クライアントの最小インターフェース（テスト差し替え用）。"""

    def __call__(
        self,
        prompt: str,
        *,
        model: str,
        size: str,
        quality: str,
    ) -> bytes: ...


@dataclass
class GenerateResult:
    total: int = 0
    generated: int = 0
    skipped_existing: int = 0
    skipped_empty_prompt: int = 0
    failed: int = 0
    failures: list[tuple[str, str]] = None  # (entry_id, error_message)

    def __post_init__(self) -> None:
        if self.failures is None:
            self.failures = []


def _now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _default_openai_generator(api_key: str) -> ImageGenerator:
    """本番用: OpenAI SDK を用いて画像バイト列を返す関数を構築。

    gpt-image-1 は b64_json を返すので base64 デコードして返す。
    """
    from openai import OpenAI

    client = OpenAI(api_key=api_key)

    def generate(
        prompt: str,
        *,
        model: str,
        size: str,
        quality: str,
    ) -> bytes:
        resp = client.images.generate(
            model=model,
            prompt=prompt,
            size=size,
            quality=quality,
            n=1,
        )
        b64 = resp.data[0].b64_json
        if not b64:
            raise RuntimeError("OpenAI response missing b64_json")
        return base64.b64decode(b64)

    return generate


def generate_images_for_batch(
    prompts_csv: Path,
    images_in_dir: Path,
    db_path: Path,
    *,
    generator: ImageGenerator,
    model: str = DEFAULT_MODEL,
    size: str = DEFAULT_SIZE,
    quality: str = DEFAULT_QUALITY,
    max_retries: int = DEFAULT_MAX_RETRIES,
    retry_delay_s: float = DEFAULT_RETRY_DELAY_S,
    progress_cb: Callable[[str], None] | None = None,
    log_path: Path | None = None,
) -> GenerateResult:
    """prompts.csv の各行について画像を生成して images/in/ に保存。"""
    result = GenerateResult()
    images_in_dir.mkdir(parents=True, exist_ok=True)

    with prompts_csv.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    result.total = len(rows)

    conn = sqlite3.connect(db_path)
    try:
        for row in rows:
            eid = row["entry_id"].strip()
            prompt = (row.get("prompt") or "").strip()
            out_path = images_in_dir / f"{eid}.png"

            if not prompt:
                result.skipped_empty_prompt += 1
                if progress_cb:
                    progress_cb(f"[SKIP empty_prompt] {eid}")
                continue

            if out_path.exists():
                result.skipped_existing += 1
                if progress_cb:
                    progress_cb(f"[SKIP existing] {eid}")
                continue

            attempt = 0
            last_err: Exception | None = None
            img_bytes: bytes | None = None
            while attempt <= max_retries:
                try:
                    img_bytes = generator(
                        prompt, model=model, size=size, quality=quality
                    )
                    break
                except Exception as exc:  # noqa: BLE001
                    last_err = exc
                    attempt += 1
                    if attempt <= max_retries:
                        if progress_cb:
                            progress_cb(
                                f"[RETRY {attempt}/{max_retries}] {eid}: {exc}"
                            )
                        time.sleep(retry_delay_s)

            if img_bytes is None:
                result.failed += 1
                msg = repr(last_err) if last_err else "unknown"
                result.failures.append((eid, msg))
                if log_path:
                    log_path.parent.mkdir(parents=True, exist_ok=True)
                    with log_path.open("a", encoding="utf-8") as lf:
                        lf.write(f"{_now_iso()}\t{eid}\t{msg}\n")
                if progress_cb:
                    progress_cb(f"[FAIL] {eid}: {msg}")
                continue

            out_path.write_bytes(img_bytes)

            conn.execute(
                """UPDATE generations
                SET image_path=?, created_at=?
                WHERE entry_id=?
                AND gen_id=(
                    SELECT gen_id FROM generations
                    WHERE entry_id=? ORDER BY attempt DESC LIMIT 1
                )""",
                (str(out_path), _now_iso(), eid, eid),
            )
            conn.commit()

            result.generated += 1
            if progress_cb:
                progress_cb(f"[OK {result.generated}/{result.total}] {eid}")

    finally:
        conn.close()

    return result


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default=None, help="YYYY-MM-DD (default: today)")
    parser.add_argument("--model", default=DEFAULT_MODEL)
    parser.add_argument("--size", default=DEFAULT_SIZE)
    parser.add_argument(
        "--quality", default=DEFAULT_QUALITY, choices=["low", "medium", "high"]
    )
    args = parser.parse_args()

    batch_date = args.date or date.today().isoformat()
    root = Path(__file__).resolve().parent.parent
    prompts_csv = root / "data" / "batches" / f"{batch_date}_prompts.csv"
    if not prompts_csv.exists():
        print(f"[ERROR] バッチCSVが見つかりません: {prompts_csv}", file=sys.stderr)
        return 2

    images_in = root / "images" / "in"
    db_path = root / "data" / "state.sqlite"
    log_path = root / "logs" / "image_gen.log"

    from scripts.hatena_client import load_env

    env_file = Path(
        r"C:\Users\jsber\OneDrive\Documents\Claude_task"
        r"\youtube-slides\食事指導シリーズ\_shared\.env"
    )
    env = load_env(env_file)
    api_key = env.get("OPENAI_API_KEY")
    if not api_key:
        print(
            "[ERROR] OPENAI_API_KEY が .env に設定されていません。\n"
            f"  .env: {env_file}",
            file=sys.stderr,
        )
        return 3

    generator = _default_openai_generator(api_key)
    result = generate_images_for_batch(
        prompts_csv,
        images_in,
        db_path,
        generator=generator,
        model=args.model,
        size=args.size,
        quality=args.quality,
        progress_cb=print,
        log_path=log_path,
    )

    print()
    print(
        f"完了: 生成={result.generated} 既存スキップ={result.skipped_existing} "
        f"プロンプト空={result.skipped_empty_prompt} 失敗={result.failed} "
        f"合計={result.total}"
    )
    if result.failures:
        print("失敗:")
        for eid, msg in result.failures:
            print(f"  {eid}: {msg}")
    return 0 if result.failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
