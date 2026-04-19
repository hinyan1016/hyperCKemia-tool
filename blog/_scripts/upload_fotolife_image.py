#!/usr/bin/env python3
"""Hatena Fotolife AtomPub API で画像をアップロード。

Usage:
    python upload_fotolife_image.py <image_path> [--title <title>]

Returns (prints) the Fotolife image URL on success.
"""

from __future__ import annotations

import base64
import hashlib
import random
import re
import string
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Final

ENV_FILE: Final = Path(
    r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env"
)
FOTOLIFE_ENDPOINT: Final = "https://f.hatena.ne.jp/atom/post"


def load_env() -> dict[str, str]:
    env: dict[str, str] = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def wsse_header(user: str, api_key: str) -> str:
    nonce_raw = "".join(random.choices(string.ascii_letters + string.digits, k=16)).encode()
    nonce_b64 = base64.b64encode(nonce_raw).decode()
    created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    digest_src = nonce_raw + created.encode() + api_key.encode()
    digest_b64 = base64.b64encode(hashlib.sha1(digest_src).digest()).decode()
    return (
        f'UsernameToken Username="{user}", '
        f'PasswordDigest="{digest_b64}", '
        f'Nonce="{nonce_b64}", '
        f'Created="{created}"'
    )


def build_entry_xml(title: str, image_b64: str, mime: str) -> str:
    ext = {"image/png": "png", "image/jpeg": "jpeg"}.get(mime, "png")
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<entry xmlns="http://purl.org/atom/ns#">'
        f"<title>{title}</title>"
        f'<content mode="base64" type="{mime}">{image_b64}</content>'
        f'<dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/">blog</dc:subject>'
        "</entry>"
    )


def upload(image_path: Path, title: str) -> str:
    env = load_env()
    hatena_id = env["HATENA_ID"]
    api_key = env["HATENA_API_KEY"]

    mime = "image/png" if image_path.suffix.lower() == ".png" else "image/jpeg"
    image_bytes = image_path.read_bytes()
    image_b64 = base64.b64encode(image_bytes).decode()
    xml = build_entry_xml(title, image_b64, mime)

    req = urllib.request.Request(
        FOTOLIFE_ENDPOINT,
        data=xml.encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/atom+xml",
            "Authorization": f"WSSE profile=\"UsernameToken\"",
            "X-WSSE": wsse_header(hatena_id, api_key),
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code}: {body[:500]}") from e

    # Try to extract the image URL from the response
    # Fotolife returns <hatena:imageurl>, <hatena:syntax>, or similar
    image_url_match = re.search(r"<hatena:imageurl[^>]*>([^<]+)</hatena:imageurl>", body)
    syntax_match = re.search(r"<hatena:syntax[^>]*>([^<]+)</hatena:syntax>", body)
    image_id_match = re.search(r"<hatena:imageid[^>]*>([^<]+)</hatena:imageid>", body)
    id_match = re.search(r"<id>([^<]+)</id>", body)

    # Also look for hatena:imageurlmedium / imageurloriginal
    orig_match = re.search(r"<hatena:imageurl original=\"([^\"]+)\"", body)
    medium_match = re.search(r"<hatena:imageurl medium=\"([^\"]+)\"", body)
    # Dump full response to stderr for inspection
    import sys
    print("--- RAW RESPONSE (first 2KB) ---", file=sys.stderr)
    print(body[:2000], file=sys.stderr)
    print("--- END ---", file=sys.stderr)

    if orig_match:
        return orig_match.group(1)
    if image_url_match:
        return image_url_match.group(1)
    if id_match:
        return id_match.group(1)
    return f"(could not parse response)\n{body[:1000]}"


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: upload_fotolife_image.py <image_path> [--title <title>]")
        sys.exit(1)

    image_path = Path(sys.argv[1])
    title = "blog_image"
    if "--title" in sys.argv:
        idx = sys.argv.index("--title")
        if idx + 1 < len(sys.argv):
            title = sys.argv[idx + 1]

    print(f"Uploading {image_path} as '{title}'...")
    url = upload(image_path, title)
    print(f"IMAGE_URL={url}")


if __name__ == "__main__":
    main()
