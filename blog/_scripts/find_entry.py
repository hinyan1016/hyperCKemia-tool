#!/usr/bin/env python3
"""全エントリのURLとedit URLを一覧表示"""

import re
import base64
import urllib.request
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")

def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env

def main():
    env = load_env()
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]
    auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)

    # ページ1だけ取得
    req = urllib.request.Request(url, headers={"Authorization": "Basic {}".format(auth_b64)})
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")

    # エントリを抽出 - feedレベルのlinkを除外するため、<entry>で分割
    # フィード全体の構造: <feed>...<entry>...</entry><entry>...</entry>...</feed>
    parts = body.split("<entry>")
    print("エントリ数: {} (フィードヘッダー除く)".format(len(parts) - 1))
    print()

    for i, part in enumerate(parts[1:], 1):
        alt_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', part)
        edit_match = re.search(r'<link rel="edit"[^>]*href="([^"]+)"', part)
        title_match = re.search(r'<title[^>]*>([^<]*)</title>', part)

        alt_url = alt_match.group(1) if alt_match else "N/A"
        edit_url = edit_match.group(1) if edit_match else "N/A"
        title = title_match.group(1) if title_match else "N/A"

        has_xnbx = "XNBX" if "XNBX" in part else ""
        print("{:2d}. {} {}".format(i, alt_url, has_xnbx))
        print("    title: {}".format(title[:60]))
        print("    edit:  {}".format(edit_url.split("/")[-1] if edit_url != "N/A" else "N/A"))
        print()

if __name__ == "__main__":
    main()
