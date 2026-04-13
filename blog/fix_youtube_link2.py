#!/usr/bin/env python3
"""全エントリからOLD_IDを含むものを全て修正"""

import re
import base64
import urllib.request
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
OLD_ID = "YrVR1Iwh3rA"
NEW_ID = "Yl28Tu1NU_c"

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
    auth_str = "{}:{}".format(hatena_id, api_key)
    auth_b64 = base64.b64encode(auth_str.encode()).decode()

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
    visited = set()
    fixed = []

    for page in range(15):
        if url in visited:
            break
        visited.add(url)
        print("ページ {} ...".format(page + 1))

        req = urllib.request.Request(url, headers={"Authorization": "Basic {}".format(auth_b64)})
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")

        parts = body.split("</entry>")
        for part in parts:
            if OLD_ID in part:
                edit_match = re.search(r'<link rel="edit"[^>]*href="([^"]+)"', part)
                alt_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', part)
                if edit_match:
                    edit_url = edit_match.group(1)
                    alt_url = alt_match.group(1) if alt_match else "?"
                    print("  発見: {}".format(alt_url))

                    # GET
                    req2 = urllib.request.Request(edit_url, headers={"Authorization": "Basic {}".format(auth_b64)})
                    with urllib.request.urlopen(req2) as resp2:
                        full_entry = resp2.read().decode("utf-8")

                    count = full_entry.count(OLD_ID)
                    updated = full_entry.replace(OLD_ID, NEW_ID)

                    # PUT
                    data = updated.encode("utf-8")
                    req3 = urllib.request.Request(
                        edit_url, data=data, method="PUT",
                        headers={
                            "Content-Type": "application/xml; charset=utf-8",
                            "Authorization": "Basic {}".format(auth_b64),
                        },
                    )
                    with urllib.request.urlopen(req3) as resp3:
                        resp3.read()
                    print("    修正完了 ({} 箇所)".format(count))
                    fixed.append(alt_url)

        # 次ページ
        next_links = re.findall(r'<link rel="next"[^>]*href="([^"]+)"', body)
        if not next_links:
            break
        url = next_links[0].replace("&amp;", "&")

    print("\n合計 {} 記事を修正".format(len(fixed)))
    for u in fixed:
        print("  {}".format(u))

if __name__ == "__main__":
    main()
