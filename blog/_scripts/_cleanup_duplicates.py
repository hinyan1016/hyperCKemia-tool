#!/usr/bin/env python3
"""低活動性せん妄の重複下書きを検索し、新URL以外を削除する"""
import re
import urllib.request
import urllib.error
import base64
import xml.etree.ElementTree as ET
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
KEEP_URL = "https://hinyan1016.hatenablog.com/entry/2026/04/09/160554"

def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env

env = load_env()
hatena_id = env["HATENA_ID"]
blog_domain = env["HATENA_BLOG_DOMAIN"]
api_key = env["HATENA_API_KEY"]
auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()

ns = {"atom": "http://www.w3.org/2005/Atom", "app": "http://www.w3.org/2007/app"}

# Fetch recent entries (including drafts)
api_url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
req = urllib.request.Request(api_url, headers={"Authorization": "Basic " + auth_b64})
with urllib.request.urlopen(req) as resp:
    body = resp.read().decode("utf-8")

root = ET.fromstring(body)
found = []
for entry in root.findall("atom:entry", ns):
    title_el = entry.find("atom:title", ns)
    title = title_el.text if title_el is not None else ""
    if "低活動性せん妄" not in title:
        continue
    alt_url = ""
    edit_url = ""
    for link in entry.findall("atom:link", ns):
        if link.get("rel") == "alternate":
            alt_url = link.get("href", "")
        if link.get("rel") == "edit":
            edit_url = link.get("href", "")
    draft_el = entry.find(".//app:draft", ns)
    is_draft = draft_el.text == "yes" if draft_el is not None else False
    found.append({"title": title, "url": alt_url, "edit": edit_url, "draft": is_draft})

print("「低活動性せん妄」を含む記事: {}件".format(len(found)))
for f in found:
    status = "下書き" if f["draft"] else "公開"
    keep = " [KEEP]" if f["url"] == KEEP_URL else ""
    print("  {} {} {}{}".format(status, f["url"], f["title"], keep))

# Delete all except KEEP_URL
deleted = 0
for f in found:
    if f["url"] == KEEP_URL:
        continue
    print("\n削除中: {}".format(f["url"]))
    req = urllib.request.Request(f["edit"], method="DELETE", headers={
        "Authorization": "Basic " + auth_b64,
    })
    try:
        with urllib.request.urlopen(req) as resp:
            print("  削除成功: {}".format(resp.status))
            deleted += 1
    except urllib.error.HTTPError as e:
        print("  削除失敗: {} {}".format(e.code, e.read().decode("utf-8", errors="replace")[:200]))

print("\n完了: {}件削除、1件保持({})".format(deleted, KEEP_URL))
