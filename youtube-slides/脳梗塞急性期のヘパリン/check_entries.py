#!/usr/bin/env python3
"""脳梗塞急性期のヘパリン を含むエントリを全て列挙"""
import base64
import urllib.request
from pathlib import Path
from xml.etree import ElementTree as ET

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
SEARCH_KEY = "脳梗塞急性期のヘパリン"


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def main():
    env = load_env()
    base = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(env["HATENA_ID"], env["HATENA_BLOG_DOMAIN"])
    auth_str = "{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"])
    auth = "Basic " + base64.b64encode(auth_str.encode()).decode()
    ns = {"atom": "http://www.w3.org/2005/Atom", "app": "http://www.w3.org/2007/app"}
    url = base
    count = 0
    while url:
        req = urllib.request.Request(url, headers={"Authorization": auth})
        with urllib.request.urlopen(req) as r:
            body = r.read().decode("utf-8")
        root = ET.fromstring(body)
        for entry in root.findall("atom:entry", ns):
            t = entry.find("atom:title", ns)
            if t is not None and SEARCH_KEY in (t.text or ""):
                count += 1
                edit_url = ""
                alt_url = ""
                for link in entry.findall("atom:link", ns):
                    rel = link.get("rel")
                    if rel == "edit":
                        edit_url = link.get("href")
                    elif rel == "alternate":
                        alt_url = link.get("href")
                draft = entry.find(".//app:draft", ns)
                draft_val = draft.text if draft is not None else "?"
                print("--- entry #{} ---".format(count))
                print("title: " + (t.text or "")[:80])
                print("alt:   " + alt_url)
                print("edit:  " + edit_url)
                print("draft: " + draft_val)
        nxt = None
        for link in root.findall("atom:link", ns):
            if link.get("rel") == "next":
                nxt = link.get("href")
                break
        url = nxt
    print("\nTotal: {}".format(count))


if __name__ == "__main__":
    main()
