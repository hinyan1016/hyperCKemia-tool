#!/usr/bin/env python3
"""既存はてなブログ記事（113203）をローカルarticle.htmlで上書きPUT更新"""

import sys
import base64
import urllib.request
import urllib.error
import json
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SCRIPT_DIR = Path(__file__).parent
ENV_FILE = SCRIPT_DIR.parent / "食事指導シリーズ" / "_shared" / ".env"
HTML_FILE = SCRIPT_DIR / "blog" / "article.html"
STATUS_FILE = SCRIPT_DIR / "post_status.json"

TITLE = "2026年上半期 内科Practice-Changing論文14選 — Less is More / 上流介入 / 予防医療進化の3潮流"
CATEGORIES = ["医学論文", "内科", "NEJM", "Lancet", "JAMA"]


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def build_entry(title, content, categories, draft=False):
    draft_val = "yes" if draft else "no"
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in categories)
    return """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{categories}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(title=title, content=content, categories=cat_xml, draft=draft_val)


def main():
    env = load_env()
    with open(STATUS_FILE, encoding="utf-8") as f:
        status = json.load(f)
    edit_url = status.get("edit") or status.get("edit_url")

    with open(HTML_FILE, encoding="utf-8") as f:
        content = f.read()

    print("=" * 60)
    print("はてなブログ 既存エントリ PUT更新")
    print("=" * 60)
    print("編集URL: {}".format(edit_url))
    print("HTMLサイズ: {:,} bytes".format(len(content)))
    print()

    xml = build_entry(TITLE, content, CATEGORIES, draft=False)
    auth_b64 = base64.b64encode(
        "{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"]).encode()
    ).decode()
    req = urllib.request.Request(
        edit_url,
        data=xml.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth_b64),
        },
    )
    try:
        with urllib.request.urlopen(req) as resp:
            print("[成功] HTTP {}".format(resp.status))
            print("  記事URL: {}".format(status.get("alternate") or status.get("published_url")))
    except urllib.error.HTTPError as e:
        print("[失敗] HTTP {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:400]))
        sys.exit(1)


if __name__ == "__main__":
    main()
