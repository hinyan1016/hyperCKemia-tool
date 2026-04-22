#!/usr/bin/env python3
"""
2026年上半期 内科Practice-Changing論文14選 はてなブログ下書き投稿
"""

import os
import re
import sys
import json
import base64
import urllib.request
import urllib.error
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


def build_atom_entry(title, content, categories, draft=True):
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


def post_entry(env, xml_body):
    hatena_id = env["HATENA_ID"]
    blog_domain = env["HATENA_BLOG_DOMAIN"]
    api_key = env["HATENA_API_KEY"]
    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(hatena_id, blog_domain)
    auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()
    req = urllib.request.Request(
        url,
        data=xml_body.encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth_b64),
        },
    )
    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            url_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            edit_match = re.search(r'<link rel="edit"[^>]*href="([^"]+)"', body)
            return True, {
                "alternate": url_match.group(1) if url_match else None,
                "edit": edit_match.group(1) if edit_match else None,
            }
    except urllib.error.HTTPError as e:
        return False, "HTTP {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:300])


def main():
    publish = "--publish" in sys.argv
    env = load_env()

    with open(HTML_FILE, encoding="utf-8") as f:
        content = f.read()

    mode = "公開" if publish else "下書き"
    print("=" * 60)
    print("はてなブログ {} 投稿".format(mode))
    print("=" * 60)
    print("タイトル: {}".format(TITLE))
    print("HTMLサイズ: {:,} bytes".format(len(content)))
    print("カテゴリ: {}".format(", ".join(CATEGORIES)))
    print()

    xml = build_atom_entry(TITLE, content, CATEGORIES, draft=not publish)
    ok, result = post_entry(env, xml)

    if ok:
        print("[成功] 投稿完了")
        print("  記事URL: {}".format(result.get("alternate")))
        print("  編集URL: {}".format(result.get("edit")))
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump({"title": TITLE, "draft": not publish, **result}, f, ensure_ascii=False, indent=2)
        print("  → post_status.json に記録")
    else:
        print("[失敗] {}".format(result))
        sys.exit(1)


if __name__ == "__main__":
    main()
