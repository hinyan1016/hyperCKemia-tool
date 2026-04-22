#!/usr/bin/env python3
"""指定したはてなブログ記事の末尾に1000記事突破記事への内部リンクを追加する。

- AtomPub feed を辿って alternate URL で対象記事を特定
- content (XHTML/HTML) の末尾に callout box を追記
- 同じ記事に既にミルストーンリンクが入っていればスキップ（冪等）
- PUT で更新
"""
import base64
import re
import sys
import urllib.request
import urllib.error
from pathlib import Path
from xml.etree import ElementTree as ET

SCRIPT_DIR = Path(__file__).parent
ENV_FILE = SCRIPT_DIR.parent / "youtube-slides" / "食事指導シリーズ" / "_shared" / ".env"

MILESTONE_URL = "https://hinyan1016.hatenablog.com/entry/2026/04/22/093852"
MILESTONE_TITLE = "ブログ開始から1年余りで1000記事を突破。AIとともに続けてきた発信の現在地"

# 追加対象の記事 (alternate URL)
TARGETS = [
    "https://hinyan1016.hatenablog.com/entry/2026/04/20/180033",  # その症状大丈夫 #01 手のふるえ
    "https://hinyan1016.hatenablog.com/entry/2026/04/18/085038",  # からだの不思議 #17
]

# 末尾に挿入するcallout (インラインスタイル)
CALLOUT_HTML = """
<div style="background:#f3f7fc;border:1px solid #c8d6e8;border-left:5px solid #2C5AA0;border-radius:6px;padding:14px 18px;margin:28px 0 8px 0;color:#2c3e50;font-size:0.94rem;line-height:1.75;">
<div style="font-weight:bold;color:#1B3A5C;margin-bottom:6px;">📣 おしらせ</div>
ブログ開始から1年余りで累計記事数が1000件を超えました。AIとともに続けてきた発信の歩みを、こちらの記事で振り返っています ⇒ <a href="{url}" style="color:#2c5aa0;font-weight:bold;text-decoration:underline;">{title}</a>
</div>
""".format(url=MILESTONE_URL, title=MILESTONE_TITLE)

ATOM_NS = "http://www.w3.org/2005/Atom"
APP_NS = "http://www.w3.org/2007/app"

ET.register_namespace("", ATOM_NS)
ET.register_namespace("app", APP_NS)


def load_env():
    env = {}
    with open(ENV_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def auth_header(env):
    raw = "{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"])
    return "Basic " + base64.b64encode(raw.encode()).decode()


def http_get(url, env):
    req = urllib.request.Request(
        url, method="GET", headers={"Authorization": auth_header(env)}
    )
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8")


def http_put(url, body, env):
    req = urllib.request.Request(
        url,
        data=body.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": auth_header(env),
        },
    )
    with urllib.request.urlopen(req) as resp:
        return resp.status, resp.read().decode("utf-8")


def find_entry_xml(env, target_url, max_pages=20):
    """フィードを辿って対象記事のentry XMLを返す。"""
    feed_url = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(
        env["HATENA_ID"], env["HATENA_BLOG_DOMAIN"]
    )
    page = 0
    while feed_url and page < max_pages:
        page += 1
        print("  feed page {}: {}".format(page, feed_url))
        feed_xml = http_get(feed_url, env)
        root = ET.fromstring(feed_xml)
        for entry in root.findall("{{{}}}entry".format(ATOM_NS)):
            for link in entry.findall("{{{}}}link".format(ATOM_NS)):
                if (
                    link.get("rel") == "alternate"
                    and link.get("href") == target_url
                ):
                    return entry
        # next link
        next_url = None
        for link in root.findall("{{{}}}link".format(ATOM_NS)):
            if link.get("rel") == "next":
                next_url = link.get("href")
                break
        feed_url = next_url
    return None


def get_edit_url(entry):
    for link in entry.findall("{{{}}}link".format(ATOM_NS)):
        if link.get("rel") == "edit":
            return link.get("href")
    return None


def get_text(entry, tag):
    el = entry.find("{{{}}}{}".format(ATOM_NS, tag))
    return el.text if el is not None else None


def get_categories(entry):
    cats = []
    for c in entry.findall("{{{}}}category".format(ATOM_NS)):
        term = c.get("term")
        if term:
            cats.append(term)
    return cats


def get_draft_flag(entry):
    ctrl = entry.find("{{{}}}control".format(APP_NS))
    if ctrl is None:
        return "no"
    draft = ctrl.find("{{{}}}draft".format(APP_NS))
    if draft is not None and draft.text:
        return draft.text
    return "no"


def build_atom(title, content, cats, draft):
    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in cats)
    return """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>{author}</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{cats}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(
        title=escape_xml(title),
        author="hinyan1016",
        content=content,
        cats=cat_xml,
        draft=draft,
    )


def escape_xml(s):
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def already_has_link(content):
    return MILESTONE_URL in content


def append_callout(content):
    return content.rstrip() + "\n" + CALLOUT_HTML


def process_target(env, target_url):
    print("\n=== {} ===".format(target_url))
    entry = find_entry_xml(env, target_url)
    if entry is None:
        print("  [SKIP] entry not found in feed")
        return False
    title = get_text(entry, "title") or ""
    content = get_text(entry, "content") or ""
    cats = get_categories(entry)
    draft = get_draft_flag(entry)
    edit_url = get_edit_url(entry)
    print("  title: {}".format(title))
    print("  draft: {}".format(draft))
    print("  cats : {}".format(cats))
    print("  edit : {}".format(edit_url))
    print("  size before: {:,} bytes".format(len(content.encode("utf-8"))))

    if already_has_link(content):
        print("  [SKIP] milestone link already present")
        return False

    new_content = append_callout(content)
    print("  size after : {:,} bytes".format(len(new_content.encode("utf-8"))))

    xml = build_atom(title, new_content, cats, draft)
    if "--dry-run" in sys.argv:
        print("  [DRY-RUN] would PUT {} bytes".format(len(xml.encode("utf-8"))))
        return True
    try:
        status, _ = http_put(edit_url, xml, env)
        print("  [OK] HTTP {}".format(status))
        return True
    except urllib.error.HTTPError as e:
        print("  [FAIL] HTTP {}".format(e.code))
        print(e.read().decode("utf-8", errors="replace")[:400])
        return False


def main():
    env = load_env()
    print("Adding milestone backlink to {} entries".format(len(TARGETS)))
    if "--dry-run" in sys.argv:
        print("(DRY-RUN mode)")
    ok = 0
    for url in TARGETS:
        if process_target(env, url):
            ok += 1
    print("\nDone: {}/{} updated".format(ok, len(TARGETS)))


if __name__ == "__main__":
    main()
