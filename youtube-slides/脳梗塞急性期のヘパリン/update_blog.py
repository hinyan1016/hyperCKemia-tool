#!/usr/bin/env python3
"""脳梗塞急性期のヘパリン 記事の下書き更新"""
import base64
import re
import urllib.request
import urllib.error
from pathlib import Path
from xml.etree import ElementTree as ET

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(__file__).parent / "ブログ記事.html"
TITLE = "脳梗塞急性期のヘパリン、結局どう使う？ 1万単位 vs APTT管理・ブリッジング不要論・HIT対応【ガイドライン2025対応】"
CATEGORIES = ["脳神経内科", "脳卒中", "医師向け", "医学情報", "ガイドライン"]
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


def auth_header(env):
    s = "{}:{}".format(env["HATENA_ID"], env["HATENA_API_KEY"])
    return "Basic " + base64.b64encode(s.encode()).decode()


def find_edit_url(env):
    base = "https://blog.hatena.ne.jp/{}/{}/atom/entry".format(env["HATENA_ID"], env["HATENA_BLOG_DOMAIN"])
    auth = auth_header(env)
    url = base
    ns = {"atom": "http://www.w3.org/2005/Atom", "app": "http://www.w3.org/2007/app"}
    while url:
        req = urllib.request.Request(url, headers={"Authorization": auth})
        with urllib.request.urlopen(req) as r:
            body = r.read().decode("utf-8")
        root = ET.fromstring(body)
        for entry in root.findall("atom:entry", ns):
            t = entry.find("atom:title", ns)
            if t is not None and SEARCH_KEY in (t.text or ""):
                for link in entry.findall("atom:link", ns):
                    if link.get("rel") == "edit":
                        print("Found: {}".format(t.text))
                        return link.get("href")
        nxt = None
        for link in root.findall("atom:link", ns):
            if link.get("rel") == "next":
                nxt = link.get("href")
                break
        url = nxt
    return None


def build_xml(title, content, categories, draft=True):
    cats = "\n".join('  <category term="{}" />'.format(c) for c in categories)
    return """<?xml version="1.0" encoding="utf-8"?>
<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">
  <title>{title}</title>
  <author><name>hinyan1016</name></author>
  <content type="text/html"><![CDATA[{content}]]></content>
{cats}
  <app:control>
    <app:draft>{draft}</app:draft>
  </app:control>
</entry>""".format(title=title, content=content, cats=cats, draft="yes" if draft else "no")


def put_update(env, edit_url, xml_body):
    req = urllib.request.Request(
        edit_url, data=xml_body.encode("utf-8"), method="PUT",
        headers={"Content-Type": "application/xml; charset=utf-8", "Authorization": auth_header(env)},
    )
    try:
        with urllib.request.urlopen(req) as r:
            body = r.read().decode("utf-8")
            m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
            return True, m.group(1) if m else "(URL not found)"
    except urllib.error.HTTPError as e:
        return False, "HTTP {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:500])


def main():
    env = load_env()
    html = HTML_FILE.read_text(encoding="utf-8")
    print("HTML: {} bytes".format(len(html)))
    edit_url = find_edit_url(env)
    if not edit_url:
        print("[ERROR] Article not found")
        return
    xml = build_xml(TITLE, html, CATEGORIES, draft=False)
    ok, result = put_update(env, edit_url, xml)
    print("[OK] " + result if ok else "[NG] " + result)


if __name__ == "__main__":
    main()
