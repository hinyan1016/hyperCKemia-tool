#!/usr/bin/env python3
"""
誤って上書きした entry 17179246901352271304（高齢者の手の震え）を
backup_uploaddate の内容で復元する。
"""

import sys
import re
import base64
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

HATENA_ID = "hinyan1016"
BLOG_DOMAIN = "hinyan1016.hatenablog.com"
API_KEY = "y8n4v3sioq"
ENTRY_ID = "17179246901352271304"
TITLE = "高齢者の手の震え – 原因の見分け方と対処法を専門医が解説"
CATEGORIES = ["パーキンソン病", "一般向け", "医学情報", "脳神経内科"]
BACKUP = r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\seo-improvement\backup_uploaddate\17179246901352271304.html"


def auth():
    return "Basic " + base64.b64encode("{}:{}".format(HATENA_ID, API_KEY).encode()).decode()


def main():
    with open(BACKUP, encoding="utf-8") as f:
        content = f.read()

    cat_xml = "\n".join('  <category term="{}" />'.format(c) for c in CATEGORIES)
    xml = '<?xml version="1.0" encoding="utf-8"?>\n'
    xml += '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:app="http://www.w3.org/2007/app">\n'
    xml += '  <title>{}</title>\n'.format(TITLE)
    xml += '  <author><name>{}</name></author>\n'.format(HATENA_ID)
    xml += '  <content type="text/html"><![CDATA[{}]]></content>\n'.format(content)
    xml += '{}\n'.format(cat_xml)
    xml += '  <app:control>\n    <app:draft>no</app:draft>\n  </app:control>\n'
    xml += '</entry>'

    url = "https://blog.hatena.ne.jp/{}/{}/atom/entry/{}".format(HATENA_ID, BLOG_DOMAIN, ENTRY_ID)
    req = urllib.request.Request(
        url, data=xml.encode("utf-8"), method="PUT",
        headers={"Content-Type": "application/xml; charset=utf-8", "Authorization": auth()},
    )
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode("utf-8")
    m = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', body)
    print("復元成功: {}".format(m.group(1) if m else "(URL不明)"))


if __name__ == "__main__":
    main()
