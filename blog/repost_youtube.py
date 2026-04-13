#!/usr/bin/env python3
"""YouTube記事をローカルHTMLから丸ごと再投稿（PUT）"""

import re
import sys
import base64
import urllib.request
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from post_summary_drafts import convert_to_inline, YT_CLASS_MAP, resolve_var

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
HTML_FILE = Path(__file__).parent / "youtube_featured_april9.html"
EDIT_URL = "https://blog.hatena.ne.jp/hinyan1016/hinyan1016.hatenablog.com/atom/entry/17179246901374671149"

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
    api_key = env["HATENA_API_KEY"]
    auth_b64 = base64.b64encode("{}:{}".format(hatena_id, api_key).encode()).decode()

    # ローカルHTMLを読み込み・変換
    print("ローカルHTML読み込み中...")
    with open(HTML_FILE, encoding="utf-8") as f:
        raw_html = f.read()

    # XNBX確認
    print("ローカルHTML内のXNBX-zDf3g0: {}".format(raw_html.count("XNBX-zDf3g0")))

    title_match = re.search(r'<title>(.*?)</title>', raw_html)
    title = title_match.group(1) if title_match else "YouTube動画まとめ"

    content = convert_to_inline(raw_html, YT_CLASS_MAP)

    # aタグのスタイル追加
    content = re.sub(
        r'(<a\s)(?!.*?style=)',
        r'\1style="color:#1e6ba5;font-weight:600;text-decoration:none;" ',
        content
    )

    print("変換後のXNBX-zDf3g0: {}".format(content.count("XNBX-zDf3g0")))
    print("変換後サイズ: {:.1f} KB".format(len(content.encode("utf-8")) / 1024))

    # 既存エントリをGET
    print("\n既存エントリを取得中...")
    req = urllib.request.Request(EDIT_URL, headers={"Authorization": "Basic {}".format(auth_b64)})
    with urllib.request.urlopen(req) as resp:
        existing = resp.read().decode("utf-8")

    # content部分を丸ごと置換
    # <content type="text/html"> or <content type="text/x-hatena-syntax">
    new_entry = re.sub(
        r'(<content[^>]*>).*?(</content>)',
        r'\1<![CDATA[{}]]>\2'.format(content),
        existing,
        flags=re.DOTALL
    )

    # タイトルも更新
    new_entry = re.sub(
        r'<title[^>]*>.*?</title>',
        '<title>{}</title>'.format(title),
        new_entry,
        count=1
    )

    print("更新後エントリ内のXNBX-zDf3g0: {}".format(new_entry.count("XNBX-zDf3g0")))

    # PUT
    print("PUT更新中...")
    data = new_entry.encode("utf-8")
    req2 = urllib.request.Request(
        EDIT_URL, data=data, method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": "Basic {}".format(auth_b64),
        },
    )
    try:
        with urllib.request.urlopen(req2) as resp2:
            resp2.read()
        print("[成功] 記事を更新しました")
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8", errors="replace")
        print("[失敗] HTTP {}: {}".format(e.code, error_body[:500]))

if __name__ == "__main__":
    main()
