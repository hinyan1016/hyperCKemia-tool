#!/usr/bin/env python3
"""公開済みYouTubeまとめ記事から#10を削除 - alternate linkで正確にマッチ"""

import re
import base64
import urllib.request
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
TARGET_ALT_URL = "https://hinyan1016.hatenablog.com/entry/2026/04/09/174119"

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
    visited = set()

    for page in range(10):
        if url in visited:
            break
        visited.add(url)
        print("ページ {}...".format(page + 1))

        req = urllib.request.Request(url, headers={"Authorization": "Basic {}".format(auth_b64)})
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")

        # <entry>...</entry> を正確に抽出
        entries = re.findall(r'<entry>(.*?)</entry>', body, re.DOTALL)
        for entry_xml in entries:
            # alternate link を抽出
            alt_match = re.search(r'<link rel="alternate"[^>]*href="([^"]+)"', entry_xml)
            if not alt_match:
                continue
            alt_url = alt_match.group(1)

            if alt_url == TARGET_ALT_URL:
                edit_match = re.search(r'<link rel="edit"[^>]*href="([^"]+)"', entry_xml)
                if not edit_match:
                    print("edit linkが見つかりません")
                    return
                edit_url = edit_match.group(1)
                title_match = re.search(r'<title[^>]*>([^<]*)</title>', entry_xml)
                title = title_match.group(1) if title_match else "?"
                print("\n発見: {} ({})".format(title, alt_url))
                print("edit URL: {}".format(edit_url))

                # GET full entry
                req2 = urllib.request.Request(edit_url, headers={"Authorization": "Basic {}".format(auth_b64)})
                with urllib.request.urlopen(req2) as resp2:
                    full_entry = resp2.read().decode("utf-8")

                # XNBX-zDf3g0 が含まれるか確認
                count = full_entry.count("XNBX-zDf3g0")
                print("XNBX-zDf3g0 の出現回数: {}".format(count))

                if count == 0:
                    print("既に#10は削除されています")
                    return

                # 修正: XNBX-zDf3g0 を含む動画カードブロックを削除
                # インラインスタイル変換後なのでclass属性はない、style属性で検索
                original_len = len(full_entry)

                # パターン: <div style="...flex..."><div style="...180px..."><a href="...XNBX...">...</a></div><div style="...flex:1...">...</div></div>
                full_entry = re.sub(
                    r'<div\s+style="[^"]*display:flex[^"]*margin-bottom:10px[^"]*">\s*<div\s+style="[^"]*180px[^"]*">\s*<a[^>]*XNBX-zDf3g0[^>]*>.*?</div>\s*</div>',
                    '',
                    full_entry,
                    flags=re.DOTALL
                )

                # もし上のパターンでマッチしなかった場合、単純にXNBX-zDf3g0を含む行を削除
                if full_entry.count("XNBX-zDf3g0") > 0:
                    print("パターン1でマッチせず、別の方法で削除します")
                    # 各行をフィルタリング
                    lines = full_entry.split("\n")
                    new_lines = [l for l in lines if "XNBX-zDf3g0" not in l]
                    full_entry = "\n".join(new_lines)

                # 数値修正
                full_entry = full_entry.replace(">34<", ">33<")
                full_entry = full_entry.replace(">12<", ">11<")
                full_entry = full_entry.replace("10 / 17", "9 / 17")
                full_entry = full_entry.replace("width:59%", "width:53%")

                new_len = len(full_entry)
                print("サイズ: {} -> {} ({} バイト削減)".format(original_len, new_len, original_len - new_len))

                # 残りのXNBX確認
                remaining = full_entry.count("XNBX-zDf3g0")
                print("残りのXNBX-zDf3g0: {}".format(remaining))

                # PUT
                print("更新中...")
                data = full_entry.encode("utf-8")
                req3 = urllib.request.Request(
                    edit_url, data=data, method="PUT",
                    headers={
                        "Content-Type": "application/xml; charset=utf-8",
                        "Authorization": "Basic {}".format(auth_b64),
                    },
                )
                with urllib.request.urlopen(req3) as resp3:
                    resp3.read()
                print("[成功] #10を削除しました")
                return

        # 次ページ
        next_links = re.findall(r'<link rel="next"[^>]*href="([^"]+)"', body)
        if not next_links:
            break
        url = next_links[0].replace("&amp;", "&")

    print("エントリが見つかりません: {}".format(TARGET_ALT_URL))

if __name__ == "__main__":
    main()
