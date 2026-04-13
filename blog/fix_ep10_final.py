#!/usr/bin/env python3
"""YouTube記事から#10カードを正確に削除"""

import re
import base64
import urllib.request
from pathlib import Path

ENV_FILE = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env")
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

    # GET
    print("記事を取得中...")
    req = urllib.request.Request(EDIT_URL, headers={"Authorization": "Basic {}".format(auth_b64)})
    with urllib.request.urlopen(req) as resp:
        full_entry = resp.read().decode("utf-8")

    count = full_entry.count("XNBX-zDf3g0")
    print("XNBX-zDf3g0 出現回数: {}".format(count))

    if count == 0:
        print("既に削除済みです")
        return

    # XNBX-zDf3g0周辺のコンテキストを表示
    for m in re.finditer(r'.{0,80}XNBX-zDf3g0.{0,80}', full_entry):
        print("\nコンテキスト: ...{}...".format(m.group()))

    # インラインスタイル版の動画カードを削除
    # パターン: display:flex を含むdivの中にXNBX-zDf3g0がある構造
    # まずXNBX-zDf3g0の位置を特定し、そこを含む最も近い外側のdivブロックを特定

    original = full_entry

    # 方法: XNBX-zDf3g0を含むimg/aタグの親divブロック全体を削除
    # カード構造: <div style="...display:flex...margin-bottom:10px...">..XNBX.....</div>
    # 最外側のdivを見つけて削除する

    # XNBX-zDf3g0が含まれる位置を探す
    pos = full_entry.find("XNBX-zDf3g0")
    if pos == -1:
        print("見つかりません")
        return

    # posから前方にさかのぼって、margin-bottom:10pxを含むdivの開始位置を探す
    # カードの特徴的なスタイル: display:flex;gap:14px;padding:12px;background:#fff;border:1px solid #e5e7eb;border-radius:12px;margin-bottom:10px
    search_start = max(0, pos - 2000)
    before = full_entry[search_start:pos]

    # 最も近い "margin-bottom:10px" を含むdiv開始を探す
    card_starts = [m.start() + search_start for m in re.finditer(r'<div\s+style="[^"]*margin-bottom:10px[^"]*"', full_entry[search_start:pos])]

    if not card_starts:
        print("カード開始位置が見つかりません。行単位で削除します。")
        lines = full_entry.split("\n")
        new_lines = [l for l in lines if "XNBX-zDf3g0" not in l]
        full_entry = "\n".join(new_lines)
    else:
        card_start = card_starts[-1]  # 最も近いもの
        # card_startから次の同レベルの</div>を見つける（ネストを考慮）
        depth = 0
        i = card_start
        card_end = None
        while i < len(full_entry):
            if full_entry[i:i+4] == "<div":
                depth += 1
            elif full_entry[i:i+6] == "</div>":
                depth -= 1
                if depth == 0:
                    card_end = i + 6
                    break
            i += 1

        if card_end:
            removed = full_entry[card_start:card_end]
            print("\n削除するブロック ({} 文字):".format(len(removed)))
            print(removed[:200] + "...")
            full_entry = full_entry[:card_start] + full_entry[card_end:]
        else:
            print("カード終了位置が見つかりません")
            return

    # 数値修正
    # 34→33 (動画数)
    full_entry = full_entry.replace(">34<", ">33<")
    # 12→11 (NEW数)
    # 慎重に: ">12<" はNEW数以外にもマッチする可能性
    # 10 / 17→9 / 17
    full_entry = full_entry.replace("10 / 17", "9 / 17")
    # 進捗バー
    full_entry = full_entry.replace("width:59%", "width:53%")

    remaining = full_entry.count("XNBX-zDf3g0")
    print("\n残りのXNBX-zDf3g0: {}".format(remaining))
    print("サイズ変化: {} -> {}".format(len(original), len(full_entry)))

    if remaining > 0:
        print("警告: まだ残っています。行単位で追加削除します。")
        lines = full_entry.split("\n")
        new_lines = [l for l in lines if "XNBX-zDf3g0" not in l]
        full_entry = "\n".join(new_lines)
        print("最終残り: {}".format(full_entry.count("XNBX-zDf3g0")))

    # PUT
    print("\n更新中...")
    data = full_entry.encode("utf-8")
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
        print("[成功]")
    except urllib.request.HTTPError as e:
        print("[失敗] HTTP {}: {}".format(e.code, e.read().decode("utf-8", errors="replace")[:300]))

if __name__ == "__main__":
    main()
