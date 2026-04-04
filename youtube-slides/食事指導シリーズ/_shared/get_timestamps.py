"""
Gemini APIで動画のスライド切り替えタイムスタンプを取得するスクリプト。
使い方: python get_timestamps.py <動画ファイルパス> [台本ファイルパス]
"""

import os
import sys
import time
from google import genai

# Windows cp932 エンコーディング問題を回避
sys.stdout.reconfigure(encoding="utf-8")

def main():
    if len(sys.argv) < 2:
        print("使い方: python get_timestamps.py <動画ファイルパス> [台本ファイルパス]")
        sys.exit(1)

    video_path = sys.argv[1]
    script_path = sys.argv[2] if len(sys.argv) >= 3 else None

    api_key = "AIzaSyAW8gS1mBZJWL0U3yJj7naoOq78tpVc0Oo"
    client = genai.Client(api_key=api_key)

    # 動画アップロード
    print(f"動画をアップロード中: {video_path}")
    video_file = client.files.upload(file=video_path)
    print(f"アップロード完了: {video_file.name}")

    # 処理完了を待機
    while video_file.state.name == "PROCESSING":
        print("処理中...")
        time.sleep(5)
        video_file = client.files.get(name=video_file.name)

    if video_file.state.name == "FAILED":
        print(f"エラー: 動画処理に失敗しました - {video_file.state}")
        sys.exit(1)

    print(f"動画の処理完了 (状態: {video_file.state.name})")

    # 台本がある場合は読み込み
    script_text = ""
    if script_path:
        with open(script_path, "r", encoding="utf-8") as f:
            script_text = f.read()
        script_context = f"\n\n以下は発表台本です。各スライドのタイトルはこちらを参考にしてください:\n{script_text}"
    else:
        script_context = ""

    prompt = f"""この動画はYouTube用のスライド発表動画です。
各スライドが切り替わるタイミングを正確に特定し、以下の形式でタイムスタンプを出力してください。

出力形式（YouTube説明欄用）:
MM:SS スライドの内容を簡潔に表すタイトル

ルール:
- 最初のスライド（タイトル）は必ず 00:00 から始める
- スライドが実際に切り替わった瞬間の時刻を記録する
- タイトルは日本語で、動画の内容に基づいて簡潔に（20文字以内）
- 最後のスライド（エンドカード等）も含める
- タイムスタンプ以外の余計なテキストは出力しない{script_context}"""

    print("\nGeminiに解析を依頼中...\n")
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[video_file, prompt]
    )

    print("=== タイムスタンプ ===")
    print(response.text)

    # 結果をファイルにも保存
    video_dir = os.path.dirname(video_path)
    out_path = os.path.join(video_dir, "timestamps.txt")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(response.text)
    print(f"\n(結果を {out_path} に保存しました)")

    # アップロードしたファイルを削除
    client.files.delete(name=video_file.name)
    print("(アップロードファイルを削除しました)")


if __name__ == "__main__":
    main()
