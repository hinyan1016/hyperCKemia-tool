import google.generativeai as genai
import os
import time
import json

API_KEY = os.environ.get("GEMINI_API_KEY", "AIzaSyAW8gS1mBZJWL0U3yJj7naoOq78tpVc0Oo")
genai.configure(api_key=API_KEY)

VIDEO_PATH = os.path.join(os.path.dirname(__file__), "japanese_origin_genetics_slides.mp4")

print(f"Uploading video: {VIDEO_PATH}")
print(f"File size: {os.path.getsize(VIDEO_PATH) / 1024 / 1024:.1f} MB")

video_file = genai.upload_file(path=VIDEO_PATH)
print(f"Upload complete. URI: {video_file.uri}")

# Wait for processing
while video_file.state.name == "PROCESSING":
    print("Processing...")
    time.sleep(5)
    video_file = genai.get_file(video_file.name)

if video_file.state.name == "FAILED":
    raise ValueError(f"Video processing failed: {video_file.state.name}")

print(f"Video ready. State: {video_file.state.name}")

model = genai.GenerativeModel(model_name="gemini-2.5-flash")

prompt = """この動画はYouTube用のプレゼンテーション動画です。
動画の内容を分析し、各トピックの切り替わりのタイムスタンプを作成してください。

以下の形式で出力してください（YouTube概要欄用）：
0:00 オープニング
X:XX トピック名
...

注意点：
- 実際の動画の時間に基づいてタイムスタンプを記載してください
- スライドの切り替わりやトピックの変化を正確に捉えてください
- トピック名は日本語で簡潔に記載してください
- 細かすぎず、大まかすぎない粒度で（15〜25項目程度）

タイムスタンプのみを出力してください。説明文は不要です。
"""

response = model.generate_content([video_file, prompt])
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

print("\n=== TIMESTAMPS ===")
print(response.text)

# Save to file
output_path = os.path.join(os.path.dirname(__file__), "timestamps.txt")
with open(output_path, "w", encoding="utf-8") as f:
    f.write(response.text)
print(f"\nSaved to: {output_path}")

# Clean up uploaded file
genai.delete_file(video_file.name)
print("Cleaned up uploaded file.")
