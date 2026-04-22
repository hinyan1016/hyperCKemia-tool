# GPT Image 2 プロンプト集 — 脳梗塞急性期のヘパリン

ChatGPT Plus/Pro の画像生成機能（GPT Image 2 / 旧DALL-E）にそのまま貼り付けて使える3パターン。

## 運用のコツ（GPT Image 2 の癖）
- **日本語は短いフレーズで**: 長文は化ける。見出し＋3〜5語の箱書きが安全
- **アスペクト比は明示**: `aspect ratio 2:3` `portrait` `1024x1536` など
- **既存インフォグラフィックを参考画像として添付** → スタイル統一（食事指導シリーズの `infographic.png` などをアップロード）
- **1発目で完璧は狙わない**: 出力後に「この箇所の文字を〇〇に修正」と会話で追い込む
- **日本語の誤字は頻発** → 出力後にPhotoshop/Canvaで軽く修正するか、テキスト部分を別途オーバーレイする運用がロバスト

---

## ① 縦長インフォグラフィック（ブログ本文埋め込み用） 推奨

**用途**: はてなブログ本文冒頭 / スマホ閲覧最適 / SNSストーリーズ転用可
**アスペクト比**: 2:3（portrait, 1024×1536）

```
Create a clean, professional medical infographic in Japanese for neurologists, portrait orientation (aspect ratio 2:3). Navy blue (#1B3A5C) and white color scheme with accent blue (#2C5AA0). Modern flat design, sans-serif typography (BIZ UDPGothic style).

TITLE AREA (top, dark navy background):
- Large title: "脳梗塞急性期のヘパリン"
- Subtitle: "結局どう使う？"
- Small tag: "ガイドライン2025対応"

SECTION 1 — "3つの論点" (white card with icons):
- ① 用量: "1万単位 vs APTT管理"
- ② ブリッジング: "もう不要？"
- ③ HIT: "低用量でも発症"

SECTION 2 — "実践フロー" (horizontal timeline with arrows):
- "0h 開始 1万〜1.5万単位/日"
- "6-12h APTT測定"
- "目標 前値の2〜2.5倍"

SECTION 3 — "心原性脳塞栓症" (highlighted green panel):
- Heading: "ブリッジング不要"
- Text: "DOAC早期開始が標準"
- Badge: "1-2-3-4日ルール"

SECTION 4 — "HIT警戒" (red alert panel):
- Heading: "4Tsスコア ≥4点で即対応"
- Bullets:
  - "全ヘパリン即中止"
  - "アルガトロバン開始"
  - "ワルファリン禁忌"
  - "血小板輸血も禁忌"

FOOTER:
- "医知創造ラボ | 脳神経内科"

Design requirements: clinical and trustworthy appearance, high readability, ample white space between sections, subtle shadows on cards, thin horizontal dividers in navy. No photographs, no illustrations of patients. Icons can be simple line-style medical symbols (drop, clock, warning triangle). Ensure Japanese text is rendered clearly and accurately.
```

---

## ② 正方形サムネイル（Xポスト / Instagram投稿用）

**用途**: SNS拡散 / 記事アイキャッチ
**アスペクト比**: 1:1（square, 1024×1024）

```
Create a square medical infographic in Japanese (1:1 aspect ratio) designed as an attention-grabbing social media post for neurologists and junior doctors.

Background: deep navy blue (#1B3A5C) with subtle diagonal geometric accent in brighter blue (#2C5AA0).

MAIN HEADLINE (upper half, large white typography):
- "脳梗塞ヘパリン"
- "3つの誤解"

THREE NUMBERED CARDS (lower half, stacked or arranged in a row, white cards on navy background):
1. "1万単位で漫然投与"
   → 小さく: "APTTで微調整が必須"
2. "心原性ブリッジング必須"
   → 小さく: "DOAC早期で不要 (ELAN/OPTIMAS)"
3. "低用量ならHIT安全"
   → 小さく: "5千単位でも発症しうる"

BOTTOM:
- small badge: "ガイドライン2025"
- logo text: "医知創造ラボ"

Style: bold modern editorial, clean typography, high contrast for feed visibility, flat design (no gradients on main elements). Japanese text must be clear and accurately rendered. No medical photographs.
```

---

## ③ 横長YouTubeサムネイル風（ブログOGP兼用）

**用途**: はてなブログ OGP画像 / YouTube関連ポスト / Twitter Card
**アスペクト比**: 16:9（landscape, 1536×1024）

```
Create a 16:9 landscape medical infographic in Japanese for a YouTube thumbnail and blog OGP image, designed to drive clicks among neurology trainees and junior doctors.

Background: split design — left 60% navy blue (#1B3A5C), right 40% white.

LEFT SIDE (navy, white text):
- Eyebrow: "脳神経内科"
- Giant headline: "ヘパリン、"
- Second line: "結局どう使う？"
- Supporting: "1万単位 vs APTT管理"
- Badge bottom-left: "ガイドライン2025"

RIGHT SIDE (white, navy text):
- Three numbered bullets in large navy text:
  ▸ "ブリッジング不要"
  ▸ "APTT 2〜2.5倍"
  ▸ "4Tsスコアで警戒"
- Small accent: "ELAN / TIMING / OPTIMAS"

STYLE: bold editorial, high-contrast typography, clean flat design, minimal line icons (optional blood drop, clock, warning symbol). Japanese characters must be rendered accurately. No photographic imagery. Thumbnail-style with strong hierarchy so the main message is readable even at small size.
```

---

## ワークフロー統合（今後のパイプライン）

```
[ブログHTML作成 / medical-content-pipeline]
   ↓
[本スクリプトを参照してChatGPT Plus/Proで画像生成]
   ↓
   • 既存の食事指導シリーズ infographic.png を参考画像として添付
   • 上記プロンプトをそのまま貼付
   • 出力確認 → 文字化け箇所があれば会話で修正依頼
   • 完成PNGをダウンロード
   ↓
[ローカル保存]
   youtube-slides/<トピック>/infographic_vertical.png
   youtube-slides/<トピック>/infographic_square.png
   youtube-slides/<トピック>/infographic_landscape.png
   ↓
[はてなフォトライフへアップロード]
   blog/_scripts/upload_fotolife_image.py を使用
   ↓
[ブログ記事HTMLへ挿入]
   blog/_scripts/upload_and_insert_infographic.py を使用
   or 手動でHTMLに <img> タグ追加
   ↓
[ブログ記事更新 — 本トピックの update_blog.py を実行]
```

---

## 追加Tips

- **ブログの既存スタイルに寄せたい場合**: ChatGPTチャット内で `youtube-slides/食事指導シリーズ/02_高血圧/blog/infographic.png` をアップロードし、「このスタイルを踏襲して」と一言添える
- **Japanese text化け対策**: 出力後に「画像内の『ヘパリンブリッジング』を『ブリッジング不要』に書き換えて」のように**テキスト単位で追加指示**すると修正可
- **著作権**: GPT Image 2の出力はOpenAI利用規約上、商用・ブログ掲載ともに問題なし
- **医学的免責**: インフォグラフィック下部に「本図は教育目的。個別診療の代替ではない」の但し書きを別途追加するのが無難

---

## 検証ポイント（ChatGPT出力を受け取った後）

- [ ] 数字の誤り（"1万単位" が "10000単位" に化けてないか）
- [ ] 薬剤名スペル（"アルガトロバン" が正しく）
- [ ] 推奨度表記（"推奨度B" → "推奨度8" 等に化けてないか）
- [ ] 全体のバランス・読みやすさ
- [ ] ブランドカラー（Navy #1B3A5C が維持されているか）
