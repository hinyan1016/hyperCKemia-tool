# 外来でそのまま使える食事指導シリーズ — YouTube動画 制作ガイド

## シリーズ概要

- **シリーズ名**: 外来でそのまま使える食事指導シリーズ
- **全10回構成**
- **対象**: 医療従事者（医師・看護師・管理栄養士）
- **コンセプト**: 疾患別の食事指導を、外来で即実践できる形で解説

## 各回タイトル・対応表

| # | フォルダ | YouTube動画タイトル（仮） | ブログ記事 | 状態 |
|---|---------|--------------------------|-----------|------|
| 1 | 01_総論 | 食事指導は「何を禁止するか」ではなく「何を守るか」から始める | diet_guidance_series_01.html | 未着手 |
| 2 | 02_高血圧 | 高血圧の食事指導 ― 減塩だけでは不十分な理由 | diet_guidance_series_02_hypertension.html | 未着手 |
| 3 | 03_脂質異常症 | 脂質異常症の食事指導 ― LDLと中性脂肪で戦略が違う | diet_guidance_series_03_dyslipidemia.html | 未着手 |
| 4 | 04_糖尿病 | 糖尿病の食事指導 ― 血糖を安定させる3つのルール | diet_guidance_series_04_diabetes.html | 未着手 |
| 5 | 05_肥満 | 肥満の食事指導 ― 食べ方を変える行動療法的アプローチ | diet_guidance_series_05_obesity.html | 未着手 |
| 6 | 06_CKD | CKDの食事指導 ― たんぱく質・塩分・カリウムの三本柱 | diet_guidance_series_06_ckd.html | 未着手 |
| 7 | 07_肝疾患 | 肝疾患の食事指導 ― 脂肪肝と肝硬変で真逆になる理由 | diet_guidance_series_07_liver.html | 未着手 |
| 8 | 08_嚥下障害 | 嚥下障害の食事指導 ― 食形態選択と誤嚥予防の実際 | diet_guidance_series_08_dysphagia.html | 未着手 |
| 9 | 09_術後消化管 | 消化管術後の食事指導 ― 胃切除・膵切除・短腸の栄養管理 | diet_guidance_series_09_postop_gi.html | 未着手 |
| 10 | 10_フレイル | フレイル・多疾患併存の食事指導 ― 制限よりも「守る」優先順位 | diet_guidance_series_10_frailty_multimorbidity.html | 未着手 |

## デザイン統一設定

### テーマカラー
- **メインカラー**: #2E7D32（緑系 — 栄養・食事のイメージ）
- **サブカラー**: #1B5E20（ダーク緑）
- **アクセント**: #FF8F00（オレンジ — 注意喚起・強調）
- **背景**: #FAFAFA

### フォント
- **日本語**: BIZ UDPGothic
- **英語**: Segoe UI
- **文字コントラスト**: 背景との比率 4.5:1 以上を確保

### 共通スライド構成
1. **タイトルスライド** — シリーズ名 + 第N回 + 回タイトル
2. **本編スライド** — 各回の内容
3. **まとめスライド** — Key Takeaways（3〜5項目）
4. **次回予告スライド** — 次回テーマの予告（最終回は全体振り返り）
5. **エンドスライド** — チャンネル登録・再生リスト案内

## 各回フォルダの標準ファイル構成

```
NN_テーマ/
├── create_slides.js        ← pptxgenjs スクリプト
├── slides.pptx             ← 生成されたスライド
├── slides.pdf              ← PDF変換版
├── script.md               ← 発表台本
├── thumbnail.jpg           ← サムネイル画像
└── YouTube説明欄.md         ← 説明文・タイムスタンプ・関連リンク
```

## ブログ素材の所在

ブログ記事・画像素材は以下に格納済み:
```
blog/dietary-guidance-series/
├── diet_guidance_series_NN_*.html   ← はてなブログ用HTML
├── diet_guidance_series_NN_*.md     ← Markdown原稿
├── NN_テーマ/
│   ├── infographic.png              ← Canva図解
│   └── eyecatch.png                 ← Geminiアイキャッチ画像
└── images_index.md                  ← 画像管理台帳
```

※ ブログのインフォグラフィック・アイキャッチはスライドに転用可能

## YouTube説明欄テンプレート

```markdown
外来でそのまま使える食事指導シリーズ 第N回「タイトル」

▼ 目次
00:00 はじめに
XX:XX セクション1
XX:XX セクション2
XX:XX まとめ

▼ 関連ブログ記事
https://hinyan1016.hatenablog.com/entry/YYYY/MM/DD/XXXXXX

▼ 再生リスト
（再生リストURL）

▼ シリーズ全10回
第1回: 総論 ― 食事指導の基本フレーム
第2回: 高血圧
第3回: 脂質異常症
第4回: 糖尿病
第5回: 肥満
第6回: CKD（慢性腎臓病）
第7回: 肝疾患
第8回: 嚥下障害
第9回: 術後消化管
第10回: フレイル・多疾患併存

#食事指導 #医学教育 #栄養管理
```

## PDF変換コマンド

PowerPoint COM オートメーション（VBScript）で変換:
```bash
cscript //nologo "../_shared/convert_to_pdf.vbs" "C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\NN_テーマ\slides.pptx"
```
※ パスはWindowsフルパス（バックスラッシュ）で指定すること

## 制作ワークフロー（1回あたり）

1. ブログ記事（`.md`）を読み込み、スライド構成を決定
2. `create_slides.js` を作成（`_shared/series_template.js` を参照）
3. pptxgenjs でスライド生成
4. `_shared/convert_to_pdf.vbs` で PDF変換
5. `script.md` に発表台本を執筆
6. サムネイル作成
7. `YouTube説明欄.md` を作成
8. 動画収録・編集 → アップロード
