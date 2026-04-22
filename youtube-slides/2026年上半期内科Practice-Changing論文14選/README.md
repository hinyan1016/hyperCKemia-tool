# 2026年上半期 内科Practice-Changing論文14選

医療従事者向けコンテンツの一気通貫パイプライン成果物。

## 現在のステータス（2026-04-21 時点）

| 工程 | 状態 | URL / ファイル |
|------|------|---------------|
| 元素材 PDF | 完了 | `c:\Users\jsber\Downloads\internal_medicine_papers_2026.pdf` |
| 参考文献URL収集 | 完了 | `references.md`（14論文DOI/PubMed + 元PDF誤記訂正2件） |
| はてなブログHTML生成 | 完了 | `blog/article.html`（FAQ JSON-LD含む） |
| **はてなブログ公開** | **✅ 公開済み** | **https://hinyan1016.hatenablog.com/entry/2026/04/21/113203** |
| PPTXスライド生成 | 完了（22枚） | `slides/2026_internal_medicine_top14.pptx` |
| スライド視覚検証 | 完了 | `slides/2026_internal_medicine_top14.pdf` |
| YouTube台本（Markdown） | 完了 | `script/2026年上半期内科Practice-Changing論文14選_発表台本.md` |
| SlideCast JSON台本 | 完了（TTS読み替え済） | `script/slidecast.json` |
| YouTube概要欄テンプレ | 完了（ブログURL反映済） | `youtube_description.md` |
| 動画撮影・公開 | 未着手 | — |
| ブログVideoObject追加 | 未着手 | 動画公開後 |

## 元素材
- `c:\Users\jsber\Downloads\internal_medicine_papers_2026.pdf`（ユーザー提供PDF・14論文）

## ディレクトリ構成

```
2026年上半期内科Practice-Changing論文14選/
├── README.md                     ← このファイル（全体ガイド）
├── HANDOFF.md                    ← セッション再開用ハンドオフ
├── references.md                 ← 14論文のDOI・PubMed URL一覧（元PDF誤記訂正含む）
├── youtube_description.md        ← YouTube概要欄テンプレ（ブログURL反映済）
├── post_blog.py                  ← はてなブログAtomPub投稿スクリプト
├── post_status.json              ← 投稿状態記録（公開済み）
├── blog/
│   └── article.html              ← はてなブログ投稿用HTML
├── slides/
│   ├── generate_slides.js        ← PptxGenJSスライド生成スクリプト
│   ├── 2026_internal_medicine_top14.pptx  ← 22枚スライド本体
│   └── 2026_internal_medicine_top14.pdf   ← LibreOffice生成PDF
├── script/
│   ├── 2026年上半期内科Practice-Changing論文14選_発表台本.md  ← 通常台本
│   └── slidecast.json            ← SlideCast Studio用JSON（TTS読み替え済）
└── assets/                       ← 予備
```

## ファイル別の使い方

### 1. はてなブログ投稿（完了済み）
- 投稿スクリプト: `python post_blog.py`（下書き）/ `python post_blog.py --publish`（公開）
- 認証情報は `youtube-slides/食事指導シリーズ/_shared/.env` を参照
- 注意: Python stdout は UTF-8 再設定済み（`PYTHONIOENCODING=utf-8` 推奨）

### 2. PPTXスライド
```bash
cd slides/
NODE_PATH="C:/Users/jsber/AppData/Roaming/npm/node_modules" node generate_slides.js
"/c/Program Files/LibreOffice/program/soffice.exe" --headless --convert-to pdf --outdir . 2026_internal_medicine_top14.pptx
```

### 3. 台本
- **Markdown版**: `script/2026年上半期内科Practice-Changing論文14選_発表台本.md`（人間の朗読・読み合わせ用）
- **SlideCast JSON**: `script/slidecast.json`（TTS自動読み上げ用。NEJM→ニューイングランドジャーナル、DOAC→ドアック に置換済み）

### 4. YouTube概要欄
- `youtube_description.md` をコピーして概要欄へ
- 公開後にチャプタータイムスタンプを実時間に合わせて修正

## 元PDFの誤記訂正
ブログ・スライド・台本すべてに反映済み：
1. **#7 水分摂取試験**: PDFは「結石再発リスク低減に寄与」 → 正しくは「尿量は増加したが再発抑制効果は確認できず」
2. **#13 病院前全血輸血**: PDFは「RePHILL-2」 → 正しくは **SWiFT試験**

## 残タスク
- [ ] YouTube動画撮影・編集（SlideCast JSON or 台本Markdownを使用）
- [ ] 動画公開後、ブログのVideoObject JSON-LDプレースホルダに動画ID埋め込み
- [ ] YouTube概要欄のチャプタータイムスタンプを実時間に修正
- [ ] 必要に応じて、はてなブログに関連記事の正規URL内部リンク追加（SEO改善）
