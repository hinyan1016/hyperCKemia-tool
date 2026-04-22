# セッション再開用ハンドオフ

**作成日**: 2026-04-21
**プロジェクト**: 2026年上半期 内科Practice-Changing論文14選
**作業フォルダ**: `C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\2026年上半期内科Practice-Changing論文14選\`

---

## このセッションで完了したこと

ユーザー提供の `c:\Users\jsber\Downloads\internal_medicine_papers_2026.pdf`（NEJM/Lancet/JAMA掲載14論文サマリー）から、`/oh-my-claudecode:medical-content-pipeline` 相当の一気通貫ワークフローで以下を生成・公開：

1. **参考文献URL収集** → `references.md`
2. **はてなブログ記事** → `blog/article.html` を生成 → `post_blog.py` で下書き投稿 → ユーザーが管理画面から手動公開
3. **PPTXスライド22枚** → `slides/generate_slides.js` で生成、PDF/PNG視覚検証済み
4. **YouTube台本Markdown** → `script/...発表台本.md`
5. **SlideCast Studio用JSON** → `script/slidecast.json`（NEJM→ニューイングランドジャーナル、DOAC→ドアックに置換済み、subtitleは略語のまま維持）
6. **YouTube概要欄テンプレ** → `youtube_description.md`（公開済みブログURL反映）

**公開済みブログURL**: https://hinyan1016.hatenablog.com/entry/2026/04/21/113203

---

## 別マシン/別セッションで再開する手順

### 1. 状況把握（5分）
```
cd "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/2026年上半期内科Practice-Changing論文14選"
cat README.md         # 全体ステータス表
cat HANDOFF.md        # このファイル
cat post_status.json  # 公開状態
```

### 2. ファイル全体像
すべての成果物パスは `README.md` のディレクトリ構成セクション参照。

### 3. 残タスク（優先順）

| # | タスク | 必要ファイル | 備考 |
|---|--------|-------------|------|
| ~~1~~ | ~~YouTube動画撮影~~ | ~~`slides/...pptx` + `script/slidecast.json`~~ | **✅ 2026-04-22 完了 — 動画ID `ksjQHzu43Vo`（https://youtu.be/ksjQHzu43Vo）** |
| ~~2~~ | ~~動画公開後、ブログにVideoObject JSON-LDを追加~~ | ~~`blog/article.html`~~ | **✅ 2026-04-22 完了 — `update_blog.py` で公開記事を上書きPUT** |
| ~~3~~ | ~~YouTube概要欄のチャプタータイムスタンプを実時間に更新~~ | ~~`youtube_description.md`~~ | **✅ 2026-04-22 完了 — Claude in Chrome経由でStudioに直接投入（タイトル＋概要欄3,068字）** |
| 4 | （任意）はてなブログ記事に関連記事の内部リンク追加 | 公開済みブログ管理画面 | SEO改善 |

### 2026-04-22 セッションの完了内容

- ブログ `article.html` に以下を追加し `update_blog.py` で公開記事を差し替え：
  - 動画レスポンシブiframe埋込（16:9）
  - VideoObject JSON-LD（thumbnailUrl／contentUrl／embedUrl／uploadDate 2026-04-22）
  - ※ チャプター一覧はユーザー指示により**非設置**（SEOはJSON-LD側で担保）
- YouTube Studio を Claude in Chrome で開き、以下を自動入力・保存：
  - タイトル：`2026年上半期 内科Practice-Changing論文14選 ― Less is More／上流介入／予防医療進化の3潮流`
  - 概要欄：3,068字（リード＋ブログURL＋20行チャプター＋参考文献14本＋チャンネル誘導＋免責＋ハッシュタグ）
- 新規スクリプト `update_blog.py`（既存エントリPUT更新用）追加
- `post_status.json` に YouTube公開情報とブログ側の動画統合状態を記録

### 再開用スクリプト

- `post_blog.py` — 新規下書き／公開POST用（既存）
- `update_blog.py` — 既存エントリを `article.html` から上書きPUT（新規）
  - 使い方：`PYTHONIOENCODING=utf-8 python update_blog.py`（`post_status.json` の `edit_url` を参照）

### 4. 元PDFの誤記訂正（重要・全成果物に反映済み）
- **#7 水分摂取試験**: PDFは「再発リスク低減」→ 実際は「再発抑制効果なし」
- **#13 病院前全血輸血**: PDFは「RePHILL-2」→ 正しくは **SWiFT試験**

## 環境情報

| 用途 | パス |
|------|------|
| Python 3.13 | `/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe` |
| PptxGenJS（NODE_PATH） | `C:/Users/jsber/AppData/Roaming/npm/node_modules` |
| LibreOffice | `/c/Program Files/LibreOffice/program/soffice.exe` |
| pdftoppm | `/c/Users/jsber/AppData/Local/Microsoft/WinGet/Packages/oschwartz10612.Poppler.../bin/pdftoppm.exe` |
| はてなAPI認証 | `youtube-slides/食事指導シリーズ/_shared/.env`（HATENA_ID/DOMAIN/API_KEY） |

## 既知のハマりどころ

1. **Windows cp932 エンコーディング**: Python のprint で `—`（emダッシュ）等が `UnicodeEncodeError` になる。`PYTHONIOENCODING=utf-8` を環境変数で設定するか、スクリプト先頭で `sys.stdout.reconfigure(encoding="utf-8")` する。
2. **はてなAtomPub URL**: 投稿時刻が変わると URL末尾も変わる（例: 112859 → 113203）。投稿後は必ず `https://hinyan1016.hatenablog.com/` トップから正規URL確認。
3. **SlideCast JSON 改行ルール**:
   - `script` には `\n\n` を入れない（TTSが間として解釈し不自然）
   - `subtitle` は `\n\n` で字幕切替
4. **TTS読み替え対応**: `slidecast.json` の `script` フィールドのみ `NEJM→ニューイングランドジャーナル`、`DOAC→ドアック` に置換済み。今後追加で読み間違えが発覚した略語があれば同パターンで対応（subtitle側は略語維持）。

## 投稿スクリプト再利用

`post_blog.py` はこの記事専用に作ったが、他の単発記事にも転用可能：
- `TITLE` と `CATEGORIES` と `HTML_FILE` を書き換えるだけで使える
- `--publish` 引数で公開モード（ただし管理画面プレビュー後の公開推奨）

## 次回作業を素早く再開するためのプロンプト例

```
2026年上半期内科Practice-Changing論文14選 の続き。
作業フォルダの README.md と HANDOFF.md を読んで、現状を把握してください。
次に [動画撮影 / VideoObject追加 / 概要欄更新 / 等] に進みたい。
```
