# プロジェクト概要

医学教育コンテンツ（YouTube動画・ブログ記事・診断ツール）の開発プロジェクト。

## 主な成果物
- **医学診断ツール (HTML)**: 低ナトリウム血症、全身浮腫、高CK血症、高BUN、低カリウム血症、神経局在診断など
- **YouTube用スライド (PPTX)**: pptxgenjsで作成、LibreOfficeでPDF変換
- **はてなブログ記事 (HTML)**: 医療従事者向け
- **GitHub Pages**: 診断ツールの公開

## フォルダ構成
```
Claude_task/
├── medical-ddx-tools/    ← 診断ツール本体（GitHub Pagesデプロイ対象）
├── youtube-slides/       ← YouTube動画トピック別フォルダ（スライド・台本・ブログ）
├── hospital-data/        ← 病院統計データ・ダッシュボード
├── nerve-excel/          ← 神経ツールExcel・生成スクリプト
├── blog/                 ← 単発ブログ記事
├── scripts/              ← 汎用スクリプト（スライド変換等）
├── docs/                 ← ドキュメント類
└── archive/              ← 旧版ツール（medical-ddx-toolsに移行済み）
```

## 技術スタック
- HTML/CSS/JavaScript（診断ツール、単一ファイル構成）
- Node.js + pptxgenjs（スライド作成）
- Python 3.13（`C:/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe`）
- GitHub CLI（`export PATH="$PATH:/c/Program Files/GitHub CLI"` が必要）
- GitHub Pages（hinyan1016アカウント）

## よく使うコマンド
- スライド作成: `NODE_PATH="C:/Users/jsber/AppData/Roaming/npm/node_modules" node <script>.js`
- Python実行: `/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe`
- GitHub CLI: `export PATH="$PATH:/c/Program Files/GitHub CLI" && gh ...`
- LibreOffice (headless): `"/c/Program Files/LibreOffice/program/soffice.exe" --headless --convert-to <format> --outdir <dir> <file>`
- pdftoppm: `"/c/Users/jsber/AppData/Local/Microsoft/WinGet/Packages/oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe/poppler-25.07.0/Library/bin/pdftoppm.exe" -png -r 150 <pdf> <output_prefix>`

## スライド視覚検証手順（LibreOffice + pdftoppm）
PPTXスライド作成後、以下の手順でレンダリング品質を確認する：
1. **PPTX→PDF変換**: `soffice --headless --convert-to pdf --outdir <dir> <pptx>`
2. **PDF→ページ別PNG変換**: `pdftoppm -png -r 150 -f <from> -l <to> <pdf> <prefix>`
3. **PNG画像をReadツールで視覚確認**: レイアウト崩れ・文字切れ・配色・フォントをチェック
4. **問題があればスクリプトを修正して再生成**

※ サムネイル（1スライド）の場合はPPTX→PNG直接変換も可: `soffice --headless --convert-to png`

## コーディング規約
- 日本語UI・日本語コメント
- 診断ツールはHTML単一ファイル（CSS/JS内包）
- レスポンシブデザイン対応

---

## 医学診断ツール作成ガイド

### ファイル構成
- ツール本体: `C:\Users\jsber\OneDrive\Documents\Claude_task\medical-ddx-tools\`
- 全ツール単一HTMLファイル（CSS・JS内包、外部依存なし）
- `index.html` — ツール一覧ページ（カード型リンク集）
- `sw.js` — Service Worker（キャッシュ管理）
- `manifest.json` — PWA設定

### 新規ツール作成時の全手順
1. **HTMLファイル作成** — 下記アーキテクチャに従い `<tool_name>.html` を作成
2. **index.htmlにカード追加** — 適切な位置に `<a class="tool-card">` ブロックを追加
3. **sw.js更新** — ASSETS配列に `'./<tool_name>.html'` を追加し、`CACHE_NAME` のバージョンを1つ上げる（例: `ddx-tools-v8` → `ddx-tools-v9`）
4. **git commit & push** — `git add <files> && git commit && git push origin main`
5. **デプロイ確認** — 45秒待機後、WebFetchで `https://hinyan1016.github.io/medical-ddx-tools/<tool_name>.html` を確認

### HTMLテンプレート構造（`<head>`必須要素）
```html
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<script>if('serviceWorker'in navigator){navigator.serviceWorker.getRegistrations().then(function(r){r.forEach(function(reg){reg.update()})});navigator.serviceWorker.addEventListener('message',function(e){if(e.data&&e.data.type==='SW_UPDATED')window.location.reload()})}</script>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="theme-color" content="#XXXXXX">
<meta name="apple-mobile-web-app-capable" content="yes">
<link rel="manifest" href="manifest.json">
<link rel="apple-touch-icon" href="icons/icon-192.png">
<title>ツール名</title>
```
- `<meta charset>` 直後にSW強制更新スクリプト（必須、省略不可）
- `theme-color` はツールごとのテーマカラーに合わせる

### JSアーキテクチャ
```
DISEASES = { disease_id: { name, category, description, tags: {step_id: [tag_ids]}, keyFindings, workup } }
STEPS = [ { id, title, subtitle, multi, options: [{id, label, icon}] } ]
calcScores(answers) → {disease_id: score}
rankDiseases(scores) → [{id, score, ...}]
getProbLevel(score) → "high"/"moderate"/"low"/"very_low"
render() → DOM更新
```

### JS記法ルール（ブラウザ互換性）
- `var` を使用（`const`/`let` ではなく）
- `function(){}` を使用（アロー関数 `=>` ではなく）
- render()内は文字列連結（テンプレートリテラルではなく）
- これによりiOS Safari古いバージョンなどでの互換性を確保

### テーマカラー一覧（既存ツール）
| ツール | primary-dark | primary | 用途 |
|--------|-------------|---------|------|
| AKI | #1A5276 | #2E86C1 | 腎臓系 |
| 浮腫 | #1A5276 | #2E86C1 | 腎臓系 |
| 低Na | #1A5276 | #2E86C1 | 電解質 |
| 高Na | #1A5276 | #2E86C1 | 電解質 |
| 高CK | #1A5276 | #2E86C1 | 筋疾患 |
| 低K | #1A5276 | #2E86C1 | 電解質 |
| 高BUN | #1A5276 | #2E86C1 | 腎臓系 |
| 腹痛 | #1A5276 | #2E86C1 | 消化器 |
| 両側視床 | #4A0E4E | #7B2D8E | 神経画像 |
| 感覚性PN | #1A5276 | #2E86C1 | 神経 |
| 二次性高血圧 | #8B1A1A | #C0392B | 循環器 |
| 外転神経麻痺 | #1A5276 | #2E86C1 | 神経 |
| PN鑑別・検査 | #2C6E49 | #4C956C | 神経 |
| 神経局在 | #1A5276 | #2E86C1 | 神経 |
| 多発脳神経 | #1A5276 | #2E86C1 | 神経 |

### index.htmlカード形式
```html
<a class="tool-card" href="<tool_name>.html">
  <div class="tool-emoji">絵文字</div>
  <div class="tool-name">ツール名</div>
  <div class="tool-desc">説明文</div>
  <span class="tool-tag">Nステップ / N疾患鑑別</span>
</a>
```

### Service Worker (sw.js) 仕様
- CACHE_NAME: `ddx-tools-vN`（新ツール追加ごとにN+1）
- ASSETS配列: 全HTML + manifest.json + icons
- フェッチ戦略: HTML=network-first / 静的アセット=cache-first
- activateイベントで古いキャッシュ削除 + `SW_UPDATED` postMessage
- 各HTMLのheadにSW更新リスナー（上記テンプレート参照）

### 参考ブログ記事からツール作成する際の流れ
1. WebFetchでブログ記事の内容を取得
2. 疾患リスト・鑑別ステップ・検査項目を抽出
3. 医学的正確性を確認（ガイドライン・基準値の照合）
4. 上記アーキテクチャに従いHTMLファイルを作成
5. index.html・sw.js更新 → git push → デプロイ確認

### 既知の注意点
- render()内でダブルクォートを含む文字列（例: "vanishing tumor"）は文字列連結で安全に処理する
- 空ページ問題の主原因はSWキャッシュ。sw.jsのバージョンを必ず上げること
- GitHub Pagesデプロイには約30-60秒かかる

---

## はてなブログ記事 SEO必須項目

新規ブログ記事の作成・投稿時に必ず含める要素（SEO改善プロジェクト2026-04の知見）。

### 記事作成時
1. **導入文（リード文）**: 記事冒頭に120文字以上。主キーワードを含む。GSCのメタディスクリプションになる
2. **本文500文字以上**: 薄いコンテンツはGoogleにインデックスされにくい
3. **内部リンク2〜5件**: 関連記事への正規はてなURL（`https://hinyan1016.hatenablog.com/entry/...`）
4. **VideoObject JSON-LD**: YouTube埋め込み記事には必須（動画インデックス登録用）
5. **FAQ構造化データ**: 検索結果のリッチスニペット対応

### リンクのルール
- 内部リンクは必ず**正規はてなURL**を使用（`/entry/YYYY/MM/DD/HHMMSS` 形式）
- `.html`ファイル名や相対パスは使わない（404の原因）
- リンク先が公開中であることを確認

### 記事の下書き化・削除時
記事を下書きにする場合、**必ず**以下を同時実施:
1. 他の記事からの内部リンク（関連記事ブロック含む）を検索・除去
2. まとめページ等でのリンクも除去
3. リンク除去後に空になるページは連鎖的に下書き化

### VideoObject JSON-LDテンプレート
```html
<script type="application/ld+json">
{"@context":"https://schema.org","@type":"VideoObject","name":"動画タイトル","description":"説明","thumbnailUrl":"https://img.youtube.com/vi/VIDEO_ID/maxresdefault.jpg","uploadDate":"YYYY-MM-DD","contentUrl":"https://www.youtube.com/watch?v=VIDEO_ID","embedUrl":"https://www.youtube.com/embed/VIDEO_ID"}
</script>
```

---

## ブログ文献リンク有効化プロジェクト進捗

はてなブログ(hinyan1016)の参考文献にDOI/URLリンクを追加する継続作業。
「ブログ文献リンクの続きをお願いします」で再開可能。

### 完了月
- 2026年3月/2月/1月、2025年12月/11月/10月/9月 — 完了
- 2025年8月 — 完了（2025-03-26）。全61記事スキャン、68件リンク追加
- 2025年7月 — 完了（2025-03-26）。282件リンク追加（P1:145+P2:137）。6月以前は対応不要

### 累計実績
**862件のリンク追加 + 9件のリンク修正**

### 技術メモ（要点）
- iframe直接変更→submit[0].click()で保存。beforeunload対策必須
- バッチスキャン: アーカイブページから`a.entry-title-link`のhrefでURL取得→fetchでref-itemチェック
- 一部記事はref-contentスパンなし（テキスト直接div内）→テキスト直接置換方式
- 詳細はローカルメモリ `~/.claude/projects/.../memory/project_blog_reference_links.md` 参照
