# ヤンキー7語の違いプロジェクト 引き継ぎメモ

**最終更新**: 2026-04-21
**再開合言葉**: 「ヤンキー7語の続きをお願いします」／「youtube-slides/ヤンキー7語の違い/HANDOFF.md を読んで」

## プロジェクト概要
- ヤンキー・ツッパリ・不良・やんちゃ・チンピラ・暴走族・半グレの7語の違いを解説する雑学コンテンツ
- hinyan1016の**既存医療チャンネル／ブログ**で雑学回として公開する方針（ユーザー選択）
- 医学とは無関係だが、医学系視聴者向けに一般向けトーンで制作

## 作業ディレクトリ
`C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/ヤンキー7語の違い/`

## 成果物ファイル一覧
| ファイル | 役割 | 状態 |
|---|---|---|
| `blog_markdown.md` | ブログ本文（Markdown, 約5200字） | 完成 |
| `blog_hatena.html` | はてな投稿用HTML（インラインスタイル＋FAQ JSON-LD） | 完成／投稿済 |
| `create_slides.js` | pptxgenjsスクリプト（16枚） | 完成 |
| `ヤンキー7語_スライド.pptx` | スライド本体 | 完成 |
| `ヤンキー7語_スライド.pdf` | 視覚検証用PDF | 完成 |
| `slide_check-01〜16.png` | 視覚検証PNG（16枚全て品質OK） | 完成 |
| `YouTube台本.md` | 1人語り台本（約9分） | 完成 |
| `台本.json` | SlideCast用JSON（TTS最適化済） | 完成 |
| `post_draft.py` | はてな下書き投稿スクリプト | 実行済 |

## 投稿済みURL
- **はてな下書き記事**: https://hinyan1016.hatenablog.com/entry/2026/04/21/182553
- **はてな編集URL**: https://blog.hatena.ne.jp/hinyan1016/hinyan1016.hatenablog.com/atom/entry/17179246901379066755
- **状態**: 下書き（未公開）
- **カテゴリ**: `日本語雑学` / `意味の違い` / `言葉の歴史`
- **タイトル**: ヤンキー・ツッパリ・不良・やんちゃ・チンピラ・暴走族・半グレの違いとは

## 未完了タスク（次セッションの候補）
1. **YouTube動画の収録・公開** — スライドPPTXと台本.json利用
2. **サムネイル作成** — 既存の`thumbnail_slide.pptx`パターン参照（例: `からだの不思議17`フォルダ）
3. **YouTubeタイトル・説明欄・タグ案** — 未着手
4. **VideoObject JSON-LD追加** — 動画公開後にはてな記事本文末尾に追加必要
5. **はてな記事を下書き→公開切り替え** — YouTube公開と同期してタイミング決定
6. **YouTube概要欄にブログURL記載** — 投稿後、ブログの正規URLを反映

## 技術メモ
### スライドデザイン（ブランド整合）
- カラー: ネイビー系（`dark=1B3A5C`, `primary=2C5AA0`, `accent=4A90D9`）
- フォント: 日本語=BIZ UDPGothic / 英語=Segoe UI
- 16:9レイアウト、16スライド構成

### はてなブログ投稿時の注意
- **`<style>`ブロック除去必須**、全装飾インラインスタイル（遵守済）
- `<th>`タグは必ずインラインで `background:#2C5AA0;color:#fff;...` 付与（遵守済）
- AtomPub API使用、.envは `youtube-slides/食事指導シリーズ/_shared/.env` を流用

### SlideCast JSON仕様
- `script`: TTS用、`\n\n`禁止（間延び防止）、難読漢字はひらがな化
- `subtitle`: 字幕用、`\n\n`で表示切替
- 今回のひらがな化箇所: きょせい／えんきょく／ぞくしょう／そりこみ／とっこうふく／ぶべつ／じゅんぼうりょくだん／きゅうしゃかい／はくほうどう／とくめい／りゅうどうがた／やみバイト／いちそうぞう／シーニー／エスエヌエス／エムアールアイ／なんばこうじ／さいとうたまき／うちこしまさゆき

### ファクトチェック済の要点
- やんちゃ（1702年）/ 不良（1749年一般・1905年青少年）/ チンピラ（1915-30）/ ヤンキー（1904アメリカ人・1973若者語）/ 暴走族（1973前後警察白書）/ ツッパリ（1984名詞）/ 半グレ（2010年代）
- 出典: 精選版日本国語大辞典、デジタル大辞泉、警察白書（昭和49年・令和2年・令和5年）、大阪府警察、CiNii論文、博報堂生活総合研究所

## ビルド・再生成コマンド
```bash
# スライド再生成
cd "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/ヤンキー7語の違い"
NODE_PATH="C:/Users/jsber/AppData/Roaming/npm/node_modules" node create_slides.js

# PDF・PNG変換
"/c/Program Files/LibreOffice/program/soffice.exe" --headless --convert-to pdf --outdir . "ヤンキー7語_スライド.pptx"
"/c/Users/jsber/AppData/Local/Microsoft/WinGet/Packages/oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe/poppler-25.07.0/Library/bin/pdftoppm.exe" -png -r 100 "ヤンキー7語_スライド.pdf" slide_check

# はてな投稿（下書き → 公開切り替えは post_draft.py の draft=True を False に）
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe post_draft.py
```

## 再開時の動作フロー
1. 本ファイル（HANDOFF.md）を読む
2. `ls`でディレクトリ内ファイル確認
3. ユーザーに次のアクションを確認（サムネ制作／YouTube公開準備など上記「未完了タスク」から選択）
