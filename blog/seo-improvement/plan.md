# SEO改善プロジェクト計画

## 背景
Google Search Consoleの分析結果に基づくSEO改善。
- インデックス登録済み: 1,414ページ
- インデックス未登録: 992ページ（うちクロール済み未登録820件が最大課題）
- 動画未インデックス: 483件
- 404エラー: 46件
- 平均CTR: 1.9%（改善余地大）

## フェーズ構成

### Phase 1: 現状把握・データ収集（1セッション）
全記事をスキャンして問題を分類する。改善の土台。

**タスク:**
1. 全記事メタデータ取得スクリプト作成
   - 記事URL、タイトル、文字数、カテゴリ、YouTube埋め込み有無、内部リンク数
   - CSVで出力
2. 結果を分析してレポート作成
   - 薄いコンテンツ（文字数少ない記事）リスト
   - YouTube埋め込み記事リスト（→ Phase 2の対象）
   - 内部リンクゼロの孤立記事リスト（→ Phase 4の対象）

**成果物:** `blog/seo-improvement/audit_report.csv`

---

### Phase 2: VideoObject構造化データ追加（1-2セッション）
483件の動画未インデックスに対応。効果が大きく、機械的に処理できる。

**タスク:**
1. YouTube埋め込み記事を特定（Phase 1のデータ利用）
2. 各記事からYouTube動画ID・タイトルを抽出
3. VideoObject JSON-LDスニペットを生成
4. 記事HTMLの末尾に挿入してAtomPub APIで更新
5. 数件テスト → 問題なければバッチ実行

**JSON-LD例:**
```json
{
  "@context": "https://schema.org",
  "@type": "VideoObject",
  "name": "動画タイトル",
  "description": "...",
  "thumbnailUrl": "https://img.youtube.com/vi/VIDEO_ID/maxresdefault.jpg",
  "uploadDate": "2025-01-01",
  "contentUrl": "https://www.youtube.com/watch?v=VIDEO_ID",
  "embedUrl": "https://www.youtube.com/embed/VIDEO_ID"
}
```

**成果物:** 更新スクリプト + 更新済み記事リスト

---

### Phase 3: 薄いコンテンツ対策（2-3セッション）
820件のクロール済み未登録の主因に対処。

**タスク:**
1. Phase 1で特定した薄い記事（目安: 本文500文字未満）をリストアップ
2. 記事の内容を確認し、以下に分類:
   - A: リライト・加筆で改善できる → 加筆
   - B: 他の記事と統合すべき → 統合してリダイレクト
   - C: 価値が低い → noindex追加を検討
3. 優先度の高い記事から順にリライト

**注意:** これは完全自動化できない。1セッションあたり5-10記事ずつ対応。

---

### Phase 4: 内部リンク強化（1-2セッション）
クロールバジェット最適化。孤立記事を減らす。

**タスク:**
1. 全記事の内部リンク構造をマッピング
2. 被リンクゼロの孤立記事を特定
3. 関連記事同士のリンク追加候補を提案
4. 関連記事リンクブロックを自動生成して挿入

---

### Phase 5: メタディスクリプション最適化（1セッション）
CTR 1.9% → 3%以上を目標。

**タスク:**
1. 既存の `scan_meta_descriptions.py` を拡張
2. 概要欄が未設定 or YouTube URLが入っている記事を特定
3. 表示回数が多くCTRが低い記事を優先（GSCデータと突合）
4. 改善案を生成してバッチ更新

---

### Phase 6: 404対応（1セッション）
46件の404エラーを解消。

**タスク:**
1. GSCから404 URLリストを取得
2. 各URLが過去に存在した記事かを調査
3. 移転先がある → リダイレクト設定
4. 復活可能 → 記事を再作成
5. 不要 → GSCから削除リクエスト

---

## 進捗（2026-04-15更新）

### Phase 1: 完了
- `audit_scan.py` で全1,066記事をスキャン済み
- `audit_report.csv` に全データ出力済み
- 結果: 公開997件、下書き69件、YouTube埋め込み319件、500文字未満9件、孤立記事790件

### Phase 2: 完了（2026-04-14）
- `add_video_jsonld_v2.py` で318件にVideoObject JSON-LD追加（1件は既存でスキップ）
- エラー0件、破損0件、全件検証OK
- `video_jsonld_v2_log.csv` にログ記録済み

### Phase 3: 完了（2026-04-14）
- C分類5件を下書きに変更（本文のないリンク集・自己紹介）
- B分類3件を加筆（ボツ記事イラスト・低Na鑑別ツール・自立支援シミュレーター）
- A分類1件は現状維持（病院統計ダッシュボード）

### Phase 4: 完了（2026-04-15）
- `add_internal_links.py` で560件に関連記事リンクブロックを挿入
- 230件はカテゴリ不足でスキップ
- エラー0件、破損0件

### Phase 5: 完了（2026-04-15）
- `fix_meta_desc.py` で120件にリード文を挿入（YouTube URL混入対策）
- 全120件成功、エラー0件、破損0件

### Phase 6: 完了（2026-04-15）
- GSCの404エラー46件を分析・分類
- 全1,066記事をスキャンし、22件の壊れた内部リンクを検出
- 9記事から34件の壊れたリンクを除去・修正（エラー0件）
- 内訳: .htmlファイル名リンク7件、削除entry3件、削除custom entry12件×2重複、/alzheimer-disease 1件
- 残り24件（ページネーション超過9件、削除済みentry15件）は内部リンクなし→Google自然消滅待ち

---

## ファイル一覧（blog/seo-improvement/）

| ファイル | 用途 |
|---------|------|
| `plan.md` | この計画書 |
| `audit_scan.py` | Phase 1: 全記事スキャンスクリプト |
| `audit_report.csv` | Phase 1: スキャン結果（1,066記事） |
| `add_video_jsonld.py` | Phase 2: JSON-LD追加（バグあり・使用禁止） |
| `video_jsonld_log.csv` | Phase 2: 追加ログ（319件のentry_id一覧として有用） |
| `revert_video_jsonld.py` | 修復スクリプト旧版（不完全・使用禁止） |
| `revert_final.py` | 修復スクリプト最終版（正常動作確認済み） |
| `revert_final_log.csv` | 修復ログ（ok:90 skip:229） |
| `revert_log.csv` | 旧修復ログ（参考用） |
| `add_internal_links.py` | Phase 4: 内部リンク挿入スクリプト |
| `internal_links_log.csv` | Phase 4: 挿入ログ（560件） |
| `fix_meta_desc.py` | Phase 5: メタディスクリプション修正 |
| `meta_desc_log.csv` | Phase 5: 修正ログ（120件） |
| `404_urls.csv` | Phase 6: GSCから取得した404 URLリスト（46件） |
| `check_404_urls.py` | Phase 6: 404 URL生存確認スクリプト |
| `404_check_result.csv` | Phase 6: 生存確認結果 |
| `find_broken_links.py` | Phase 6: 壊れた内部リンク検出スクリプト |
| `broken_links_found.csv` | Phase 6: 検出された壊れたリンク（22件） |
| `fix_broken_links.py` | Phase 6: 壊れたリンク修正スクリプト |
| `fix_broken_links_log.csv` | Phase 6: 修正ログ（9記事34件） |

---

## 進行ルール
- 1セッション = 1フェーズ（または半フェーズ）ずつ進める
- 各フェーズ完了時にGSCの数値変化を確認（反映に1-2週間かかる）
- 「SEO改善の続き」で再開可能
