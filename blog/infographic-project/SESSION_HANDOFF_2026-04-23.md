# Session Handoff — 2026-04-23（Day 3 完了時点）

インフォグラフィック・バックフィルプロジェクト、Day 3 まで完走。
本運用ツール（`--date` CLI、sync_generations、export_for_codex）一式を整備。

---

## 累計進捗（2026-04-23 セッション終了時点）

| 指標 | 値 |
|------|-----|
| articles: done | **39 / 227** (17.2%) |
| articles: excluded | 2（1件は削除済み判定） |
| articles: pending | 186 |
| articles: in_batch | 0 |
| generations | 41件（approved 39 + attempt未処理0 + 初期21手動） |

**残り 186 件、ペース 20件/日で 10営業日程度**で完走見込み。

---

## 本セッション（2026-04-23）の成果

### Day 1（18件本番反映、前回セッション実績）

### Day 2（1件試行・完走、本セッション）
- Codex ルートでの画像生成を初めて本番投入
- `6802418398336571447`（2025年感染症トレンド記事）反映 → HTTP 200、画像挿入確認
- 反映中に build_dashboard の相対パスバグ発覚 → 修正

### Day 3（20件一括、本セッション）
- batch_date = `2026-04-25` で20件バッチ生成 → Codex で画像生成 → 一括反映
- 処理結果: OK=20 NG=0 失敗=0
- サンプル公開URL 3件で HTTP 200 + Fotolife img 挿入確認

### 本セッションのコミット（7件）

| hash | 種別 | 内容 |
|------|------|------|
| `7482d21` | fix | build_dashboard の画像相対パス計算修正（Day 2 発覚） |
| `7c6943b` | feat | Codex向け画像生成指示書エクスポート（`export_for_codex.py`）+ handoff doc |
| `fb921e1` | feat | build_dashboard `--date` CLI 引数 |
| `d3f346a` | feat | `sync_generations_from_csv` スクリプト + 5テスト |
| `20b18a7` | docs | RUNBOOK に Codex ルート / sync / --date フロー反映 |
| `2ec48ab` | feat | generate_batch_csv `--batch-date` CLI 引数 |
| -- | -- | Day 3 本番反映（DBと公開記事のみ、コード変更なし） |

### テスト状態
`pytest tests/` = **44件全pass**（元39 + sync_generations 5）。

---

## Day 4 以降の本運用フロー（全CLIで一貫）

```bash
cd blog/infographic-project

# ① 20件バッチ生成（batch_date 指定、AtomPub 存在確認つき）
python -m scripts.generate_batch_csv --batch-date <YYYY-MM-DD> --batch-size 20

# ② Claude がプロンプトを記入
#    `data/entries_for_claude/<date>.md` を読み、
#    `data/batches/<date>_prompts.csv` の style_summary / prompt / alt を埋める
#    （Day 3 では _fill_prompts_<date>.py でPythonスクリプト化、.gitignore対象）

# ③ generations テーブル同期（必須）
python -m scripts.sync_generations_from_csv --date <YYYY-MM-DD>

# ④ Codex 指示書生成
python -m scripts.export_for_codex --date <YYYY-MM-DD>
# → data/batches/<date>_codex_instructions.md をCodexに渡す
# → 画像を images/in/<entry_id>.png に保存（20件）

# ⑤ ダッシュボード生成 → ブラウザでレビュー
python -m scripts.build_dashboard --date <YYYY-MM-DD>
start data/reviews/<date>_review.html
# → OK/NG/Skip選択 → CSVエクスポート → data/decisions/<date>_decisions.csv に配置

# ⑥ dry-run → 本番反映
python -m scripts.process_decisions data/decisions/<date>_decisions.csv --dry-run
python -m scripts.process_decisions data/decisions/<date>_decisions.csv --yes

# ⑦ 検証
curl -s -o /dev/null -w "%{http_code}\n" https://hinyan1016.hatenablog.com/entry/YYYY/MM/DD/HHMMSS
```

### 注意点（Day 3 で学んだこと）

- **ダッシュボードの CSV エクスポート漏れ**: 1カードだけ未選択だったため、CSV が 19件 ＜ バッチ20件 になる事象。対処は(A) ブラウザで再選択、(B) CSV に手動追記、(C) スキップ扱いのいずれか。選ばれたら「ユーザに選択肢を提示」する運用で対応済。
- **`sync_generations_from_csv` を忘れると `process_decisions` は `generations に <eid> がない` で失敗**: Day 3 以降は必須ステップ。
- **batch_date は重複不可**: `batches` テーブルのPKが batch_date。既存日付指定は INSERT OR REPLACE で prompts.csv ごと上書きになるので注意。

---

## 未消化の低優先度タスク

- **Codex ルートのレート制限検証**: Day 3 で20件 Plus範囲内で問題なく回ったが、連日実行時の累積制限は未検証（Day 4-5で要観察）
- **.omc 等のスペック/アーキドキュ整備**: 未着手。現運用が安定すれば不要

---

## 現在の Git 状態（参考）

```
最新: 2ec48ab feat(infographic): generate_batch_csv に --batch-date 引数を追加
ブランチ: main
未コミット: .claude/settings.local.json のみ（プロジェクト外、無関係）
```

### 現在のDB状態

```sql
articles: done=39, excluded=2, pending=186
generations: 41件（全 in_batch 処理済み）
batches:
  2026-04-23 / 2026-04-24 / 2026-04-25 いずれも処理完了
```

---

## 次セッション再開フレーズ

- **「Day 4 バッチ準備」** → batch_date=2026-04-26 で 20件バッチ生成以降のフロー
- **「インフォグラフィックの続き」** → このファイルを読んで現状確認
- **「ブログ文献リンクの続き」** → 別プロジェクト（文献リンク有効化、残561件）

---

**セッション終了時刻**: 2026-04-23（Day 3 完了、39/227件 done、次は Day 4）
