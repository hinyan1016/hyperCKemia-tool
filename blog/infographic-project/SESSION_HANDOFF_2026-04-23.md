# Session Handoff — 2026-04-23

インフォグラフィック・バックフィルプロジェクト、Day 1 反映 + API実装 + Day 2 Codex試行準備まで。

---

## Accomplished

### ① Day 1 本番反映（2026-04-23）
- 19件中 **18件成功**: Fotolifeへ画像アップロード + はてなブログ本文へ`max-width:100%; box-shadow`スタイル付きimg挿入
- **1件失敗**: `6802418398334380887`（抗てんかん薬の妊娠に対する催奇形リスクまとめ） → AtomPub 404。はてな側で記事削除済みと確認（AtomPub 108ページスキャンで該当なし）。`articles.status='excluded', excluded_reason='deleted_from_hatena'` マーク、画像は `images/archive/<eid>_deleted_<ts>.png` に退避
- 代表2件（NSAIDs/BBE）で公開URL上のimg挿入を確認

### ② OneDriveコンフリクト整理
- 別PC（Envy14）との同期で発生した `-Envy14faHI.*` 27ファイルを削除（内容は全てCRLFのみの差、ロジック差なし）

### ③ API統合実装（コミット `87f9542`）
| ファイル | 変更 |
|---------|------|
| `scripts/generate_images.py` | **新規** — OpenAI Images API (`gpt-image-1`) で prompts.csv 全件を自動生成。既存ファイル/空プロンプトはスキップ、失敗はリトライ後ログ記録して継続 |
| `scripts/generate_batch_csv.py` | `select_batch_with_existence_check()` 追加。AtomPub GETで404の記事は `excluded_reason='deleted_from_hatena'` マーク＋代替選出。`--skip-existence-check` フラグ追加 |
| `scripts/hatena_client.py` | `entry_exists()` 追加 |
| `requirements.txt` | `openai>=1.40.0` 追加 |
| `tests/integration/test_generate_images.py` | **新規** 4テスト |
| `tests/integration/test_batch_generation.py` | 存在確認テスト 1件追加 |
| `RUNBOOK.md` | 新ワークフロー反映、OPENAI_API_KEY必須化、Day 1チェックリスト更新 |

テスト結果: **39件全pass**（元32 + 新7）

### ④ `.gitignore` 整理（コミット `c65ca6b`）
- `.coverage` / `.coverage.*` / `htmlcov/` 追加
- `data/batches/*.md` 追加（`prompts_for_copy.md`等）
- `data/inventory.csv` / `inventory_summary.txt` 追加（v1は除外、v2のみcommit対象）
- `scripts/_fill_prompts_*.py` 追加（日付固定のワンショット）

### ⑤ Day 2 Codex試行準備（**未コミット**）
- 1件バッチ生成: `2026-04-24_prompts.csv` (entry_id `6802418398336571447`, 「2025年最新の一般内科トレンド：感染症動向と対策」)
- Claudeがプロンプト記入済（感染症4分割: 百日咳/COVID-19/インフルエンザ/エムポックス&AMR）
- `generations` テーブルへ1件手動INSERT（alt_text保存）
- `scripts/export_for_codex.py` 新規作成 — Codex用指示書生成
- `data/batches/2026-04-24_codex_instructions.md` 出力済

---

## Current State

### 動いているもの
- パイプライン全体（generate_batch_csv → build_dashboard → process_decisions → rollback）
- 39テスト全pass
- OpenAI API経由の画像生成スクリプト（API key設定後に動作可能）
- AtomPub存在確認つきバッチ選出

### 部分的な実装
- **Codexルート**: 指示書生成は完了、Codex側での画像生成はユーザ作業待ち
- **build_dashboard main()**: `date.today()` 固定。未来日（2026-04-24）を指定する CLI 引数がない

### 意図的に後回し
- **`sync_generations_from_csv.py` の恒久化**: Day 2 試行のために1件手動INSERTで凌いだ。Day 3以降に本格運用するには必須
- **build_dashboard への `--date` 引数追加**: Codex結果が揃ったら必要
- **`.env` への `OPENAI_API_KEY` 追加**: Codexルートで問題なく進めばAPI routeは保険扱いでOK

---

## Pending

### 優先度高
1. **Codex画像生成の結果受け取り** — ユーザがCodexで1件生成して `images/in/6802418398336571447.png` に保存
2. **build_dashboard `--date` 引数追加** — Day 2以降のバッチをCLIで処理するため
3. **Day 2 の反映完走テスト** — ダッシュボード→レビュー→process_decisions

### 優先度中
4. **`sync_generations_from_csv.py` 実装** — CSVのprompt/alt/styleをgenerationsテーブルに自動同期（UPSERT）
5. **`export_for_codex.py` のcommit** — 機能確認後にcommit

### 優先度低
6. **RUNBOOK に Codex ルート追記** — API/Codex両方の分岐を明記
7. **Codexルートのレート制限検証** — 20件連続でもPlus範囲内で回るか

### 既知の課題
- **OneDriveロックによるgit失敗** → 回避済み（ユーザが `git config windows.appendAtomically false` 設定）
- **state.sqlite はOneDrive上**: ロック問題で同期衝突のリスクあり。別PCで作業する場合は注意

---

## Context for Next Session

### 最初に読むべきファイル
1. `blog/infographic-project/RUNBOOK.md` — 日次運用手順（最新）
2. `blog/infographic-project/SESSION_HANDOFF_2026-04-23.md` — これ
3. `blog/infographic-project/data/batches/2026-04-24_codex_instructions.md` — Codex向け指示書
4. `blog/infographic-project/scripts/generate_images.py` — OpenAI APIルート（実装済）

### 重要な決定事項
- **画像生成は当面Codexルートを本命に**（ChatGPT Plus範囲、追加課金ゼロ）。OpenAI APIルートは保険として温存
- **VS Code上で Codex と Claude Code が同じワークスペースを共有** → ファイルリレーで分業可能
- **generations テーブルはプロセス全体の中核** — alt_text格納が process_decisions の前提条件
- **batch_date キーは日付。当日のCSV/MDを上書きするため、テストは未来日で行う**

### 非自明なGotchas
- **process_decisions は generations 行を必須とする**。alt_text がないと `RuntimeError("generations に {eid} がない")` で失敗。現状 generate_batch_csv はarticlesしか触らないので、CSV記入と generations 書込みは別ステップ
- **CRLF警告は無視してOK** — Windows環境下でgitの改行変換警告が常時出るが、実害なし
- **build_dashboard の HTML には `awaiting_image` がCSS内に1回出現** — ステータス検出のとき正規表現要注意
- **はてな側で削除された記事は AtomPub でも404** — 棚卸しv2時点のsnapshotに対してリアルタイム存在確認が必要

### 現在のDB状態（`data/state.sqlite`）
```
articles:
  done:      18  (Day 1 で反映完了)
  excluded:   2  (1件はv2除外 + 1件は削除済み判定)
  pending:  207
  in_batch:   1  (Day 2試行中: 6802418398336571447)

generations:
  計21件 (Day 1分20 + Day 2試行1)
```

### 現在のGit状態
```
commits:
  c65ca6b chore(infographic): .gitignore に .coverage / batches/*.md / v1棚卸し / ワンショット追加
  87f9542 feat(infographic): 画像生成自動化 - OpenAI Images APIと記事存在確認
  60399b8 fix(infographic): ... (Day 1セッション以前)

未コミット（プロジェクト内）:
  - blog/infographic-project/scripts/export_for_codex.py   ← 新規、未追跡
  - blog/infographic-project/data/batches/2026-04-24_*.md  ← gitignore対象
  - blog/infographic-project/data/state.sqlite             ← gitignore対象

ブランチ: main
```

---

## Resume Commands

### 状態確認（新セッション開始時）
```bash
cd "C:/Users/jsber/OneDrive/Documents/Claude_task"

# gitとDB状態
git log --oneline -5
git status --short | grep infographic-project

# DB確認
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -c "
import sqlite3
conn = sqlite3.connect(r'blog/infographic-project/data/state.sqlite')
for r in conn.execute(\"SELECT status, COUNT(*) FROM articles GROUP BY status\"):
    print(r)
"

# Codex側の画像到着確認
ls blog/infographic-project/images/in/
# → 6802418398336571447.png があれば Codex生成済
```

### テスト（健全性チェック）
```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest -v
# 39件全passが期待値
```

### Codex画像が揃ったら（Day 2 反映フロー）
```bash
cd blog/infographic-project

# 1. build_dashboard を未来日で実行（--date引数未実装なのでPython直叩き）
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -c "
from pathlib import Path
from scripts.build_dashboard import build_dashboard
build_dashboard(
    Path('data/batches/2026-04-24_prompts.csv'),
    Path('images/in'),
    Path('data/reviews/2026-04-24_review.html'),
    batch_date='2026-04-24',
)
print('OK')
"

# 2. ブラウザで開いてレビュー → decisions.csv ダウンロード → data/decisions/ に配置
start "" "data/reviews/2026-04-24_review.html"

# 3. dry-run 確認
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m scripts.process_decisions data/decisions/2026-04-24_decisions.csv --dry-run

# 4. 本番反映
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m scripts.process_decisions data/decisions/2026-04-24_decisions.csv --yes

# 5. 検証
curl -s -o /dev/null -w "%{http_code}\n" https://hinyan1016.hatenablog.com/entry/2025/03/15/075022
```

### Day 3 以降の本運用
```bash
cd blog/infographic-project

# バッチ生成（存在確認つき）
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m scripts.generate_batch_csv --batch-size 20

# Claudeがプロンプト記入
# → 2026-04-NN_prompts.csv の空欄を手で/会話で埋める

# (未実装) sync_generations — 現状は手動INSERTで凌ぐ必要あり
# 1件ずつ手動:
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -c "... 前回手動でやった処理 ..."

# Codex用指示書生成
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m scripts.export_for_codex

# → Codex に data/batches/<date>_codex_instructions.md を渡す
# → images/in/ に画像が揃う

# 以降 build_dashboard → レビュー → process_decisions
```

### 緊急ロールバック（反映後にNGが見つかった場合）
```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m scripts.rollback <entry_id>
# ※ body_html_before が保存されている7日以内のみ有効
```

---

## 会話での再開フレーズ

次セッションで以下のいずれかを言えば流れを戻せます:

- 「インフォグラフィックの続き」 → このファイルを読んで現状確認から
- 「Codex画像保存完了」 → Day 2 build_dashboard以降へ直行
- 「Day 3 バッチ準備」 → 20件バッチ生成から

---

**セッション終了時刻**: 2026-04-23（主要コミット2件、全テスト39件pass、Day 2試行準備完了）
