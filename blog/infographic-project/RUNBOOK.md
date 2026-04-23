# インフォグラフィック・バックフィル運用手順

**対象**: インフォグラフィック未対応記事227件へのAI生成画像追加
**ペース**: 20件/日 × 約12日で完走予定
**関連設計書**: [`docs/superpowers/specs/2026-04-23-infographic-backfill-design.md`](../../docs/superpowers/specs/2026-04-23-infographic-backfill-design.md)

**2026-04-23 更新**: ChatGPT UIでの手動画像生成 → OpenAI Images API による自動生成に切り替え。
人間の拘束時間が 20-30分 → 5-7分（ダッシュボードレビューのみ）に短縮。

**2026-04-23 更新 (Day 2)**: 画像生成は **Codex ルートを本命**に（ChatGPT Plus範囲、追加課金ゼロ）。
OpenAI API ルートは保険として温存。`sync_generations_from_csv` で Day 3 以降の本運用ワークフローを整備。

---

## 日次ルーチン（10〜15分、人間の拘束は5〜7分のみ）

### ① バッチ生成 + 存在確認（3分）

Claude Codeセッションで:

> 「今日の20件のバッチ生成して」

Claudeが実行:

```bash
cd blog/infographic-project
python -m scripts.generate_batch_csv
```

これにより:
- 候補各記事に対して **AtomPub GET で存在確認**（404なら`excluded=1, excluded_reason='deleted_from_hatena'` に自動マーク、代替を選出）
- `data/batches/<today>_prompts.csv` を生成（プロンプト欄は空）
- `data/entries_for_claude/<today>.md` を生成

オフライン/テスト時は `--skip-existence-check` フラグで存在確認を省略可能。

### ② プロンプト記入（2分）

Claudeが `<today>.md` の各記事について:
- **style_summary**（例: 「神経疾患・青系・解剖図レイアウト」）
- **prompt**（画像生成用、日本語、約300-500字、医学的正確さを確保）
- **alt**（120字前後、主キーワード含む）

を決定し、`<today>_prompts.csv` の空欄を埋める。

### ③ generations テーブル同期（30秒、**必須**）

プロンプト記入後に CSV の内容を `generations` テーブルへ UPSERT:

```bash
python -m scripts.sync_generations_from_csv --date <today>
```

動作:
- `(entry_id, attempt)` キーで既存行を検索
- 空の既存行は UPDATE、存在しなければ INSERT
- 既に `decision` が set された行（approved/ng）は上書きしない
- 省略すると⑥ process_decisions が `generations に <eid> がない` で失敗する

### ④ 画像生成（5〜10分）

画像生成は2ルート。**本命は A）Codex ルート**（追加課金ゼロ）、B）API ルートは保険。

#### A）Codex ルート（本命、ChatGPT Plus範囲）

Codex向け指示書を生成:

```bash
python -m scripts.export_for_codex --date <today>
# → data/batches/<today>_codex_instructions.md
```

生成されたmdを:
- VS Code の Codex タブにドラッグ、または
- 内容をコピペしてChatGPT Plus に依頼

指示書は entry_id ごとに保存ファイル名・参考URL・プロンプト・命名規約を自動整形。
画像は `images/in/<entry_id>.png` に保存。完了後は Claude Code セッションに「Codex画像保存完了」と伝える。

#### B）API ルート（保険、大量・完全自動化向け）

```bash
python -m scripts.generate_images --date <today>
```

動作:
- OpenAI Images API (`gpt-image-1`, 1024x1024, quality=`medium`) で全行を生成
- 出力先: `images/in/<entry_id>.png`
- 既に存在するファイルはスキップ（idempotency）
- 失敗は `logs/image_gen.log` に記録、他の行は継続
- デフォルトは `medium` 品質（$0.07/枚、19件で約$1.33）

品質パラメータをオプションで調整:

```bash
python -m scripts.generate_images --quality high       # $0.19/枚
python -m scripts.generate_images --size 1536x1024     # 横長
```

### ⑤ ダッシュボードレビュー（5〜7分、人間担当）

> 「ダッシュボード作って」

Claudeが実行:

```bash
python -m scripts.build_dashboard --date <today>
```

1. 生成された `data/reviews/<today>_review.html` をブラウザで開く
2. 各カードで OK/NG/スキップを選択
3. NGの場合は理由タグ（medical_error/typo/design/mismatch/other）と自由記述を入力
4. 下部の「決定をCSVにエクスポート」ボタン押下
5. ダウンロードされた CSV を `data/decisions/<today>_decisions.csv` として保存

### ⑥ 反映（3分）

> 「決定を反映して」

Claudeが実行:

```bash
python -m scripts.process_decisions data/decisions/<today>_decisions.csv
# "本番のはてなブログを更新します。続行しますか?" → yes
```

結果確認:

```
完了: OK=17 NG=2 Skip=1 既処理=0 失敗=0
```

### ⑦ 進捗確認

```bash
python -c "
import sqlite3
conn = sqlite3.connect('data/state.sqlite')
done = conn.execute(\"SELECT COUNT(*) FROM articles WHERE status='done'\").fetchone()[0]
total = conn.execute(\"SELECT COUNT(*) FROM articles\").fetchone()[0]
print(f'{done}/{total} 完了、残 {total-done} 件')
"
```

---

## 前提設定

### `.env` 認証情報

場所: `C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env`

必要キー:
```
HATENA_ID=hinyan1016
HATENA_BLOG_DOMAIN=hinyan1016.hatenablog.com
HATENA_API_KEY=<your-key>
OPENAI_API_KEY=<your-openai-key>    # 2026-04-23 追加: 画像自動生成用
```

### 依存

```bash
pip install -r requirements.txt
# openai>=1.40.0, pillow==11.0.0, pytest==8.3.4, pytest-cov==6.0.0
```

---

## Day 1（初回運用）チェックリスト

本番の1件目を処理する前に:

- [ ] `git status` でクリーン
- [ ] `pytest --cov=scripts` 全pass（44テスト）
- [ ] `.env` に `HATENA_API_KEY` と（API route使うなら）`OPENAI_API_KEY` が設定済
- [ ] `python -m scripts.generate_batch_csv --batch-size 1`（1件だけ選定＋存在確認）
- [ ] Claudeが `<today>_prompts.csv` のprompt/alt/style_summary を記入
- [ ] `python -m scripts.sync_generations_from_csv --date <today>` — generations へUPSERT
- [ ] 画像生成: A) `python -m scripts.export_for_codex --date <today>` → Codexで画像作成 → `images/in/<eid>.png` に配置、または B) `python -m scripts.generate_images --date <today>`（API route）
- [ ] `python -m scripts.build_dashboard --date <today>` — ダッシュボード表示確認
- [ ] 1件にOK判定 → CSV エクスポート → `data/decisions/<today>_decisions.csv` に保存
- [ ] `python -m scripts.process_decisions <path> --dry-run` — 変更内容確認
- [ ] `python -m scripts.process_decisions <path>` — yesで実行
- [ ] ブラウザで記事を開き、インフォグラフィック表示を確認
- [ ] 記事本文の非破壊（目次・本文・参考文献の順序保持）を確認

初回成功後、残り19件も同じフローで処理 → Day 2から20件/日ペースで運用開始。

---

## トラブル対応

### 画像アップロードが失敗する

- `logs/errors.log` 確認
- Fotolife API の一時障害の可能性、30分後に再試行
- 同じ decisions.csv を**再実行可能**（idempotency保証、Fotolifeへ重複アップされない）

### 誤って更新してしまった

```bash
python -m scripts.rollback <entry_id>
```

※ `body_html_before` が保存されている7日以内のみ有効

### NGが3回繰り返された記事

`status='manual_intervention'` に自動変更されるので、手動でChatGPT生成＋既存の `blog/_scripts/upload_and_insert_infographic.py` で対応。

### バッチCSVでpromptが空のまま `generate_images` を実行してしまった

`generate_images` は prompt が空の行を `skipped_empty_prompt` としてスキップする。Claudeに「今日のCSVにプロンプトを埋めて」と依頼してから再実行する。`generate_batch_csv.py` は再実行しない（既に in_batch ステータスになっているため）。

### OpenAI APIで生成が失敗する

- `logs/image_gen.log` を確認
- レート制限 / 内容ポリシー違反が主因
- ポリシー違反の場合はプロンプトを修正して再実行（既存画像はスキップされる）
- レート制限の場合は `--max-retries 5 --retry-delay-s 10` などで粘り強く再試行
- `OPENAI_API_KEY` が `.env` にない場合はエラー終了

### ブラウザダッシュボードで画像が表示されない

- `images/in/<entry_id>.png` のファイル名が `entry_id` と一致しているか確認
- 大文字/小文字の違い、拡張子の違い（.PNG, .jpg等）を確認
- build_dashboard の相対パス生成は `../../images/in/<eid>.png` が出ることが期待値（2026-04-23 修正）。古いHTMLが残っている場合は `python -m scripts.build_dashboard --date <today>` で再生成

### process_decisions が `generations に <eid> がない` で失敗

- ③ の `sync_generations_from_csv --date <today>` をスキップしている
- 同コマンドを走らせてから再実行すればOK（idempotent）

---

## 定期メンテナンス

### 週1回

- `logs/errors.log` を grep して ERROR/WARN の傾向確認
- NG率が高いパターン（特定カテゴリなど）がないかチェック
- 再生成を繰り返している記事を抽出し、プロンプト設計を見直す:

```bash
python -c "
import sqlite3
conn = sqlite3.connect('data/state.sqlite')
for eid, cnt in conn.execute('''
    SELECT entry_id, COUNT(*) FROM generations
    WHERE decision='ng' GROUP BY entry_id HAVING COUNT(*) >= 2
    ORDER BY COUNT(*) DESC
'''):
    print(f'{eid}: {cnt}回NG')
"
```

### プロジェクト完了時

- `SELECT COUNT(*) WHERE status='done'` = 227 を確認
- 全エラーログを棚卸し
- 最終レポート `COMPLETION_REPORT.md` を作成してコミット
- `data/state.sqlite` を `data/state_final.sqlite` にバックアップ

---

## ディレクトリ構造

```
blog/infographic-project/
├── RUNBOOK.md                       ← このファイル
├── requirements.txt
├── pytest.ini
├── .gitignore                       ← キャッシュ/画像/DBを除外
├── scripts/
│   ├── init_db.py                  # SQLite初期化
│   ├── inventory.py                # 判定ロジック (detect_v2, should_exclude)
│   ├── html_inserter.py            # HTML挿入（3段階フォールバック）
│   ├── hatena_client.py            # AtomPub/Fotolife クライアント
│   ├── generate_batch_csv.py       # 日次バッチ生成
│   ├── sync_generations_from_csv.py # prompts.csv → generations UPSERT
│   ├── generate_images.py          # OpenAI API 画像生成（保険ルート）
│   ├── export_for_codex.py         # Codex向け指示書生成（本命ルート）
│   ├── build_dashboard.py          # レビューダッシュボード生成
│   ├── process_decisions.py        # OK/NG/Skip反映
│   ├── rollback.py                 # 緊急ロールバック
│   └── inventory_scan.py           # 初回AtomPub取得（実行済・保持）
│
├── data/
│   ├── state.sqlite                # 進捗DB（唯一の真実、.gitignore対象）
│   ├── entries/                    # AtomPubキャッシュ（.gitignore対象）
│   ├── targets_v2.csv              # 対象227件の確定リスト
│   ├── inventory_v2.csv            # 全1003件の判定結果
│   ├── inventory_summary_v2.txt    # 集計サマリ
│   ├── batches/                    # 日次プロンプトCSV
│   ├── entries_for_claude/         # Claudeが読む記事情報MD
│   ├── reviews/                    # レビューダッシュボードHTML
│   └── decisions/                  # ユーザーのレビュー結果CSV
│
├── images/
│   ├── in/                         # ChatGPT生成物の受け入れ先
│   ├── approved/                   # （未使用、将来拡張用）
│   └── archive/                    # 処理済み画像
│
├── logs/                           # 実行ログ（.gitignore対象）
│
└── tests/
    ├── unit/                       # detect, exclude, insert, hatena_client, init_db, dashboard
    ├── integration/                # batch, process_decisions, rollback
    └── fixtures/entries/           # 実データから抽出した判定用サンプル
```

---

## よくある質問（FAQ）

**Q: state.sqlite を他のPCに持っていきたい**
A: `blog/infographic-project/data/state.sqlite` を手動コピー。Gitには含めない運用。

**Q: NG理由タグの意味を変えたい**
A: `scripts/build_dashboard.py` の `ng_tags_html` と、`scripts/process_decisions.py` の`_handle_ng` で整合性を取って変更。

**Q: 1日のバッチサイズを変えたい**
A: `scripts/generate_batch_csv.py` の main() で `batch_size=20` を変更、または `generate_daily_batch(..., batch_size=N)` を直接呼ぶ。

**Q: 処理を一時中断したい**
A: SQLiteと `images/in/` のファイル状態で中断可能。任意の日に再開可。
