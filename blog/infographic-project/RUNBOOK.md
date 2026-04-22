# インフォグラフィック・バックフィル運用手順

**対象**: インフォグラフィック未対応記事227件へのChatGPT生成画像追加
**ペース**: 20件/日 × 約12日で完走予定
**関連設計書**: [`docs/superpowers/specs/2026-04-23-infographic-backfill-design.md`](../../docs/superpowers/specs/2026-04-23-infographic-backfill-design.md)

---

## 日次ルーチン（15〜20分）

### ① バッチ生成（2分）

Claude Codeセッションで:

> 「今日の20件のプロンプト生成して」

Claudeが以下を実行:

```bash
cd blog/infographic-project
python -m scripts.generate_batch_csv
```

これにより生成:
- `data/batches/<today>_prompts.csv` — プロンプト記入用CSV（空欄テンプレ）
- `data/entries_for_claude/<today>.md` — Claudeが読む記事情報

次にClaudeが `<today>.md` を読み、各記事について:
- **style_summary**（例: 「神経疾患・青系・解剖図レイアウト」）
- **prompt**（ChatGPT GPT Image用、日本語、約300-500字）
- **alt**（120字前後、主キーワード含む）

を決定し、`<today>_prompts.csv` の空欄を埋める。

### ② ChatGPT画像生成（5〜10分）

1. `data/batches/<today>_prompts.csv` をExcelで開く
2. 1行ずつ `prompt` 列をコピー → ChatGPT UI（画像生成）に貼付
3. 生成画像を右クリック → 保存 → `blog/infographic-project/images/in/<entry_id>.png` にリネーム
4. 20件全部保存できたらファイルマネージャで件数確認

### ③ ダッシュボードレビュー（5〜7分）

> 「ダッシュボード作って」

Claudeが実行:

```bash
python -m scripts.build_dashboard
```

1. 生成された `data/reviews/<today>_review.html` をブラウザで開く
2. 各カードで OK/NG/スキップを選択
3. NGの場合は理由タグ（medical_error/typo/design/mismatch/other）と自由記述を入力
4. 下部の「決定をCSVにエクスポート」ボタン押下
5. ダウンロードされた CSV を `data/decisions/<today>_decisions.csv` として保存

### ④ 反映（3分）

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

### ⑤ 進捗確認

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

## Day 1（初回運用）チェックリスト

本番の1件目を処理する前に:

- [ ] `git status` でクリーン
- [ ] `pytest --cov=scripts` 全pass（32テスト）
- [ ] `python -m scripts.generate_batch_csv`（ただし `batch_size=1` でテスト）
- [ ] ChatGPTで1件だけ生成、`images/in/<entry_id>.png` に配置
- [ ] `python -m scripts.build_dashboard` — ダッシュボード表示確認
- [ ] 1件にOK判定 → CSV エクスポート → `data/decisions/` に保存
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

### バッチCSVでpromptが空のまま ChatGPT手順に移ってしまった

Claudeに「今日のCSVにプロンプトを埋めて」と依頼。`generate_batch_csv.py` は再実行しない（既に in_batch ステータスになっているため）。

### ブラウザダッシュボードで画像が表示されない

- `images/in/<entry_id>.png` のファイル名が `entry_id` と一致しているか確認
- 大文字/小文字の違い、拡張子の違い（.PNG, .jpg等）を確認

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
