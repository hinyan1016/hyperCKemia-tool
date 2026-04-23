# Session Handoff — 2026-04-23（Day 5 完了時点）

インフォグラフィック・バックフィルプロジェクト、Day 5 まで完走。
Day 4 完了後の別PC環境で Day 5 の20件を全フロー通過。対象ベースで32.0%に到達。

---

## 累計進捗（Day 5 終了時点）

| 指標 | 値 |
|------|-----|
| articles: done | **72 / 227**（対象ベース 72 / 218 = **33.0%**） |
| articles: excluded | 9（deleted 1 + deletion_planned 1 + skipped_by_review 7） |
| articles: pending | 146 |
| articles: in_batch | 0 |
| generations | 81件 |

**残り 146 件、ペース 17〜20件/日で約 8〜9営業日**で完走見込み。

注: 対象ベース分母は `227 - (deleted_from_hatena 1 + deletion_planned 1 + skipped_by_review 7) = 218`。
skipped_by_review はレビューで再選定不要と判断した記事群。

---

## Day 5（2026-04-23、本セッション）の成果

- batch_date = `2026-04-27` で20件バッチ生成 → Codex で画像生成 → 一括反映
- **処理結果: OK=16 NG=0 Skip=4 失敗=0**
- サンプル公開URL 3件で HTTP 200 + Fotolife img 挿入確認
- Skip 4件を `status='excluded' / excluded_reason='skipped_by_review'` に昇格

### Day 5 Skip 詳細（今後選定不要）

| # | entry_id | タイトル |
|---|---|---|
| 6 | `6802418398342753435` | 腕神経引き抜き損傷の電気生理検査 |
| 12 | `6802418398343144851` | 頸部椎骨動脈解離（eVAD） |
| 13 | `6802418398343295524` | Bow Hunter症候群（Gemini版） |
| 17 | `6802418398343407214` | 頭蓋内椎骨動脈解離（ICVAD） |

### Day 5 フロー実績

1. `generate_batch_csv --batch-date 2026-04-27 --batch-size 20` → 20件選出、pending 166→146
2. `_fill_prompts_20260427.py` で20件全てに style_summary / prompt / alt を投入（BOM対応で `utf-8-sig` 使用）
3. `sync_generations_from_csv --date 2026-04-27` → INSERT=20 UPDATE=0
4. `export_for_codex --date 2026-04-27` → Codex指示書生成
5. Codex で20枚画像生成（1-1.6MB each）
6. `build_dashboard --date 2026-04-27` → ブラウザレビュー
7. decisions CSV → 19件（1件選択漏れ） → 手動補修して20件化 → dry-run OK → `--yes` 本番反映
8. Skip 4件を手動で `excluded / skipped_by_review` に昇格

---

## Day 6 以降の本運用フロー（変更なし）

```bash
cd blog/infographic-project

# ① バッチ生成
python -m scripts.generate_batch_csv --batch-date <YYYY-MM-DD> --batch-size 20

# ② Claude がプロンプトを記入（_fill_prompts_<date>.py スクリプト、utf-8-sig推奨）

# ③ generations 同期（必須）
python -m scripts.sync_generations_from_csv --date <YYYY-MM-DD>

# ④ Codex 指示書生成
python -m scripts.export_for_codex --date <YYYY-MM-DD>
# → Codex で画像生成 → images/in/<entry_id>.png に保存

# ⑤ ダッシュボード生成 → ブラウザでレビュー
python -m scripts.build_dashboard --date <YYYY-MM-DD>
start data/reviews/<date>_review.html

# ⑥ dry-run → 本番反映
python -m scripts.process_decisions data/decisions/<date>_decisions.csv --dry-run
python -m scripts.process_decisions data/decisions/<date>_decisions.csv --yes

# ⑦ Skip 昇格（再選定不要な場合）
#   status='excluded' / excluded=1 / excluded_reason='skipped_by_review' に UPDATE

# ⑧ 検証
curl -s -o /dev/null -w "%{http_code}\n" https://hinyan1016.hatenablog.com/entry/YYYY/MM/DD/HHMMSS
```

### Day 5 で学んだこと

- **CSVエクスポート漏れの再発**: ダッシュボードの判定が1件不足した際、ユーザ確認で「手動追記」パスを採用すると最速。Day 5 では `6802418398342199347` (#1 医療AI/DX) を手動で ok として追記。
- **BOM付きUTF-8の罠**: `generate_batch_csv` の出力CSVは `utf-8-sig` (BOM付き)。Day 5 の `_fill_prompts` スクリプトで最初にBOMを読めず `KeyError: 'entry_id'` が出た。今後の fill_prompts では `encoding='utf-8-sig'` をデフォルトに。

---

## 現在のDB状態

```sql
articles:
  done=72, excluded=9, pending=146
  (excluded内訳: deleted_from_hatena 1 / deletion_planned 1 / skipped_by_review 7)
generations: 81件（全 in_batch 処理済み）
batches:
  2026-04-23 / 2026-04-24 / 2026-04-25 / 2026-04-26 / 2026-04-27 いずれも処理完了
```

---

## 次セッション再開フレーズ

- **「Day 6 バッチ準備」** → batch_date=2026-04-28 で 20件バッチ生成以降のフロー
- **「インフォグラフィックの続き」** → このファイルを読んで現状確認
- **「ブログ文献リンクの続き」** → 別プロジェクト（文献リンク有効化、残561件）

---

**セッション終了時刻**: 2026-04-23 22:30頃（Day 5 完了、72/227件 done、次は Day 6）
