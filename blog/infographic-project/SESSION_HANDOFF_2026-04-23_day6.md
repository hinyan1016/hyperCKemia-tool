# Session Handoff — 2026-04-23（Day 6 完了時点）

インフォグラフィック・バックフィルプロジェクト、Day 6 まで完走。
Day 6 は**初の9:16縦長ポートレート対応**バッチ。対象ベースで39.7%に到達。

---

## 累計進捗（Day 6 終了時点）

| 指標 | 値 |
|------|-----|
| articles: done | **83 / 227**（対象ベース 83 / 209 = **39.7%**） |
| articles: excluded | 18（deleted 1 + deletion_planned 1 + skipped_by_review 16） |
| articles: pending | 126 |
| articles: in_batch | 0 |
| generations | 101件 |

**残り 126 件、ペース 17〜20件/日で約 7〜8営業日**で完走見込み。

注: 対象ベース分母は `227 - (deleted 1 + deletion_planned 1 + skipped_by_review 16) = 209`。

---

## Day 6（2026-04-23、本セッション）の成果

- batch_date = `2026-04-28` で20件バッチ生成 → Codex で画像生成 → 一括反映
- **処理結果: OK=11 NG=0 Skip=9 失敗=0**
- サンプル公開URL 3件で HTTP 200 + Fotolife img 挿入確認
- Skip 9件を `status='excluded' / excluded_reason='skipped_by_review'` に昇格
- **9:16ポートレート仕様を `export_for_codex.py` HEADER に恒久化**（Day 7以降自動適用）

### Day 6 Skip 詳細（今後選定不要）

| # | entry_id | タイトル |
|---|---|---|
| 4 | `6802418398402104253` | FSHD（顔面肩甲上腕型筋ジストロフィー） |
| 5 | `6802418398404086674` | CGRP片頭痛治療 2025 |
| 7 | `6802418398416132613` | ラクナ梗塞+心房細動 抗血栓 |
| 10 | `6802418398419039140` | 低Na血症 鑑別支援ツール |
| 11 | `6802418398431254895` | アルツハイマー治療20年 |
| 12 | `6802418398432035702` | 多発性硬化症治療20年 |
| 13 | `6802418398432125402` | NMOSD診療20年 |
| 16 | `6802418398435175649` | 臨床栄養療法 |
| 17 | `6802418398436017605` | 心房細動治療20年 |

### Day 6 OK採用（11件）

| # | entry_id | タイトル |
|---|---|---|
| 1 | `6802418398375457494` | 2025脳神経内科注目論文TOP4 |
| 2 | `6802418398375720899` | 2025糖尿病注目論文 |
| 3 | `6802418398380904949` | 処方日数と頓用制限 |
| 6 | `6802418398407929045` | 中高生スマホ依存 |
| 8 | `6802418398418048298` | 更年期障害（女性） |
| 9 | `6802418398418049747` | 男性更年期（LOH） |
| 14 | `6802418398432272973` | ニホンイシガメ食餌ガイド |
| 15 | `6802418398435001826` | 成人の悪夢 |
| 18 | `6802418398436019240` | 高齢者身体障害認定 |
| 19 | `6802418398438298962` | 胃管vs胃瘻 |
| 20 | `6802418398438453131` | トルーソー症候群 |

---

## Day 6 での重要変更: 9:16ポートレート仕様化

**背景**: ユーザー要望「スマホ対応の9:16動画用として生成してほしい。後から1024x1024にリサイズするとレイアウトが崩れる」

**対応**: `scripts/export_for_codex.py` のHEADERテンプレート更新（Day 7以降も自動適用）

```
## 画像仕様（重要）
- アスペクト比: 9:16（縦長ポートレート・スマホ/ショート動画対応）
- 推奨解像度: 1080x1920 以上
- 正方形（1024x1024）では生成しない
- レイアウト再解釈: 「左上/右上/左下/右下」は
  「上段→中上段→中下段→下段」の4段縦積みに再配置
- タイトル帯は上部または各段の間に配置
```

プロンプト本文（DATA辞書）は変更不要。指示書ヘッダーで統一ルールとして適用。

---

## Day 7 以降の本運用フロー（変更なし）

```bash
cd blog/infographic-project

# ① バッチ生成
python -m scripts.generate_batch_csv --batch-date <YYYY-MM-DD> --batch-size 20

# ② Claude がプロンプトを記入（_fill_prompts_<date>.py スクリプト、utf-8-sig）

# ③ generations 同期
python -m scripts.sync_generations_from_csv --date <YYYY-MM-DD>

# ④ Codex 指示書生成（9:16 ヘッダー自動付与）
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

---

## 現在のDB状態

```sql
articles:
  done=83, excluded=18, pending=126
  (excluded内訳: deleted_from_hatena 1 / deletion_planned 1 / skipped_by_review 16)
generations: 101件（全 in_batch 処理済み）
batches:
  2026-04-23 / 2026-04-24 / 2026-04-25 / 2026-04-26 / 2026-04-27 / 2026-04-28 いずれも処理完了
```

---

## 次セッション再開フレーズ

- **「Day 7 バッチ準備」** → batch_date=2026-04-29 で 20件バッチ生成以降のフロー
- **「インフォグラフィックの続き」** → このファイルを読んで現状確認

---

**セッション終了時刻**: 2026-04-24 早朝（Day 6 完了、83/227件 done、次は Day 7）
