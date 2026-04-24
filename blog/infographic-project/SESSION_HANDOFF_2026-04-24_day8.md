# Session Handoff — 2026-04-24（Day 8 完了時点）

インフォグラフィック・バックフィルプロジェクト、Day 8 完走。
Day 8 は**初の NG 再生成ラウンド付きバッチ**。5件を再生成で救済し、結果 OK=20/20。
対象ベースで **58.9%**（折り返し後）。

---

## 累計進捗（Day 8 終了時点）

| 指標 | 値 |
|------|-----|
| articles: done | **123 / 227**（対象ベース 123 / 209 = **58.9%**） |
| articles: excluded | 18（deleted 1 + deletion_planned 1 + skipped_by_review 16） |
| articles: pending | 86 |
| articles: in_batch | 0 |
| generations | 141件 |

**残り 86 件、ペース 20件/日で約 4〜5営業日**で完走見込み。

注: 対象ベース分母は `227 - (deleted 1 + deletion_planned 1 + skipped_by_review 16) = 209`。

---

## Day 8（2026-04-24 夕方、本セッション）の成果

- batch_date = `2026-04-30` で20件バッチ生成 → Codex で画像生成 → 初回レビューで **NG=5**
- NG 5件に対し**修正指示書 `2026-04-30_ng_retry_instructions.md`** を生成して Codex 再送
- 再生成後の5件をユーザーが1件ずつ視覚確認 → 全件 OK に更新
- **最終結果: OK=20 NG=0 Skip=0 失敗=0**
- サンプル公開URL 3件で HTTP 200 + `hatena-fotolife` img 挿入確認

### Day 8 初回 NG → 再生成で救済した5件

| # | entry_id | タイトル | 初回NGタグ | 再生成の要点 |
|---|---|---|---|---|
| 10 | `6802418398477263320` | 尾状核梗塞 | medical_error | 「内頭動脈」誤記 → 「内頚動脈」厳守 |
| 13 | `6802418398479342287` | CGRP 3剤比較 | design | ピンク単色 → ピンクはアクセント限定＋本文は濃紺/黒でWCAG AAコントラスト |
| 17 | `6802418398480130015` | 高齢者NCSE/IIC | design | 赤単色 → 赤はアクセント限定＋本文は濃紺/黒で白基調 |
| 19 | `6802418398481548886` | 前脈絡叢動脈梗塞 | medical_error | 視野図で両耳側半盲を描画 → 同名半盲（両眼同側欠損）を厳守 |
| 20 | `6802418398482404611` | Gerstmann症候群 | medical_error | 角回位置不正確 → BA39・上側頭溝後端を囲む弓状回を厳守 |

### Day 8 OK採用（20件、カテゴリ別）

**20年の変遷シリーズ（5件）— 全件採用**
| # | entry_id | タイトル |
|---|---|---|
| 2 | `6802418398471356830` | 脳神経内科領域20年総論 |
| 3 | `6802418398472288800` | 薬剤抵抗性てんかん20年 |
| 4 | `6802418398472332707` | AD型認知症20年 |
| 5 | `6802418398472747666` | MS 20年 |
| 16 | `6802418398480060052` | PD 20年 |

※ Day 6 で「AD治療20年」「MS治療20年」等がSkipパターンだったが、Day 8 では9:16仕様・俯瞰構成で4段レイアウトに収まり全件採用可と判断

**片頭痛シリーズ（4件）**
| # | entry_id | タイトル |
|---|---|---|
| 12 | `6802418398478906157` | 片頭痛治療薬使い分け |
| 13 | `6802418398479342287` | CGRP 3剤比較（再生成で救済） |
| 14 | `6802418398479360267` | CGRP以外予防薬 |
| 15 | `6802418398479987938` | 急性期治療薬まとめ |

**単一疾患・症候群（11件）**
| # | entry_id | タイトル |
|---|---|---|
| 1 | `6802418398470532882` | 封入体筋炎(IBM) |
| 6 | `6802418398472842011` | 妊娠とてんかん（抗ASM+葉酸） |
| 7 | `6802418398473403361` | イソプロピルアセチル尿素の危険性 |
| 8 | `6802418398474663886` | 神経保険適応遺伝子検査 |
| 9 | `6802418398477234476` | ハンチントン病 |
| 10 | `6802418398477263320` | 尾状核梗塞（再生成で救済） |
| 11 | `6802418398477844602` | CJD |
| 17 | `6802418398480130015` | 高齢者NCSE/IIC（再生成で救済） |
| 18 | `6802418398480451565` | p-tau217 |
| 19 | `6802418398481548886` | 前脈絡叢動脈梗塞（再生成で救済） |
| 20 | `6802418398482404611` | Gerstmann症候群（再生成で救済） |

---

## Day 8 での重要学習: NG再生成フローの実用化

### NG再生成プロンプトの作り方（次回以降のテンプレ）

1. NG判定された entry の元プロンプトを CSV から取得
2. 各 NG について以下の構造で再生成指示書を作成:
   - **前回NGタグ**: `medical_error` / `design` / `layout`
   - **NG理由**: 具体的な誤り内容（ユーザーの `ng_note` を引用）
   - **修正ポイント**: 正しい解剖・配色・レイアウトの具体指示
   - **改訂プロンプト**: 元プロンプト本体 + `【厳守】` 節で NG回避指示を末尾に追加

### 典型的な NG パターンと対策

| NGタグ | 代表例 | 対策 |
|-------|--------|------|
| `medical_error` (文字誤記) | 内頚動脈→内頭動脈 | `【厳守】`節で正しい表記を強調・漢字指定 |
| `medical_error` (解剖図誤り) | 同名半盲が両耳側半盲、角回の位置違い | 病態の詳細定義+視野/解剖図の厳守ルール明記 |
| `design` (視認性) | 赤/ピンク単色で本文読めない | アクセント色と本文テキスト色を分離・WCAG AA 4.5:1 指定 |

### 本日の教訓

- **単色指定（ピンク系・赤系）は視認性 NG を招きやすい**
  - 今後は `style_summary` で「本文は濃紺/黒、アクセントは◯◯」と明示する方針
- **解剖図の模式化でCodexが苦手な3領域**
  1. 視野欠損パターン（同名半盲 vs 両耳側半盲 vs 両鼻側半盲）
  2. 脳回の正確な位置（角回・縁上回・下前頭回など）
  3. 血管名の漢字（内頚動脈・前脈絡叢動脈など）
  - 上記3点は `【厳守】`節でプロンプトに明示すると再生成成功率が高い

---

## Day 9 以降の本運用フロー（NG再生成対応版）

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

# ⑥ decisions CSV を手動作成（data/decisions/<date>_decisions.csv）
#    形式: entry_id, decision(ok/ng/skip), ng_tags, ng_note

# ⑥' ★NG再生成フロー（NG判定があった場合）
#    1. data/batches/<date>_ng_retry_instructions.md を Claude に作成依頼
#    2. Codex で5件等を再生成 → images/in/<entry_id>.png 上書き
#    3. 再レビュー後 decisions CSV を OK に更新

# ⑦ dry-run → 本番反映
python -m scripts.process_decisions data/decisions/<date>_decisions.csv --dry-run
python -m scripts.process_decisions data/decisions/<date>_decisions.csv --yes

# ⑧ Skip 昇格（必要な場合のみ）
#   status='excluded' / excluded=1 / excluded_reason='skipped_by_review' に UPDATE

# ⑨ 検証
curl -s -o /dev/null -w "%{http_code}\n" https://hinyan1016.hatenablog.com/entry/YYYY/MM/DD/HHMMSS
```

---

## 現在のDB状態

```sql
articles:
  done=123, excluded=18, pending=86
  (excluded内訳: deleted_from_hatena 1 / deletion_planned 1 / skipped_by_review 16)
generations: 141件（全 in_batch 処理済み）
batches:
  2026-04-23 / 2026-04-24 / 2026-04-25 / 2026-04-26 / 2026-04-27 / 2026-04-28 / 2026-04-29 / 2026-04-30
  いずれも処理完了
```

---

## 次セッション再開フレーズ

- **「Day 9 バッチ準備」** → batch_date=2026-05-01 で 20件バッチ生成以降のフロー
- **「インフォグラフィックの続き」** → このファイルを読んで現状確認

残 86 件。4〜5営業日で完走圏内。

---

**セッション終了時刻**: 2026-04-24 夕方（Day 8 完了、123/227件 done、次は Day 9）
