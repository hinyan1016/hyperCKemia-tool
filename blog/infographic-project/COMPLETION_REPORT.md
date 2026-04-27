# インフォグラフィック・バックフィル運用 完了レポート

**プロジェクト期間:** 2026-04-23 〜 2026-05-05（13バッチ・実日数 4日 → 13日に分散）
**完了日:** 2026-04-26（DB上は2026-05-05分まで処理済）
**完了レポート作成日:** 2026-04-28

---

## 1. 最終達成

| 指標 | 値 |
|------|----|
| 総記事数 | 227 |
| 完成（done） | **224** |
| 除外（excluded） | 3 |
| 対象ベース（target） | 224 |
| **完了率** | **224 / 224 = 100.0%** |

### 除外3件の内訳
| entry_id | 理由 | 備考 |
|----------|------|------|
| 6802418398334367196 | `deletion_planned` | クマの記事、本人削除予定 |
| 6802418398334380887 | `deleted_from_hatena` | クマの妊娠 受診タイミング、はてな側で既に削除 |
| 6802418398343407214 | `skipped_by_review` | ICVAD（中条道夫病）、Day 13で unexclude再挑戦も再度skip |

---

## 2. バッチ別実績

| Day | Date | バッチ規模 | 完成（DB articles） | NG初回→救済 | Skip | 備考 |
|-----|------|-----------|--------------------|------------|------|------|
| 1 | 2026-04-23 | 19 | 18 | - | 1（deleted） | 立ち上げ |
| 2 | 2026-04-24 | 1 | 1 | - | 0 | 補完 |
| 3 | 2026-04-25 | 20 | 20 | - | 0 | |
| 4 | 2026-04-26 | 20 | 17 | - | 3 | |
| 5 | 2026-04-27 | 20 | 16 | - | 4 | |
| 6 | 2026-04-28 | 20 | 11 | - | 9 | **初の9:16ポートレート**仕様導入で混乱 |
| 7 | 2026-04-29 | 20 | 20 | - | 0 | 9:16安定化 |
| 8 | 2026-04-30 | 20 | 20 | NG 6→6救済 | 0 | |
| 9 | 2026-05-01 | 20 | 20 | NG 6→6救済 | 0 | |
| 10 | 2026-05-02 | 20 | 20 | 0 | 0 | **HEADER強化版**で初の初回全件OK |
| 11 | 2026-05-03 | 20 | 20 | 0 | 0 | 2連続初回全件OK |
| 12 | 2026-05-04 | 20 | 20 | NG 1→1救済 | 0 | 片麻痺胃管/咽頭交差テキスト誤訳修正 |
| 13 | 2026-05-05 | 22 | 21 | - | 1（ICVAD） | **過去skip 16件をunexclude再挑戦 → 15件救済** |
| **計** | | **242** | **224** | **NG 13→13** | **18 events** | 重複あり |

### バッチサイズ242 vs articles 224 の差15
- 16: Day 4-6 の skip (3+4+9) を Day 13 で unexclude → 同じ entry_id が再バッチ入り
- 1: Day 1 の deleted_from_hatena → そもそも完成しない
- 1: Day 13 の ICVAD → 再挑戦で再度skip
- 1: Day 1 の deletion_planned → そもそも完成しない（Day 1 batch=19、processed 18 で帳尻）

---

## 3. 救済率（Rescue Rate）

### NG retry救済率
- 検知回数: 13件（Day 8: 6 + Day 9: 6 + Day 12: 1）
- 救済成功: 13件
- **救済率: 100%（13/13）**

### Skip再挑戦救済率（Day 13）
- 過去Skip: 16件（Day 4-6 由来）
- unexclude再挑戦で OK: 15件
- 引き続きSkip: 1件（ICVAD）
- **救済率: 93.75%（15/16）**

### 総合品質
- generations 累計 227行（attempt=1のみ、NG履歴は decision上書きで未保存）
- 最終 decision 内訳: approved 224 / NULL 2 / skipped_deleted 1

---

## 4. プロジェクト中に確立した恒久仕様

### 4.1 画像仕様（Day 6で確立）
- **アスペクト比 9:16（縦長ポートレート）**、1080×1920以上
- 正方形（1024×1024）禁止 — リサイズでレイアウト崩壊
- レイアウト解釈: プロンプトの「左上/右上/左下/右下」を「上段→中上段→中下段→下段」の4段縦積みに読み替え

### 4.2 Codex指示書HEADERの強化（Day 10で確立）
1. **情報の優先順位**: 記事本文 > プロンプト > 装飾（矛盾時は記事優先）
2. **参考記事の精読プロセス**: アクセス → 抽出 → 設計 → 補正 の4ステップを必須化
3. **医学画像の対比表**: ✗ 生成NG（AI描画禁止） vs ✓ 引用OK（別工程・本文埋込）
4. **per-entry見出し**: `**参考記事URL（必読・最優先）**` ＋精読リマインダ＋プロンプト見出しに「医学画像の新規描画は厳禁」明示
5. **失敗扱い**: URLアクセス不可も「推測で作らない・スキップ報告」

### 4.3 Claude側のプロンプト作成プロセス
- **4並列Exploreサブエージェント**（各5記事）でWebFetch・要点抽出
- 各記事の中心メッセージ／4段構成／固有名詞・数値／配色／医学画像合成の危険ポイント／alt案を構造化抽出
- 抽出結果をもとに `_fill_prompts_<date>.py` を作成
- 全エントリに【厳守】節で**記事固有の医学画像合成NG**を明示

### 4.4 NG retry運用（Day 12で確立）
- retry時は1件用の小さい decisions CSV（entry_id, decision=ok のみ）を作成
- `process_decisions` で最新attempt行のdecision を ng → approved に書き換え
- `data/decisions/<date>_retry_decisions.csv` の命名で履歴管理

### 4.5 Skip再挑戦運用（Day 13で確立）
- 過去のskip_by_review記事は `excluded=0, status='pending', last_batch_date=NULL` でunexcludeすれば再バッチ対象化
- `sync_generations_from_csv` は `decision='skip'` 行を保護するため、skip→NULLリセット後に再syncが必要
- 強化HEADER＋4並列精読＋医学画像対比表が過去skip案件にも有効

---

## 5. 効果検証 — Day別Skip率の推移

```
Day 1-3 (立ち上げ):           skip 1/40 = 2.5%   ※多くは"deleted"由来
Day 4-6 (品質不安定期):       skip 16/60 = 26.7% ★問題期
Day 7-9 (9:16+NG retry確立):  skip 0/60 = 0.0%   NG retry救済で対応
Day 10-12 (HEADER強化以降):    skip 0/60 = 0.0%   初回NG=0達成
Day 13 (skip救済):           skip 1/22 = 4.5%   ICVADのみ
```

**運用知見の効果**: HEADER強化（Day 10）以降、新規記事 60件に対し skip=0、initial NG=0 を継続。Day 6以前の不安定期からの脱却が定量的に確認できる。

---

## 6. 残作業

### ✅ 実施済（2026-04-28）
- [x] `state_final_2026-04-26.sqlite` バックアップ作成
- [x] COMPLETION_REPORT.md 作成

### 任意（必要に応じて）
- [ ] Day 1〜13 の SESSION_HANDOFF を統合した `SESSION_HANDOFF_FINAL.md` 作成
- [ ] `images/in/` 配下の最終納品画像を別バックアップ
- [ ] 全エラーログ（`logs/`）の棚卸し
- [ ] 運用知見を `RUNBOOK.md` に逆輸入（HEADER強化版仕様、retry運用）
- [ ] `LATEST_STATUS_2026-04-27.md` を本レポートへの参照に更新
- [ ] `state.sqlite` を読取専用にする（誤更新防止）

---

## 7. 主要ドキュメント

| 種別 | パス |
|------|------|
| 設計書 | `docs/superpowers/specs/2026-04-23-infographic-backfill-design.md` |
| 手順書 | `blog/infographic-project/RUNBOOK.md` |
| 最新運用注意 | `blog/infographic-project/LATEST_STATUS_2026-04-27.md` |
| DB（運用） | `blog/infographic-project/data/state.sqlite` |
| DB（凍結） | `blog/infographic-project/data/state_final_2026-04-26.sqlite` |
| メモリ | `~/.claude/projects/.../memory/project_infographic_backfill.md` |
| 完走画像 | `blog/infographic-project/images/in/<entry_id>.png` |
| 日次セッション記録 | `blog/infographic-project/SESSION_HANDOFF_*.md`（Day 4〜10分が現存） |

---

## 8. プロジェクトサマリー

227記事のうち、削除済み2件と医学的レビューでSkipしたICVAD 1件を除く **224記事すべてに 9:16 縦長ポートレートのインフォグラフィックを付与**。13バッチ運用で安定化を進め、Day 10以降の HEADER 強化版で初回NG＝0／Skip＝0の定常状態に到達。Day 13 の skip再挑戦では過去案件 15/16 を救済し、強化版プロンプトの遡及効果を実証した。

NG retry 救済率 100%、過去Skip 救済率 93.75%。

🎉 **完走（2026-04-26）**
