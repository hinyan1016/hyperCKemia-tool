# インフォグラフィック バックフィル プロジェクト 設計書

**作成日**: 2026-04-23
**対象ブログ**: `https://hinyan1016.hatenablog.com/` (hinyan1016)
**対象記事数**: 227件（実データ棚卸し済み）
**想定期間**: 約12日（20件/日）

---

## 1. 目的と背景

はてなブログ hinyan1016 には約1000記事が蓄積されており、大半はインフォグラフィック（記事内容を1枚で要約するサマリー画像）付きで運用されている。しかし初期の記事（主に2025年3月〜2026年3月のうち一部）と、YouTubeコンパニオン記事の一部には**サマリー型インフォグラフィック画像が未挿入**の記事が残っている。

本プロジェクトは、これら未対応記事に対して**ChatGPTで生成した高品質インフォグラフィック画像**を追加し、以下を達成する:

1. ブログ全体の視覚的統一感の向上
2. SEO効果（画像検索・滞在時間・CTR）の底上げ
3. 読者の情報到達効率の向上（テキスト読破前に概要把握）

---

## 2. 前提条件と決定事項

### 2.1 ユーザー決定事項（ブレインストーミング結果）

| # | 項目 | 決定 |
|---|------|------|
| 1 | 対象定義 | サマリー型インフォグラフィックが冒頭にない記事すべて |
| 2 | 処理順 | 全件・公開日古い順（下限なし） |
| 3 | 分担 | セミ自動・**20件/日・約12日** |
| 4 | スタイル | 記事ごとにClaudeが最適案を提案（統一テンプレートは使わない） |
| 5 | 本文取得 | AtomPub API一括DL → ローカルキャッシュ |
| 6 | 品質管理 | ダッシュボード型レビュー（全件） |
| 7 | 判定 | マーカー厳格＋グレーゾーン手動（v2で緩和済み） |
| 8 | NG処理 | 理由タグ＋自動改訂プロンプト |
| 9 | 除外 | 日記/雑記カテゴリ・500字未満・インデックスページ |

### 2.2 確定した対象件数（実データ検証済み）

| 判定 | 件数 | 内訳 |
|------|------|------|
| インフォグラフィックあり（強マーカー） | 27 | コメント・alt・ファイル名 |
| インフォグラフィックあり（推定） | 734 | Fotolife figure 485 + Fotolife img 249 |
| **インフォグラフィック要（対象）** | **227** | 2025=205件、2026=22件 |
| 除外（日記/短文/インデックス） | 15 | index_page 14、category 1 |
| **合計** | **1003** | |

### 2.3 追加料金に関する制約

- ChatGPT Proプラン利用（追加料金なし）で画像生成は**ChatGPT UI手動**
- Claude Code 内のプロンプト生成・スクリプト作成は現契約内で完結
- はてなAtomPub API / Fotolife API は無料
- OpenAI API（gpt-image-1）等の従量課金サービスは**使用しない**

---

## 3. アーキテクチャ

### 3.1 3層構造

```
┌─────────────────────────────────────────────────────────────┐
│  データ層（Python・決定的処理）                               │
│  ・AtomPub取得/PUT  ・Fotolifeアップ                         │
│  ・SQLite状態DB     ・CSV読み書き                             │
└──────────────┬──────────────────────────┬───────────────────┘
               │                          │
               ↓                          ↑
┌─────────────────────────────────────────────────────────────┐
│  創造層（Claude Code セッション・対話実行）                   │
│  ・記事解析  ・スタイル提案  ・プロンプト生成                │
│  ・NG改訂プロンプト生成  ・ダッシュボード生成                │
└──────────────┬──────────────────────────┬───────────────────┘
               │                          │
               ↓                          ↑
┌─────────────────────────────────────────────────────────────┐
│  UI層（ユーザー操作）                                          │
│  ・ChatGPT UI（画像生成・手動）                               │
│  ・レビューダッシュボードHTML（ブラウザで操作）                │
│  ・file explorer（画像ファイルのドラッグ&ドロップ）            │
└─────────────────────────────────────────────────────────────┘
```

### 3.2 責務境界

| レイヤー | **する** | **しない** |
|---------|---------|-----------|
| Python | HTTP通信、DB更新、ファイル操作、HTML機械処理 | 医学的判断、プロンプト設計、ユーザー対話 |
| Claude | 記事内容の解釈、スタイル提案、プロンプト文面、改訂判断 | DB直接書込（スクリプト経由のみ）、API直接呼出 |
| ユーザー | ChatGPT操作、画像ダウンロード、ダッシュボードOK/NG判定 | プロンプト作文（Claudeのものをコピペ） |

### 3.3 ディレクトリ構成

```
blog/infographic-project/
├── scripts/                          # Python実装
│   ├── fetch_atompub.py             # 全エントリ取得（初回のみ・差分再取得可、inventory_scan.pyから分離）
│   ├── inventory.py                 # 判定＋除外→targets.csv（inventory_rescan_v2.pyのdetect_v2()を流用）
│   ├── init_db.py                   # state.sqliteのスキーマ初期化＋targets_v2.csvをarticlesに投入
│   ├── generate_batch_csv.py        # 次の20件を取り出すユーティリティ
│   ├── build_dashboard.py           # レビュー用HTML生成
│   ├── process_decisions.py         # OK→upload+insert、NG→改訂要求
│   ├── upload_fotolife.py           # Fotolifeアップ（既存流用）
│   ├── update_entry.py              # AtomPub PUT（既存流用）
│   └── rollback.py                  # 緊急ロールバック
│
├── data/
│   ├── state.sqlite                 # 進捗DB（唯一の真実）
│   ├── entries/                     # 取得済み記事XML（entry_id毎）
│   ├── targets_v2.csv               # 対象227件の確定リスト（生成済み）
│   ├── batches/                     # 日次CSV
│   │   └── 2026-04-23_prompts.csv
│   ├── reviews/                     # レビューダッシュボード
│   │   └── 2026-04-23_review.html
│   ├── decisions/                   # ユーザーのレビュー結果CSV
│   │   └── 2026-04-23_decisions.csv
│   └── entries_for_claude/          # Claudeが読む本文束
│       └── 2026-04-23.md
│
├── images/
│   ├── in/                          # ユーザーがChatGPT生成物を置く
│   │   └── <entry_id>.png
│   ├── approved/                    # OK判定後（アップロード前）
│   └── archive/                     # 使用済み・Fotolifeアップ完了
│
├── logs/
│   ├── fetch_YYYY-MM-DD.log
│   ├── batch_YYYY-MM-DD.log
│   ├── dashboard_YYYY-MM-DD.log
│   ├── process_YYYY-MM-DD.log
│   └── errors.log                   # すべてのERRORを集約
│
└── tests/
    ├── unit/
    ├── integration/
    └── fixtures/
```

### 3.4 SQLiteスキーマ（state.sqlite）

```sql
-- 記事台帳
CREATE TABLE articles (
  entry_id TEXT PRIMARY KEY,
  title TEXT,
  category TEXT,
  published_date TEXT,
  body_html_path TEXT,          -- data/entries/<id>.xml へのパス
  body_html_before TEXT,        -- 更新前の本文（7日保持、ロールバック用）
  word_count INTEGER,
  has_infographic INTEGER,      -- 0=なし, 1=あり(強), 2=あり(推定)
  excluded INTEGER,             -- 0=対象, 1=除外
  excluded_reason TEXT,         -- "short" | "category" | "index_page" など
  status TEXT,                  -- "pending" | "in_batch" | "image_received"
                                --  | "approved" | "uploaded" | "done"
                                --  | "ng_pending" | "upload_failed"
                                --  | "update_failed" | "conflict"
                                --  | "manual_intervention"
  last_batch_date TEXT,
  last_ng_date TEXT,
  edited_at TEXT                -- 取得時のapp:edited（競合検出用）
);

-- 生成履歴（1記事で複数回試行可）
CREATE TABLE generations (
  gen_id INTEGER PRIMARY KEY AUTOINCREMENT,
  entry_id TEXT,
  attempt INTEGER,              -- 1, 2, 3...
  prompt TEXT,
  alt_text TEXT,
  style_summary TEXT,
  image_path TEXT,
  decision TEXT,                -- "approved" | "ng" | NULL
  ng_tags TEXT,                 -- カンマ区切り
  ng_note TEXT,
  fotolife_url TEXT,
  created_at TEXT,
  processed_at TEXT,
  FOREIGN KEY (entry_id) REFERENCES articles(entry_id)
);

-- バッチ管理
CREATE TABLE batches (
  batch_date TEXT PRIMARY KEY,
  entry_ids TEXT,               -- カンマ区切り
  prompts_csv TEXT,
  dashboard_html TEXT,
  decisions_csv TEXT,
  status TEXT                   -- "preparing" | "generating" | "reviewing" | "closed"
);

-- 処理済み decisions の追跡（idempotency保証）
CREATE TABLE processed_decisions (
  decisions_csv_path TEXT,
  entry_id TEXT,
  decision TEXT,
  processed_at TEXT,
  PRIMARY KEY (decisions_csv_path, entry_id)
);
```

---

## 4. コンポーネント詳細

### 4.1 `fetch_atompub.py` — 全記事取得

**入力**: `HATENA_API_KEY`（.envから）
**出力**: `data/entries/<entry_id>.xml`（N件）、`articles`テーブル初期化

- AtomPub `collection` エンドポイント `?page=N` で全走査
- リクエスト間0.5秒スリープ、429発生時は間隔を2秒→5秒に自動延長（最大3リトライ）
- 再実行時 `--since-days N` オプションで差分更新可能
- **本プロジェクトでは2026-04-23に初回実行済み、1003件取得完了**

### 4.2 `inventory.py` — 判定＋除外（v2）

**入力**: `articles` テーブル（本文HTML含む）
**出力**: `articles` テーブルの `has_infographic`, `excluded` 更新 + `data/targets_v2.csv`

**判定ロジック（v2）**:

```python
def detect_v2(body_html: str) -> tuple[str, str]:
    # iframeとscriptを除去してから判定
    cleaned = re.sub(r"<iframe[^>]*>.*?</iframe>", "", body_html, flags=re.DOTALL | re.I)
    cleaned = re.sub(r"<script[^>]*>.*?</script>", "", cleaned, flags=re.DOTALL | re.I)

    # Strong signal: 明示マーカー
    if "<!-- インフォグラフィック -->" in cleaned: return "has_ig_strong", "comment_marker"
    if re.search(r'<img[^>]+alt="[^"]*インフォグラフィック', cleaned): return "has_ig_strong", "alt_marker"
    if re.search(r'<img[^>]+alt="[^"]*(まとめ|要約)', cleaned): return "has_ig_strong", "alt_marker_summary"
    if re.search(r'<img[^>]+src="[^"]*infograph', cleaned, re.I): return "has_ig_strong", "filename_infographic"

    # Likely signal（本文冒頭4000文字）
    head = cleaned[:4000]
    if re.search(r'<figure[^>]*figure-image-fotolife', head): return "has_ig_likely", "fotolife_figure_head"
    if re.search(r'<img[^>]+src="https?://cdn-ak\.f\.st-hatena\.com/images/fotolife', head):
        return "has_ig_likely", "fotolife_img_head"
    for m in re.finditer(r'<img[^>]+width="(\d+)"', head):
        if int(m.group(1)) >= 600: return "has_ig_likely", f"big_img_head_w{m.group(1)}"
    if re.search(r"<figure[^>]*>.*?<img[^>]+>.*?</figure>", head, re.DOTALL):
        return "has_ig_likely", "figure_img_head"

    return "needs_ig", "no_image_in_head"

def should_exclude(entry, body_text) -> tuple[bool, str]:
    if any(cat in EXCLUDED_CATEGORIES for cat in entry.categories): return True, "category"
    if len(body_text) < 500: return True, "short_content"
    total = count_all_links(entry.body_html)
    internal = count_internal_hatena_links(entry.body_html)
    if total >= 5 and internal / total >= 0.8: return True, "index_page"
    return False, ""
```

`EXCLUDED_CATEGORIES = {"日記", "雑記", "お知らせ", "ブログ運営", "プライベート"}`

### 4.3 `generate_batch_csv.py` — 日次20件抽出

**入力**: `state.sqlite`（`status='pending'` の対象記事）
**出力**: `data/batches/<today>_prompts.csv`（テンプレート）+ `data/entries_for_claude/<today>.md`

**バッチ選定ロジック**:

```python
def select_daily_batch(conn, batch_size=20) -> list[str]:
    # 優先度1: 前日NG判定された記事（20%割当）
    ng_revision = conn.execute("""
        SELECT entry_id FROM articles
        WHERE status='ng_pending'
        ORDER BY last_ng_date ASC
        LIMIT ?
    """, (batch_size // 5,)).fetchall()

    # 優先度2: 未着手・古い順
    remaining = batch_size - len(ng_revision)
    pending = conn.execute("""
        SELECT entry_id FROM articles
        WHERE status='pending' AND excluded=0
        ORDER BY published_date ASC
        LIMIT ?
    """, (remaining,)).fetchall()

    return [r[0] for r in ng_revision + pending]
```

### 4.4 Claude Code 対話ステップ（スクリプト化しない部分）

Claudeがセッション内で実行:

1. `data/entries_for_claude/<today>.md` を読む
2. 各記事について決定:
   - **スタイル選定**: 「神経疾患・青系・解剖図レイアウト」等
   - **プロンプト作文**: ChatGPT GPT Image 1 用、約300〜500字、「日本語のみ、英数字最小限」指示を含む
   - **alt文**: 120字前後、主キーワード含む
3. 結果を `<today>_prompts.csv` の空欄に埋めて確定

### 4.5 `build_dashboard.py` — レビューダッシュボード生成

**入力**: `data/batches/<today>_prompts.csv` + `images/in/<entry_id>.png`
**出力**: `data/reviews/<today>_review.html`

- 1カード1記事（タイトル・本文冒頭300字・生成画像・記事リンク・OK/NG/Skipラジオ）
- NG選択時: 理由タグチェックボックス＋自由記述テキストエリア（JSで表示切替）
- 画像未到着の記事はグレーアウト
- エクスポートボタンで `<today>_decisions.csv` をダウンロード（クライアントサイドJS）

### 4.6 `process_decisions.py` — OK→反映 / NG→改訂要求

**入力**: `data/decisions/<today>_decisions.csv`
**出力**: Fotolifeアップ + 記事HTML更新 + DBステータス更新

```python
# idempotencyガード: 既処理CSVの行はスキップ（processed_decisionsテーブルで管理）
def is_already_processed(conn, csv_path: str, entry_id: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM processed_decisions WHERE decisions_csv_path=? AND entry_id=?",
        (csv_path, entry_id),
    ).fetchone()
    return row is not None

def mark_as_processed(conn, csv_path: str, entry_id: str, decision: str) -> None:
    conn.execute(
        "INSERT OR IGNORE INTO processed_decisions VALUES (?, ?, ?, ?)",
        (csv_path, entry_id, decision, now_iso()),
    )
    conn.commit()

# メインループ（トランザクション内でatomic処理）
for row in decisions:
    if is_already_processed(conn, decisions_csv_path, row.entry_id):
        continue

    with conn:  # 自動コミット/ロールバック
        if row.decision == "ok":
            # INV-1: 既に fotolife_url が埋まっていれば再アップしない
            existing_url = get_fotolife_url(conn, row.entry_id)
            image_url = existing_url or upload_fotolife(f"images/in/{row.entry_id}.png")
            insert_infographic_into_entry(row.entry_id, image_url, row.alt_text)
            move_image(f"images/in/{row.entry_id}.png",
                       f"images/archive/{row.entry_id}.png")
            update_db(conn, entry_id=row.entry_id, status="done",
                      fotolife_url=image_url, processed_at=now_iso())
        elif row.decision == "ng":
            update_generation(conn, entry_id=row.entry_id, decision="ng",
                              ng_tags=row.ng_tags, ng_note=row.ng_note)
            move_image(f"images/in/{row.entry_id}.png",
                       f"images/archive/{row.entry_id}_ng_{ts}.png")
            mark_for_revision(conn, row.entry_id)  # 翌日バッチで再出
        elif row.decision == "skip":
            pass  # status='pending' のまま

        mark_as_processed(conn, decisions_csv_path, row.entry_id, row.decision)
```

### 4.7 HTML挿入ロジック（3段階フォールバック）

```python
def insert_infographic(body_html: str, img_tag: str) -> str:
    infographic_block = (
        '\n<!-- インフォグラフィック -->\n'
        '<div style="text-align:center; margin:20px 0;">\n'
        f'{img_tag}\n'
        '</div>\n\n'
    )

    # Level 1: <!-- 目次 --> マーカー
    if "<!-- 目次 -->" in body_html:
        return body_html.replace("<!-- 目次 -->", infographic_block + "<!-- 目次 -->")

    # Level 2: 最初の <h2> の直前
    h2_match = re.search(r'(\n|^)<h2[^>]*>', body_html)
    if h2_match:
        pos = h2_match.start()
        return body_html[:pos] + infographic_block + body_html[pos:]

    # Level 3: 最初の </p> の直後
    p_match = re.search(r'</p>', body_html)
    if p_match:
        pos = p_match.end()
        return body_html[:pos] + infographic_block + body_html[pos:]

    # Level 4 (最終): 本文先頭
    return infographic_block + body_html
```

---

## 5. データフロー・日次ルーチン

### 5.1 状態遷移

```
  [pending] ──(generate_batch_csv)──→ [in_batch]
                                        │
                                        │ ChatGPTで生成 → images/in/ に配置
                                        ↓
                                  [image_received]
                                        │
                                        │ build_dashboard → レビュー
                                        ↓
                                ┌───[approved]───→ [uploaded] ───→ [done] ✓
                                │
                                ├───[ng]────→ attempt+1、NGタグ記録
                                │              └─→ [ng_pending]（翌日再生成）
                                │
                                └───[skip]──→ [pending]（翌日以降）
```

### 5.2 日次ルーチン（20件/日、15-20分作業）

```
┌─ DAY N ────────────────────────────────────────────────────────┐
│ ① 朝（2分）─ Claude Codeセッション開始                         │
│    U: 「今日の20件お願いします」                                 │
│    C: generate_batch_csv.py → entries_for_claude/<today>.md 読込│
│       → 各記事を解析・プロンプト作文 → prompts.csv 確定         │
│                                                                │
│ ② 日中（5-10分）─ ChatGPT手動                                  │
│    U: prompts.csv 1行ずつコピー → ChatGPT生成 → ダウンロード    │
│       → images/in/<entry_id>.png にリネーム保存                 │
│                                                                │
│ ③ レビュー（5-7分）─ ダッシュボード                            │
│    U: 「ダッシュボード作って」                                   │
│    C: build_dashboard.py → review_<date>.html                  │
│    U: ブラウザで20件確認 → OK/NG → エクスポート                 │
│       → decisions/<date>_decisions.csv 保存                     │
│                                                                │
│ ④ 反映（3分）─ バッチ処理                                      │
│    U: 「決定を反映して」                                         │
│    C: process_decisions.py --yes                              │
│       → OK: Fotolife+AtomPub更新、NG: DB記録、Skip: 維持       │
│    C: 「完了。OK=17, NG=2, Skip=1 / 累計 X/227」               │
└────────────────────────────────────────────────────────────────┘
```

### 5.3 プロンプトCSV形式（`<today>_prompts.csv`）

列: `entry_id,title,category,url,style_summary,prompt,alt,attempt,previous_ng_tags,previous_ng_note`

### 5.4 decisions.csv形式（ユーザーエクスポート）

列: `entry_id,decision,ng_tags,ng_note`

---

## 6. エラー処理・リカバリ

### 6.1 主要失敗パターン

| # | 失敗 | 検知 | リカバリ |
|---|------|------|----------|
| 1 | AtomPub取得中API制限(429) | HTTPError | 間隔2→5秒延長、3回リトライ |
| 2 | エントリXMLパース失敗 | None返却 | `data/entries/failed/` へ移動、次へ |
| 3 | Fotolifeアップ失敗 | "ERROR"プレフィクス | 3リトライ、失敗時 `status='upload_failed'` |
| 4 | AtomPub PUT失敗 | HTTP 4xx/5xx | リトライ、`status='update_failed'`、Fotolife URLは保持（再アップ回避） |
| 5 | 挿入位置マーカーなし | regex不一致 | 3段階フォールバック（目次→h2→p） |
| 6 | 画像ファイル未到着 | ファイル不在 | ダッシュボードでグレーアウト、自動skip |
| 7 | decisions.csv不正 | CSV parseエラー | 処理中断、ユーザー通知 |
| 8 | 画像フォーマット異常 | Pillow (PIL) `Image.open` 失敗 or サイズ<100KB/>10MB | ダッシュボード警告、差し替え依頼（新規依存: `pillow` をrequirementsに追加） |
| 9 | 記事競合（他PC編集） | app:edited更新 | `status='conflict'`、ユーザー確認後続行 |
| 10 | NG再生成ループ（3回） | `attempt >= 3` | `status='manual_intervention'`、手動キュー |

### 6.2 キーインバリアント

- **INV-1: 画像重複アップ禁止**: `fotolife_url IS NOT NULL`なら再アップ不可
- **INV-2: 部分適用禁止**: 「Fotolifeアップ成功＋HTML更新成功」をATOMICに扱う
- **INV-3: DBが真実の源**: CSV/ファイルシステム状態は DB から再生成可能
- **INV-4: idempotency**: 同一decisions.csvの2回実行で副作用なし

### 6.3 ロールバック

- `articles.body_html_before` を7日保持
- `python scripts/rollback.py <entry_id>` で手動巻き戻し

### 6.4 安全装置

1. **Dry-run**: `--dry-run` で実API呼出なし
2. **バッチサイズ上限**: 1バッチ最大100件
3. **明示確認**: 本番実行時「続行しますか? (yes/no)」、`--yes`で省略可
4. **バックアップ**: 処理前に元本文をDB保存

---

## 7. テスト戦略

### 7.1 優先度付きテスト対象

| # | 対象 | 優先度 |
|---|------|--------|
| 1 | `detect_v2()` 判定ロジック | ★★★ |
| 2 | `should_exclude()` 除外ロジック | ★★★ |
| 3 | HTML挿入位置（3段階フォールバック） | ★★★ |
| 4 | idempotency（再実行安全性） | ★★★ |
| 5 | AtomPub XML生成 | ★★ |
| 6 | WSSE認証ヘッダ生成 | ★★ |
| 7 | CSV読み書き（エスケープ） | ★★ |

### 7.2 テスト配置

```
tests/
├── unit/
│   ├── test_detect.py           # detect_v2() パラメタライズ
│   ├── test_exclude.py
│   ├── test_insert.py           # 3段階フォールバック
│   ├── test_atompub_xml.py
│   └── test_wsse.py
├── integration/
│   ├── test_idempotent.py       # 同じCSV 2回実行 → 副作用なし
│   ├── test_rollback.py
│   └── test_full_batch_dry.py
└── fixtures/
    └── entries/                 # 実データ抜粋20件
```

### 7.3 カバレッジ目標

- Unit: 95%+
- Integration: 80%+
- E2E: 実施せず（手動QAで代替）

### 7.4 手動QAチェックリスト（各バッチ後）

```
□ SELECT COUNT(*) FROM articles WHERE status='done' が期待値と一致
□ Fotolife URL全件アクセス可能（HEAD 200）
□ Hatena で3〜5記事ランダムにブラウザ確認
□ alt属性が日本語で設定されている
□ 目次の位置関係（インフォグラフィックが目次の直前 or h2前）
□ 記事本文の非破壊（目次後コンテンツの順序維持）
□ カテゴリ保持
□ errors.log に新規ERRORなし
```

### 7.5 テストを書かない部分

- **Claude Code対話ステップ**（プロンプト生成・スタイル選定）: 創造的判断のため自動テスト困難、初日バッチを全件目視で代替
- **ダッシュボードJS挙動**: 単機能、初日手動確認で十分

---

## 8. 初日（Day 1）チェックリスト

プロジェクト本格開始前の必須ステップ:

```
□ 1. fetch_atompub.py 実行済み（完了: inventory_scan.py で代行、1003件キャッシュ取得済み）
□ 2. inventory.py 実行済み、targets_v2.csv = 227件確認済み（完了）
□ 3. requirements.txt に pillow, pytest を追加しpip install
□ 4. init_db.py 実行 → state.sqlite 初期化 + targets_v2.csv を articles テーブルに投入
□ 5. 既存 inventory_scan.py / inventory_rescan_v2.py を設計通りに fetch_atompub.py / inventory.py に分離
□ 6. tests/unit/ のユニットテストを全て通す（detect, exclude, insert の3本）
□ 7. generate_batch_csv.py --limit 1 で1件のみ抽出テスト
□ 8. ChatGPTで1件だけ生成、images/in/ に配置
□ 9. build_dashboard.py 実行、ダッシュボード表示確認
□10. decisions.csv を手動作成（1件、OK判定）
□11. process_decisions.py --dry-run で何が起きるかログ確認
□12. process_decisions.py 実行（本番1件、自分の下書き記事推奨）
□13. ブラウザで更新結果確認
□14. OK なら Day 2 から 20件/日の本番運用開始
```

---

## 9. 成功基準（Definition of Done）

プロジェクト完了の判定:

1. `SELECT COUNT(*) FROM articles WHERE status='done'` が **227** に到達
2. `errors.log` にクリティカルなERRORが残っていない
3. サンプル10記事をブラウザで目視確認し、表示崩れなし
4. Fotolife に孤児画像（記事から参照されていない画像）がない or `fotolife_cleanup.py` で確認済み
5. Tolosa-Hunt や Practice-Changing論文記事など、**GSCインプレッション上位の記事**で画像検索流入が発生していることを1ヶ月後に確認

---

## 10. リスクと対応

| リスク | 影響度 | 対応 |
|--------|--------|------|
| ChatGPTの画像生成品質が期待以下 | 高 | NGタグ＋自動改訂プロンプトで品質向上。初週終了時点で誤り率を集計し、プロンプト設計を見直し |
| Fotolifeの容量上限到達 | 中 | 有料プランで月5GB、1枚500KB想定で1万枚アップ可。227枚なら余裕 |
| AtomPub API制限での長時間停止 | 中 | 再試行＋間隔延長で吸収。最悪ケースでも1日中断で復旧 |
| 本番記事への誤更新事故 | 高 | Dry-run必須化、`body_html_before`保存、rollback.py 整備 |
| ユーザーが途中で離脱（継続性） | 中 | SQLite＋`images/in/`ファイル状態で中断可能、任意日数後に再開可能 |
| 医学的誤りが公開されてしまう | 高 | **ダッシュボード全件レビュー**が一次防御、事後監査チェックリストが二次防御 |

---

## 11. スコープ外（Non-Goals）

本プロジェクトで**やらないこと**を明示:

- 既にインフォグラフィック付きの734件の再生成・差し替え
- 除外15件（日記・インデックス）への対応
- medical-ddx-tools の HTML インフォグラフィックとの連携（記事内の`<a>`リンク挿入）
- 英語圏向け記事（hinyan1016は日本語専用）
- 画像検索SEO以外のSEO施策（メタディスクリプション・構造化データ）
- ChatGPT APIの自動化（UI手動運用を維持）

---

## 12. 付録

### 12.1 実測データ参照

- `blog/infographic-project/data/inventory_v2.csv` — 全1003件の判定結果
- `blog/infographic-project/data/targets_v2.csv` — 対象227件（古い順ソート済み）
- `blog/infographic-project/data/inventory_summary_v2.txt` — 集計サマリ

### 12.2 関連既存資産

- `blog/_scripts/upload_and_insert_infographic.py` — 1記事用の既存実装、WSSE/AtomPubロジックを流用
- `blog/infographic-project/scripts/inventory_scan.py` — 初回fetch+v1判定（実行済み、`fetch_atompub.py` と `inventory.py` へ分離予定）
- `blog/infographic-project/scripts/inventory_rescan_v2.py` — v2判定ロジック（実行済み、`inventory.py` の detect_v2() 本体）
- `youtube-slides/食事指導シリーズ/_shared/.env` — `HATENA_API_KEY` 格納
- `CLAUDE.md`（プロジェクトルート） — ブログSEO必須項目・インフォグラフィック挿入位置規約

### 12.4 新規依存パッケージ

- `pillow` — 画像フォーマット/サイズ検証（エラーパターン#8）
- `pytest` — ユニット・統合テスト実行
- その他の機能は Python 3.13 標準ライブラリで完結（urllib.request, sqlite3, csv, re, base64, hashlib）

### 12.3 用語

- **エントリID**: はてなAtomPubで使う編集用ID（17桁の数字）
- **AtomPub**: はてなの記事取得・更新用REST API（XMLベース）
- **Fotolife**: はてなの画像ホスティングサービス（記事内`<img>`のsrc先）
- **グレーゾーン**: 判定ロジックでインフォグラフィックの有無が確定できない記事

---

**本設計書は 2026-04-23 のブレインストーミングセッション結果に基づく。実装着手前にユーザー最終レビューを経て確定する。**
