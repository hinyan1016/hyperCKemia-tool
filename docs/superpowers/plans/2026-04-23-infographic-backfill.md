# インフォグラフィック・バックフィル 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** はてなブログ hinyan1016 のインフォグラフィック未対応記事227件に対して、ChatGPT生成画像を追加する半自動バックフィルパイプラインを構築する。

**Architecture:** Python決定層（AtomPub/Fotolife/SQLite）+ Claude Code創造層（記事解析・プロンプト生成）+ ユーザー手動層（ChatGPT画像生成・レビュー）の3層ハイブリッド。SQLite が唯一の真実源、idempotent処理で事故耐性を確保。

**Tech Stack:** Python 3.13 標準ライブラリ（urllib, sqlite3, csv, re, hashlib, base64）+ pillow（画像検証）+ pytest（テスト）。フロントエンドは手書きHTML+vanilla JS。

**Spec:** [`docs/superpowers/specs/2026-04-23-infographic-backfill-design.md`](../specs/2026-04-23-infographic-backfill-design.md)

**Conventions:**
- すべての新規ファイルは `blog/infographic-project/` 配下
- コメント・UI文言は日本語
- Python PEP 8 + 型ヒント必須、black/ruff整形
- テストは pytest + fixtures ベース、カバレッジ Unit 95%+ / Integration 80%+
- 各タスク完了ごとにcommit（Co-Authored-By 等の追記なし）

**既存資産と本プランのスコープ境界:**
- `blog/infographic-project/scripts/inventory_scan.py` — 既に存在、AtomPub取得ロジック + v1判定。**本プランでは再利用（削除しない・変更しない）**。fetch_atompub.py として機能。差分取得が必要になった将来時点でリファクタ予定
- `blog/infographic-project/scripts/inventory_rescan_v2.py` — 既に存在、v2判定の初版。**Task 3で `scripts/inventory.py` にロジックを移植**した後、旧ファイルは残置（実行履歴として保存）
- `blog/_scripts/upload_and_insert_infographic.py` — 既存1記事用スクリプト、Task 6 で `scripts/hatena_client.py` に WSSE/AtomPubロジックを抽出。旧ファイルは他用途で使われる可能性があるため残す

---

## Task 1: 環境セットアップと依存定義

**Files:**
- Create: `blog/infographic-project/requirements.txt`
- Create: `blog/infographic-project/scripts/__init__.py`
- Create: `blog/infographic-project/tests/__init__.py`
- Create: `blog/infographic-project/tests/unit/__init__.py`
- Create: `blog/infographic-project/tests/integration/__init__.py`
- Create: `blog/infographic-project/tests/fixtures/__init__.py`
- Create: `blog/infographic-project/pytest.ini`

- [ ] **Step 1: requirements.txt を作成**

```
pillow==11.0.0
pytest==8.3.4
pytest-cov==6.0.0
```

- [ ] **Step 2: pytest.ini を作成**

```ini
[pytest]
testpaths = tests
python_files = test_*.py
python_classes = Test*
python_functions = test_*
addopts = -v --strict-markers
markers =
    unit: 単体テスト（外部依存なし）
    integration: 統合テスト（SQLite・fakeサーバー使用）
```

- [ ] **Step 3: 空の `__init__.py` ファイル群を作成**

```bash
touch blog/infographic-project/scripts/__init__.py
touch blog/infographic-project/tests/__init__.py
touch blog/infographic-project/tests/unit/__init__.py
touch blog/infographic-project/tests/integration/__init__.py
touch blog/infographic-project/tests/fixtures/__init__.py
```

- [ ] **Step 4: 依存インストールと動作確認**

```bash
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pip install -r blog/infographic-project/requirements.txt
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest blog/infographic-project/tests/ --collect-only
```

Expected: `no tests ran in X.XXs` (コレクト成功、まだテストなし)

- [ ] **Step 5: コミット**

```bash
git add blog/infographic-project/requirements.txt blog/infographic-project/pytest.ini blog/infographic-project/scripts/__init__.py blog/infographic-project/tests/
git commit -m "chore(infographic): セットアップ - requirements.txt と pytest構成"
```

---

## Task 2: SQLite初期化スクリプト

**Files:**
- Create: `blog/infographic-project/scripts/init_db.py`
- Create: `blog/infographic-project/tests/unit/test_init_db.py`

- [ ] **Step 1: テストを書く（失敗確認用）**

`blog/infographic-project/tests/unit/test_init_db.py`:

```python
from __future__ import annotations

import sqlite3
import tempfile
from pathlib import Path

from scripts.init_db import initialize_schema, load_targets_csv


def test_schema_creates_four_tables():
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.sqlite"
        initialize_schema(db_path)
        conn = sqlite3.connect(db_path)
        tables = {
            row[0]
            for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
        }
        assert tables == {
            "articles",
            "generations",
            "batches",
            "processed_decisions",
        }


def test_load_targets_marks_articles_pending():
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.sqlite"
        csv_path = Path(tmpdir) / "targets.csv"
        csv_path.write_text(
            "entry_id,title,category,published,word_count,reason,url\n"
            "123,テスト記事,脳神経内科,2025-03-07T04:20:53+09:00,8582,no_image_in_head,https://x.y/z\n",
            encoding="utf-8-sig",
        )
        initialize_schema(db_path)
        count = load_targets_csv(db_path, csv_path)
        assert count == 1
        conn = sqlite3.connect(db_path)
        row = conn.execute(
            "SELECT entry_id, status, title FROM articles WHERE entry_id='123'"
        ).fetchone()
        assert row == ("123", "pending", "テスト記事")


def test_load_targets_idempotent():
    with tempfile.TemporaryDirectory() as tmpdir:
        db_path = Path(tmpdir) / "test.sqlite"
        csv_path = Path(tmpdir) / "targets.csv"
        csv_path.write_text(
            "entry_id,title,category,published,word_count,reason,url\n"
            "123,テスト,cat,2025-01-01T00:00:00+09:00,1000,no_image,https://x\n",
            encoding="utf-8-sig",
        )
        initialize_schema(db_path)
        load_targets_csv(db_path, csv_path)
        load_targets_csv(db_path, csv_path)
        conn = sqlite3.connect(db_path)
        count = conn.execute("SELECT COUNT(*) FROM articles").fetchone()[0]
        assert count == 1
```

- [ ] **Step 2: テストを実行して失敗を確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_init_db.py -v
```

Expected: `ModuleNotFoundError: No module named 'scripts.init_db'`

- [ ] **Step 3: 最小実装を書く**

`blog/infographic-project/scripts/init_db.py`:

```python
"""SQLite state.sqlite の初期化と targets_v2.csv の投入。"""

from __future__ import annotations

import csv
import sqlite3
import sys
from pathlib import Path


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS articles (
  entry_id TEXT PRIMARY KEY,
  title TEXT,
  category TEXT,
  published_date TEXT,
  body_html_path TEXT,
  body_html_before TEXT,
  word_count INTEGER,
  has_infographic INTEGER DEFAULT 0,
  excluded INTEGER DEFAULT 0,
  excluded_reason TEXT,
  status TEXT DEFAULT 'pending',
  last_batch_date TEXT,
  last_ng_date TEXT,
  edited_at TEXT
);

CREATE TABLE IF NOT EXISTS generations (
  gen_id INTEGER PRIMARY KEY AUTOINCREMENT,
  entry_id TEXT,
  attempt INTEGER,
  prompt TEXT,
  alt_text TEXT,
  style_summary TEXT,
  image_path TEXT,
  decision TEXT,
  ng_tags TEXT,
  ng_note TEXT,
  fotolife_url TEXT,
  created_at TEXT,
  processed_at TEXT,
  FOREIGN KEY (entry_id) REFERENCES articles(entry_id)
);

CREATE TABLE IF NOT EXISTS batches (
  batch_date TEXT PRIMARY KEY,
  entry_ids TEXT,
  prompts_csv TEXT,
  dashboard_html TEXT,
  decisions_csv TEXT,
  status TEXT
);

CREATE TABLE IF NOT EXISTS processed_decisions (
  decisions_csv_path TEXT,
  entry_id TEXT,
  decision TEXT,
  processed_at TEXT,
  PRIMARY KEY (decisions_csv_path, entry_id)
);
"""


def initialize_schema(db_path: Path) -> None:
    """state.sqlite を作成、スキーマを適用。既存DBは破壊しない。"""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(SCHEMA_SQL)
        conn.commit()
    finally:
        conn.close()


def load_targets_csv(db_path: Path, csv_path: Path) -> int:
    """targets_v2.csv の行を articles テーブルに INSERT OR IGNORE で投入。"""
    conn = sqlite3.connect(db_path)
    inserted = 0
    try:
        with csv_path.open(encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                cur = conn.execute(
                    """INSERT OR IGNORE INTO articles
                    (entry_id, title, category, published_date, word_count, status)
                    VALUES (?, ?, ?, ?, ?, 'pending')""",
                    (
                        row["entry_id"],
                        row["title"],
                        row["category"],
                        row["published"],
                        int(row.get("word_count") or 0),
                    ),
                )
                if cur.rowcount > 0:
                    inserted += 1
        conn.commit()
    finally:
        conn.close()
    return inserted


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    csv_path = root / "data" / "targets_v2.csv"

    initialize_schema(db_path)
    print(f"スキーマ初期化: {db_path}")

    if csv_path.exists():
        count = load_targets_csv(db_path, csv_path)
        print(f"articles テーブルに {count} 件投入")
    else:
        print(f"[警告] {csv_path} がありません。inventory スキャンを先に実行してください。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: テスト再実行して成功を確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_init_db.py -v
```

Expected: `3 passed`

- [ ] **Step 5: 本番DBを初期化して実データを投入、コミット**

```bash
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe blog/infographic-project/scripts/init_db.py
```

Expected: `articles テーブルに 227 件投入`

```bash
git add blog/infographic-project/scripts/init_db.py blog/infographic-project/tests/unit/test_init_db.py
git commit -m "feat(infographic): init_db.py - SQLite スキーマと targets 投入"
```

---

## Task 3: 判定ロジック（detect_v2）のテスト整備

**Files:**
- Create: `blog/infographic-project/tests/fixtures/entries/` 配下 8〜10個のXML
- Create: `blog/infographic-project/tests/unit/test_detect.py`

- [ ] **Step 1: fixtures を実データから抽出（シェルワンライナー）**

```bash
mkdir -p blog/infographic-project/tests/fixtures/entries
# has_ig_strong (comment_marker)
cp blog/infographic-project/data/entries/17179246901379369429.xml blog/infographic-project/tests/fixtures/entries/with_infographic_comment.xml
# has_ig_likely (fotolife_figure)
cp blog/infographic-project/data/entries/17179246901379066755.xml blog/infographic-project/tests/fixtures/entries/with_fotolife_figure.xml
# has_ig_likely (fotolife_img after iframe)
cp blog/infographic-project/data/entries/17179246901379326820.xml blog/infographic-project/tests/fixtures/entries/iframe_then_fotolife.xml
# needs_ig (no image)
cp blog/infographic-project/data/entries/17179246901378951746.xml blog/infographic-project/tests/fixtures/entries/no_image.xml
# needs_ig (2026-03-31 Tolosa-Hunt)
cp blog/infographic-project/data/entries/17179246901371022240.xml blog/infographic-project/tests/fixtures/entries/only_external_link.xml
```

- [ ] **Step 2: パラメタライズドテストを書く**

`blog/infographic-project/tests/unit/test_detect.py`:

```python
from __future__ import annotations

from pathlib import Path

import pytest

from scripts.inventory import detect_v2, parse_entry_xml


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "entries"


def _body_from_fixture(name: str) -> str:
    xml = (FIXTURES / name).read_text(encoding="utf-8")
    entry = parse_entry_xml(xml)
    assert entry is not None, f"fixture parse failed: {name}"
    return entry.body_html


@pytest.mark.unit
@pytest.mark.parametrize(
    "fixture_name, expected_status, expected_reason_part",
    [
        ("with_infographic_comment.xml", "has_ig_strong", "comment_marker"),
        ("with_fotolife_figure.xml", "has_ig_likely", "fotolife"),
        ("iframe_then_fotolife.xml", "has_ig_likely", "fotolife"),
        ("no_image.xml", "needs_ig", "no_image"),
        ("only_external_link.xml", "needs_ig", "no_image"),
    ],
)
def test_detect_v2_classifies_correctly(
    fixture_name: str, expected_status: str, expected_reason_part: str
) -> None:
    body = _body_from_fixture(fixture_name)
    status, reason = detect_v2(body)
    assert status == expected_status, (
        f"{fixture_name}: expected {expected_status}, got {status} ({reason})"
    )
    assert expected_reason_part in reason, (
        f"{fixture_name}: reason {reason!r} missing {expected_reason_part!r}"
    )


@pytest.mark.unit
def test_detect_v2_ignores_iframe_content() -> None:
    # iframe内の<img>タグは判定に影響しない
    body = (
        '<p>lead</p>'
        '<iframe width="560" height="315" src="https://youtube.com/...">'
        '<img src="fake.png" width="600"></iframe>'
        '<h2>本文</h2>'
    )
    status, _ = detect_v2(body)
    assert status == "needs_ig"
```

- [ ] **Step 3: テスト実行（初回は import エラー）**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_detect.py -v
```

Expected: `ModuleNotFoundError: No module named 'scripts.inventory'`

- [ ] **Step 4: inventory.py を作成（detect_v2 + parse_entry_xml を移植）**

`blog/infographic-project/scripts/inventory.py`:

```python
"""インフォグラフィック有無の判定ロジック（v2）。

inventory_rescan_v2.py を整理してモジュール化したもの。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional


EXCLUDED_CATEGORIES = {"日記", "雑記", "お知らせ", "ブログ運営", "プライベート"}
SHORT_CONTENT_THRESHOLD = 500
HEAD_WINDOW = 4000


@dataclass(frozen=True)
class Entry:
    entry_id: str
    title: str
    categories: tuple[str, ...]
    published: str
    body_html: str
    alternate_url: str


def parse_entry_xml(entry_xml: str) -> Optional[Entry]:
    """AtomPubのentry XMLをEntryに変換。"""
    id_match = re.search(r"/atom/entry/(\d+)", entry_xml)
    if not id_match:
        return None
    entry_id = id_match.group(1)

    title_match = re.search(r"<title[^>]*>([^<]*)</title>", entry_xml)
    title = title_match.group(1) if title_match else ""

    cats = re.findall(r'<category[^>]*term="([^"]+)"', entry_xml)

    published_match = re.search(r"<published>([^<]+)</published>", entry_xml)
    published = published_match.group(1) if published_match else ""

    content_match = re.search(
        r'<content[^>]*type="text/html"[^>]*>(.*?)</content>',
        entry_xml,
        re.DOTALL,
    )
    body_html = ""
    if content_match:
        raw = content_match.group(1).strip()
        if raw.startswith("<![CDATA["):
            raw = raw[len("<![CDATA[") :]
            if raw.endswith("]]>"):
                raw = raw[:-3]
        body_html = (
            raw.replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&amp;", "&")
            .replace("&quot;", '"')
            .replace("&#39;", "'")
        )

    alt_match = re.search(
        r'<link[^>]+rel="alternate"[^>]+href="([^"]+)"', entry_xml
    )
    alternate_url = alt_match.group(1) if alt_match else ""

    return Entry(
        entry_id=entry_id,
        title=title,
        categories=tuple(cats),
        published=published,
        body_html=body_html,
        alternate_url=alternate_url,
    )


def detect_v2(body_html: str) -> tuple[str, str]:
    """インフォグラフィックの有無を判定。

    Returns:
        (status, reason) where status in {'has_ig_strong', 'has_ig_likely', 'needs_ig'}
    """
    cleaned = re.sub(
        r"<iframe[^>]*>.*?</iframe>", "", body_html, flags=re.DOTALL | re.I
    )
    cleaned = re.sub(
        r"<script[^>]*>.*?</script>", "", cleaned, flags=re.DOTALL | re.I
    )

    # Strong signals
    if "<!-- インフォグラフィック -->" in cleaned:
        return "has_ig_strong", "comment_marker"
    if re.search(r'<img[^>]+alt="[^"]*インフォグラフィック', cleaned):
        return "has_ig_strong", "alt_marker_infographic"
    if re.search(r'<img[^>]+alt="[^"]*(まとめ|要約)', cleaned):
        return "has_ig_strong", "alt_marker_summary"
    if re.search(r'<img[^>]+src="[^"]*infograph', cleaned, flags=re.I):
        return "has_ig_strong", "filename_infographic"

    head = cleaned[:HEAD_WINDOW]

    if re.search(r"<figure[^>]*figure-image-fotolife", head, flags=re.I):
        return "has_ig_likely", "fotolife_figure_head"
    if re.search(
        r'<img[^>]+src="https?://cdn-ak\.f\.st-hatena\.com/images/fotolife',
        head,
    ):
        return "has_ig_likely", "fotolife_img_head"
    for m in re.finditer(r'<img[^>]+width="(\d+)"', head):
        try:
            if int(m.group(1)) >= 600:
                return "has_ig_likely", f"big_img_head_w{m.group(1)}"
        except ValueError:
            continue
    if re.search(r"<figure[^>]*>.*?<img[^>]+>.*?</figure>", head, re.DOTALL):
        return "has_ig_likely", "figure_img_head"

    return "needs_ig", "no_image_in_head"
```

- [ ] **Step 5: テスト再実行してグリーン、コミット**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_detect.py -v
```

Expected: `6 passed`

```bash
git add blog/infographic-project/scripts/inventory.py blog/infographic-project/tests/unit/test_detect.py blog/infographic-project/tests/fixtures/entries/
git commit -m "feat(infographic): inventory.py - detect_v2 + parse_entry_xml モジュール化"
```

---

## Task 4: 除外ロジック（should_exclude）

**Files:**
- Modify: `blog/infographic-project/scripts/inventory.py`
- Create: `blog/infographic-project/tests/unit/test_exclude.py`

- [ ] **Step 1: 失敗するテストを書く**

`blog/infographic-project/tests/unit/test_exclude.py`:

```python
from __future__ import annotations

import pytest

from scripts.inventory import Entry, should_exclude, strip_tags


def _entry(
    cats: tuple[str, ...] = (),
    body_html: str = "<p>" + "本文" * 300 + "</p>",
) -> Entry:
    return Entry(
        entry_id="1",
        title="t",
        categories=cats,
        published="2025-01-01T00:00:00+09:00",
        body_html=body_html,
        alternate_url="",
    )


@pytest.mark.unit
def test_exclude_diary_category() -> None:
    e = _entry(cats=("日記",))
    excluded, reason = should_exclude(e, strip_tags(e.body_html))
    assert excluded
    assert reason.startswith("category")


@pytest.mark.unit
def test_exclude_short_content() -> None:
    e = _entry(body_html="<p>短すぎる本文</p>")
    excluded, reason = should_exclude(e, strip_tags(e.body_html))
    assert excluded
    assert "short" in reason


@pytest.mark.unit
def test_exclude_index_page() -> None:
    internal_links = "\n".join(
        f'<a href="https://hinyan1016.hatenablog.com/entry/x/{i}">link {i}</a>'
        for i in range(10)
    )
    e = _entry(body_html="<p>" + "目次" * 300 + "</p>" + internal_links)
    excluded, reason = should_exclude(e, strip_tags(e.body_html))
    assert excluded
    assert "index" in reason


@pytest.mark.unit
def test_no_exclude_normal_article() -> None:
    e = _entry(cats=("脳神経内科",))
    excluded, _ = should_exclude(e, strip_tags(e.body_html))
    assert not excluded
```

- [ ] **Step 2: テスト実行して失敗確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_exclude.py -v
```

Expected: `ImportError: cannot import name 'should_exclude'`

- [ ] **Step 3: inventory.py に should_exclude と strip_tags を追加**

`blog/infographic-project/scripts/inventory.py` の末尾に追加:

```python
def strip_tags(html: str) -> str:
    text = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.DOTALL | re.I)
    text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.DOTALL | re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def count_internal_hatena_links(html: str) -> int:
    return len(
        re.findall(
            r'<a[^>]+href="https?://hinyan1016\.hatenablog\.com/', html
        )
    )


def count_all_links(html: str) -> int:
    return len(re.findall(r"<a\s[^>]*href=", html))


def should_exclude(entry: Entry, body_text: str) -> tuple[bool, str]:
    for cat in entry.categories:
        if cat in EXCLUDED_CATEGORIES:
            return True, f"category:{cat}"
    if len(body_text) < SHORT_CONTENT_THRESHOLD:
        return True, f"short:{len(body_text)}chars"
    total = count_all_links(entry.body_html)
    internal = count_internal_hatena_links(entry.body_html)
    if total >= 5 and internal / total >= 0.8:
        return True, f"index_page:{internal}/{total}"
    return False, ""
```

- [ ] **Step 4: テスト再実行してグリーン**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_exclude.py -v
```

Expected: `4 passed`

- [ ] **Step 5: コミット**

```bash
git add blog/infographic-project/scripts/inventory.py blog/infographic-project/tests/unit/test_exclude.py
git commit -m "feat(infographic): should_exclude - 除外ロジック（カテゴリ/短文/インデックス）"
```

---

## Task 5: HTML挿入ロジック（3段階フォールバック）

**Files:**
- Create: `blog/infographic-project/scripts/html_inserter.py`
- Create: `blog/infographic-project/tests/unit/test_insert.py`

- [ ] **Step 1: 失敗するテストを書く**

`blog/infographic-project/tests/unit/test_insert.py`:

```python
from __future__ import annotations

import pytest

from scripts.html_inserter import insert_infographic


IMG_TAG = '<img src="https://cdn-ak.f.st-hatena.com/images/fotolife/h/x.png" alt="テストalt">'


@pytest.mark.unit
def test_insert_with_toc_marker() -> None:
    body = "<p>リード文</p>\n<!-- 目次 -->\n<h2>セクション1</h2>"
    result = insert_infographic(body, IMG_TAG)
    assert "<!-- インフォグラフィック -->" in result
    assert result.index("<img") < result.index("<!-- 目次 -->")


@pytest.mark.unit
def test_insert_fallback_to_h2() -> None:
    body = "<p>リード文</p>\n<p>続き</p>\n<h2>セクション1</h2>"
    result = insert_infographic(body, IMG_TAG)
    assert result.index("<img") < result.index("<h2>")


@pytest.mark.unit
def test_insert_fallback_to_first_p() -> None:
    body = "<p>リード文だけ</p>\n<p>本文</p>"
    result = insert_infographic(body, IMG_TAG)
    first_p_end = result.index("</p>") + len("</p>")
    rest = result[first_p_end:].lstrip()
    assert rest.startswith("<!-- インフォグラフィック -->")


@pytest.mark.unit
def test_insert_fallback_to_prepend() -> None:
    body = "<div>pタグも見出しもない</div>"
    result = insert_infographic(body, IMG_TAG)
    assert result.startswith("\n<!-- インフォグラフィック -->")


@pytest.mark.unit
def test_insert_preserves_existing_content() -> None:
    body = "<p>リード</p>\n<!-- 目次 -->\n<h2>本文</h2>\n<p>結び</p>"
    result = insert_infographic(body, IMG_TAG)
    assert "<p>リード</p>" in result
    assert "<h2>本文</h2>" in result
    assert "<p>結び</p>" in result
```

- [ ] **Step 2: テスト実行して失敗確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_insert.py -v
```

Expected: `ModuleNotFoundError: No module named 'scripts.html_inserter'`

- [ ] **Step 3: html_inserter.py を実装**

`blog/infographic-project/scripts/html_inserter.py`:

```python
"""インフォグラフィック<img>タグを記事HTMLの適切な位置に挿入する。"""

from __future__ import annotations

import re


INFOGRAPHIC_TEMPLATE = (
    "\n<!-- インフォグラフィック -->\n"
    '<div style="text-align:center; margin:20px 0;">\n'
    "{img_tag}\n"
    "</div>\n\n"
)


def insert_infographic(body_html: str, img_tag: str) -> str:
    """3段階フォールバックで img_tag を body_html に挿入。

    優先順位:
      Level 1: <!-- 目次 --> コメントの直前
      Level 2: 最初の <h2> の直前
      Level 3: 最初の </p> の直後
      Level 4 (最終): 本文先頭
    """
    block = INFOGRAPHIC_TEMPLATE.format(img_tag=img_tag)

    if "<!-- 目次 -->" in body_html:
        return body_html.replace("<!-- 目次 -->", block + "<!-- 目次 -->", 1)

    h2_match = re.search(r"<h2[^>]*>", body_html)
    if h2_match:
        pos = h2_match.start()
        return body_html[:pos] + block + body_html[pos:]

    p_match = re.search(r"</p>", body_html)
    if p_match:
        pos = p_match.end()
        return body_html[:pos] + block + body_html[pos:]

    return block + body_html
```

- [ ] **Step 4: テスト再実行してグリーン**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_insert.py -v
```

Expected: `5 passed`

- [ ] **Step 5: コミット**

```bash
git add blog/infographic-project/scripts/html_inserter.py blog/infographic-project/tests/unit/test_insert.py
git commit -m "feat(infographic): html_inserter - 3段階フォールバック付き挿入ロジック"
```

---

## Task 6: AtomPub/Fotolife クライアント抽出

**Files:**
- Create: `blog/infographic-project/scripts/hatena_client.py`
- Create: `blog/infographic-project/tests/unit/test_hatena_client.py`

- [ ] **Step 1: WSSEヘッダ生成のテストを書く**

`blog/infographic-project/tests/unit/test_hatena_client.py`:

```python
from __future__ import annotations

import base64
import hashlib

import pytest

from scripts.hatena_client import (
    basic_auth_header,
    build_entry_put_xml,
    make_wsse_header,
)


@pytest.mark.unit
def test_make_wsse_header_format() -> None:
    header = make_wsse_header(
        "testuser",
        "testkey",
        nonce_str="abc",
        created="2026-04-23T00:00:00Z",
    )
    assert header.startswith('UsernameToken Username="testuser",')
    assert 'PasswordDigest="' in header
    assert 'Nonce="' in header
    assert 'Created="2026-04-23T00:00:00Z"' in header


@pytest.mark.unit
def test_make_wsse_header_digest_reproducible() -> None:
    h1 = make_wsse_header(
        "u", "k", nonce_str="xyz", created="2026-01-01T00:00:00Z"
    )
    h2 = make_wsse_header(
        "u", "k", nonce_str="xyz", created="2026-01-01T00:00:00Z"
    )
    assert h1 == h2


@pytest.mark.unit
def test_basic_auth_header() -> None:
    header = basic_auth_header("user", "key")
    assert header.startswith("Basic ")
    encoded = header[len("Basic ") :]
    assert base64.b64decode(encoded).decode() == "user:key"


@pytest.mark.unit
def test_build_entry_put_xml_preserves_title_and_categories() -> None:
    xml = build_entry_put_xml(
        title="テストタイトル & test",
        author="hinyan1016",
        body_html="<p>本文</p>",
        categories=("脳神経内科", "医師向け"),
    )
    assert "<title>テストタイトル &amp; test</title>" in xml
    assert '<category term="脳神経内科"' in xml
    assert '<category term="医師向け"' in xml
    assert "<![CDATA[<p>本文</p>]]>" in xml
    assert "<app:draft>no</app:draft>" in xml
```

- [ ] **Step 2: テスト実行、失敗確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_hatena_client.py -v
```

Expected: `ModuleNotFoundError: No module named 'scripts.hatena_client'`

- [ ] **Step 3: hatena_client.py 実装**

`blog/infographic-project/scripts/hatena_client.py`:

```python
"""はてなブログ AtomPub / Fotolife クライアント。

既存の blog/_scripts/upload_and_insert_infographic.py のロジックを
再利用可能な形に整理したもの。
"""

from __future__ import annotations

import base64
import hashlib
import random
import re
import string
import time
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


HATENA_ID_DEFAULT = "hinyan1016"
BLOG_DOMAIN_DEFAULT = "hinyan1016.hatenablog.com"


def load_env(env_file: Path) -> dict[str, str]:
    env: dict[str, str] = {}
    with env_file.open(encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                env[key.strip()] = val.strip()
    return env


def _escape_xml_text(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def make_wsse_header(
    username: str,
    api_key: str,
    *,
    nonce_str: Optional[str] = None,
    created: Optional[str] = None,
) -> str:
    if nonce_str is None:
        nonce_str = "".join(
            random.choices(string.ascii_letters + string.digits, k=40)
        )
    if created is None:
        created = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    nonce_bytes = nonce_str.encode("utf-8")
    digest_raw = hashlib.sha1(
        nonce_bytes + created.encode("utf-8") + api_key.encode("utf-8")
    ).digest()
    password_digest = base64.b64encode(digest_raw).decode("utf-8")
    nonce_b64 = base64.b64encode(nonce_bytes).decode("utf-8")
    return (
        f'UsernameToken Username="{username}", '
        f'PasswordDigest="{password_digest}", '
        f'Nonce="{nonce_b64}", '
        f'Created="{created}"'
    )


def basic_auth_header(username: str, api_key: str) -> str:
    token = base64.b64encode(f"{username}:{api_key}".encode()).decode()
    return f"Basic {token}"


def build_entry_put_xml(
    *,
    title: str,
    author: str,
    body_html: str,
    categories: tuple[str, ...],
    draft: bool = False,
) -> str:
    cats_xml = "\n".join(
        f'  <category term="{_escape_xml_text(c)}" />' for c in categories
    )
    draft_value = "yes" if draft else "no"
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<entry xmlns="http://www.w3.org/2005/Atom" '
        'xmlns:app="http://www.w3.org/2007/app">\n'
        f"  <title>{_escape_xml_text(title)}</title>\n"
        f"  <author><name>{_escape_xml_text(author)}</name></author>\n"
        f"  <content type=\"text/html\"><![CDATA[{body_html}]]></content>\n"
        f"{cats_xml}\n"
        "  <app:control>\n"
        f"    <app:draft>{draft_value}</app:draft>\n"
        "  </app:control>\n"
        "</entry>"
    )


def fetch_entry(
    entry_id: str,
    api_key: str,
    *,
    hatena_id: str = HATENA_ID_DEFAULT,
    blog_domain: str = BLOG_DOMAIN_DEFAULT,
) -> str:
    url = (
        f"https://blog.hatena.ne.jp/{hatena_id}/{blog_domain}"
        f"/atom/entry/{entry_id}"
    )
    req = urllib.request.Request(
        url,
        method="GET",
        headers={"Authorization": basic_auth_header(hatena_id, api_key)},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def put_entry(
    entry_id: str,
    xml_body: str,
    api_key: str,
    *,
    hatena_id: str = HATENA_ID_DEFAULT,
    blog_domain: str = BLOG_DOMAIN_DEFAULT,
) -> str:
    url = (
        f"https://blog.hatena.ne.jp/{hatena_id}/{blog_domain}"
        f"/atom/entry/{entry_id}"
    )
    req = urllib.request.Request(
        url,
        data=xml_body.encode("utf-8"),
        method="PUT",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "Authorization": basic_auth_header(hatena_id, api_key),
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def upload_to_fotolife(
    image_path: Path,
    api_key: str,
    *,
    title: str = "infographic",
    hatena_id: str = HATENA_ID_DEFAULT,
) -> str:
    """Fotolife に PNG をアップロードし、画像URLを返す。

    失敗時は 'ERROR:' または 'URL not found:' プレフィクスの文字列を返す。
    """
    image_data = image_path.read_bytes()
    image_b64 = base64.b64encode(image_data).decode("utf-8")

    xml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<entry xmlns="http://purl.org/atom/ns#">\n'
        f"  <title>{_escape_xml_text(title)}</title>\n"
        f'  <content mode="base64" type="image/png">{image_b64}</content>\n'
        "</entry>"
    )

    url = "https://f.hatena.ne.jp/atom/post"
    wsse = make_wsse_header(hatena_id, api_key)

    req = urllib.request.Request(
        url,
        data=xml.encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/xml; charset=utf-8",
            "X-WSSE": wsse,
            "Authorization": 'WSSE profile="UsernameToken"',
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="replace")
        return f"ERROR:HTTP {e.code}: {err[:500]}"

    m = re.search(r"<hatena:imageurl>([^<]+)</hatena:imageurl>", body)
    if m:
        return m.group(1)
    m = re.search(r"https://cdn-ak\.f\.st-hatena\.com/images/fotolife/[^<\"]+", body)
    if m:
        return m.group(0)
    return f"URL not found:{body[:500]}"
```

- [ ] **Step 4: テスト再実行してグリーン**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_hatena_client.py -v
```

Expected: `4 passed`

- [ ] **Step 5: コミット**

```bash
git add blog/infographic-project/scripts/hatena_client.py blog/infographic-project/tests/unit/test_hatena_client.py
git commit -m "feat(infographic): hatena_client - AtomPub/Fotolife クライアント分離"
```

---

## Task 7: バッチ生成スクリプト

**Files:**
- Create: `blog/infographic-project/scripts/generate_batch_csv.py`
- Create: `blog/infographic-project/tests/integration/test_batch_generation.py`

- [ ] **Step 1: 統合テストを書く**

`blog/infographic-project/tests/integration/test_batch_generation.py`:

```python
from __future__ import annotations

import csv
import sqlite3
import tempfile
from datetime import date
from pathlib import Path

import pytest

from scripts.generate_batch_csv import generate_daily_batch
from scripts.init_db import initialize_schema


def _insert_article(
    conn: sqlite3.Connection,
    entry_id: str,
    *,
    status: str = "pending",
    published: str = "2025-03-01T00:00:00+09:00",
    title: str = "title",
) -> None:
    conn.execute(
        """INSERT INTO articles
        (entry_id, title, category, published_date, status, excluded, word_count)
        VALUES (?, ?, 'cat', ?, ?, 0, 1000)""",
        (entry_id, title, published, status),
    )


@pytest.mark.integration
def test_generate_batch_picks_oldest_pending_first() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db_path = tmp / "s.sqlite"
        out_dir = tmp / "batches"
        initialize_schema(db_path)

        conn = sqlite3.connect(db_path)
        _insert_article(conn, "A", published="2025-01-01T00:00:00+09:00")
        _insert_article(conn, "B", published="2025-02-01T00:00:00+09:00")
        _insert_article(conn, "C", published="2025-03-01T00:00:00+09:00")
        conn.commit()
        conn.close()

        csv_path, md_path = generate_daily_batch(
            db_path, out_dir, batch_size=2, batch_date="2026-04-23"
        )
        with csv_path.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        assert [r["entry_id"] for r in rows] == ["A", "B"]

        conn = sqlite3.connect(db_path)
        statuses = dict(
            conn.execute("SELECT entry_id, status FROM articles").fetchall()
        )
        assert statuses == {
            "A": "in_batch",
            "B": "in_batch",
            "C": "pending",
        }


@pytest.mark.integration
def test_generate_batch_includes_ng_revision() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        db_path = tmp / "s.sqlite"
        out_dir = tmp / "batches"
        initialize_schema(db_path)

        conn = sqlite3.connect(db_path)
        # 4件 pending, 1件 ng_pending
        for i, eid in enumerate(["A", "B", "C", "D"]):
            _insert_article(
                conn,
                eid,
                published=f"2025-0{i + 1}-01T00:00:00+09:00",
            )
        _insert_article(
            conn, "NG1", status="ng_pending",
            published="2024-12-01T00:00:00+09:00",
        )
        conn.execute(
            "UPDATE articles SET last_ng_date='2026-04-22' WHERE entry_id='NG1'"
        )
        conn.commit()
        conn.close()

        csv_path, _ = generate_daily_batch(
            db_path, out_dir, batch_size=5, batch_date="2026-04-23"
        )
        with csv_path.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        entry_ids = [r["entry_id"] for r in rows]
        # NG が最低1件、残りは古い順pending
        assert "NG1" in entry_ids
        assert entry_ids.index("NG1") == 0  # NG優先
        assert "A" in entry_ids and "B" in entry_ids
```

- [ ] **Step 2: テスト実行して失敗確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/integration/test_batch_generation.py -v
```

Expected: `ModuleNotFoundError: No module named 'scripts.generate_batch_csv'`

- [ ] **Step 3: generate_batch_csv.py を実装**

`blog/infographic-project/scripts/generate_batch_csv.py`:

```python
"""日次バッチCSV生成：DBから20件選定してCSVとClaude読込用MDを作る。"""

from __future__ import annotations

import csv
import sqlite3
import sys
from datetime import date
from pathlib import Path


CSV_HEADER = [
    "entry_id",
    "title",
    "category",
    "url",
    "style_summary",
    "prompt",
    "alt",
    "attempt",
    "previous_ng_tags",
    "previous_ng_note",
]


def select_batch_entry_ids(
    conn: sqlite3.Connection, batch_size: int
) -> list[str]:
    ng_quota = max(1, batch_size // 5)
    ng_rows = conn.execute(
        """SELECT entry_id FROM articles
        WHERE status='ng_pending'
        ORDER BY last_ng_date ASC
        LIMIT ?""",
        (ng_quota,),
    ).fetchall()
    ng_ids = [r[0] for r in ng_rows]

    remaining = batch_size - len(ng_ids)
    pending_rows = conn.execute(
        """SELECT entry_id FROM articles
        WHERE status='pending' AND excluded=0
        ORDER BY published_date ASC
        LIMIT ?""",
        (remaining,),
    ).fetchall()
    pending_ids = [r[0] for r in pending_rows]

    return ng_ids + pending_ids


def generate_daily_batch(
    db_path: Path,
    out_dir: Path,
    *,
    batch_size: int = 20,
    batch_date: str | None = None,
) -> tuple[Path, Path]:
    """対象記事を選定し、prompts.csvテンプレと entries_for_claude/<date>.md を生成。

    Returns:
        (prompts_csv_path, entries_md_path)
    """
    batch_date = batch_date or date.today().isoformat()
    out_dir.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(db_path)
    try:
        entry_ids = select_batch_entry_ids(conn, batch_size)
        if not entry_ids:
            raise RuntimeError("選定対象がありません。全完了または在庫切れ。")

        rows: list[dict[str, object]] = []
        md_sections: list[str] = []
        for eid in entry_ids:
            row = conn.execute(
                """SELECT title, category, published_date FROM articles
                WHERE entry_id=?""",
                (eid,),
            ).fetchone()
            if not row:
                continue
            title, category, published = row

            # 前回NG情報
            prev_ng = conn.execute(
                """SELECT ng_tags, ng_note, attempt FROM generations
                WHERE entry_id=? AND decision='ng'
                ORDER BY attempt DESC LIMIT 1""",
                (eid,),
            ).fetchone()
            prev_tags = prev_ng[0] if prev_ng else ""
            prev_note = prev_ng[1] if prev_ng else ""
            next_attempt = (prev_ng[2] if prev_ng else 0) + 1

            rows.append(
                {
                    "entry_id": eid,
                    "title": title or "",
                    "category": category or "",
                    "url": f"https://hinyan1016.hatenablog.com/entry/{published[:10].replace('-', '/')}"
                    if published
                    else "",
                    "style_summary": "",
                    "prompt": "",
                    "alt": "",
                    "attempt": next_attempt,
                    "previous_ng_tags": prev_tags,
                    "previous_ng_note": prev_note,
                }
            )

            conn.execute(
                """UPDATE articles SET status='in_batch', last_batch_date=?
                WHERE entry_id=?""",
                (batch_date, eid),
            )

            md_sections.append(f"## entry_id: {eid}\n- title: {title}\n- category: {category}\n- published: {published}\n")

        conn.commit()

        csv_path = out_dir / f"{batch_date}_prompts.csv"
        with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_HEADER)
            writer.writeheader()
            writer.writerows(rows)

        md_path = out_dir.parent / "entries_for_claude" / f"{batch_date}.md"
        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text("\n".join(md_sections), encoding="utf-8")

        # batches テーブル記録
        conn.execute(
            """INSERT OR REPLACE INTO batches
            (batch_date, entry_ids, prompts_csv, status)
            VALUES (?, ?, ?, 'generating')""",
            (batch_date, ",".join(entry_ids), str(csv_path)),
        )
        conn.commit()

        return csv_path, md_path
    finally:
        conn.close()


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    out_dir = root / "data" / "batches"
    csv_path, md_path = generate_daily_batch(db_path, out_dir, batch_size=20)
    print(f"バッチCSV: {csv_path}")
    print(f"Claude参照MD: {md_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: テスト実行してグリーン**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/integration/test_batch_generation.py -v
```

Expected: `2 passed`

- [ ] **Step 5: コミット**

```bash
git add blog/infographic-project/scripts/generate_batch_csv.py blog/infographic-project/tests/integration/test_batch_generation.py
git commit -m "feat(infographic): generate_batch_csv - 日次20件選定とCSVテンプレ出力"
```

---

## Task 8: レビューダッシュボード生成

**Files:**
- Create: `blog/infographic-project/scripts/build_dashboard.py`
- Create: `blog/infographic-project/tests/unit/test_dashboard.py`

- [ ] **Step 1: ダッシュボードHTML内容のテストを書く**

`blog/infographic-project/tests/unit/test_dashboard.py`:

```python
from __future__ import annotations

import csv
import tempfile
from pathlib import Path

import pytest

from scripts.build_dashboard import build_dashboard


SAMPLE_PROMPTS_CSV = (
    "entry_id,title,category,url,style_summary,prompt,alt,attempt,previous_ng_tags,previous_ng_note\n"
    "111,片頭痛の5つのこと,神経,https://x/y,青系,img prompt,alt text,1,,\n"
    "222,ヘパリンの使い方,循環器,https://x/y2,緑系,img prompt2,alt text2,1,,\n"
)


@pytest.mark.unit
def test_dashboard_includes_all_entries() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        csv_path = tmp / "prompts.csv"
        csv_path.write_text(SAMPLE_PROMPTS_CSV, encoding="utf-8-sig")
        images_dir = tmp / "images" / "in"
        images_dir.mkdir(parents=True)
        (images_dir / "111.png").write_bytes(b"fake")

        out_html = tmp / "review.html"
        build_dashboard(csv_path, images_dir, out_html)

        html = out_html.read_text(encoding="utf-8")
        assert "片頭痛の5つのこと" in html
        assert "ヘパリンの使い方" in html
        assert 'data-entry-id="111"' in html
        assert 'data-entry-id="222"' in html


@pytest.mark.unit
def test_dashboard_marks_missing_images() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        csv_path = tmp / "prompts.csv"
        csv_path.write_text(SAMPLE_PROMPTS_CSV, encoding="utf-8-sig")
        images_dir = tmp / "images" / "in"
        images_dir.mkdir(parents=True)
        # 111.png のみ配置、222.png は未到着

        out_html = tmp / "review.html"
        build_dashboard(csv_path, images_dir, out_html)

        html = out_html.read_text(encoding="utf-8")
        # 222は画像未到着マーク付き
        assert 'data-status="awaiting_image"' in html


@pytest.mark.unit
def test_dashboard_has_export_button_and_ng_tags() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        csv_path = tmp / "prompts.csv"
        csv_path.write_text(SAMPLE_PROMPTS_CSV, encoding="utf-8-sig")
        images_dir = tmp / "images" / "in"
        images_dir.mkdir(parents=True)

        out_html = tmp / "review.html"
        build_dashboard(csv_path, images_dir, out_html)

        html = out_html.read_text(encoding="utf-8")
        assert 'id="export-button"' in html
        assert "medical_error" in html
        assert "typo" in html
        assert "design" in html
        assert "mismatch" in html
```

- [ ] **Step 2: テスト実行、失敗確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_dashboard.py -v
```

Expected: `ModuleNotFoundError: No module named 'scripts.build_dashboard'`

- [ ] **Step 3: build_dashboard.py 実装**

`blog/infographic-project/scripts/build_dashboard.py`:

```python
"""レビューダッシュボードHTML生成。"""

from __future__ import annotations

import csv
import html
import sys
from datetime import date
from pathlib import Path


HTML_HEAD = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>インフォグラフィック レビュー {batch_date}</title>
<style>
body { font-family: -apple-system, "Hiragino Sans", sans-serif; background: #f5f5f5; margin: 0; padding: 20px; }
.card { background: white; border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
.card[data-status="awaiting_image"] { opacity: 0.5; background: #fffbf0; }
.card h3 { margin-top: 0; }
.meta { color: #888; font-size: 0.9em; margin-bottom: 15px; }
.grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
.article-summary { border-left: 4px solid #2E86C1; padding-left: 15px; }
.generated-image img { max-width: 100%; height: auto; border: 1px solid #ddd; }
.decision { margin-top: 15px; padding-top: 15px; border-top: 1px solid #eee; }
.decision label { margin-right: 15px; }
.ng-tags { display: none; margin-top: 10px; padding: 10px; background: #fff3cd; border-radius: 4px; }
.ng-tags.visible { display: block; }
.ng-tags label { display: inline-block; margin: 3px; }
.ng-tags textarea { width: 100%; margin-top: 8px; min-height: 50px; }
#export-button { position: sticky; bottom: 20px; padding: 14px 28px; background: #2E86C1; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }
#summary { font-weight: bold; margin: 20px 0; }
</style>
</head>
<body>
<h1>インフォグラフィック レビュー — {batch_date}</h1>
<div id="summary">対象: {count}件</div>
"""

HTML_TAIL = """
<button id="export-button" onclick="exportDecisions()">決定をCSVにエクスポート</button>
<script>
document.querySelectorAll('input[type=radio][value=ng]').forEach(r => {
  r.addEventListener('change', e => {
    const card = e.target.closest('.card');
    card.querySelector('.ng-tags').classList.add('visible');
  });
});
document.querySelectorAll('input[type=radio]:not([value=ng])').forEach(r => {
  r.addEventListener('change', e => {
    const card = e.target.closest('.card');
    card.querySelector('.ng-tags').classList.remove('visible');
  });
});

function exportDecisions() {
  const rows = [['entry_id', 'decision', 'ng_tags', 'ng_note']];
  document.querySelectorAll('.card').forEach(card => {
    const eid = card.dataset.entryId;
    const dec = card.querySelector('input[type=radio]:checked');
    if (!dec) return;
    const tags = Array.from(
      card.querySelectorAll('.ng-tags input[type=checkbox]:checked')
    ).map(c => c.value).join(',');
    const note = (card.querySelector('.ng-tags textarea')?.value || '').replace(/\\n/g, ' ');
    rows.push([eid, dec.value, tags, note]);
  });
  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\\n');
  const blob = new Blob(["\\uFEFF" + csv], {type: 'text/csv'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = '{batch_date}_decisions.csv';
  a.click();
}
</script>
</body>
</html>
"""


def _card_html(row: dict[str, str], image_available: bool, image_rel_path: str) -> str:
    eid = html.escape(row["entry_id"])
    title = html.escape(row["title"])
    category = html.escape(row["category"])
    url = html.escape(row["url"])
    style = html.escape(row.get("style_summary", ""))
    alt = html.escape(row.get("alt", ""))
    status = "ready" if image_available else "awaiting_image"

    img_html = (
        f'<img src="{html.escape(image_rel_path)}" alt="生成画像">'
        if image_available
        else '<div style="padding:40px; background:#f0f0f0; text-align:center;">画像未到着</div>'
    )

    ng_tags_html = """
    <div class="ng-tags">
      <strong>NG理由（複数可）:</strong><br>
      <label><input type="checkbox" value="medical_error">医学的誤り</label>
      <label><input type="checkbox" value="typo">文字化け/誤字</label>
      <label><input type="checkbox" value="design">デザイン不満</label>
      <label><input type="checkbox" value="mismatch">内容不一致</label>
      <label><input type="checkbox" value="other">その他</label>
      <textarea placeholder="自由記述（任意）"></textarea>
    </div>
    """

    return f"""
<div class="card" data-entry-id="{eid}" data-status="{status}">
  <h3>{title}</h3>
  <div class="meta">カテゴリ: {category} | <a href="{url}" target="_blank">記事を開く</a></div>
  <div class="grid">
    <div class="article-summary">
      <h4>提案スタイル</h4>
      <p>{style}</p>
      <h4>alt文</h4>
      <p style="font-size:0.9em; color:#555;">{alt}</p>
    </div>
    <div class="generated-image">
      {img_html}
    </div>
  </div>
  <div class="decision">
    <label><input type="radio" name="d_{eid}" value="ok">OK</label>
    <label><input type="radio" name="d_{eid}" value="ng">NG</label>
    <label><input type="radio" name="d_{eid}" value="skip">スキップ</label>
    {ng_tags_html}
  </div>
</div>
"""


def build_dashboard(
    prompts_csv: Path,
    images_dir: Path,
    out_html: Path,
    *,
    batch_date: str | None = None,
) -> None:
    batch_date = batch_date or date.today().isoformat()
    with prompts_csv.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))

    cards = []
    for row in rows:
        eid = row["entry_id"]
        image_path = images_dir / f"{eid}.png"
        rel_path = str(image_path.relative_to(out_html.parent).as_posix())
        cards.append(_card_html(row, image_path.exists(), rel_path))

    body = "".join(cards)
    head = HTML_HEAD.format(batch_date=batch_date, count=len(rows))
    tail = HTML_TAIL.format(batch_date=batch_date)

    out_html.parent.mkdir(parents=True, exist_ok=True)
    out_html.write_text(head + body + tail, encoding="utf-8")


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    today = date.today().isoformat()
    prompts_csv = root / "data" / "batches" / f"{today}_prompts.csv"
    images_dir = root / "images" / "in"
    out_html = root / "data" / "reviews" / f"{today}_review.html"

    build_dashboard(prompts_csv, images_dir, out_html, batch_date=today)
    print(f"ダッシュボード: {out_html}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: テスト実行してグリーン**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/unit/test_dashboard.py -v
```

Expected: `3 passed`

- [ ] **Step 5: コミット**

```bash
git add blog/infographic-project/scripts/build_dashboard.py blog/infographic-project/tests/unit/test_dashboard.py
git commit -m "feat(infographic): build_dashboard - レビュー用HTML生成とJSエクスポート"
```

---

## Task 9: 決定処理スクリプト（最重要・idempotency保証）

**Files:**
- Create: `blog/infographic-project/scripts/process_decisions.py`
- Create: `blog/infographic-project/tests/integration/test_process_decisions.py`

- [ ] **Step 1: 統合テスト（idempotency含む）を書く**

`blog/infographic-project/tests/integration/test_process_decisions.py`:

```python
from __future__ import annotations

import csv
import sqlite3
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

from scripts.init_db import initialize_schema
from scripts.process_decisions import ProcessResult, process_decisions


def _seed_article(conn: sqlite3.Connection, eid: str) -> None:
    conn.execute(
        """INSERT INTO articles
        (entry_id, title, category, published_date, status, excluded)
        VALUES (?, 'タイトル', 'カテゴリ', '2025-01-01T00:00:00+09:00', 'in_batch', 0)""",
        (eid,),
    )


def _seed_generation(
    conn: sqlite3.Connection, eid: str, attempt: int = 1
) -> None:
    conn.execute(
        """INSERT INTO generations (entry_id, attempt, alt_text, created_at)
        VALUES (?, ?, 'alt', '2026-04-23T00:00:00+09:00')""",
        (eid, attempt),
    )


def _write_decisions_csv(path: Path, rows: list[tuple[str, str, str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["entry_id", "decision", "ng_tags", "ng_note"])
        w.writerows(rows)


def _make_fake_clients():
    class FakeClient:
        def __init__(self) -> None:
            self.upload_count = 0
            self.put_count = 0

        def upload(self, image_path: Path, api_key: str, **_: object) -> str:
            self.upload_count += 1
            return f"https://cdn-ak.f.st-hatena.com/fake/{image_path.stem}.png"

        def fetch(self, entry_id: str, api_key: str, **_: object) -> str:
            return (
                '<?xml version="1.0" encoding="utf-8"?>'
                '<entry xmlns="http://www.w3.org/2005/Atom">'
                f'<title>t{entry_id}</title>'
                '<content type="text/html">&lt;p&gt;本文&lt;/p&gt;&lt;!-- 目次 --&gt;</content>'
                '<category term="c" />'
                '</entry>'
            )

        def put(self, entry_id: str, xml: str, api_key: str, **_: object) -> str:
            self.put_count += 1
            return "OK"

    return FakeClient()


@pytest.mark.integration
def test_ok_decision_uploads_and_updates(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    _seed_article(conn, "E1")
    _seed_generation(conn, "E1")
    conn.commit()
    conn.close()

    images_in = tmp_path / "images" / "in"
    images_in.mkdir(parents=True)
    (images_in / "E1.png").write_bytes(b"fake-png")
    archive = tmp_path / "images" / "archive"
    archive.mkdir()

    decisions = tmp_path / "dec.csv"
    _write_decisions_csv(decisions, [("E1", "ok", "", "")])

    fake = _make_fake_clients()
    result = process_decisions(
        db_path,
        decisions,
        images_in,
        archive,
        api_key="k",
        upload_fn=fake.upload,
        fetch_fn=fake.fetch,
        put_fn=fake.put,
    )
    assert result.uploaded == 1
    assert result.ng_recorded == 0
    assert not (images_in / "E1.png").exists()
    assert (archive / "E1.png").exists()

    conn = sqlite3.connect(db_path)
    row = conn.execute(
        "SELECT status FROM articles WHERE entry_id='E1'"
    ).fetchone()
    assert row == ("done",)


@pytest.mark.integration
def test_idempotent_same_csv_twice(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    _seed_article(conn, "E1")
    _seed_generation(conn, "E1")
    conn.commit()
    conn.close()

    images_in = tmp_path / "images" / "in"
    images_in.mkdir(parents=True)
    (images_in / "E1.png").write_bytes(b"fake-png")
    archive = tmp_path / "images" / "archive"
    archive.mkdir()

    decisions = tmp_path / "dec.csv"
    _write_decisions_csv(decisions, [("E1", "ok", "", "")])

    fake = _make_fake_clients()
    upload_fn = fake.upload
    fetch_fn = fake.fetch
    put_fn = fake.put

    r1 = process_decisions(
        db_path, decisions, images_in, archive, api_key="k",
        upload_fn=upload_fn, fetch_fn=fetch_fn, put_fn=put_fn,
    )
    r2 = process_decisions(
        db_path, decisions, images_in, archive, api_key="k",
        upload_fn=upload_fn, fetch_fn=fetch_fn, put_fn=put_fn,
    )
    assert r1.uploaded == 1
    assert r2.uploaded == 0
    assert r2.already_processed == 1
    # INV-4: 2回目は put_fn を呼ばない（副作用なし）
    assert fake.put_count == 1
    assert fake.upload_count == 1


@pytest.mark.integration
def test_ng_decision_records_and_marks_revision(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    _seed_article(conn, "E2")
    _seed_generation(conn, "E2")
    conn.commit()
    conn.close()

    images_in = tmp_path / "images" / "in"
    images_in.mkdir(parents=True)
    (images_in / "E2.png").write_bytes(b"fake")
    archive = tmp_path / "images" / "archive"
    archive.mkdir()

    decisions = tmp_path / "dec.csv"
    _write_decisions_csv(
        decisions, [("E2", "ng", "medical_error,typo", "神経名の綴り誤り")]
    )

    fake = _make_fake_clients()
    result = process_decisions(
        db_path, decisions, images_in, archive, api_key="k",
        upload_fn=fake.upload, fetch_fn=fake.fetch, put_fn=fake.put,
    )
    assert result.ng_recorded == 1

    conn = sqlite3.connect(db_path)
    article = conn.execute(
        "SELECT status FROM articles WHERE entry_id='E2'"
    ).fetchone()
    assert article == ("ng_pending",)
    gen = conn.execute(
        "SELECT decision, ng_tags, ng_note FROM generations WHERE entry_id='E2'"
    ).fetchone()
    assert gen == ("ng", "medical_error,typo", "神経名の綴り誤り")
    assert not (images_in / "E2.png").exists()
```

- [ ] **Step 2: テスト実行して失敗確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/integration/test_process_decisions.py -v
```

Expected: `ModuleNotFoundError: No module named 'scripts.process_decisions'`

- [ ] **Step 3: process_decisions.py 実装**

`blog/infographic-project/scripts/process_decisions.py`:

```python
"""decisions.csv を読んで OK/NG/Skip を反映する処理スクリプト。

idempotencyを processed_decisions テーブルで保証。
"""

from __future__ import annotations

import csv
import re
import shutil
import sqlite3
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from scripts.hatena_client import (
    build_entry_put_xml,
    fetch_entry,
    put_entry,
    upload_to_fotolife,
)
from scripts.html_inserter import insert_infographic


@dataclass
class ProcessResult:
    uploaded: int = 0
    ng_recorded: int = 0
    skipped: int = 0
    already_processed: int = 0
    failed: int = 0


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _build_img_tag(image_url: str, alt_text: str) -> str:
    return (
        f'<img src="{image_url}" alt="{alt_text}" '
        'style="max-width:100%; height:auto; border-radius:8px; '
        'box-shadow:0 2px 8px rgba(0,0,0,0.1);" loading="lazy">'
    )


def _extract_current_body(entry_xml: str) -> tuple[str, str, tuple[str, ...]]:
    """既存entry XMLから (title, body_html, categories) を抽出。"""
    title_match = re.search(r"<title[^>]*>([^<]*)</title>", entry_xml)
    title = title_match.group(1) if title_match else ""
    cats = tuple(re.findall(r'<category[^>]*term="([^"]+)"', entry_xml))
    content_match = re.search(
        r'<content[^>]*type="text/html"[^>]*>(.*?)</content>',
        entry_xml,
        re.DOTALL,
    )
    body = ""
    if content_match:
        raw = content_match.group(1).strip()
        if raw.startswith("<![CDATA["):
            raw = raw[len("<![CDATA[") :]
            if raw.endswith("]]>"):
                raw = raw[:-3]
        body = (
            raw.replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&amp;", "&")
            .replace("&quot;", '"')
            .replace("&#39;", "'")
        )
    return title, body, cats


def process_decisions(
    db_path: Path,
    decisions_csv: Path,
    images_in_dir: Path,
    images_archive_dir: Path,
    *,
    api_key: str,
    upload_fn: Callable = upload_to_fotolife,
    fetch_fn: Callable = fetch_entry,
    put_fn: Callable = put_entry,
    dry_run: bool = False,
) -> ProcessResult:
    result = ProcessResult()
    conn = sqlite3.connect(db_path)
    csv_id = str(decisions_csv.resolve())

    try:
        with decisions_csv.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))

        for row in rows:
            eid = row["entry_id"]
            decision = row["decision"]

            # idempotency check
            already = conn.execute(
                "SELECT 1 FROM processed_decisions WHERE decisions_csv_path=? AND entry_id=?",
                (csv_id, eid),
            ).fetchone()
            if already:
                result.already_processed += 1
                continue

            try:
                if decision == "ok":
                    _handle_ok(
                        conn,
                        eid,
                        images_in_dir,
                        images_archive_dir,
                        api_key=api_key,
                        upload_fn=upload_fn,
                        fetch_fn=fetch_fn,
                        put_fn=put_fn,
                        dry_run=dry_run,
                    )
                    result.uploaded += 1
                elif decision == "ng":
                    _handle_ng(
                        conn,
                        eid,
                        row.get("ng_tags", ""),
                        row.get("ng_note", ""),
                        images_in_dir,
                        images_archive_dir,
                        dry_run=dry_run,
                    )
                    result.ng_recorded += 1
                elif decision == "skip":
                    result.skipped += 1
                else:
                    continue

                if not dry_run:
                    conn.execute(
                        """INSERT OR IGNORE INTO processed_decisions
                        (decisions_csv_path, entry_id, decision, processed_at)
                        VALUES (?, ?, ?, ?)""",
                        (csv_id, eid, decision, _now()),
                    )
                    conn.commit()
            except Exception as exc:
                result.failed += 1
                print(f"[ERROR] entry={eid}: {exc}", file=sys.stderr)
                conn.rollback()
    finally:
        conn.close()

    return result


def _handle_ok(
    conn: sqlite3.Connection,
    entry_id: str,
    images_in: Path,
    archive: Path,
    *,
    api_key: str,
    upload_fn: Callable,
    fetch_fn: Callable,
    put_fn: Callable,
    dry_run: bool,
) -> None:
    gen_row = conn.execute(
        """SELECT alt_text, fotolife_url FROM generations
        WHERE entry_id=? ORDER BY attempt DESC LIMIT 1""",
        (entry_id,),
    ).fetchone()
    if not gen_row:
        raise RuntimeError(f"generations に {entry_id} がない")
    alt_text, existing_url = gen_row

    image_path = images_in / f"{entry_id}.png"
    if not image_path.exists():
        raise FileNotFoundError(str(image_path))

    if dry_run:
        print(f"[DRY-RUN] would upload {image_path}")
        print(f"[DRY-RUN] would PUT entry {entry_id}")
        return

    # INV-1: 既アップ済みなら既URLを流用
    image_url = existing_url or upload_fn(image_path, api_key)
    if str(image_url).startswith(("ERROR", "URL not found")):
        raise RuntimeError(f"Fotolife upload failed: {image_url}")

    entry_xml = fetch_fn(entry_id, api_key)
    title, body_html, categories = _extract_current_body(entry_xml)

    # ロールバック用に変更前の本文を保存
    conn.execute(
        "UPDATE articles SET body_html_before=? WHERE entry_id=?",
        (body_html, entry_id),
    )

    img_tag = _build_img_tag(image_url, alt_text)
    new_body = insert_infographic(body_html, img_tag)

    put_xml = build_entry_put_xml(
        title=title, author="hinyan1016", body_html=new_body, categories=categories
    )
    put_fn(entry_id, put_xml, api_key)

    # 画像アーカイブ移動
    archive.mkdir(parents=True, exist_ok=True)
    shutil.move(str(image_path), str(archive / f"{entry_id}.png"))

    conn.execute(
        """UPDATE generations
        SET fotolife_url=?, processed_at=?, decision='approved'
        WHERE entry_id=?
        AND gen_id = (SELECT gen_id FROM generations WHERE entry_id=? ORDER BY attempt DESC LIMIT 1)""",
        (image_url, _now(), entry_id, entry_id),
    )
    conn.execute(
        "UPDATE articles SET status='done' WHERE entry_id=?",
        (entry_id,),
    )


def _handle_ng(
    conn: sqlite3.Connection,
    entry_id: str,
    ng_tags: str,
    ng_note: str,
    images_in: Path,
    archive: Path,
    *,
    dry_run: bool,
) -> None:
    if dry_run:
        print(f"[DRY-RUN] would record NG for {entry_id}: tags={ng_tags}")
        return

    archive.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    image_path = images_in / f"{entry_id}.png"
    if image_path.exists():
        shutil.move(str(image_path), str(archive / f"{entry_id}_ng_{ts}.png"))

    conn.execute(
        """UPDATE generations
        SET decision='ng', ng_tags=?, ng_note=?, processed_at=?
        WHERE entry_id=?
        AND gen_id = (SELECT gen_id FROM generations WHERE entry_id=? ORDER BY attempt DESC LIMIT 1)""",
        (ng_tags, ng_note, _now(), entry_id, entry_id),
    )
    conn.execute(
        "UPDATE articles SET status='ng_pending', last_ng_date=? WHERE entry_id=?",
        (_now(), entry_id),
    )


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("decisions_csv")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--yes", action="store_true")
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    images_in = root / "images" / "in"
    archive = root / "images" / "archive"

    env_file = Path(
        r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env"
    )
    from scripts.hatena_client import load_env
    api_key = load_env(env_file)["HATENA_API_KEY"]

    if not args.dry_run and not args.yes:
        ans = input("本番のはてなブログを更新します。続行しますか? (yes/no): ")
        if ans.strip().lower() not in ("yes", "y"):
            print("中止しました。")
            return 1

    result = process_decisions(
        db_path,
        Path(args.decisions_csv),
        images_in,
        archive,
        api_key=api_key,
        dry_run=args.dry_run,
    )
    print(
        f"完了: OK={result.uploaded} NG={result.ng_recorded} "
        f"Skip={result.skipped} 既処理={result.already_processed} "
        f"失敗={result.failed}"
    )
    return 0 if result.failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: テスト実行してグリーン**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/integration/test_process_decisions.py -v
```

Expected: `3 passed`

- [ ] **Step 5: コミット**

```bash
git add blog/infographic-project/scripts/process_decisions.py blog/infographic-project/tests/integration/test_process_decisions.py
git commit -m "feat(infographic): process_decisions - OK/NG/Skip反映とidempotency保証"
```

---

## Task 10: ロールバックスクリプト

**Files:**
- Create: `blog/infographic-project/scripts/rollback.py`
- Create: `blog/infographic-project/tests/integration/test_rollback.py`

- [ ] **Step 1: テストを書く**

`blog/infographic-project/tests/integration/test_rollback.py`:

```python
from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from scripts.init_db import initialize_schema
from scripts.rollback import rollback_entry


@pytest.mark.integration
def test_rollback_restores_previous_body(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    conn.execute(
        """INSERT INTO articles
        (entry_id, title, status, body_html_before)
        VALUES ('E1', 't', 'done', '<p>元の本文</p>')"""
    )
    conn.commit()
    conn.close()

    puts: list[tuple[str, str]] = []

    def fake_put(entry_id: str, xml: str, api_key: str, **_: object) -> str:
        puts.append((entry_id, xml))
        return "OK"

    def fake_fetch(entry_id: str, api_key: str, **_: object) -> str:
        return (
            '<entry xmlns="http://www.w3.org/2005/Atom">'
            '<title>t</title>'
            '<category term="cat" />'
            '<content type="text/html">&lt;p&gt;新本文&lt;/p&gt;</content>'
            '</entry>'
        )

    rollback_entry(
        db_path, "E1", api_key="k", fetch_fn=fake_fetch, put_fn=fake_put
    )
    assert len(puts) == 1
    assert "<p>元の本文</p>" in puts[0][1]


@pytest.mark.integration
def test_rollback_raises_without_backup(tmp_path: Path) -> None:
    db_path = tmp_path / "s.sqlite"
    initialize_schema(db_path)
    conn = sqlite3.connect(db_path)
    conn.execute(
        "INSERT INTO articles (entry_id, title, status) VALUES ('E2', 't', 'done')"
    )
    conn.commit()
    conn.close()

    with pytest.raises(RuntimeError, match="body_html_before"):
        rollback_entry(
            db_path, "E2", api_key="k",
            fetch_fn=lambda *a, **k: "<entry></entry>",
            put_fn=lambda *a, **k: "OK",
        )
```

- [ ] **Step 2: テスト実行、失敗確認**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/integration/test_rollback.py -v
```

Expected: `ModuleNotFoundError: No module named 'scripts.rollback'`

- [ ] **Step 3: rollback.py 実装**

`blog/infographic-project/scripts/rollback.py`:

```python
"""緊急ロールバック：誤って更新した記事を body_html_before から復元する。"""

from __future__ import annotations

import argparse
import re
import sqlite3
import sys
from pathlib import Path
from typing import Callable

from scripts.hatena_client import (
    build_entry_put_xml,
    fetch_entry,
    load_env,
    put_entry,
)


def rollback_entry(
    db_path: Path,
    entry_id: str,
    *,
    api_key: str,
    fetch_fn: Callable = fetch_entry,
    put_fn: Callable = put_entry,
) -> None:
    conn = sqlite3.connect(db_path)
    try:
        row = conn.execute(
            "SELECT body_html_before FROM articles WHERE entry_id=?",
            (entry_id,),
        ).fetchone()
        if not row or not row[0]:
            raise RuntimeError(
                f"entry_id={entry_id} には body_html_before がないのでロールバック不可"
            )
        previous_body = row[0]

        entry_xml = fetch_fn(entry_id, api_key)
        title_match = re.search(r"<title[^>]*>([^<]*)</title>", entry_xml)
        title = title_match.group(1) if title_match else ""
        cats = tuple(re.findall(r'<category[^>]*term="([^"]+)"', entry_xml))

        put_xml = build_entry_put_xml(
            title=title,
            author="hinyan1016",
            body_html=previous_body,
            categories=cats,
        )
        put_fn(entry_id, put_xml, api_key)

        # ロールバック成功後は body_html_before をNULLにして二重ロールバック防止
        conn.execute(
            """UPDATE articles
            SET status='rolled_back', body_html_before=NULL
            WHERE entry_id=?""",
            (entry_id,),
        )
        conn.commit()
    finally:
        conn.close()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("entry_id")
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    db_path = root / "data" / "state.sqlite"
    env_file = Path(
        r"C:\Users\jsber\OneDrive\Documents\Claude_task\youtube-slides\食事指導シリーズ\_shared\.env"
    )
    api_key = load_env(env_file)["HATENA_API_KEY"]

    ans = input(f"entry_id={args.entry_id} をロールバックしますか? (yes/no): ")
    if ans.strip().lower() not in ("yes", "y"):
        print("中止しました。")
        return 1

    rollback_entry(db_path, args.entry_id, api_key=api_key)
    print("ロールバック完了")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: テスト実行してグリーン**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest tests/integration/test_rollback.py -v
```

Expected: `2 passed`

- [ ] **Step 5: コミット**

```bash
git add blog/infographic-project/scripts/rollback.py blog/infographic-project/tests/integration/test_rollback.py
git commit -m "feat(infographic): rollback.py - 緊急ロールバックスクリプト"
```

---

## Task 11: 全テスト実行と Day 1 ドライラン

**Files:**
- すべて既存（テスト実行のみ、本番データ1件で検証）

- [ ] **Step 1: 全ユニット・統合テストを実行**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m pytest --cov=scripts --cov-report=term-missing
```

Expected: 全テストpass、カバレッジ unit 95%+ / integration 80%+

- [ ] **Step 2: 本番DBの状態確認**

```bash
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -c "
import sqlite3
from pathlib import Path
p = Path('blog/infographic-project/data/state.sqlite')
conn = sqlite3.connect(p)
for st, n in conn.execute('SELECT status, COUNT(*) FROM articles GROUP BY status'):
    print(f'{st}: {n}')
"
```

Expected: `pending: 227`

- [ ] **Step 3: バッチ生成を1件のみでテスト**

```bash
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -c "
from pathlib import Path
from scripts.generate_batch_csv import generate_daily_batch
root = Path.cwd()
csv_p, md_p = generate_daily_batch(
    root / 'data' / 'state.sqlite',
    root / 'data' / 'batches',
    batch_size=1,
    batch_date='2026-04-23-test'
)
print(f'CSV: {csv_p}')
print(f'MD: {md_p}')
print(csv_p.read_text(encoding='utf-8-sig'))
"
```

Expected: 1行のprompts CSV出力、1件目の古い記事が in_batch 状態へ

- [ ] **Step 4: ドライランで process_decisions を試行**

```bash
# 空のdecisions.csvでdry-run確認
echo 'entry_id,decision,ng_tags,ng_note' > /tmp/empty_decisions.csv
cd blog/infographic-project
/c/Users/jsber/AppData/Local/Programs/Python/Python313/python.exe -m scripts.process_decisions /tmp/empty_decisions.csv --dry-run
```

Expected: `完了: OK=0 NG=0 Skip=0 既処理=0 失敗=0`

- [ ] **Step 5: テスト完了をコミット**

```bash
# 実行で変わった state.sqlite は gitignored なのでコミット対象外
# このタスクで追加するファイルはない（検証のみ）
git status
# 変更が出ているか確認（カバレッジ結果やログ以外は出ないはず）
```

このタスクでは commit を追加で行わない（Task 1-10 の実装が正常であることの検証のみ）。

---

## Task 12: 運用Runbook作成

**Files:**
- Create: `blog/infographic-project/RUNBOOK.md`

- [ ] **Step 1: RUNBOOK.md を作成**

`blog/infographic-project/RUNBOOK.md`:

```markdown
# インフォグラフィック・バックフィル運用手順

## 日次ルーチン（15〜20分）

### ① バッチ生成（2分）

Claude Code セッションで:

> 「今日の20件のプロンプト生成して」

Claude が実行するコマンド:
```bash
cd blog/infographic-project
python -m scripts.generate_batch_csv
# data/entries_for_claude/<today>.md を読む
# 各記事を解析してプロンプト・alt・style_summaryを作成
# data/batches/<today>_prompts.csv を上書き保存
```

### ② ChatGPT画像生成（5〜10分）

1. `data/batches/<today>_prompts.csv` をExcelで開く
2. 1行ずつ `prompt` 列をコピーして ChatGPT UI（画像生成モード）へ貼り付け
3. 生成画像を右クリック→保存、`blog/infographic-project/images/in/<entry_id>.png` にリネーム
4. 20件全部保存できたらファイルマネージャで件数確認

### ③ ダッシュボードレビュー（5〜7分）

> 「ダッシュボード作って」

Claude が実行:
```bash
python -m scripts.build_dashboard
```

1. 生成された `data/reviews/<today>_review.html` をブラウザで開く
2. 各カードで OK/NG/スキップを選択
3. NGの場合は理由タグと自由記述を入力
4. 下部の「決定をCSVにエクスポート」を押す
5. ダウンロードされた CSV を `data/decisions/<today>_decisions.csv` として保存

### ④ 反映（3分）

> 「決定を反映して」

Claude が実行:
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

## トラブル対応

### 画像アップロードが失敗する
- ログ `logs/errors.log` を確認
- Fotolife API の一時障害の可能性、30分後に再試行
- 同じdecisions.csvを再実行可能（idempotency保証）

### 誤って更新してしまった
```bash
python -m scripts.rollback <entry_id>
```
※ `body_html_before` が保存されている7日以内のみ有効

### NGが3回繰り返された記事
`status='manual_intervention'` に自動変更されるので、手動でChatGPT生成＋既存スクリプトで対応

## 定期メンテナンス

### 週1回
- `logs/errors.log` をgrepして ERROR / WARN の傾向確認
- NG率が高いパターン（特定カテゴリなど）がないかチェック

### プロジェクト完了時
- `SELECT COUNT(*) WHERE status='done'` = 227 を確認
- 全エラーログを棚卸し
- 最終レポートを作成してコミット
```

- [ ] **Step 2: RUNBOOK を確認**

```bash
cat blog/infographic-project/RUNBOOK.md | head -30
```

- [ ] **Step 3: コミット**

```bash
git add blog/infographic-project/RUNBOOK.md
git commit -m "docs(infographic): 日次運用手順書 RUNBOOK.md を追加"
```

---

## 実装完了後のDay 1チェックリスト（本番1件目）

Task 1-12 完了後、**本番の1件目を実際に処理する前**に以下を順に実施:

1. [ ] `git status` でクリーン（未コミット変更なし）
2. [ ] `pytest --cov=scripts` 全passかつカバレッジ目標達成
3. [ ] `python -m scripts.generate_batch_csv` — 20件のCSV生成確認
4. [ ] `entries_for_claude/<today>.md` をClaudeが読み、プロンプト・alt・style_summary を埋めて `<today>_prompts.csv` を保存
5. [ ] ChatGPT UIで1件だけ先に生成テスト（1行目の記事）
6. [ ] `images/in/<entry_id>.png` に配置
7. [ ] `python -m scripts.build_dashboard` — ダッシュボード生成
8. [ ] ブラウザで1件のカード表示を確認
9. [ ] この1件にOK判定して CSV をエクスポート
10. [ ] `python -m scripts.process_decisions <path> --dry-run` — 変更内容確認
11. [ ] `python -m scripts.process_decisions <path>` — yesで実行
12. [ ] ブラウザではてなブログの記事を開き、インフォグラフィックが表示されることを確認
13. [ ] 記事本文に破壊がないか（目次・本文・参考文献などが保持されているか）を確認
14. [ ] OKであれば、残り19件についてもChatGPTで生成→同じフローで処理
15. [ ] Day 2 から20件/日ペースで運用開始

---

## 補足：完了判定

全227件の処理完了は以下で判定:

```bash
python -c "
import sqlite3
conn = sqlite3.connect('blog/infographic-project/data/state.sqlite')
done = conn.execute(\"SELECT COUNT(*) FROM articles WHERE status='done'\").fetchone()[0]
assert done == 227, f'{done}/227'
print('プロジェクト完了')
"
```

完了後のアーカイブ:
- `data/state.sqlite` のスナップショットを `data/state_final.sqlite` にコピー
- `images/archive/` の全画像（Fotolife同期済み）
- `logs/` の全ログ
- これらを `blog/infographic-project/COMPLETION_REPORT.md` にまとめてコミット
