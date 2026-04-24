"""Codex (VS Code) に渡すタスク指示書を生成する。

generate_batch_csv でCSVを作り、Claudeがプロンプトを記入した後に実行。
出力: data/batches/<date>_codex_instructions.md
これを VS Code の Codex タブにドラッグするか、内容をコピペして依頼する。
"""

from __future__ import annotations

import csv
import sys
from datetime import date
from pathlib import Path


HEADER = """# Codex向けタスク指示 — {batch_date} 画像生成（{count}件）

## 目的
各エントリの**参考記事URLを精読**したうえで、その記事内容を正確に反映した
日本語インフォグラフィック画像を1枚ずつ生成し、指定パスに保存してください。
単にプロンプトをそのまま画像化するのではなく、**記事の要点を咀嚼したうえで最適なビジュアル構成を設計する**ことが目的です。

## 情報の優先順位（必ず守ること）
1. **参考記事の本文内容**（最優先・真実の情報源）
2. **下記プロンプト本文**（記事読了を前提にした構成指針）
3. 配色・アイコン等の装飾指定（仕上げ）

プロンプトの記述と記事本文に矛盾がある場合は **記事本文を優先** してください。
プロンプトに薬剤承認年・製品名・数値などの事実誤認が残っていても、記事本文や公式情報の方が
正しければそちらに合わせて **修正して生成** すること。
（過去バッチで「日本未承認」と書いた薬剤が実際は承認済みだった等の事故があったため）

## 参考記事の精読（必須・最優先プロセス）
各エントリには `**参考記事URL（必読・最優先）**` が付与されています。
**画像生成の前に**、必ず以下の4ステップを実行してください。

1. **アクセス**: 参考記事URLを開き、本文を最後まで読む（見出し・本文・箇条書きすべて）
2. **抽出**: 固有名詞（薬剤名／承認年／基準値／ガイドライン年／疾患名）と中心メッセージを抽出
3. **設計**: 記事内容を踏まえて、9:16縦長の4段構成に最適なトピック配分を決める
   （プロンプトの分割と記事の構成が食い違う場合、記事側の構成を優先して差し替えてよい）
4. **補正**: プロンプトと記事の齟齬や不足情報を補正してから画像化する
   （プロンプトは Claude が記事を読んで作った「補助指針」であり、絶対的な命令ではない）

**記事にアクセスできない場合**（404／非公開／権限エラー等）は、その `entry_id` は
**生成せず飛ばし**、最後の完了報告に「アクセス不可」として記載してください（推測で作らない）。

## 保存先（ワークスペース内）
```
blog/infographic-project/images/in/<entry_id>.png
```

フルパス例:
```
C:\\Users\\jsber\\OneDrive\\Documents\\Claude_task\\blog\\infographic-project\\images\\in\\<entry_id>.png
```

## 命名規約
- ファイル名: `<entry_id>.png` （各セクションの「保存ファイル名」をそのまま使用）
- 拡張子必ず `.png`
- 既にファイルが存在する場合はスキップ（上書きしないこと）

## 画像仕様（重要）
- **アスペクト比: 9:16（縦長ポートレート・スマホ/ショート動画対応）**
- 推奨解像度: 1080x1920 以上（生成エンジンの9:16ポートレート最大サイズを使用）
- 正方形（1024x1024）では生成しない — 後から縦長にリサイズするとレイアウトが崩れるため最初から9:16で生成すること
- **レイアウトの再解釈ルール**: プロンプト内の「左上/右上/左下/右下」は2×2グリッドを意味するが、9:16縦長では
  `上段(①左上)→中上段(②右上)→中下段(③左下)→下段(④右下)` の4段縦積みに再配置する
- 中央のタイトル帯は上部または各段の間に配置。日本語テキストの可読性を優先し、余白を十分に取る

## 医学画像の制約（最重要・必読）
AIによる医学画像の **新規生成は厳禁** とする（合成図は臨床的に不正確になりやすいため）。
ただし、**原典画像の引用は問題ない**（ブログ本文側でリンク・埋め込み・後工程で別挿入する想定）。
このインフォグラフィック内で **AI が描くこと自体が NG** であって、医学画像の利用そのものを禁じるものではない。

### ✗ 生成NG（AIで描画してはいけないもの） vs ✓ 引用OK（別工程・ブログ本文で処理）

| ✗ 生成NG（AI描画禁止） | ✓ 引用OK（別経路で処理） |
|---|---|
| 人体解剖図・臓器イラスト（脳実質・心臓内腔・血管走行・筋肉・神経経路等） | 教科書・ガイドラインの原典図の引用 |
| MRI/CT/XP/エコー等の模擬画像 | 実画像のブログ本文埋め込み（別工程） |
| 組織像・病理像・細胞構造・分子構造の詳細描画 | 原著論文／公式データベース（PubChem等）の図参照 |
| 心電図・脳波などの医用波形 | 実波形のスクリーンショット引用（別工程） |

**プロンプトに「〇〇の模式図」「解剖図」「断面図」「MRI所見」「波形」などの指示が残っていても
解剖的・医用画像は一切描かないこと**。本ルールがプロンプトの記述より優先される。

### 代替表現（必ずこちらを使う）
- 幾何学的ピクトグラム（丸角シルエットのハート／頭部アイコン程度のシンプル抽象化まで）
- 文字・数字・キーワード主体のカード／パネル
- 色分けブロック、矢印、フローチャート、比較表、棒グラフ／円グラフ
- 薬剤カプセル／錠剤・カレンダー／時計・チェックリストなど非解剖的アイコン

**医学画像が本当に必要と判断される場合は生成せず、このインフォグラフィックには含めない**
（ブログ本文側で原典画像を引用する前提で省略する）。

テキスト情報（疾患名／薬剤名／数値／分類）を主役にし、視覚要素は装飾に留めること。
薬事情報（承認年・製品名・適応）は、プロンプトではなく **参考記事または公式情報** に一致させる。

## 失敗時の扱い
- 画像生成が拒否された／レート制限に達した場合はその `entry_id` を飛ばし、
  最後に「未生成リスト」として報告してください
- **記事アクセス不可** の場合も同様にスキップして報告（推測で作らない）
- 処理継続が最優先。1件失敗しても他の生成を止めないこと

---

## 生成指示

"""

FOOTER = """---

## 完了報告フォーマット

```
- 成功: {{N}} 件
- 失敗: {{M}} 件
  - <entry_id>: <理由>
  - ...
```

完了したら Claude Code セッションに「Codex画像保存完了」と伝えてください。
Claude Code が build_dashboard 以降を自動で進めます。
"""


def build_instructions(rows: list[dict[str, str]], batch_date: str) -> str:
    parts = [HEADER.format(batch_date=batch_date, count=len(rows))]
    for i, row in enumerate(rows, 1):
        eid = row["entry_id"].strip()
        title = row["title"].strip()
        url = row.get("url", "").strip()
        prompt = row["prompt"].strip()
        if not prompt:
            continue

        parts.append(f"### #{i}: {title}\n")
        parts.append(f"- **保存ファイル名**: `{eid}.png`\n")
        if url:
            parts.append(f"- **参考記事URL（必読・最優先）**: {url}\n")
            parts.append(
                "  → まずこの記事を精読し、要点を抽出してから画像を設計してください。"
                "プロンプトはあくまで補助指針であり、記事本文と矛盾する場合は記事側を優先。\n"
            )
        else:
            parts.append(
                "- **参考記事URLなし**: この entry_id は生成をスキップし、"
                "完了報告に「URL欠落」として記載してください（推測で作らない）。\n"
            )
        parts.append(
            "- **プロンプト（記事精読の上で構成指針として使用／医学画像の新規描画は厳禁）**:\n"
        )
        parts.append("\n")
        parts.append("```\n")
        parts.append(prompt + "\n")
        parts.append("```\n")
        parts.append("\n")

    parts.append(FOOTER)
    return "".join(parts)


def export(prompts_csv: Path, out_md: Path, *, batch_date: str) -> Path:
    with prompts_csv.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    out_md.parent.mkdir(parents=True, exist_ok=True)
    out_md.write_text(build_instructions(rows, batch_date), encoding="utf-8")
    return out_md


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default=None, help="YYYY-MM-DD (default: today)")
    args = parser.parse_args()

    batch_date = args.date or date.today().isoformat()
    root = Path(__file__).resolve().parent.parent
    prompts_csv = root / "data" / "batches" / f"{batch_date}_prompts.csv"
    if not prompts_csv.exists():
        print(f"[ERROR] バッチCSVが見つかりません: {prompts_csv}", file=sys.stderr)
        return 2
    out_md = root / "data" / "batches" / f"{batch_date}_codex_instructions.md"
    export(prompts_csv, out_md, batch_date=batch_date)
    print(f"Codex指示書: {out_md}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
