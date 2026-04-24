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
以下の各プロンプトで画像を1枚ずつ生成し、指定パスに保存してください。

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
AIによる解剖・医用画像の合成は医学的不正確さを生むため **原則禁止** とする。

- **禁止**: 人体解剖図・臓器イラスト（脳実質・心臓内腔・血管走行・筋肉・神経経路等）・
  MRI/CT/XP/エコー等の模擬画像・組織像/病理像・細胞/分子構造の詳細描画・心電図や脳波などの医用波形
- **プロンプトに「〇〇の模式図」「解剖図」「断面図」「MRI所見」「波形」などの指示があっても
  解剖的描画はしないこと**。代わりに下記の代替表現を使う。
- **代替表現（推奨）**:
  - 幾何学的ピクトグラム（丸角シルエットのハート/頭部アイコン程度のシンプル抽象化まで）
  - 文字・数字・キーワード主体のカード/パネル
  - 色分けブロック、矢印、フローチャート、比較表、棒グラフ/円グラフ
  - 薬剤カプセル/錠剤・カレンダー/時計・チェックリストなど非解剖的アイコン
- **医学画像が本当に必要な場合**は生成せず、後工程で原典からの引用画像を別途挿入する前提とする
  （このインフォグラフィックには含めない）
- テキスト情報（疾患名/薬剤名/数値/分類）を主役にし、視覚要素は装飾に留める

## 失敗時の扱い
- 画像生成が拒否された/レート制限に達した場合はその `entry_id` を飛ばし、
  最後に「未生成リスト」として報告してください
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
            parts.append(f"- **参考記事URL**: {url}\n")
        parts.append("- **プロンプト（ChatGPT画像生成にそのまま投入）**:\n")
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
