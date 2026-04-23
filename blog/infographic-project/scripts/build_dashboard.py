"""レビューダッシュボードHTML生成。"""

from __future__ import annotations

import csv
import html
import os
import sys
from datetime import date
from pathlib import Path


HTML_HEAD = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>インフォグラフィック レビュー {batch_date}</title>
<style>
body {{ font-family: -apple-system, "Hiragino Sans", sans-serif; background: #f5f5f5; margin: 0; padding: 20px; }}
.card {{ background: white; border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
.card[data-status="awaiting_image"] {{ opacity: 0.5; background: #fffbf0; }}
.card h3 {{ margin-top: 0; }}
.meta {{ color: #888; font-size: 0.9em; margin-bottom: 15px; }}
.grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
.article-summary {{ border-left: 4px solid #2E86C1; padding-left: 15px; }}
.generated-image img {{ max-width: 100%; height: auto; border: 1px solid #ddd; }}
.decision {{ margin-top: 15px; padding-top: 15px; border-top: 1px solid #eee; }}
.decision label {{ margin-right: 15px; }}
.ng-tags {{ display: none; margin-top: 10px; padding: 10px; background: #fff3cd; border-radius: 4px; }}
.ng-tags.visible {{ display: block; }}
.ng-tags label {{ display: inline-block; margin: 3px; }}
.ng-tags textarea {{ width: 100%; margin-top: 8px; min-height: 50px; }}
#export-button {{ position: sticky; bottom: 20px; padding: 14px 28px; background: #2E86C1; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }}
#summary {{ font-weight: bold; margin: 20px 0; }}
</style>
</head>
<body>
<h1>インフォグラフィック レビュー — {batch_date}</h1>
<div id="summary">対象: {count}件</div>
"""

HTML_TAIL = """
<button id="export-button" onclick="exportDecisions()">決定をCSVにエクスポート</button>
<script>
document.querySelectorAll('input[type=radio][value=ng]').forEach(r => {{
  r.addEventListener('change', e => {{
    const card = e.target.closest('.card');
    card.querySelector('.ng-tags').classList.add('visible');
  }});
}});
document.querySelectorAll('input[type=radio]:not([value=ng])').forEach(r => {{
  r.addEventListener('change', e => {{
    const card = e.target.closest('.card');
    card.querySelector('.ng-tags').classList.remove('visible');
  }});
}});

function exportDecisions() {{
  const rows = [['entry_id', 'decision', 'ng_tags', 'ng_note']];
  document.querySelectorAll('.card').forEach(card => {{
    const eid = card.dataset.entryId;
    const dec = card.querySelector('input[type=radio]:checked');
    if (!dec) return;
    const tags = Array.from(
      card.querySelectorAll('.ng-tags input[type=checkbox]:checked')
    ).map(c => c.value).join(',');
    const note = (card.querySelector('.ng-tags textarea')?.value || '').replace(/\\n/g, ' ');
    rows.push([eid, dec.value, tags, note]);
  }});
  const csv = rows.map(r => r.map(v => `"${{String(v).replace(/"/g, '""')}}"`).join(',')).join('\\n');
  const blob = new Blob(["\\uFEFF" + csv], {{type: 'text/csv'}});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = '{batch_date}_decisions.csv';
  a.click();
}}
</script>
</body>
</html>
"""


def _card_html(
    row: dict[str, str], image_available: bool, image_rel_path: str
) -> str:
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
    out_html.parent.mkdir(parents=True, exist_ok=True)
    for row in rows:
        eid = row["entry_id"]
        image_path = images_dir / f"{eid}.png"
        try:
            rel_path = Path(
                os.path.relpath(image_path.resolve(), out_html.parent.resolve())
            ).as_posix()
        except ValueError:
            # different drives: fallback to absolute
            rel_path = image_path.resolve().as_posix()
        cards.append(_card_html(row, image_path.exists(), rel_path))

    body = "".join(cards)
    head = HTML_HEAD.format(batch_date=batch_date, count=len(rows))
    tail = HTML_TAIL.format(batch_date=batch_date)

    out_html.write_text(head + body + tail, encoding="utf-8")


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--date",
        default=None,
        help="YYYY-MM-DD (default: today)。指定日のバッチCSV/ダッシュボードを処理。",
    )
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    batch_date = args.date or date.today().isoformat()
    prompts_csv = root / "data" / "batches" / f"{batch_date}_prompts.csv"
    images_dir = root / "images" / "in"
    out_html = root / "data" / "reviews" / f"{batch_date}_review.html"

    if not prompts_csv.exists():
        print(f"[ERROR] バッチCSVが見つかりません: {prompts_csv}", file=sys.stderr)
        return 2

    build_dashboard(prompts_csv, images_dir, out_html, batch_date=batch_date)
    print(f"ダッシュボード: {out_html}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
