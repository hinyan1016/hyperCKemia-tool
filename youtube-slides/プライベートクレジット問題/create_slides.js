var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "プライベートクレジット問題とは？ 3兆ドル市場のリスクをわかりやすく解説";

// --- Color Palette (金融テーマ: ダークネイビー + レッドアクセント) ---
var C = {
  dark: "0D1B2A",
  primary: "1B3A5C",
  accent: "C0392B",
  accentLight: "E74C3C",
  light: "F0F4F8",
  warmBg: "F5F7FA",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "C0392B",
  orange: "E67E22",
  yellow: "F5C518",
  green: "27AE60",
  blue: "2C5AA0",
  gold: "D4A017",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var FONT_H = FONT_JP;
var FONT_B = FONT_JP;

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
}

function addCard(s, x, y, w, h, fillColor) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: fillColor || C.white }, shadow: makeShadow(), rectRadius: 0.08 });
}

function addPageNum(s, num) {
  s.addText(num, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_B, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("金融リスク解説シリーズ", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("プライベートクレジット問題\nとは何か？", {
    x: 0.8, y: 1.1, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 3兆ドル市場のリスクをわかりやすく解説 ―", {
    x: 0.8, y: 2.8, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.6, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.9, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: プライベートクレジットとは？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "プライベートクレジットとは？");

  // 定義ボックス
  addCard(s, 0.5, 1.1, 9.0, 1.2, C.light);
  s.addText("銀行を介さず、投資ファンドが企業に直接融資する仕組み", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.0,
    fontSize: 20, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // フロー図
  // 投資家
  addCard(s, 0.5, 2.7, 2.4, 1.4, C.blue);
  s.addText("投資家", { x: 0.5, y: 2.8, w: 2.4, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText("年金基金\n保険会社\n富裕層", { x: 0.5, y: 3.2, w: 2.4, h: 0.8, fontSize: 13, fontFace: FONT_JP, color: C.light, align: "center", margin: 0, lineSpacingMultiple: 1.1 });

  // 矢印1
  s.addText("出資 →", { x: 3.0, y: 3.0, w: 1.2, h: 0.8, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0 });

  // ファンド
  addCard(s, 4.0, 2.7, 2.2, 1.4, C.accent);
  s.addText("PC ファンド", { x: 4.0, y: 2.8, w: 2.2, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText("ブラックストーン\nアポロ\nブルー・アウル 等", { x: 4.0, y: 3.2, w: 2.2, h: 0.8, fontSize: 12, fontFace: FONT_JP, color: C.light, align: "center", margin: 0, lineSpacingMultiple: 1.1 });

  // 矢印2
  s.addText("融資 →", { x: 6.3, y: 3.0, w: 1.2, h: 0.8, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0 });

  // 企業
  addCard(s, 7.3, 2.7, 2.2, 1.4, C.primary);
  s.addText("借り手企業", { x: 7.3, y: 2.8, w: 2.2, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText("中堅・非上場企業\n（銀行融資が困難）", { x: 7.3, y: 3.2, w: 2.2, h: 0.8, fontSize: 12, fontFace: FONT_JP, color: C.light, align: "center", margin: 0, lineSpacingMultiple: 1.1 });

  // 特徴
  var features = [
    { icon: "🔒", text: "非公開（市場で取引されない）" },
    { icon: "💰", text: "高金利（8〜12%）" },
    { icon: "⏳", text: "低流動性（3〜7年ロック）" },
    { icon: "📋", text: "規制が緩い" },
  ];
  features.forEach(function(f, i) {
    var xPos = 0.5 + i * 2.35;
    s.addText(f.icon + " " + f.text, {
      x: xPos, y: 4.5, w: 2.2, h: 0.7,
      fontSize: 12, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0,
    });
  });

  addPageNum(s, "2/10");
})();

// ============================================================
// SLIDE 3: 市場の急成長
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "爆発的に成長した市場規模");

  // 成長バー（疑似チャート）
  var data = [
    { year: "2000年頃", val: "数百億$", w: 0.5, color: C.blue },
    { year: "2023年", val: "2.1兆$", w: 3.5, color: C.blue },
    { year: "2025年", val: "3〜3.4兆$", w: 5.5, color: C.orange },
    { year: "2029年\n(予測)", val: "4.9兆$", w: 8.0, color: C.accent },
  ];
  data.forEach(function(d, i) {
    var yPos = 1.2 + i * 1.0;
    s.addText(d.year, { x: 0.3, y: yPos, w: 1.5, h: 0.7, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "right", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 2.0, y: yPos + 0.1, w: d.w, h: 0.5, fill: { color: d.color }, rectRadius: 0.05 });
    s.addText(d.val, { x: 2.0 + d.w + 0.15, y: yPos, w: 1.5, h: 0.7, fontSize: 14, fontFace: FONT_EN, color: C.text, bold: true, valign: "middle", margin: 0 });
  });

  // 成長の背景
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9.0, h: 0.5, fill: { color: C.light } });
  s.addText("25年間で数十倍  ―  IMF「監視の強化が必要」(2024年4月)", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "3/10");
})();

// ============================================================
// SLIDE 4: 成長の3つの背景
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "なぜここまで巨大化したのか？");

  var items = [
    { num: "1", title: "リーマン後の銀行規制強化", desc: "バーゼルIII等で銀行がリスク融資を縮小\n→ その「空白」をファンドが埋めた", color: C.blue },
    { num: "2", title: "長期ゼロ金利と利回り追求", desc: "年金・生保が目標利回り達成のため\n高リターンのPCに資金を集中", color: C.orange },
    { num: "3", title: "セミリキッド型で個人にも開放", desc: "四半期5%まで解約可能な商品で\n富裕層マネーが大量流入", color: C.accent },
  ];
  items.forEach(function(item, i) {
    var yPos = 1.2 + i * 1.35;
    addCard(s, 0.5, yPos, 9.0, 1.15);
    // 番号
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.2, w: 0.7, h: 0.7, fill: { color: item.color } });
    s.addText(item.num, { x: 0.7, y: yPos + 0.2, w: 0.7, h: 0.7, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // テキスト
    s.addText(item.title, { x: 1.6, y: yPos + 0.1, w: 7.6, h: 0.4, fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true, margin: 0 });
    s.addText(item.desc, { x: 1.6, y: yPos + 0.5, w: 7.6, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0, lineSpacingMultiple: 1.1 });
  });

  addPageNum(s, "4/10");
})();

// ============================================================
// SLIDE 5: 3つの構造リスク
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "3つの構造リスク ― 問題の本質");

  var risks = [
    { icon: "👁️", title: "評価の不透明性", desc: "非上場 → 市場価格なし\n損失の表面化が遅れる\n表面デフォルト率<2% vs 実質約5%", color: C.accent },
    { icon: "💧", title: "流動性ミスマッチ", desc: "「四半期5%解約OK」のはずが\n融資先は3〜7年満期\n解約殺到時に取り付けリスク", color: C.orange },
    { icon: "🏦", title: "銀行との相互依存", desc: "米銀→ノンバンク融資 約3,000億$\nJPM単独で1,600億$（3倍増）\nファンド破綻 → 銀行に波及", color: C.accent },
  ];
  risks.forEach(function(r, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.2, 3.0, 3.8);
    // アイコン
    s.addText(r.icon, { x: xPos, y: 1.3, w: 3.0, h: 0.6, fontSize: 28, align: "center", valign: "middle", margin: 0 });
    // タイトル帯
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.1, y: 2.0, w: 2.8, h: 0.5, fill: { color: r.color }, rectRadius: 0.05 });
    s.addText(r.title, { x: xPos + 0.1, y: 2.0, w: 2.8, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // 説明
    s.addText(r.desc, { x: xPos + 0.2, y: 2.7, w: 2.6, h: 2.2, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3 });
  });

  addPageNum(s, "5/10");
})();

// ============================================================
// SLIDE 6: 時系列 — 何が起きたか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "2025〜2026年 何が起きたか");

  // タイムライン縦線
  s.addShape(pres.shapes.LINE, { x: 2.0, y: 1.1, w: 0, h: 4.2, line: { color: C.accent, width: 3 } });

  var events = [
    { date: "2025年10月", text: "First Brands・Tricolor 破綻\nダイモンCEO「まだゴキブリが出てくる」", y: 1.1 },
    { date: "2026年2月", text: "Blue Owl 解約停止\n英MFS 経営破綻", y: 2.15 },
    { date: "2026年2-3月", text: "BCRED 解約率7.9%（過去最大）\n月次損失 -0.4%（3年超ぶり）", y: 3.2 },
    { date: "2026年3月", text: "BDC株価 → NAVの80%に下落\n規制強化の議論が本格化", y: 4.25 },
  ];
  events.forEach(function(e) {
    // ドット
    s.addShape(pres.shapes.OVAL, { x: 1.75, y: e.y + 0.2, w: 0.5, h: 0.5, fill: { color: C.accent } });
    // 日付
    s.addText(e.date, { x: 0.2, y: e.y + 0.15, w: 1.5, h: 0.5, fontSize: 12, fontFace: FONT_JP, color: C.accent, bold: true, align: "right", valign: "middle", margin: 0 });
    // カード
    addCard(s, 2.6, e.y, 6.9, 0.9);
    s.addText(e.text, { x: 2.8, y: e.y + 0.05, w: 6.5, h: 0.8, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "6/10");
})();

// ============================================================
// SLIDE 7: リーマンとの比較
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "「第2のリーマン」になるのか？");

  // テーブルヘッダー
  var colW = [2.5, 3.0, 3.0];
  var colX = [0.7, 0.7 + colW[0], 0.7 + colW[0] + colW[1]];

  s.addShape(pres.shapes.RECTANGLE, { x: colX[0], y: 1.1, w: colW[0], h: 0.5, fill: { color: C.primary } });
  s.addShape(pres.shapes.RECTANGLE, { x: colX[1], y: 1.1, w: colW[1], h: 0.5, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: colX[2], y: 1.1, w: colW[2], h: 0.5, fill: { color: C.primary } });

  s.addText("比較項目", { x: colX[0], y: 1.1, w: colW[0], h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("2008 サブプライム", { x: colX[1], y: 1.1, w: colW[1], h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("2026 PC問題", { x: colX[2], y: 1.1, w: colW[2], h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var rows = [
    ["規模", "1.5兆ドル", "2〜3兆ドル"],
    ["構造の複雑さ", "CDO多重証券化\nリスク所在不明", "直接融資主流\n比較的シンプル"],
    ["レバレッジ", "30〜40倍（極端）", "相対的に低い"],
    ["銀行の備え", "引当金不十分", "現時点で大幅\n積み増しなし"],
    ["規制", "格付AAA乱発\n監視機能不全", "IMF・中央銀行が\n早期警戒中"],
  ];
  rows.forEach(function(row, i) {
    var yPos = 1.65 + i * 0.7;
    var bgColor = (i % 2 === 0) ? C.light : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: colX[0], y: yPos, w: colW[0] + colW[1] + colW[2], h: 0.7, fill: { color: bgColor } });
    s.addText(row[0], { x: colX[0] + 0.1, y: yPos, w: colW[0] - 0.2, h: 0.7, fontSize: 12, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(row[1], { x: colX[1] + 0.1, y: yPos, w: colW[1] - 0.2, h: 0.7, fontSize: 11, fontFace: FONT_JP, color: C.accent, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
    s.addText(row[2], { x: colX[2] + 0.1, y: yPos, w: colW[2] - 0.2, h: 0.7, fontSize: 11, fontFace: FONT_JP, color: C.primary, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  // 結論
  s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 5.15, w: 8.5, h: 0.45, fill: { color: C.light } });
  s.addText("現時点で直ちにリーマン級の危機に発展する可能性は低い ― ただし未検証", {
    x: 0.7, y: 5.15, w: 8.5, h: 0.45,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "7/10");
})();

// ============================================================
// SLIDE 8: 日本への影響
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "日本への影響は？");

  // 機関投資家
  addCard(s, 0.5, 1.2, 4.3, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.45, fill: { color: C.primary }, rectRadius: 0.08 });
  s.addText("🏢 機関投資家（生保・年金）", { x: 0.5, y: 1.2, w: 4.3, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("• 大手生保がPC投資を拡大してきた\n• 破綻案件にエクスポージャーあり\n• 現在は対象を厳選する方針に転換", {
    x: 0.7, y: 1.8, w: 3.9, h: 1.3, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // 個人投資家
  addCard(s, 5.2, 1.2, 4.3, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.45, fill: { color: C.accent }, rectRadius: 0.08 });
  s.addText("👤 個人投資家への間接影響", { x: 5.2, y: 1.2, w: 4.3, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("• 株式市場 ― 波及売りリスク\n• 金融株 ― 銀行の融資損失\n• 円相場 ― リスクオフで円高\n• 生命保険 ― 運用損失の影響", {
    x: 5.4, y: 1.8, w: 3.9, h: 1.3, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // 影響フロー
  addCard(s, 0.5, 3.5, 9.0, 1.8);
  s.addText("影響の波及経路", { x: 0.5, y: 3.55, w: 9.0, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0 });

  var flowItems = [
    { text: "米PC市場\nストレス", color: C.accent, x: 0.8 },
    { text: "→", color: "", x: 2.5 },
    { text: "米銀損失\n金融株下落", color: C.orange, x: 3.0 },
    { text: "→", color: "", x: 4.7 },
    { text: "リスクオフ\n円高圧力", color: C.orange, x: 5.2 },
    { text: "→", color: "", x: 6.9 },
    { text: "日本株\n生保運用", color: C.primary, x: 7.4 },
  ];
  flowItems.forEach(function(f) {
    if (f.color === "") {
      s.addText(f.text, { x: f.x, y: 4.1, w: 0.5, h: 0.9, fontSize: 18, fontFace: FONT_JP, color: C.sub, align: "center", valign: "middle", margin: 0 });
    } else {
      s.addShape(pres.shapes.RECTANGLE, { x: f.x, y: 4.15, w: 1.6, h: 0.8, fill: { color: f.color }, rectRadius: 0.08 });
      s.addText(f.text, { x: f.x, y: 4.15, w: 1.6, h: 0.8, fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
    }
  });

  addPageNum(s, "8/10");
})();

// ============================================================
// SLIDE 9: 今後の3つのシナリオ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "今後のシナリオ ― 3つのケース");

  var scenarios = [
    { title: "ソフトランディング", prob: "やや高い", desc: "個別破綻は続くが波及は限定的\n規制強化で市場が「成熟」へ", color: C.green, y: 1.2 },
    { title: "中規模ストレス", prob: "一定の可能性", desc: "景気減速＋複数ファンドで解約停止連鎖\n金融株全体に売り圧力", color: C.orange, y: 2.65 },
    { title: "システミック危機", prob: "現時点では低い", desc: "大手ファンド破綻 → 銀行の連鎖損失\n深刻な景気後退が必要条件", color: C.accent, y: 4.1 },
  ];
  scenarios.forEach(function(sc) {
    addCard(s, 0.5, sc.y, 9.0, 1.2);
    // シナリオラベル
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: sc.y + 0.15, w: 2.4, h: 0.9, fill: { color: sc.color }, rectRadius: 0.08 });
    s.addText(sc.title, { x: 0.7, y: sc.y + 0.15, w: 2.4, h: 0.55, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(sc.prob, { x: 0.7, y: sc.y + 0.65, w: 2.4, h: 0.35, fontSize: 11, fontFace: FONT_JP, color: C.white, align: "center", valign: "middle", margin: 0 });
    // 説明
    s.addText(sc.desc, { x: 3.4, y: sc.y + 0.1, w: 5.8, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });
  });

  addPageNum(s, "9/10");
})();

// ============================================================
// SLIDE 10: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("まとめ", {
    x: 0.8, y: 0.3, w: 8.4, h: 0.6,
    fontSize: 26, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var points = [
    "プライベートクレジット = ファンドが企業に直接融資する約3兆ドル市場",
    "3大リスク: 評価の不透明性 / 流動性ミスマッチ / 銀行との相互依存",
    "2025年末〜2026年: 破綻・解約急増で脆弱性が表面化",
    "現時点で「第2のリーマン」の可能性は低いが、景気後退なら要注意",
    "日本: 生保・年金を通じた間接影響 / 株式・為替への波及に注意",
  ];
  points.forEach(function(p, i) {
    var yPos = 1.2 + i * 0.85;
    addCard(s, 0.8, yPos, 8.4, 0.65, C.primary);
    s.addText((i + 1) + ".", { x: 0.9, y: yPos + 0.05, w: 0.5, h: 0.55, fontSize: 18, fontFace: FONT_EN, color: C.accent, bold: true, valign: "middle", margin: 0 });
    s.addText(p, { x: 1.4, y: yPos + 0.05, w: 7.5, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 5.5, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("パニックは不要 ― でも目を離さないこと", {
    x: 0.8, y: 5.55, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 6.2, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  addPageNum(s, "10/10");
})();

// ============================================================
// 出力
// ============================================================
var outPath = __dirname + "/private_credit_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
