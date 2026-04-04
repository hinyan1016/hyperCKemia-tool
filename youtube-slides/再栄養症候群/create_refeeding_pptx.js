var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "低栄養患者に「どんどん栄養を入れる」のは正しいか";

// --- Color Palette ---
var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "E8850C",
  light: "E8F4FD",
  warmBg: "F8F9FA",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "DC3545",
  green: "28A745",
  yellow: "F5C518",
  lightRed: "F8D7DA",
  lightYellow: "FFF3CD",
  lightGreen: "D4EDDA",
};

var FJ = "BIZ UDPGothic";
var FE = "Segoe UI";

function shd() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

// helper: header bar
function hdr(s, title, bgColor) {
  bgColor = bgColor || C.primary;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: bgColor } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, fontFace: FJ, color: C.white, bold: true, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("低栄養患者に「どんどん栄養を入れる」のは正しいか", {
    x: 0.6, y: 1.0, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("再栄養症候群と急性期栄養管理の考え方", {
    x: 0.6, y: 3.0, w: 8.8, h: 0.6,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("必要カロリーは「初日投与量」ではなく「安全に到達すべき目標量」", {
    x: 0.6, y: 4.0, w: 8.8, h: 0.6,
    fontSize: 18, fontFace: FJ, color: "B0BEC5", align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: よくある疑問
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある疑問");

  // Question card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 8.4, h: 1.4, fill: { color: C.light }, shadow: shd() });
  s.addText("低栄養の患者に、計算上の必要カロリーを最初から入れたほうが早く良くなるのでは？", {
    x: 1.2, y: 1.4, w: 7.6, h: 1.2,
    fontSize: 22, fontFace: FJ, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Answer card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.1, w: 8.4, h: 2.2, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.1, w: 0.15, h: 2.2, fill: { color: C.red } });
  s.addText([
    { text: "答え：必ずしもそうではない", options: { fontSize: 26, bold: true, color: C.red, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "急速な栄養投与は再栄養症候群や過栄養を招き、かえって有害になることがある", options: { fontSize: 20, color: C.text } },
  ], { x: 1.3, y: 3.3, w: 7.5, h: 1.8, fontFace: FJ, align: "center", valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 3: 必要カロリーの考え方
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "必要カロリーの正しい考え方");

  // Left: Wrong
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 3.8, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.6, fill: { color: C.red } });
  s.addText("誤った考え方", { x: 0.5, y: 1.3, w: 4.3, h: 0.6, fontSize: 20, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "必要カロリー ＝ 初日投与量", options: { fontSize: 22, bold: true, color: C.red, breakLine: true } },
    { text: "", options: { fontSize: 16, breakLine: true } },
    { text: "「早く到達 ＝ 早く回復」", options: { fontSize: 20, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "過剰負荷のリスク", options: { fontSize: 20, bold: true, color: C.red } },
  ], { x: 0.8, y: 2.2, w: 3.7, h: 2.6, fontFace: FJ, align: "center", valign: "middle", margin: 0 });

  // Right: Correct
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 3.8, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.6, fill: { color: C.green } });
  s.addText("正しい考え方", { x: 5.2, y: 1.3, w: 4.3, h: 0.6, fontSize: 20, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "必要カロリー ＝ 最終到達目標", options: { fontSize: 22, bold: true, color: C.green, breakLine: true } },
    { text: "", options: { fontSize: 16, breakLine: true } },
    { text: "「安全に到達する」", options: { fontSize: 20, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "段階的に増量", options: { fontSize: 20, bold: true, color: C.green } },
  ], { x: 5.5, y: 2.2, w: 3.7, h: 2.6, fontFace: FJ, align: "center", valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 4: 再栄養症候群とは
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.red } });
  s.addText([
    { text: "再栄養症候群（", options: { fontFace: FJ } },
    { text: "Refeeding Syndrome", options: { fontFace: FE } },
    { text: "）とは", options: { fontFace: FJ } },
  ], { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, color: C.white, bold: true, margin: 0 });

  // Main card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 9, h: 4.2, fill: { color: C.white }, shadow: shd() });

  // Step 1
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.3, w: 4.0, h: 0.7, fill: { color: C.primary } });
  s.addText("長期絶食・摂取不良", { x: 0.8, y: 1.3, w: 4.0, h: 0.7, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("体内の P・K・Mg・チアミンが枯渇（血清値は見かけ上正常のことも）", {
    x: 5.0, y: 1.3, w: 4.2, h: 0.7, fontSize: 14, fontFace: FJ, color: C.sub, align: "left", valign: "middle", margin: 0,
  });

  // Arrow
  s.addText("▼", { x: 2.5, y: 2.05, w: 0.8, h: 0.35, fontSize: 16, fontFace: FE, color: C.sub, align: "center", margin: 0 });

  // Step 2
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.4, w: 4.0, h: 0.7, fill: { color: C.accent } });
  s.addText("急速な栄養投与（特に糖質）", { x: 0.8, y: 2.4, w: 4.0, h: 0.7, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("インスリン分泌増加", { x: 5.0, y: 2.4, w: 4.2, h: 0.7, fontSize: 16, fontFace: FJ, color: C.text, align: "left", valign: "middle", margin: 0 });

  // Arrow
  s.addText("▼", { x: 2.5, y: 3.15, w: 0.8, h: 0.35, fontSize: 16, fontFace: FE, color: C.sub, align: "center", margin: 0 });

  // Step 3
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.5, w: 4.0, h: 0.7, fill: { color: C.red } });
  s.addText("電解質の急速な細胞内移動", { x: 0.8, y: 3.5, w: 4.0, h: 0.7, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText([
    { text: "P\u2193  K\u2193  Mg\u2193", options: { fontFace: FE } },
    { text: " ＋ 水・", options: { fontFace: FJ } },
    { text: "Na", options: { fontFace: FE } },
    { text: "貯留", options: { fontFace: FJ } },
  ], { x: 5.0, y: 3.5, w: 4.2, h: 0.7, fontSize: 16, color: C.red, bold: true, align: "left", valign: "middle", margin: 0 });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.4, w: 8.4, h: 0.7, fill: { color: C.lightRed } });
  s.addText([
    { text: "重症例：臓器障害・不整脈・心不全・", options: { fontFace: FJ } },
    { text: "Wernicke", options: { fontFace: FE } },
    { text: "脳症", options: { fontFace: FJ } },
  ], { x: 0.8, y: 4.4, w: 8.4, h: 0.7, fontSize: 18, color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 5: 急性期の過栄養
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "急性期の過栄養 ― もうひとつの問題");

  // Key concept card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.5, fill: { color: C.light }, shadow: shd() });
  s.addText([
    { text: "急性期では内因性エネルギー産生（糖新生・脂肪分解・蛋白異化）が残る", options: { fontSize: 17, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "計算式どおりに入れると「内因性＋外因性」で過栄養になる", options: { fontSize: 20, bold: true, color: C.red } },
  ], { x: 0.8, y: 1.3, w: 8.4, h: 1.3, fontFace: FJ, align: "center", valign: "middle", margin: 0 });

  // ESPEN card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 4.3, h: 0.5, fill: { color: C.primary } });
  s.addText([
    { text: "ESPEN", options: { fontFace: FE } },
    { text: "推奨", options: { fontFace: FJ } },
  ], { x: 0.5, y: 3.0, w: 4.3, h: 0.5, fontSize: 18, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "ICU", options: { fontSize: 16, fontFace: FE, color: C.sub } },
    { text: "最初の1週間：", options: { fontSize: 16, fontFace: FJ, color: C.sub, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "推定必要量の", options: { fontSize: 24, fontFace: FJ, bold: true, color: C.primary } },
    { text: "70%", options: { fontSize: 24, fontFace: FE, bold: true, color: C.primary } },
    { text: "未満", options: { fontSize: 24, fontFace: FJ, bold: true, color: C.primary } },
  ], { x: 0.7, y: 3.7, w: 3.9, h: 1.3, align: "center", valign: "middle", margin: 0 });

  // NUTRIREA-3 card
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.0, w: 4.3, h: 2.2, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.0, w: 4.3, h: 0.5, fill: { color: C.accent } });
  s.addText([
    { text: "NUTRIREA-3", options: { fontFace: FE } },
    { text: "試験", options: { fontFace: FJ } },
  ], { x: 5.2, y: 3.0, w: 4.3, h: 0.5, fontSize: 18, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "標準投与（", options: { fontSize: 15, fontFace: FJ, color: C.text } },
    { text: "25 kcal/kg/", options: { fontSize: 15, fontFace: FE, color: C.text } },
    { text: "日）は", options: { fontSize: 15, fontFace: FJ, color: C.text, breakLine: true } },
    { text: "死亡率を改善せず", options: { fontSize: 17, fontFace: FJ, bold: true, color: C.red, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "嘔吐・下痢・腸管虚血・肝機能障害が増加", options: { fontSize: 15, fontFace: FJ, color: C.text } },
  ], { x: 5.4, y: 3.6, w: 3.9, h: 1.4, align: "center", valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 6: 高リスク患者の見分け方（NICE基準）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText([
    { text: "再栄養リスクの高い患者（", options: { fontFace: FJ } },
    { text: "NICE", options: { fontFace: FE } },
    { text: "基準）", options: { fontFace: FJ } },
  ], { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, color: C.white, bold: true, margin: 0 });

  // Left: 1項目
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 4.0, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.6, fill: { color: C.red } });
  s.addText("1項目でも該当 → 高リスク", { x: 0.5, y: 1.2, w: 4.3, h: 0.6, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "BMI < 16", options: { fontSize: 18, fontFace: FE, bold: true, color: C.text, breakLine: true, bullet: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "3\u20136か月で15%以上の体重減少", options: { fontSize: 18, fontFace: FJ, bold: true, color: C.text, breakLine: true, bullet: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "10日超のほぼ絶食", options: { fontSize: 18, fontFace: FJ, bold: true, color: C.text, breakLine: true, bullet: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "開始前から低K・低P・低Mg", options: { fontSize: 18, fontFace: FJ, bold: true, color: C.text, bullet: true } },
  ], { x: 0.9, y: 2.0, w: 3.6, h: 3.0, valign: "middle", margin: 0 });

  // Right: 2項目以上
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 4.0, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.6, fill: { color: C.accent } });
  s.addText("2項目以上 → 高リスク", { x: 5.2, y: 1.2, w: 4.3, h: 0.6, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "BMI < 18.5", options: { fontSize: 18, fontFace: FE, bold: true, color: C.text, breakLine: true, bullet: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "3\u20136か月で10%以上の体重減少", options: { fontSize: 18, fontFace: FJ, bold: true, color: C.text, breakLine: true, bullet: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "5日超の摂取不良", options: { fontSize: 18, fontFace: FJ, bold: true, color: C.text, breakLine: true, bullet: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "アルコール多飲・利尿薬・制酸薬・インスリン使用歴", options: { fontSize: 16, fontFace: FJ, bold: true, color: C.text, bullet: true } },
  ], { x: 5.6, y: 2.0, w: 3.6, h: 3.0, valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 7: 導入量の目安
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText([
    { text: "導入量の目安（", options: { fontFace: FJ } },
    { text: "NICE / ASPEN", options: { fontFace: FE } },
    { text: "）", options: { fontFace: FJ } },
  ], { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, color: C.white, bold: true, margin: 0 });

  var rowY = 1.15;
  var rowH = 0.58;
  var colL = 0.5;
  var colLW = 3.8;
  var gap = 0.05;
  var colR = colL + colLW + gap;
  var colRW = 10 - colR - 0.5;

  // Header
  s.addShape(pres.shapes.RECTANGLE, { x: colL, y: rowY, w: colLW, h: rowH, fill: { color: C.primary } });
  s.addText("患者像", { x: colL, y: rowY, w: colLW, h: rowH, fontSize: 16, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: colR, y: rowY, w: colRW, h: rowH, fill: { color: C.primary } });
  s.addText("導入の考え方", { x: colR, y: rowY, w: colRW, h: rowH, fontSize: 16, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var rows = [
    { left: "再栄養リスクなし", right: "25\u201335 kcal/kg/日を目標", bg: C.white, lc: C.text, rc: C.text, rb: false },
    { left: "重症・外傷（EN/PN）", right: "目標の50%以下から開始", bg: "F0F4F8", lc: C.text, rc: C.text, rb: true },
    { left: "5日以上ほぼ絶食", right: "最初の2日間は必要量の50%以下", bg: C.white, lc: C.text, rc: C.text, rb: true },
    { left: "再栄養 高リスク", right: "10 kcal/kg/日以下 → 4\u20137日で増量", bg: C.lightYellow, lc: C.accent, rc: C.accent, rb: true },
    { left: "極めて高リスク", right: "5 kcal/kg/日から開始", bg: C.lightRed, lc: C.red, rc: C.red, rb: true },
  ];

  for (var i = 0; i < rows.length; i++) {
    var ry = rowY + rowH + i * rowH;
    var r = rows[i];
    s.addShape(pres.shapes.RECTANGLE, { x: colL, y: ry, w: colLW, h: rowH, fill: { color: r.bg }, line: { color: "DEE2E6", width: 0.5 } });
    s.addText(r.left, { x: colL + 0.2, y: ry, w: colLW - 0.4, h: rowH, fontSize: 15, fontFace: FJ, color: r.lc, bold: r.rb, valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: colR, y: ry, w: colRW, h: rowH, fill: { color: r.bg }, line: { color: "DEE2E6", width: 0.5 } });
    s.addText(r.right, { x: colR + 0.2, y: ry, w: colRW - 0.4, h: rowH, fontSize: 15, fontFace: FJ, color: r.rc, bold: r.rb, valign: "middle", margin: 0 });
  }

  // Bottom message
  var msgY = rowY + rowH * (rows.length + 1) + 0.25;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: msgY, w: 9, h: 0.7, fill: { color: C.light } });
  s.addText("必要量はゴールであって、開始量ではない", {
    x: 0.5, y: msgY, w: 9, h: 0.7,
    fontSize: 22, fontFace: FJ, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });
})();

// ============================================================
// SLIDE 8: カロリーより先に考えるべきこと
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "カロリーより先に考えるべきこと");

  // Card 1: チアミン
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 2.0, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fill: { color: C.green } });
  s.addText("チアミン投与", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "チアミン ", options: { fontSize: 20, fontFace: FJ, bold: true, color: C.green } },
    { text: "100 mg", options: { fontSize: 20, fontFace: FE, bold: true, color: C.green, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "栄養開始前 or ブドウ糖投与前に投与", options: { fontSize: 15, fontFace: FJ, color: C.text, breakLine: true } },
    { text: "5\u20137日以上継続", options: { fontSize: 15, fontFace: FJ, color: C.text } },
  ], { x: 0.7, y: 1.85, w: 3.9, h: 1.2, align: "center", valign: "middle", margin: 0 });

  // Card 2: 電解質モニタリング
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 2.0, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fill: { color: C.primary } });
  s.addText("電解質モニタリング", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "K・Mg・Pを開始前に確認", options: { fontSize: 16, fontFace: FJ, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "高リスク：最初3日間 ", options: { fontSize: 16, fontFace: FJ, color: C.text } },
    { text: "12", options: { fontSize: 18, fontFace: FE, bold: true, color: C.primary } },
    { text: "時間ごと", options: { fontSize: 18, fontFace: FJ, bold: true, color: C.primary } },
  ], { x: 5.4, y: 1.85, w: 3.9, h: 1.2, align: "center", valign: "middle", margin: 0 });

  // Card 3: 落とし穴
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9, h: 1.8, fill: { color: C.lightYellow }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 0.15, h: 1.8, fill: { color: C.yellow } });
  s.addText([
    { text: "見落としやすい落とし穴", options: { fontSize: 20, bold: true, color: C.accent, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "維持輸液・薬剤中のブドウ糖も再栄養負荷として数える", options: { fontSize: 17, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "電解質が急速に低下 → カロリーを", options: { fontSize: 17, color: C.text } },
    { text: "50%", options: { fontSize: 17, fontFace: FE, bold: true, color: C.accent } },
    { text: "減量して再増量", options: { fontSize: 17, color: C.text } },
  ], { x: 1.0, y: 3.6, w: 8.2, h: 1.6, fontFace: FJ, valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 9: 常に「low and slow」か？
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText([
    { text: "常に「", options: { fontFace: FJ } },
    { text: "low and slow", options: { fontFace: FE, italic: true } },
    { text: "」が正解か？", options: { fontFace: FJ } },
  ], { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, color: C.white, bold: true, margin: 0 });

  // Study card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 2.0, fill: { color: C.lightGreen }, shadow: shd() });
  s.addText([
    { text: "摂食障害の若年患者（12\u201324歳）のRCT", options: { fontSize: 18, fontFace: FJ, bold: true, color: C.green, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "2000 kcal/日開始  vs  1400 kcal/日開始", options: { fontSize: 20, fontFace: FJ, bold: true, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "高カロリー群：内科的安定化が早い・在院日数約4日短縮・電解質異常の増加なし", options: { fontSize: 15, fontFace: FJ, color: C.text } },
  ], { x: 0.8, y: 1.3, w: 8.4, h: 1.8, align: "center", valign: "middle", margin: 0 });

  // Caution card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9, h: 1.8, fill: { color: C.lightYellow }, shadow: shd() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 0.15, h: 1.8, fill: { color: C.accent } });
  s.addText([
    { text: "ただし外挿に注意", options: { fontSize: 20, bold: true, color: C.accent, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "厳密な電解質モニタリング下での若年摂食障害のデータであり、高齢内科患者・ICU患者・敗血症へのそのまま外挿は危険", options: { fontSize: 16, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "導入量は病態・年齢・炎症・臓器障害・監視体制で決める", options: { fontSize: 18, bold: true, color: C.primary } },
  ], { x: 1.0, y: 3.6, w: 8.2, h: 1.6, fontFace: FJ, valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 10: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("まとめ", {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 3.6, fill: { color: "1F4A7A" } });
  s.addText([
    { text: "1.  必要カロリーは「初日投与量」ではなく「目標量」", options: { fontSize: 19, color: C.white, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "2.  急速投与は再栄養症候群・過栄養のリスク", options: { fontSize: 19, color: C.white, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "3.  高リスク患者はNICE基準で判別", options: { fontSize: 19, color: C.white, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "4.  チアミンと電解質管理はカロリーより先に", options: { fontSize: 19, color: C.white, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "5.  導入量は病態・年齢・監視体制で個別に決める", options: { fontSize: 19, color: C.white } },
  ], { x: 1.4, y: 1.4, w: 7.4, h: 3.2, fontFace: FJ, valign: "middle", margin: 0 });

  s.addText("必要量を計算したら、その数字はゴールとして扱い、導入量は再栄養リスクと急性期病態で決める", {
    x: 0.6, y: 5.0, w: 8.8, h: 0.5,
    fontSize: 13, fontFace: FJ, color: C.accent, align: "center", valign: "middle", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 11: 参考文献
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "参考文献");

  s.addText([
    { text: "1.  NICE CG32. Nutrition support for adults. Updated 2017.", options: { fontSize: 14, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "2.  da Silva JSV, et al. ASPEN Consensus Recommendations for Refeeding Syndrome.", options: { fontSize: 14, color: C.text, breakLine: true } },
    { text: "     Nutr Clin Pract. 2020;35(2):178-195.", options: { fontSize: 13, color: C.sub, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "3.  Singer P, et al. ESPEN guideline on clinical nutrition in the ICU.", options: { fontSize: 14, color: C.text, breakLine: true } },
    { text: "     Clin Nutr. 2019;38(1):48-79.", options: { fontSize: 13, color: C.sub, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "4.  Reignier J, et al. NUTRIREA-3: Low vs standard calorie and protein feeding", options: { fontSize: 14, color: C.text, breakLine: true } },
    { text: "     in ventilated adults with shock. Lancet Respir Med. 2023;11(7):602-612.", options: { fontSize: 13, color: C.sub, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "5.  Garber AK, et al. Refeeding to Optimize Inpatient Gains for Patients With", options: { fontSize: 14, color: C.text, breakLine: true } },
    { text: "     Anorexia Nervosa: A Multicenter RCT. JAMA Pediatr. 2021;175(1):19-27.", options: { fontSize: 13, color: C.sub } },
  ], { x: 0.6, y: 1.2, w: 8.8, h: 4.2, fontFace: FE, valign: "top", margin: 0 });
})();

// ============================================================
// Generate PPTX
// ============================================================
var outputPath = "youtube-slides/再栄養症候群/refeeding_syndrome_slides.pptx";
pres.writeFile({ fileName: outputPath }).then(function() {
  console.log("PPTX created: " + outputPath);
}).catch(function(err) {
  console.error("Error:", err);
});
