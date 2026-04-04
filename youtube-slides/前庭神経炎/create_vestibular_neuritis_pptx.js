const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "前庭神経炎の鑑別診断 ― 脳卒中を見逃さないためのアプローチ";

// --- Color Palette (Medical / Neuro / Navy) ---
const C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "4A90D9",
  light: "E8F0FE",
  warmBg: "F5F7FA",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "DC3545",
  orange: "E8850C",
  yellow: "F5C518",
  green: "28A745",
};

const FONT_H = "Arial Black";
const FONT_B = "Arial";

const makeShadow = function() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; };

// ============================================================
// SLIDE 1: Title
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("前庭神経炎の鑑別診断", {
    x: 0.8, y: 1.0, w: 8.4, h: 1.2,
    fontSize: 42, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("― 脳卒中を見逃さないためのアプローチ ―", {
    x: 0.8, y: 2.2, w: 8.4, h: 0.8,
    fontSize: 24, fontFace: FONT_B, color: C.accent, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 3.2, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.6, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_B, color: C.light, align: "center", margin: 0,
  });
  s.addText("Barany Society 2022 分類 / GRACE-3 ガイドライン準拠", {
    x: 0.8, y: 4.3, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_B, color: C.sub, align: "center", italic: true, margin: 0,
  });
}

// ============================================================
// SLIDE 2: 前庭神経炎とは
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("前庭神経炎とは ― 疾患概念", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Main card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.6, fill: { color: C.light }, shadow: makeShadow() });
  s.addText([
    { text: "AUVP（Acute Unilateral Vestibulopathy）", options: { fontSize: 22, bold: true, color: C.primary, breakLine: true } },
    { text: "= 急性一側前庭障害  ≒ 前庭神経炎", options: { fontSize: 18, color: C.text } },
  ], { x: 0.8, y: 1.3, w: 8.4, h: 1.4, fontFace: FONT_B, align: "center", valign: "middle", margin: 0 });

  // 3 key points
  var items = [
    { num: "1", txt: "急性に生じた一側の\n末梢前庭機能低下\nによる持続性めまい" },
    { num: "2", txt: "急性の中枢神経症状\nがないこと" },
    { num: "3", txt: "急性の聴覚症状\nがないこと" },
  ];
  items.forEach(function(item, i) {
    var xPos = 0.5 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.1, w: 2.8, h: 2.2, fill: { color: C.white }, shadow: makeShadow() });
    s.addText(item.num, { x: xPos, y: 3.2, w: 2.8, h: 0.6, fontSize: 28, fontFace: FONT_H, color: C.accent, bold: true, align: "center", margin: 0 });
    s.addText(item.txt, { x: xPos + 0.15, y: 3.8, w: 2.5, h: 1.4, fontSize: 15, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
  });
}

// ============================================================
// SLIDE 3: 病態と原因
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("病態と原因", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });

  s.addText("病態は十分には解明されていない", {
    x: 1, y: 1.4, w: 8, h: 0.6,
    fontSize: 24, fontFace: FONT_H, color: C.red, bold: true, align: "center", margin: 0,
  });

  var hypotheses = [
    { label: "ウイルス関連炎症説", desc: "感染後の炎症反応を支持するレビューあり", color: C.accent },
    { label: "血管性機序", desc: "前庭神経への血流障害が提案されている", color: C.orange },
    { label: "免疫学的機序", desc: "自己免疫的プロセスの関与も指摘", color: C.green },
  ];
  hypotheses.forEach(function(h, i) {
    var yPos = 2.3 + i * 0.9;
    s.addShape(pres.shapes.RECTANGLE, { x: 1.2, y: yPos, w: 0.3, h: 0.7, fill: { color: h.color } });
    s.addText(h.label, { x: 1.8, y: yPos, w: 3, h: 0.7, fontSize: 18, fontFace: FONT_H, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(h.desc, { x: 5, y: yPos, w: 4.2, h: 0.7, fontSize: 16, fontFace: FONT_B, color: C.sub, valign: "middle", margin: 0 });
  });

  s.addText("→ 単一の病因では説明できず、複数の機序が関与する可能性", {
    x: 1, y: 4.6, w: 8, h: 0.5,
    fontSize: 16, fontFace: FONT_B, color: C.primary, italic: true, align: "center", margin: 0,
  });
}

// ============================================================
// SLIDE 4: 症状の特徴
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("症状の特徴", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Left: typical symptoms
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.6, fill: { color: C.primary } });
  s.addText("典型的な症状", { x: 0.5, y: 1.2, w: 4.3, h: 0.6, fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });

  s.addText([
    { text: "持続性の強いめまい", options: { fontSize: 18, bold: true, color: C.primary, breakLine: true } },
    { text: "（24時間以上持続）", options: { fontSize: 14, color: C.sub, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "・ふらつき", options: { fontSize: 16, breakLine: true } },
    { text: "・悪心・嘔吐", options: { fontSize: 16, breakLine: true } },
    { text: "・頭部回旋時の増悪", options: { fontSize: 16, breakLine: true } },
    { text: "・Oscillopsia", options: { fontSize: 16, breakLine: true } },
    { text: "  （物が揺れて見える）", options: { fontSize: 14, color: C.sub } },
  ], { x: 0.8, y: 2.0, w: 3.7, h: 3.2, fontFace: FONT_B, color: C.text, valign: "top", margin: 0 });

  // Right: red flags
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.6, fill: { color: C.red } });
  s.addText("典型的でない症状", { x: 5.2, y: 1.2, w: 4.3, h: 0.6, fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });

  s.addText([
    { text: "以下は前庭神経炎に合わない", options: { fontSize: 16, bold: true, color: C.red, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "・急性の難聴", options: { fontSize: 18, bold: true, breakLine: true } },
    { text: "・耳鳴り", options: { fontSize: 18, bold: true, breakLine: true } },
    { text: "・耳閉感", options: { fontSize: 18, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "→ 診断を立て直す", options: { fontSize: 18, bold: true, color: C.red } },
  ], { x: 5.5, y: 2.0, w: 3.7, h: 3.2, fontFace: FONT_B, color: C.text, valign: "top", margin: 0 });
}

// ============================================================
// SLIDE 5: 診察 ― 前庭神経炎に合う所見
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("診察：前庭神経炎に合う所見", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Finding 1
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.8, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("1", { x: 0.7, y: 1.3, w: 0.8, h: 1.6, fontSize: 40, fontFace: FONT_H, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText([
    { text: "方向固定性の水平回旋混合性自発眼振", options: { fontSize: 20, bold: true, color: C.primary, breakLine: true } },
    { text: "視固定を外すと増強する（Frenzel眼鏡下で確認）", options: { fontSize: 16, color: C.sub } },
  ], { x: 1.8, y: 1.4, w: 7.4, h: 1.4, fontFace: FONT_B, valign: "middle", margin: 0 });

  // Finding 2
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.3, w: 9, h: 2.0, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("2", { x: 0.7, y: 3.4, w: 0.8, h: 1.8, fontSize: 40, fontFace: FONT_H, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText([
    { text: "一側のVOR（前庭眼反射）低下", options: { fontSize: 20, bold: true, color: C.primary, breakLine: true } },
    { text: "Head Impulse Testで評価", options: { fontSize: 16, color: C.sub, breakLine: true } },
    { text: "患側に素早く頭部回旋 → catch-up saccade 出現", options: { fontSize: 16, color: C.text } },
  ], { x: 1.8, y: 3.5, w: 7.4, h: 1.6, fontFace: FONT_B, valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 6: 中枢性めまいを疑う所見
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.red } });
  s.addText("中枢性めまいを疑う所見", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  var redFlags = [
    "Skew deviation（垂直性の眼位ずれ）",
    "注視方向で変わる眼振（gaze-evoked nystagmus）",
    "垂直眼振 / 純回旋性眼振",
    "強い体幹失調 / 歩行不能",
    "明らかな神経巣症状",
  ];
  redFlags.forEach(function(flag, i) {
    var yPos = 1.2 + i * 0.8;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.65, fill: { color: i % 2 === 0 ? "FFF0F0" : C.white } });
    s.addText("!", { x: 0.9, y: yPos, w: 0.5, h: 0.65, fontSize: 20, fontFace: FONT_H, color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(flag, { x: 1.5, y: yPos, w: 7.5, h: 0.65, fontSize: 18, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
  });

  // Warning box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 5.0, w: 8.4, h: 0.5, fill: { color: C.red } });
  s.addText("Head Impulse Test 単独では脳卒中を除外できない", {
    x: 0.8, y: 5.0, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
}

// ============================================================
// SLIDE 7: HINTSの位置づけと注意点
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("HINTSの位置づけと注意点", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // HINTS components
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.2, fill: { color: C.light } });
  s.addText([
    { text: "H", options: { fontSize: 24, bold: true, color: C.primary } },
    { text: "ead-Impulse  /  ", options: { fontSize: 18, color: C.text } },
    { text: "N", options: { fontSize: 24, bold: true, color: C.primary } },
    { text: "ystagmus  /  ", options: { fontSize: 18, color: C.text } },
    { text: "T", options: { fontSize: 24, bold: true, color: C.primary } },
    { text: "est of ", options: { fontSize: 18, color: C.text } },
    { text: "S", options: { fontSize: 24, bold: true, color: C.primary } },
    { text: "kew", options: { fontSize: 18, color: C.text } },
  ], { x: 0.5, y: 1.2, w: 9, h: 1.2, fontFace: FONT_B, align: "center", valign: "middle", margin: 0 });

  // Conditions
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.7, w: 4.3, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.7, w: 4.3, h: 0.5, fill: { color: C.green } });
  s.addText("適用条件", { x: 0.5, y: 2.7, w: 4.3, h: 0.5, fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "・自発眼振を伴う持続性AVS", options: { fontSize: 15, breakLine: true } },
    { text: "・訓練された診察者が実施", options: { fontSize: 15 } },
  ], { x: 0.7, y: 3.3, w: 3.9, h: 0.8, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.7, w: 4.3, h: 1.5, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.7, w: 4.3, h: 0.5, fill: { color: C.red } });
  s.addText("適用外", { x: 5.2, y: 2.7, w: 4.3, h: 0.5, fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "・頭位誘発性のめまい（BPPV）", options: { fontSize: 15, breakLine: true } },
    { text: "・反復発作性のめまい", options: { fontSize: 15 } },
  ], { x: 5.4, y: 3.3, w: 3.9, h: 0.8, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });

  // GRACE-3 quote
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 9, h: 1.0, fill: { color: "FFF3CD" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 0.15, h: 1.0, fill: { color: C.orange } });
  s.addText("GRACE-3 (2023): 訓練なしの救急現場でroutineに行われるHINTSは不正確であり、standard of careとは言えない", {
    x: 0.9, y: 4.5, w: 8.4, h: 1.0,
    fontSize: 15, fontFace: FONT_B, color: C.text, italic: true, valign: "middle", margin: 0,
  });
}

// ============================================================
// SLIDE 8: 鑑別診断テーブル
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("鑑別診断 ― 主要5疾患", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  var headers = ["疾患", "性状", "持続時間", "鑑別のポイント"];
  var colW = [2.0, 1.8, 1.6, 3.6];
  var colX = [0.5, 2.5, 4.3, 5.9];

  // Header row
  headers.forEach(function(h, i) {
    s.addShape(pres.shapes.RECTANGLE, { x: colX[i], y: 1.15, w: colW[i], h: 0.5, fill: { color: C.primary } });
    s.addText(h, { x: colX[i], y: 1.15, w: colW[i], h: 0.5, fontSize: 13, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  var rows = [
    ["前庭神経炎", "持続性・自発性", "24時間以上", "一側VOR低下、聴覚症状なし"],
    ["BPPV", "頭位変換誘発", "数秒～1分", "Dix-Hallpike陽性"],
    ["メニエール病", "反復発作性", "20分～12時間", "変動性聴覚症状を伴う"],
    ["前庭性片頭痛", "反復性・多彩", "5分～72時間", "片頭痛の病歴・症状関連"],
    ["脳幹・小脳梗塞", "持続性・急性", "持続", "中枢性眼振、skew、失調"],
  ];
  rows.forEach(function(row, ri) {
    var yPos = 1.65 + ri * 0.62;
    var bgColor = ri === 4 ? "FFF0F0" : (ri % 2 === 0 ? C.light : C.white);
    row.forEach(function(cell, ci) {
      s.addShape(pres.shapes.RECTANGLE, { x: colX[ci], y: yPos, w: colW[ci], h: 0.62, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      var txtColor = (ri === 4 && ci === 0) ? C.red : C.text;
      var isBold = ci === 0;
      s.addText(cell, { x: colX[ci] + 0.1, y: yPos, w: colW[ci] - 0.2, h: 0.62, fontSize: 12, fontFace: FONT_B, color: txtColor, bold: isBold, valign: "middle", margin: 0 });
    });
  });

  // Emphasis (y = 1.65 + 5*0.62 = 4.75, so start at 4.85 with gap)
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.85, w: 9, h: 0.5, fill: { color: C.red } });
  s.addText("最重要の鑑別 = 脳幹・小脳梗塞（片麻痺や構音障害がなくても安心できない）", {
    x: 0.5, y: 4.85, w: 9, h: 0.5,
    fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
}

// ============================================================
// SLIDE 9: 検査の進め方
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("検査の進め方", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // 3 checkpoints
  var checks = [
    { num: "1", txt: "持続性の\n急性前庭症候群か" },
    { num: "2", txt: "一側VOR低下\nを示せるか" },
    { num: "3", txt: "聴覚症状や\n中枢所見がないか" },
  ];
  checks.forEach(function(c, i) {
    var xPos = 0.5 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 2.8, h: 1.6, fill: { color: C.white }, shadow: makeShadow() });
    s.addText(c.num, { x: xPos, y: 1.25, w: 2.8, h: 0.5, fontSize: 24, fontFace: FONT_H, color: C.accent, bold: true, align: "center", margin: 0 });
    s.addText(c.txt, { x: xPos + 0.15, y: 1.7, w: 2.5, h: 1.0, fontSize: 15, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
  });

  // Imaging guidelines
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.1, w: 9, h: 2.4, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("画像検査の位置づけ（GRACE-3）", {
    x: 0.7, y: 3.2, w: 8.6, h: 0.5,
    fontSize: 18, fontFace: FONT_H, color: C.primary, bold: true, margin: 0,
  });

  var imgRules = [
    { icon: "×", text: "AVSでのbrain CTは勧めない", color: C.red },
    { icon: "△", text: "HINTSに熟練した診察者がいればroutine MRIはfirst-lineにしない", color: C.orange },
    { icon: "○", text: "HINTSが中枢性/判定不能 → MRIをconfirmatory testに", color: C.green },
  ];
  imgRules.forEach(function(r, i) {
    var yPos = 3.8 + i * 0.55;
    s.addText(r.icon, { x: 1, y: yPos, w: 0.5, h: 0.5, fontSize: 20, fontFace: FONT_H, color: r.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(r.text, { x: 1.6, y: yPos, w: 7.5, h: 0.5, fontSize: 15, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
  });
}

// ============================================================
// SLIDE 10: 治療 ― 急性期の支持療法
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("治療：急性期の支持療法", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Goals
  var goals = [
    { icon: "1", txt: "悪心・嘔吐の軽減" },
    { icon: "2", txt: "脱水予防" },
    { icon: "3", txt: "離床のきっかけを作る" },
  ];
  goals.forEach(function(g, i) {
    var xPos = 0.5 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 2.8, h: 1.4, fill: { color: C.white }, shadow: makeShadow() });
    s.addText(g.icon, { x: xPos, y: 1.25, w: 2.8, h: 0.5, fontSize: 28, fontFace: FONT_H, color: C.accent, bold: true, align: "center", margin: 0 });
    s.addText(g.txt, { x: xPos + 0.15, y: 1.7, w: 2.5, h: 0.8, fontSize: 17, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
  });

  // Warning about vestibular suppressants
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.9, w: 9, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("制吐薬・前庭抑制薬の使い方", {
    x: 0.7, y: 3.0, w: 8.6, h: 0.5,
    fontSize: 18, fontFace: FONT_H, color: C.primary, bold: true, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1, y: 3.6, w: 3.8, h: 1.5, fill: { color: "E8FFE8" } });
  s.addText([
    { text: "初期数日", options: { fontSize: 16, bold: true, color: C.green, breakLine: true } },
    { text: "症状緩和に有用", options: { fontSize: 15, color: C.text } },
  ], { x: 1.2, y: 3.7, w: 3.4, h: 1.3, fontFace: FONT_B, align: "center", valign: "middle", margin: 0 });

  s.addText("→", { x: 4.6, y: 3.8, w: 0.8, h: 1.0, fontSize: 30, fontFace: FONT_H, color: C.accent, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.6, w: 3.8, h: 1.5, fill: { color: "FFF0F0" } });
  s.addText([
    { text: "長期使用は避ける", options: { fontSize: 16, bold: true, color: C.red, breakLine: true } },
    { text: "中枢代償を妨げる\n可能性あり", options: { fontSize: 15, color: C.text } },
  ], { x: 5.4, y: 3.7, w: 3.4, h: 1.3, fontFace: FONT_B, align: "center", valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 11: ステロイドの位置づけ
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("ステロイドの位置づけ ― エビデンスの変遷", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 28, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  var studies = [
    { year: "2004", author: "Strupp (NEJM)", result: "有意に改善", resultColor: C.green, desc: "methylprednisolone RCT\nP<0.001" },
    { year: "2021", author: "Leong (SR/MA)", result: "限定的", resultColor: C.orange, desc: "短期改善示唆\n長期根拠不十分" },
    { year: "2022", author: "Bogdanova (MA)", result: "慎重な解釈", resultColor: C.orange, desc: "一定の利益\n研究の質に問題" },
    { year: "2025", author: "Sjogren (RCT)", result: "有意差なし", resultColor: C.red, desc: "プラセボ対照\n二重盲検 N=69" },
  ];

  studies.forEach(function(st, i) {
    var yPos = 1.15 + i * 1.05;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9, h: 0.9, fill: { color: i % 2 === 0 ? C.white : C.light }, shadow: makeShadow() });
    s.addText(st.year, { x: 0.6, y: yPos, w: 0.9, h: 0.9, fontSize: 18, fontFace: FONT_H, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.author, { x: 1.6, y: yPos, w: 2.2, h: 0.9, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(st.desc, { x: 3.8, y: yPos, w: 3, h: 0.9, fontSize: 13, fontFace: FONT_B, color: C.sub, valign: "middle", margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y: yPos + 0.15, w: 2, h: 0.6, fill: { color: st.resultColor }, rectRadius: 0.1 });
    s.addText(st.result, { x: 7.2, y: yPos + 0.15, w: 2, h: 0.6, fontSize: 15, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // Conclusion
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.5, fill: { color: "FFF3CD" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 0.15, h: 0.5, fill: { color: C.orange } });
  s.addText("GRACE-3:「短期ステロイドを考慮してよい」＝ 全例標準投与ではない", {
    x: 0.9, y: 5.0, w: 8.4, h: 0.5,
    fontSize: 15, fontFace: FONT_B, color: C.text, bold: true, valign: "middle", margin: 0,
  });
}

// ============================================================
// SLIDE 12: 前庭リハビリテーション
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("前庭リハビリテーション", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Evidence card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.5, fill: { color: "E8FFE8" }, shadow: makeShadow() });
  s.addText([
    { text: "比較的しっかりしたエビデンスがある治療", options: { fontSize: 20, bold: true, color: C.green, breakLine: true } },
    { text: "Cochrane 2015 + メタ解析 2024 で有効性支持", options: { fontSize: 16, color: C.text, breakLine: true } },
    { text: "リハビリ + ステロイド併用 > ステロイド単独", options: { fontSize: 16, color: C.primary, bold: true } },
  ], { x: 0.8, y: 1.3, w: 8.4, h: 1.3, fontFace: FONT_B, align: "center", valign: "middle", margin: 0 });

  // Concept
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 4.3, h: 2.4, fill: { color: "FFF0F0" }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 4.3, h: 0.5, fill: { color: C.red } });
  s.addText("NG：安静を長引かせる", { x: 0.5, y: 3.0, w: 4.3, h: 0.5, fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "安静 → 代償遅延", options: { fontSize: 16, breakLine: true } },
    { text: "前庭抑制薬の長期使用", options: { fontSize: 16, breakLine: true } },
    { text: "→ 回復が遅れる", options: { fontSize: 16, bold: true, color: C.red } },
  ], { x: 0.7, y: 3.6, w: 3.9, h: 1.6, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.0, w: 4.3, h: 2.4, fill: { color: "E8FFE8" }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.0, w: 4.3, h: 0.5, fill: { color: C.green } });
  s.addText("OK：安全に動かして代償促進", { x: 5.2, y: 3.0, w: 4.3, h: 0.5, fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "嘔吐が落ち着いたら", options: { fontSize: 16, breakLine: true } },
    { text: "抑制薬を切り上げ", options: { fontSize: 16, breakLine: true } },
    { text: "→ 運動療法へ", options: { fontSize: 16, bold: true, color: C.green } },
  ], { x: 5.4, y: 3.6, w: 3.9, h: 1.6, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 13: 予後
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("予後", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 4.2, fill: { color: C.white }, shadow: makeShadow() });

  s.addText([
    { text: "急性期の回転感", options: { fontSize: 20, bold: true, color: C.primary, breakLine: true } },
    { text: "時間とともに軽減していくことが多い", options: { fontSize: 18, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "しかし一部の患者では…", options: { fontSize: 18, bold: true, color: C.orange, breakLine: true } },
    { text: "・慢性の浮動感が残存", options: { fontSize: 17, color: C.text, breakLine: true } },
    { text: "・Visual dependence（視覚依存性めまい）", options: { fontSize: 17, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "長期予後は末梢障害の程度だけでは説明できない", options: { fontSize: 18, bold: true, color: C.red, breakLine: true } },
    { text: "中枢代償側の因子（anxiety等）が関与", options: { fontSize: 16, color: C.sub, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "経過が非典型なら鑑別のやり直しが必要", options: { fontSize: 18, bold: true, color: C.primary } },
  ], { x: 1, y: 1.4, w: 8, h: 3.8, fontFace: FONT_B, valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 14: 見逃してはいけない所見
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.red } });
  s.addText("見逃してはいけない所見", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  var alerts = [
    "注視方向で変わる眼振 / 垂直眼振 / 純回旋性眼振 / skew deviation",
    "強い体幹失調 / 歩行不能 / 明らかな神経巣症状",
    "急性の難聴 / 耳鳴り / 耳閉感を伴う",
    "HINTSが中枢性 or 判定不能 / 適切な訓練がない",
    "MRIが早期陰性でも臨床的に疑わしい",
  ];

  alerts.forEach(function(a, i) {
    var yPos = 1.15 + i * 0.82;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9, h: 0.7, fill: { color: i % 2 === 0 ? "FFF0F0" : C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 0.5, h: 0.7, fill: { color: C.red } });
    s.addText(String(i + 1), { x: 0.5, y: yPos, w: 0.5, h: 0.7, fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(a, { x: 1.2, y: yPos, w: 8.1, h: 0.7, fontSize: 15, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.1, w: 9, h: 0.45, fill: { color: C.dark } });
  s.addText("「画像が陰性だから末梢性」と短絡しない", {
    x: 0.5, y: 5.1, w: 9, h: 0.45,
    fontSize: 17, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
}

// ============================================================
// SLIDE 15: まとめ ― 実地診療での要点
// ============================================================
{
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.dark } });
  s.addText("実地診療での要点（Take-Home Message）", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 28, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  var takeaways = [
    { num: "1", txt: "前庭神経炎は臨床診断 → 最優先は脳卒中除外", color: C.primary },
    { num: "2", txt: "難聴・skew・垂直眼振・歩行不能は典型VNから外れる", color: C.red },
    { num: "3", txt: "HINTSは「持続性AVS ＋ 訓練された診察者」が条件", color: C.orange },
    { num: "4", txt: "CT陰性≠stroke除外、早期MRI陰性でも安心しない", color: C.red },
    { num: "5", txt: "治療の柱 = 支持療法 ＋ 前庭リハビリ（ステロイドは個別判断）", color: C.green },
  ];

  takeaways.forEach(function(t, i) {
    var yPos = 1.1 + i * 0.88;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9, h: 0.75, fill: { color: C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 0.6, h: 0.75, fill: { color: t.color } });
    s.addText(t.num, { x: 0.5, y: yPos, w: 0.6, h: 0.75, fontSize: 22, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.txt, { x: 1.3, y: yPos, w: 8, h: 0.75, fontSize: 17, fontFace: FONT_B, color: C.text, bold: true, valign: "middle", margin: 0 });
  });

  // Footer
  s.addText("ご視聴ありがとうございました", {
    x: 0.5, y: 5.2, w: 9, h: 0.4,
    fontSize: 16, fontFace: FONT_B, color: C.sub, align: "center", italic: true, margin: 0,
  });
}

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/前庭神経炎_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX saved: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
