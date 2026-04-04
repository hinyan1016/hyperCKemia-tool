var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "Tolosa-Hunt症候群 ― 有痛性眼筋麻痺の鑑別と治療";

// --- Color Palette ---
var C = {
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

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var FONT_H = FONT_JP;
var FONT_B = FONT_JP;
var TOTAL = 17;

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
}

function addCard(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: C.white }, shadow: makeShadow() });
}

function addPageNum(s, num) {
  s.addText(num + "/" + TOTAL, { x: 9.0, y: 5.2, w: 0.8, h: 0.4, fontSize: 11, fontFace: FONT_B, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: Title
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Tolosa-Hunt症候群", {
    x: 0.8, y: 0.7, w: 8.4, h: 1.2,
    fontSize: 42, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("Tolosa-Hunt Syndrome (THS)", {
    x: 0.8, y: 1.8, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_EN, color: C.accent, align: "center", italic: true, margin: 0,
  });
  s.addText("― 有痛性眼筋麻痺の鑑別と治療 ―", {
    x: 0.8, y: 2.5, w: 8.4, h: 0.7,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 3.4, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.7, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("ICHD-3診断基準・最新レビューに基づく", {
    x: 0.8, y: 4.4, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 症例提示
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "症例提示");
  addPageNum(s, 2);

  addCard(s, 0.5, 1.2, 9, 3.8);

  s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.4, w: 1.8, h: 0.5, fill: { color: C.red } });
  s.addText("Case", { x: 0.7, y: 1.4, w: 1.8, h: 0.5, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0 });

  s.addText("55歳 男性", { x: 2.7, y: 1.35, w: 3, h: 0.6, fontSize: 22, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var lines = [
    "【主訴】左眼の奥の激しい痛み、物が二重に見える",
    "【現病歴】3日前から左眼窩後方に刺すような痛みが出現。",
    "　翌日から複視を自覚し、左眼瞼下垂も出現。",
    "【既往歴】特記事項なし　【服薬歴】なし",
    "【神経所見】左CN III麻痺（眼瞼下垂、瞳孔散大、上下内転制限）",
    "　左CN VI麻痺（外転制限）、左V1領域の感覚低下",
  ];
  lines.forEach(function(line, i) {
    s.addText(line, { x: 1.0, y: 2.1 + i * 0.4, w: 8, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 4.55, w: 7, h: 0.02, fill: { color: C.accent } });
  s.addText("この症例、どう考える？", { x: 0.5, y: 4.65, w: 9, h: 0.45, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0 });
})();

// ============================================================
// SLIDE 3: 有痛性眼筋麻痺とは
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "有痛性眼筋麻痺（Painful Ophthalmoplegia）とは");
  addPageNum(s, 3);

  addCard(s, 0.5, 1.2, 9, 1.5);
  s.addText("定義", { x: 0.8, y: 1.3, w: 2, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("眼窩部・眼窩周囲の疼痛 ＋ 同側の眼球運動障害（CN III, IV, VI 麻痺）を呈する症候群", {
    x: 0.8, y: 1.8, w: 8.4, h: 0.7, fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  s.addText("主な原因疾患", { x: 0.8, y: 3.0, w: 3, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var causes = [
    { cat: "炎症性", items: "Tolosa-Hunt症候群\nサルコイドーシス\nIgG4関連疾患", color: C.primary },
    { cat: "腫瘍性", items: "髄膜腫\nリンパ腫\n転移性腫瘍", color: C.red },
    { cat: "血管性", items: "動脈瘤\nCCF\n海綿静脈洞血栓", color: C.orange },
    { cat: "感染性", items: "真菌症\n帯状疱疹\n梅毒", color: C.green },
  ];
  causes.forEach(function(c, i) {
    var xPos = 0.5 + i * 2.35;
    addCard(s, xPos, 3.5, 2.15, 2.0);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.5, w: 2.15, h: 0.45, fill: { color: c.color } });
    s.addText(c.cat, { x: xPos, y: 3.5, w: 2.15, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(c.items, { x: xPos + 0.1, y: 4.05, w: 1.95, h: 1.35, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0 });
  });
})();

// ============================================================
// SLIDE 4: 海綿静脈洞の解剖
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "海綿静脈洞の解剖 ― なぜ複数の脳神経が障害されるのか");
  addPageNum(s, 4);

  // 左側：模式図的な構造の説明
  addCard(s, 0.3, 1.2, 4.5, 4.2);
  s.addText("海綿静脈洞の断面構造", { x: 0.5, y: 1.3, w: 4, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  // 外側壁の神経を上から順にカラーボックスで表示
  var wallNerves = [
    { name: "CN III（動眼神経）", y: 2.0, color: C.primary },
    { name: "CN IV（滑車神経）", y: 2.5, color: C.accent },
    { name: "V1（眼神経）", y: 3.0, color: "6C5CE7" },
    { name: "V2（上顎神経）", y: 3.5, color: "A29BFE" },
  ];

  s.addText("外側壁", { x: 0.6, y: 1.8, w: 1.5, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.sub, margin: 0 });
  wallNerves.forEach(function(n) {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: n.y, w: 3.2, h: 0.4, fill: { color: n.color }, rectRadius: 0.05 });
    s.addText(n.name, { x: 0.6, y: n.y, w: 3.2, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
  });

  s.addText("洞内", { x: 0.6, y: 4.1, w: 1.5, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.sub, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.35, w: 1.5, h: 0.4, fill: { color: C.red }, rectRadius: 0.05 });
  s.addText("内頸動脈", { x: 0.6, y: 4.35, w: 1.5, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.3, y: 4.35, w: 1.5, h: 0.4, fill: { color: C.orange }, rectRadius: 0.05 });
  s.addText("CN VI", { x: 2.3, y: 4.35, w: 1.5, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });

  // 右側：臨床的意義
  addCard(s, 5.2, 1.2, 4.5, 4.2);
  s.addText("臨床的意義", { x: 5.4, y: 1.3, w: 4, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var points = [
    "狭い空間に多数の脳神経が密集",
    "炎症・腫瘍で複数の神経が同時障害",
    "CN III + VI + V1 の複合障害\n→ 海綿静脈洞病変を強く示唆",
    "交感神経叢も通過\n→ Horner症候群を合併しうる",
    "上眼窩裂・眼窩先端部に連続\n→ 病変の進展範囲で症状が変化",
  ];
  points.forEach(function(p, i) {
    s.addText("●", { x: 5.4, y: 1.95 + i * 0.7, w: 0.3, h: 0.35, fontSize: 12, fontFace: FONT_EN, color: C.primary, margin: 0 });
    s.addText(p, { x: 5.7, y: 1.95 + i * 0.7, w: 3.8, h: 0.65, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });
  });
})();

// ============================================================
// SLIDE 5: Tolosa-Hunt症候群の病態
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "Tolosa-Hunt症候群の病態生理");
  addPageNum(s, 5);

  addCard(s, 0.5, 1.2, 9, 1.6);
  s.addText("定義", { x: 0.8, y: 1.3, w: 2, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("海綿静脈洞・上眼窩裂・眼窩先端部における原因不明の肉芽腫性炎症により、\n片側性の眼窩痛と眼筋麻痺を呈する疾患（1954年 Tolosa、1961年 Hunt ら報告）", {
    x: 0.8, y: 1.75, w: 8.4, h: 0.9, fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  s.addText("病理組織学的特徴", { x: 0.8, y: 3.1, w: 4, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var pathFeatures = [
    { icon: "1", text: "線維芽細胞の増殖", color: C.primary },
    { icon: "2", text: "リンパ球・形質細胞の浸潤", color: C.accent },
    { icon: "3", text: "肉芽腫性変化（類上皮細胞・巨細胞）", color: "6C5CE7" },
    { icon: "4", text: "硬膜の肥厚", color: C.orange },
  ];
  pathFeatures.forEach(function(f, i) {
    var xPos = 0.5 + (i % 2) * 4.6;
    var yPos = 3.6 + Math.floor(i / 2) * 0.9;
    addCard(s, xPos, yPos, 4.3, 0.75);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.15, y: yPos + 0.12, w: 0.5, h: 0.5, fill: { color: f.color } });
    s.addText(f.icon, { x: xPos + 0.15, y: yPos + 0.12, w: 0.5, h: 0.5, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.text, { x: xPos + 0.8, y: yPos + 0.1, w: 3.3, h: 0.55, fontSize: 16, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });
})();

// ============================================================
// SLIDE 6: 疫学データ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "疫学 ― 極めて稀な疾患");
  addPageNum(s, 6);

  var stats = [
    { num: "1/100万", desc: "年間発症率", color: C.primary },
    { num: "41歳", desc: "平均発症年齢", color: C.accent },
    { num: "1 : 1", desc: "男女比（性差なし）", color: "6C5CE7" },
    { num: "~5%", desc: "両側性の割合", color: C.orange },
  ];
  stats.forEach(function(item, i) {
    var xPos = 0.3 + i * 2.4;
    addCard(s, xPos, 1.2, 2.2, 2.0);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.35, y: 1.4, w: 1.5, h: 0.8, fill: { color: item.color }, rectRadius: 0.1 });
    s.addText(item.num, { x: xPos + 0.35, y: 1.4, w: 1.5, h: 0.8, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(item.desc, { x: xPos, y: 2.35, w: 2.2, h: 0.7, fontSize: 14, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
  });

  addCard(s, 0.5, 3.5, 9, 2.0);
  s.addText("疫学的特徴", { x: 0.8, y: 3.6, w: 3, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  var epiPoints = [
    "あらゆる年齢で発症するが、小児では稀",
    "人種差・地域差なし（世界中で報告あり）",
    "若年患者では再発頻度が高い傾向",
    "40〜50% が再発を経験する",
  ];
  epiPoints.forEach(function(p, i) {
    s.addText("・ " + p, { x: 0.8, y: 4.1 + i * 0.35, w: 8.4, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0 });
  });
})();

// ============================================================
// SLIDE 7: 臨床症状
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "臨床症状 ― 疼痛 → 眼筋麻痺");
  addPageNum(s, 7);

  // 疼痛の特徴
  addCard(s, 0.3, 1.2, 4.5, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 4.5, h: 0.4, fill: { color: C.red } });
  s.addText("疼痛の特徴", { x: 0.3, y: 1.2, w: 4.5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });

  var painFeatures = [
    "部位：眼窩周囲・眼窩後方",
    "性質：鋭い・刺すような・えぐるような",
    "前頭部・側頭部に放散",
    "眼筋麻痺に最大30日先行",
  ];
  painFeatures.forEach(function(p, i) {
    s.addText("● " + p, { x: 0.5, y: 1.75 + i * 0.35, w: 4, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  // 神経障害の頻度
  addCard(s, 5.2, 1.2, 4.5, 4.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.5, h: 0.4, fill: { color: C.primary } });
  s.addText("障害される神経と頻度", { x: 5.2, y: 1.2, w: 4.5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });

  var nerves = [
    { name: "CN III（動眼神経）", pct: "~80%", w: 3.4 },
    { name: "CN VI（外転神経）", pct: "~70%", w: 3.0 },
    { name: "V1（眼神経）", pct: "~30%", w: 1.3 },
    { name: "CN IV（滑車神経）", pct: "~29%", w: 1.2 },
    { name: "Horner症候群", pct: "~20%", w: 0.85 },
  ];
  nerves.forEach(function(n, i) {
    var yPos = 1.75 + i * 0.6;
    s.addText(n.name, { x: 5.4, y: yPos, w: 2.8, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: yPos + 0.28, w: n.w, h: 0.18, fill: { color: C.accent }, rectRadius: 0.05 });
    s.addText(n.pct, { x: 5.4 + n.w + 0.1, y: yPos + 0.2, w: 0.8, h: 0.3, fontSize: 12, fontFace: FONT_EN, color: C.primary, bold: true, margin: 0 });
  });

  // 時間経過
  addCard(s, 0.3, 3.5, 4.5, 1.9);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.5, w: 4.5, h: 0.4, fill: { color: C.orange } });
  s.addText("時間経過", { x: 0.3, y: 3.5, w: 4.5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });

  s.addText("眼窩痛", { x: 0.5, y: 4.1, w: 1.2, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText("→", { x: 1.7, y: 4.1, w: 0.4, h: 0.4, fontSize: 18, fontFace: FONT_EN, color: C.sub, align: "center", margin: 0 });
  s.addText("眼筋麻痺", { x: 2.1, y: 4.1, w: 1.5, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("（2週間以内）", { x: 3.5, y: 4.1, w: 1.2, h: 0.4, fontSize: 12, fontFace: FONT_JP, color: C.sub, margin: 0 });
  s.addText("未治療で最大8週間持続", { x: 0.5, y: 4.6, w: 4, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  s.addText("自然寛解もあり得る", { x: 0.5, y: 4.95, w: 4, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
})();

// ============================================================
// SLIDE 8: ICHD-3診断基準
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ICHD-3 診断基準");
  addPageNum(s, 8);

  var rows = [
    [
      { text: "基準", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 16, fontFace: FONT_JP, align: "center" } },
      { text: "内容", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 16, fontFace: FONT_JP } },
    ],
    [
      { text: "A", options: { bold: true, fontSize: 16, fontFace: FONT_EN, color: C.primary, align: "center" } },
      { text: "片側性の頭痛（同側の眉部・眼周囲に局在）", options: { fontSize: 15, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "B-1", options: { bold: true, fontSize: 16, fontFace: FONT_EN, color: C.primary, align: "center" } },
      { text: "頭痛と同側のCN III, IV, VI のうち1つ以上の麻痺", options: { fontSize: 15, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "B-2", options: { bold: true, fontSize: 16, fontFace: FONT_EN, color: C.primary, align: "center" } },
      { text: "MRI または生検で海綿静脈洞・上眼窩裂・眼窩の\n肉芽腫性炎症を証明", options: { fontSize: 15, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "C", options: { bold: true, fontSize: 16, fontFace: FONT_EN, color: C.primary, align: "center" } },
      { text: "頭痛が麻痺に2週間以内に先行、または同時に発症", options: { fontSize: 15, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "D", options: { bold: true, fontSize: 16, fontFace: FONT_EN, color: C.primary, align: "center" } },
      { text: "他の疾患では説明できない（除外診断）", options: { fontSize: 15, fontFace: FONT_JP, color: C.text } },
    ],
  ];
  s.addTable(rows, {
    x: 0.5, y: 1.2, w: 9, colW: [1.2, 7.8],
    border: { type: "solid", pt: 1, color: "DEE2E6" },
    rowH: [0.45, 0.5, 0.5, 0.65, 0.5, 0.5],
  });

  // 注意ボックス
  addCard(s, 0.5, 4.35, 9, 1.05);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.35, w: 0.15, h: 1.05, fill: { color: C.yellow } });
  s.addText("注意", { x: 0.8, y: 4.4, w: 1.0, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0 });
  s.addText("旧基準（ICHD-2）の「ステロイド反応性」は ICHD-3 で除外された\n→ リンパ腫やサルコイドーシスもステロイドに反応するため", {
    x: 1.8, y: 4.4, w: 7.5, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
  });
})();

// ============================================================
// SLIDE 9: MRI所見
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "MRI所見のポイント");
  addPageNum(s, 9);

  // 典型所見
  addCard(s, 0.3, 1.2, 5.0, 3.4);
  s.addText("典型的なMRI所見", { x: 0.5, y: 1.3, w: 4.5, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var mriFindings = [
    { label: "T1", detail: "等信号の軟部組織腫瘤", color: C.primary },
    { label: "T2", detail: "等信号〜低信号", color: C.accent },
    { label: "Gd", detail: "著明な増強効果（最重要）", color: C.red },
  ];
  mriFindings.forEach(function(f, i) {
    var yPos = 1.95 + i * 0.55;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yPos, w: 0.7, h: 0.4, fill: { color: f.color }, rectRadius: 0.05 });
    s.addText(f.label, { x: 0.6, y: yPos, w: 0.7, h: 0.4, fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(f.detail, { x: 1.5, y: yPos, w: 3.5, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  var additionalMRI = [
    "海綿静脈洞外側壁の肥厚・外方膨隆",
    "眼窩先端部への進展",
    "内頸動脈の分節的狭窄",
  ];
  additionalMRI.forEach(function(p, i) {
    s.addText("● " + p, { x: 0.6, y: 3.65 + i * 0.32, w: 4.5, h: 0.32, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  // 限界
  addCard(s, 5.5, 1.2, 4.2, 4.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 1.2, w: 4.2, h: 0.4, fill: { color: C.yellow } });
  s.addText("MRIの限界", { x: 5.5, y: 1.2, w: 4.2, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, align: "center", margin: 0 });

  s.addText("18〜57%", { x: 6.2, y: 1.8, w: 2.8, h: 0.65, fontSize: 32, fontFace: FONT_EN, color: C.red, bold: true, align: "center", margin: 0 });
  s.addText("で正常MRI", { x: 6.2, y: 2.45, w: 2.8, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.text, align: "center", margin: 0 });

  var limits = [
    "3T MRIでも1mm未満の\n病変は検出困難",
    "所見は非特異的\n（リンパ腫・髄膜腫と類似）",
    "Gd造影は必須\n（造影なしでは見逃す）",
    "薄スライス冠状断が有用",
  ];
  limits.forEach(function(l, i) {
    s.addText("● " + l, { x: 5.7, y: 3.1 + i * 0.55, w: 3.8, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });
  });
})();

// ============================================================
// SLIDE 10: 鑑別診断一覧（カテゴリ別）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "鑑別診断 ― 除外すべき疾患群");
  addPageNum(s, 10);

  var categories = [
    { title: "腫瘍性", items: "髄膜腫 / リンパ腫\n転移性腫瘍 / 下垂体腫瘍\n頭蓋底腫瘍", color: C.red, x: 0.3, y: 1.15, w: 4.5, h: 1.9 },
    { title: "血管性", items: "脳動脈瘤（ICA-Pcom）\nCCF（頸動脈海綿静脈洞瘻）\n海綿静脈洞血栓症\n頸動脈解離", color: C.orange, x: 5.2, y: 1.15, w: 4.5, h: 1.9 },
    { title: "感染性", items: "真菌症（ムコール・アスペルギルス）\n結核性髄膜炎 / 帯状疱疹\n梅毒 / ライム病", color: C.green, x: 0.3, y: 3.25, w: 4.5, h: 1.9 },
    { title: "炎症性・自己免疫", items: "サルコイドーシス / IgG4関連疾患\n肥厚性硬膜炎\nANCA関連血管炎 / SLE / PAN", color: "6C5CE7", x: 5.2, y: 3.25, w: 4.5, h: 1.9 },
  ];

  categories.forEach(function(cat) {
    addCard(s, cat.x, cat.y, cat.w, cat.h);
    s.addShape(pres.shapes.RECTANGLE, { x: cat.x, y: cat.y, w: cat.w, h: 0.4, fill: { color: cat.color } });
    s.addText(cat.title, { x: cat.x, y: cat.y, w: cat.w, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(cat.items, { x: cat.x + 0.2, y: cat.y + 0.5, w: cat.w - 0.4, h: cat.h - 0.6, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0 });
  });
})();

// ============================================================
// SLIDE 11: 鑑別のRed Flags
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "鑑別のRed Flags ― 見逃してはならないサイン");
  addPageNum(s, 11);

  var flags = [
    { sign: "急速進行＋免疫不全", dx: "真菌感染（ムコール症）", action: "緊急CT/MRI", color: C.red },
    { sign: "CN III麻痺＋瞳孔散大", dx: "動脈瘤による圧迫", action: "緊急MRA/CTA", color: C.red },
    { sign: "発熱＋両側性進展", dx: "海綿静脈洞血栓症", action: "造影CT/MRV", color: C.orange },
    { sign: "ステロイド後に再増悪", dx: "悪性リンパ腫", action: "生検・PET-CT", color: C.orange },
    { sign: "進行性の視力低下", dx: "眼窩先端部の腫瘍", action: "眼科コンサルト", color: C.yellow },
  ];

  flags.forEach(function(f, i) {
    var yPos = 1.1 + i * 0.72;
    addCard(s, 0.3, yPos, 9.4, 0.62);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.15, h: 0.62, fill: { color: f.color } });
    s.addText(f.sign, { x: 0.6, y: yPos + 0.03, w: 2.8, h: 0.56, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: 0 });
    s.addText("→ " + f.dx, { x: 3.5, y: yPos + 0.03, w: 2.8, h: 0.56, fontSize: 13, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(f.action, { x: 6.4, y: yPos + 0.03, w: 3.1, h: 0.56, fontSize: 12, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0 });
  });

  // ポイント
  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 4.65, w: 7, h: 0.02, fill: { color: C.red } });
  s.addText("THSは除外診断 → これらを系統的に否定してから診断する", {
    x: 0.5, y: 4.75, w: 9, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 12: 検査アプローチ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "検査アプローチ");
  addPageNum(s, 12);

  // 必須検査
  addCard(s, 0.3, 1.2, 4.5, 4.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 4.5, h: 0.4, fill: { color: C.primary } });
  s.addText("必須検査", { x: 0.3, y: 1.2, w: 4.5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });

  var essential = [
    { cat: "画像", items: "頭部MRI（Gd造影）\nMRA / CTA" },
    { cat: "血液", items: "CBC、ESR、CRP\n血糖 / HbA1c\n甲状腺機能\nACE、IgG4\nANA、ANCA\nRPR / TPHA" },
  ];

  var yOff = 1.8;
  essential.forEach(function(e) {
    s.addText(e.cat, { x: 0.5, y: yOff, w: 1.2, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
    var lines = e.items.split("\n");
    lines.forEach(function(line, li) {
      s.addText("・ " + line, { x: 1.5, y: yOff + li * 0.3, w: 3, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });
    });
    yOff += lines.length * 0.3 + 0.3;
  });

  // 追加検査
  addCard(s, 5.2, 1.2, 4.5, 4.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.5, h: 0.4, fill: { color: C.accent } });
  s.addText("状況に応じた追加検査", { x: 5.2, y: 1.2, w: 4.5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });

  var additional = [
    { test: "髄液検査", detail: "細胞数、蛋白、糖、培養、細胞診", when: "感染・炎症疑い" },
    { test: "胸部CT", detail: "BHL、肺病変", when: "サルコイドーシス疑い" },
    { test: "PET-CT", detail: "全身検索", when: "悪性腫瘍疑い" },
    { test: "生検", detail: "海綿静脈洞生検（リスク大）", when: "確定診断の最終手段" },
  ];

  additional.forEach(function(a, i) {
    var yPos = 1.8 + i * 0.9;
    s.addText(a.test, { x: 5.4, y: yPos, w: 2, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
    s.addText(a.detail, { x: 5.4, y: yPos + 0.35, w: 2.5, h: 0.35, fontSize: 11, fontFace: FONT_JP, color: C.text, margin: 0 });
    s.addText(a.when, { x: 7.9, y: yPos, w: 1.6, h: 0.35, fontSize: 11, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0 });
  });
})();

// ============================================================
// SLIDE 13: その他の鑑別（糖尿病性眼筋麻痺 etc.）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "その他の重要な鑑別疾患");
  addPageNum(s, 13);

  var ddx = [
    { name: "糖尿病性眼筋麻痺", points: "CN III単独（瞳孔温存が多い）またはCN VI単独\n疼痛は軽度〜中等度、数週で自然改善", color: C.orange },
    { name: "眼筋麻痺性片頭痛\n（RPON）", points: "反復する片頭痛の既往、若年発症\nMRIでCN III造影効果", color: C.accent },
    { name: "甲状腺眼症", points: "外眼筋肥大・眼球突出\n甲状腺機能異常（バセドウ病）", color: C.green },
    { name: "眼窩偽腫瘍", points: "眼窩内軟部組織腫瘤、眼球突出\nTHS類似だが海綿静脈洞外の病変", color: "6C5CE7" },
  ];

  ddx.forEach(function(d, i) {
    var yPos = 1.15 + i * 1.1;
    addCard(s, 0.3, yPos, 9.4, 0.95);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.15, h: 0.95, fill: { color: d.color } });
    s.addText(d.name, { x: 0.6, y: yPos + 0.05, w: 2.5, h: 0.85, fontSize: 15, fontFace: FONT_JP, color: d.color, bold: true, valign: "middle", margin: 0 });
    s.addText(d.points, { x: 3.2, y: yPos + 0.05, w: 6.3, h: 0.85, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });
})();

// ============================================================
// SLIDE 14: 治療プロトコル
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "治療 ― ステロイドが第一選択");
  addPageNum(s, 14);

  // ステロイド治療
  addCard(s, 0.3, 1.2, 5.5, 4.2);
  s.addText("第一選択：副腎皮質ステロイド", { x: 0.5, y: 1.3, w: 5, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var txRows = [
    [
      { text: "項目", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 14, fontFace: FONT_JP } },
      { text: "内容", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 14, fontFace: FONT_JP } },
    ],
    [
      { text: "初期用量", options: { fontSize: 13, fontFace: FONT_JP, bold: true, color: C.text } },
      { text: "PSL 0.5〜1.0 mg/kg/日", options: { fontSize: 13, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "疼痛改善", options: { fontSize: 13, fontFace: FONT_JP, bold: true, color: C.text } },
      { text: "24〜72時間以内", options: { fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true } },
    ],
    [
      { text: "神経回復", options: { fontSize: 13, fontFace: FONT_JP, bold: true, color: C.text } },
      { text: "2〜8週間かけて改善", options: { fontSize: 13, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "漸減", options: { fontSize: 13, fontFace: FONT_JP, bold: true, color: C.text } },
      { text: "数週〜数ヶ月かけて漸減", options: { fontSize: 13, fontFace: FONT_JP, color: C.text } },
    ],
  ];
  s.addTable(txRows, {
    x: 0.5, y: 1.9, w: 5.1, colW: [1.5, 3.6],
    border: { type: "solid", pt: 1, color: "DEE2E6" },
    rowH: [0.4, 0.4, 0.4, 0.4, 0.4],
  });

  // 注意ボックス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.1, w: 5.1, h: 0.05, fill: { color: C.yellow } });
  s.addText("72時間で疼痛改善なし → 診断再考！", {
    x: 0.5, y: 4.2, w: 5.1, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, align: "center", margin: 0,
  });
  s.addText("※ステロイド反応 ≠ THS確定\n（リンパ腫・サルコイドーシスも反応）", {
    x: 0.5, y: 4.6, w: 5.1, h: 0.6, fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  // タイムライン
  addCard(s, 6.0, 1.2, 3.7, 4.2);
  s.addText("治療経過イメージ", { x: 6.2, y: 1.3, w: 3.3, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var timeline = [
    { time: "Day 1", event: "PSL開始", color: C.primary },
    { time: "Day 1-3", event: "疼痛改善", color: C.green },
    { time: "Week 2-8", event: "神経障害改善", color: C.accent },
    { time: "数ヶ月", event: "漸減〜中止", color: C.sub },
  ];
  timeline.forEach(function(t, i) {
    var yPos = 2.0 + i * 0.85;
    s.addShape(pres.shapes.OVAL, { x: 6.4, y: yPos, w: 0.5, h: 0.5, fill: { color: t.color } });
    s.addText((i + 1).toString(), { x: 6.4, y: yPos, w: 0.5, h: 0.5, fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    if (i < 3) {
      s.addShape(pres.shapes.LINE, { x: 6.65, y: yPos + 0.5, w: 0, h: 0.35, line: { color: "CCCCCC", width: 2, dashType: "dash" } });
    }
    s.addText(t.time, { x: 7.1, y: yPos - 0.05, w: 1.2, h: 0.3, fontSize: 13, fontFace: FONT_EN, color: t.color, bold: true, margin: 0 });
    s.addText(t.event, { x: 7.1, y: yPos + 0.2, w: 2.3, h: 0.3, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  });
})();

// ============================================================
// SLIDE 15: 二次治療
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "難治例・ステロイド依存例の治療");
  addPageNum(s, 15);

  s.addText("ステロイド減量で再燃、またはステロイド禁忌の場合", {
    x: 0.5, y: 1.1, w: 9, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  var agents = [
    { name: "アザチオプリン", type: "免疫抑制薬", note: "従来から使用" },
    { name: "メトトレキサート", type: "免疫抑制薬", note: "低用量で使用" },
    { name: "ミコフェノール酸モフェチル", type: "免疫抑制薬", note: "比較的新しい選択肢" },
    { name: "シクロスポリン", type: "カルシニューリン阻害薬", note: "腎毒性に注意" },
    { name: "インフリキシマブ\n/アダリムマブ", type: "TNF阻害薬", note: "難治例に有効との報告あり" },
    { name: "放射線療法", type: "最終手段", note: "症例報告レベル" },
  ];

  agents.forEach(function(a, i) {
    var col = i % 3;
    var row = Math.floor(i / 3);
    var xPos = 0.3 + col * 3.2;
    var yPos = 1.7 + row * 2.05;
    addCard(s, xPos, yPos, 3.0, 1.85);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 3.0, h: 0.4, fill: { color: C.accent } });
    s.addText(a.name, { x: xPos + 0.1, y: yPos + 0.45, w: 2.8, h: 0.65, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(a.type, { x: xPos, y: yPos, w: 3.0, h: 0.4, fontSize: 12, fontFace: FONT_JP, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText(a.note, { x: xPos + 0.2, y: yPos + 1.15, w: 2.6, h: 0.55, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", valign: "middle", margin: 0 });
  });
})();

// ============================================================
// SLIDE 16: 予後と再発
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "予後と再発 ― 油断は禁物");
  addPageNum(s, 16);

  // 予後
  addCard(s, 0.3, 1.2, 4.5, 2.5);
  s.addText("予後", { x: 0.5, y: 1.3, w: 3, h: 0.5, fontSize: 20, fontFace: FONT_JP, color: C.green, bold: true, margin: 0 });

  var progPoints = [
    "基本的に予後良好（self-limiting）",
    "ステロイド治療後の残存神経障害は稀",
    "自然寛解もあり（8週間以内）",
    "致死的疾患ではない",
  ];
  progPoints.forEach(function(p, i) {
    s.addText("✓ " + p, { x: 0.6, y: 1.9 + i * 0.4, w: 4, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  // 再発
  addCard(s, 5.2, 1.2, 4.5, 2.5);
  s.addText("再発", { x: 5.4, y: 1.3, w: 3, h: 0.5, fontSize: 20, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });

  s.addText("40〜50%", { x: 5.8, y: 1.9, w: 3.5, h: 0.8, fontSize: 36, fontFace: FONT_EN, color: C.red, bold: true, align: "center", margin: 0 });
  s.addText("が再発を経験", { x: 5.8, y: 2.6, w: 3.5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.text, align: "center", margin: 0 });

  var relapsePoints = [
    "同側・対側・両側いずれも",
    "若年者で再発頻度が高い",
  ];
  relapsePoints.forEach(function(p, i) {
    s.addText("● " + p, { x: 5.6, y: 3.1 + i * 0.3, w: 3.8, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  // 重要な注意
  addCard(s, 0.3, 4.0, 9.4, 1.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.0, w: 9.4, h: 0.4, fill: { color: C.red } });
  s.addText("再発時の鉄則", { x: 0.3, y: 4.0, w: 9.4, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText("再発のたびに完全な鑑別診断ワークアップを繰り返す\n→ THSは除外診断であるため、再発時に腫瘍などが新たに発見される可能性がある", {
    x: 0.5, y: 4.5, w: 9, h: 0.8, fontSize: 15, fontFace: FONT_JP, color: C.text, align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 17: Take Home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  addPageNum(s, 17);

  s.addText("Take Home Message", {
    x: 0.5, y: 0.3, w: 9, h: 0.7, fontSize: 28, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "THSは除外診断 ― 腫瘍・血管・感染・炎症を系統的に除外", color: C.primary },
    { num: "2", text: "ICHD-3 でステロイド反応性は診断基準から除外された", color: C.accent },
    { num: "3", text: "MRI正常（18-57%）でもTHSは否定できない", color: "6C5CE7" },
    { num: "4", text: "ステロイドで24-72時間以内に疼痛改善 → 効かなければ再考", color: C.orange },
    { num: "5", text: "40-50%が再発 → 毎回フルワークアップが必要", color: C.red },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.3 + i * 0.8;
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos, w: 0.55, h: 0.55, fill: { color: m.color } });
    s.addText(m.num, { x: 1.0, y: yPos, w: 0.55, h: 0.55, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.5, h: 0.55, fontSize: 18, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.1, w: 6, h: 0, line: { color: C.accent, width: 1.5 } });
  s.addText("医知創造ラボ", {
    x: 0.5, y: 5.15, w: 9, h: 0.35, fontSize: 16, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
})();

// ============================================================
// Output
// ============================================================
var outPath = __dirname + "/Tolosa-Hunt症候群_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
