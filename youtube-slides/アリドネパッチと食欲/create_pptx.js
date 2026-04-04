var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "アリドネパッチで食欲は改善するのか？";

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
  green: "28A745",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var TOTAL = 10;

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
}

function addCard(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: C.white }, shadow: makeShadow() });
}

function addPageNum(s, n) {
  s.addText(n + "/" + TOTAL, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: Title
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("アリドネパッチで", {
    x: 0.8, y: 0.8, w: 8.4, h: 0.8,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("食欲は改善するのか？", {
    x: 0.8, y: 1.5, w: 8.4, h: 1.0,
    fontSize: 44, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });
  s.addText("ドネペジル貼付剤と食欲・体重の実際", {
    x: 0.8, y: 2.6, w: 8.4, h: 0.7,
    fontSize: 22, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 3.5, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.8, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("最新エビデンスに基づく批判的吟味", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 今日のテーマ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "今日のテーマ");

  addCard(s, 0.5, 1.2, 9, 1.8);
  s.addText("アリドネパッチで食欲が改善した", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_JP, color: C.text, align: "center", margin: 0,
  });
  s.addText("という報告があるが、本当か？", {
    x: 0.8, y: 1.9, w: 8.4, h: 0.8,
    fontSize: 28, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0,
  });

  var topics = [
    { icon: "\uD83D\uDC8A", text: "ドネペジルは食欲を上げる薬？" },
    { icon: "\uD83D\uDCC4", text: "食欲改善報告の中身と限界" },
    { icon: "\u2696\uFE0F", text: "貼付剤の消化器副作用の実態" },
    { icon: "\uD83C\uDFE5", text: "実臨床での判断基準" },
  ];

  topics.forEach(function(t, i) {
    var yPos = 3.3 + i * 0.55;
    addCard(s, 1.0, yPos, 8.0, 0.5);
    s.addText(t.icon + "  " + t.text, {
      x: 1.2, y: yPos + 0.05, w: 7.6, h: 0.4,
      fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  addPageNum(s, 2);
})();

// ============================================================
// SLIDE 3: ドネペジルと食欲 ― そもそもの前提
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ドネペジルと食欲 \u2015 そもそもの前提");

  addCard(s, 0.5, 1.2, 9, 1.2);
  s.addText("ドネペジル（アリセプト）は食欲を「下げる」方向の薬", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });
  s.addText("電子添文に食欲不振・悪心・嘔吐・下痢・体重減少を記載\n3mg開始は「消化器副作用を抑えるための非有効量」", {
    x: 0.8, y: 1.8, w: 8.4, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // Meta-analysis data
  addCard(s, 0.5, 2.7, 4.3, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.7, w: 4.3, h: 0.5, fill: { color: C.primary } });
  s.addText("Soysal 2016 \u30E1\u30BF\u30A2\u30CA\u30EA\u30B7\u30B9", {
    x: 0.5, y: 2.7, w: 4.3, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("ChEI\u4F7F\u7528\u306F\u8A8D\u77E5\u75C7\u9AD8\u9F62\u8005\u306E\n\u4F53\u91CD\u6E1B\u5C11\u3068\u6709\u610F\u306B\u95A2\u9023", {
    x: 0.7, y: 3.3, w: 3.9, h: 0.8,
    fontSize: 16, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.3,
  });
  s.addText("JNNP 2016;87:1368-74", {
    x: 0.7, y: 4.2, w: 3.9, h: 0.3,
    fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "center", italic: true, margin: 0,
  });

  // VA Cohort data
  addCard(s, 5.2, 2.7, 4.3, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.7, w: 4.3, h: 0.5, fill: { color: C.orange } });
  s.addText("Sheffrin 2015 VA\u30B3\u30DB\u30FC\u30C8", {
    x: 5.2, y: 2.7, w: 4.3, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  var vaData = [
    { label: "ChEI\u7FA4", value: "29.3%", color: C.red },
    { label: "\u5BFE\u7167\u7FA4", value: "22.8%", color: C.green },
    { label: "HR", value: "1.23", color: C.primary },
  ];
  vaData.forEach(function(d, i) {
    var yPos = 3.3 + i * 0.55;
    s.addText(d.label, { x: 5.5, y: yPos, w: 2.0, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.sub, margin: 0 });
    s.addText(d.value, { x: 7.5, y: yPos, w: 1.8, h: 0.4, fontSize: 20, fontFace: FONT_EN, color: d.color, bold: true, align: "center", margin: 0 });
  });
  s.addText("JAGS 2015;63:1512-8", {
    x: 5.5, y: 4.8, w: 3.7, h: 0.3,
    fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "center", italic: true, margin: 0,
  });

  addPageNum(s, 3);
})();

// ============================================================
// SLIDE 4: 食欲改善報告 ― 吉元2025年論文
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u98DF\u6B32\u6539\u5584\u5831\u544A \u2015 \u5409\u5143 2025\u5E74\u8AD6\u6587");

  addCard(s, 0.5, 1.2, 9, 0.7);
  s.addText("\u300E\u8A3A\u7642\u3068\u65B0\u85AC\u300F2025;62(10):701-5\u3000\u5358\u65BD\u8A2D\u30FB\u5F8C\u308D\u5411\u304D\u89B3\u5BDF\u7814\u7A76\u3000n=40", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  // Results table
  var rows = [
    { item: "\u5BFE\u8C61", result: "AD 40\u4F8B\uFF08\u524D\u6CBB\u7642\u85AC\u306A\u3057\uFF09", bg: C.primary, fc: C.white },
    { item: "MMSE", result: "24\u9031\u3067\u6709\u610F\u5909\u5316\u306A\u3057", bg: C.warmBg, fc: C.text },
    { item: "SNAQ", result: "6.7 \u2192 17.2\uFF0816\u9031\u5F8C\uFF09", bg: C.white, fc: C.red },
    { item: "\u98DF\u4E8B\u6442\u53D6\u91CF", result: "157.9%\uFF0816\u9031\u5F8C\uFF09", bg: C.warmBg, fc: C.red },
    { item: "\u526F\u4F5C\u7528", result: "\u76AE\u819A\u969C\u5BB3 12.5%\u3001\u80C3\u8178\u969C\u5BB3 0%", bg: C.white, fc: C.text },
  ];

  rows.forEach(function(r, i) {
    var yPos = 2.1 + i * 0.6;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 3.5, h: 0.55, fill: { color: i === 0 ? C.primary : C.light } });
    s.addText(r.item, {
      x: 0.7, y: yPos + 0.05, w: 3.1, h: 0.45,
      fontSize: 16, fontFace: FONT_JP, color: i === 0 ? C.white : C.primary, bold: true, margin: 0,
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.0, y: yPos, w: 5.5, h: 0.55, fill: { color: r.bg } });
    s.addText(r.result, {
      x: 4.2, y: yPos + 0.05, w: 5.1, h: 0.45,
      fontSize: 16, fontFace: FONT_JP, color: r.fc, bold: i > 1 && i < 4, margin: 0,
    });
  });

  // Note
  addCard(s, 0.5, 5.0, 9, 0.55);
  s.addText("\u2192 \u8457\u8005\u306F\u300C\u98DF\u6B32\u6539\u5584\u304C\u8A8D\u3081\u3089\u308C\u305F\u300D\u3068\u7D50\u8AD6", {
    x: 0.8, y: 5.05, w: 8.4, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  addPageNum(s, 4);
})();

// ============================================================
// SLIDE 5: この研究の限界
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u3053\u306E\u7814\u7A76\u306E\u9650\u754C \u2015 \u305D\u306E\u307E\u307E\u4FE1\u3058\u3066\u3088\u3044\u306E\u304B\uFF1F");

  var limits = [
    { icon: "\u26A0\uFE0F", title: "\u5BFE\u7167\u7FA4\u304C\u306A\u3044", desc: "\u7D4C\u53E3\u7FA4\u30FB\u672A\u6CBB\u7642\u7FA4\u3068\u306E\u6BD4\u8F03\u304C\u306A\u304F\u3001\n\u85AC\u306E\u52B9\u679C\u304B\u81EA\u7136\u7D4C\u904E\u304B\u533A\u5225\u3067\u304D\u306A\u3044" },
    { icon: "\uD83D\uDCCA", title: "\u5C11\u6570\u4F8B", desc: "SNAQ\u89E3\u6790 18\u4F8B\u3001\u98DF\u4E8B\u6442\u53D6\u91CF 10\u4F8B\n\u2192 \u7D71\u8A08\u7684\u691C\u51FA\u529B\u304C\u4F4E\u3044" },
    { icon: "\uD83C\uDFE2", title: "\u5F8C\u308D\u5411\u304D\u30FB\u5358\u65BD\u8A2D", desc: "\u75C7\u4F8B\u9078\u629E\u30D0\u30A4\u30A2\u30B9\u3092\u6392\u9664\u3067\u304D\u306A\u3044" },
    { icon: "\uD83D\uDDE3\uFE0F", title: "\u4EA4\u7D61\u56E0\u5B50", desc: "\u30B9\u30BF\u30C3\u30D5\u306E\u58F0\u304B\u3051\u3001\u4ECB\u8B77\u74B0\u5883\u306E\u5909\u5316\u304C\n\u98DF\u6B32\u306B\u5F71\u97FF\u3057\u305F\u53EF\u80FD\u6027" },
  ];

  limits.forEach(function(l, i) {
    var yPos = 1.2 + i * 1.05;
    addCard(s, 0.5, yPos, 9, 0.95);
    s.addText(l.icon, { x: 0.7, y: yPos + 0.1, w: 0.6, h: 0.7, fontSize: 28, margin: 0 });
    s.addText(l.title, { x: 1.4, y: yPos + 0.1, w: 2.5, h: 0.7, fontSize: 18, fontFace: FONT_JP, color: C.red, bold: true, margin: 0, valign: "middle" });
    s.addText(l.desc, { x: 4.0, y: yPos + 0.1, w: 5.3, h: 0.7, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, 5);
})();

// ============================================================
// SLIDE 6: 貼付剤の消化器副作用の実態
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u8CBC\u4ED8\u5264\u306E\u6D88\u5316\u5668\u526F\u4F5C\u7528\u306E\u5B9F\u614B");

  // Key message
  addCard(s, 0.5, 1.2, 9, 0.7);
  s.addText("\u300C\u8CBC\u308A\u85AC\u3060\u304B\u3089\u6D88\u5316\u5668\u526F\u4F5C\u7528\u306F\u8D77\u304D\u306A\u3044\u300D\u306F\u8AA4\u308A", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.5,
    fontSize: 22, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0,
  });

  // Comparison table header
  var tableY = 2.2;
  var cols = [
    { label: "\u6709\u5BB3\u4E8B\u8C61", w: 3.5, x: 0.5 },
    { label: "\u30D1\u30C3\u30C1\u7FA4", w: 2.8, x: 4.0 },
    { label: "\u7D4C\u53E3\u7FA4", w: 2.7, x: 6.8 },
  ];
  cols.forEach(function(c) {
    s.addShape(pres.shapes.RECTANGLE, { x: c.x, y: tableY, w: c.w, h: 0.5, fill: { color: C.primary } });
    s.addText(c.label, { x: c.x, y: tableY, w: c.w, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0, valign: "middle" });
  });

  var tRows = [
    { item: "\u80C3\u8178\u969C\u5BB3\u5168\u4F53", patch: "9.8%", oral: "8.4%", highlight: true },
    { item: "\u9069\u7528\u90E8\u4F4D\u7D05\u6591", patch: "24.9%", oral: "\u2015", highlight: false },
    { item: "\u9069\u7528\u90E8\u4F4D\u305D\u3046\u75D2\u611F", patch: "22.5%", oral: "\u2015", highlight: false },
    { item: "\u63A5\u89E6\u76AE\u819A\u708E", patch: "11.0%", oral: "\u2015", highlight: false },
  ];

  tRows.forEach(function(r, i) {
    var yPos = tableY + 0.5 + i * 0.5;
    var bg = i % 2 === 0 ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 3.5, h: 0.5, fill: { color: bg } });
    s.addText(r.item, { x: 0.7, y: yPos, w: 3.1, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.0, y: yPos, w: 2.8, h: 0.5, fill: { color: bg } });
    s.addText(r.patch, { x: 4.0, y: yPos, w: 2.8, h: 0.5, fontSize: 16, fontFace: FONT_EN, color: r.highlight ? C.red : C.text, bold: r.highlight, align: "center", margin: 0, valign: "middle" });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: yPos, w: 2.7, h: 0.5, fill: { color: bg } });
    s.addText(r.oral, { x: 6.8, y: yPos, w: 2.7, h: 0.5, fontSize: 16, fontFace: FONT_EN, color: r.highlight ? C.red : C.text, bold: r.highlight, align: "center", margin: 0, valign: "middle" });
  });

  // Takeaway
  addCard(s, 0.5, 4.7, 9, 0.8);
  s.addText("\u526F\u4F5C\u7528\u304C\u306A\u304F\u306A\u308B\u306E\u3067\u306F\u306A\u304F\u3001\u526F\u4F5C\u7528\u306E\u91CD\u5FC3\u304C\u76AE\u819A\u5074\u306B\u79FB\u308B\u88FD\u5264", {
    x: 0.8, y: 4.8, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0,
  });

  s.addText("Nakamura Y, et al. Geriatr Gerontol Int 2023", {
    x: 5.0, y: 5.2, w: 4.5, h: 0.3,
    fontSize: 10, fontFace: FONT_EN, color: C.sub, align: "right", italic: true, margin: 0,
  });

  addPageNum(s, 6);
})();

// ============================================================
// SLIDE 7: なぜ「食べるようになった」と見えるのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u306A\u305C\u300C\u98DF\u3079\u308B\u3088\u3046\u306B\u306A\u3063\u305F\u300D\u3068\u898B\u3048\u308B\u306E\u304B");

  // Key fact: study was all drug-naive, GI AE = 0
  addCard(s, 0.5, 1.1, 9, 0.9);
  s.addText("\u5409\u5143\u8AD6\u6587\u306F\u5168\u4F8B\u65B0\u898F\u958B\u59CB\uFF08\u524D\u6CBB\u7642\u85AC\u306A\u3057\uFF09\u30FB\u80C3\u8178\u969C\u5BB3 0\u4F8B", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.4,
    fontSize: 17, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0,
  });
  s.addText("\u2192 \u300C\u7D4C\u53E3\u5264\u306E\u6D88\u5316\u5668\u4E0D\u5FEB\u8EFD\u6E1B\u300D\u4EEE\u8AAC\u306F\u3053\u306E\u7814\u7A76\u306B\u306F\u5F53\u3066\u306F\u307E\u3089\u306A\u3044", {
    x: 0.8, y: 1.55, w: 8.4, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  // Author's proposed mechanisms
  s.addText("\u8457\u8005\u306E\u8003\u5BDF\uFF08\u85AC\u7406\u5B66\u7684\u4EEE\u8AAC\uFF09", {
    x: 0.5, y: 2.2, w: 4.3, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  var authored = [
    { icon: "\uD83D\uDC8A", text: "\u5687\u4E0B\u6A5F\u80FD\u306E\u6539\u5584\uFF08ACh\u6FC3\u5EA6\u2191\uFF09" },
    { icon: "\uD83D\uDD2C", text: "\u6297\u708E\u75C7\u52B9\u679C\uFF08\u30B5\u30A4\u30C8\u30AB\u30A4\u30F3\u6291\u5236\uFF09" },
    { icon: "\uD83D\uDC43", text: "\u55C5\u899A\u30FB\u5473\u899A\u306E\u6539\u5584" },
    { icon: "\uD83D\uDCC8", text: "\u8CBC\u4ED8\u5264\u306E\u7DD2\u3084\u304B\u306A\u8840\u4E2D\u6FC3\u5EA6\u4E0A\u6607" },
  ];

  authored.forEach(function(a, i) {
    var yPos = 2.7 + i * 0.5;
    addCard(s, 0.5, yPos, 4.3, 0.45);
    s.addText(a.icon + "  " + a.text, { x: 0.7, y: yPos + 0.03, w: 3.9, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  // Confounders
  s.addText("\u4EA4\u7D61\u56E0\u5B50\uFF08\u6392\u9664\u3067\u304D\u306A\u3044\uFF09", {
    x: 5.2, y: 2.2, w: 4.3, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });

  var confounders = [
    { icon: "\uD83C\uDFE5", text: "\u30B9\u30BF\u30C3\u30D5\u306E\u58F0\u304B\u3051\u30FB\u4ECB\u8B77\u74B0\u5883" },
    { icon: "\uD83D\uDD04", text: "\u5168\u8EAB\u72B6\u614B\u306E\u81EA\u7136\u5909\u52D5" },
    { icon: "\u2753", text: "\u30D7\u30E9\u30BB\u30DC\u52B9\u679C" },
  ];

  confounders.forEach(function(c, i) {
    var yPos = 2.7 + i * 0.5;
    addCard(s, 5.2, yPos, 4.3, 0.45);
    s.addText(c.icon + "  " + c.text, { x: 5.4, y: yPos + 0.03, w: 3.9, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  addCard(s, 0.5, 4.9, 9, 0.65);
  s.addText("\u203B \u7D4C\u53E3\u2192\u8CBC\u4ED8\u5264\u306E\u5207\u308A\u66FF\u3048\u5834\u9762\u3067\u306F\u3001\u6D88\u5316\u5668\u4E0D\u5FEB\u8EFD\u6E1B\u304C\u5BC4\u4E0E\u3059\u308B\u53EF\u80FD\u6027\u3042\u308A\uFF08\u5225\u306E\u4EEE\u8AAC\uFF09", {
    x: 0.8, y: 4.95, w: 8.4, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  addPageNum(s, 7);
})();

// ============================================================
// SLIDE 8: 切り替えを検討する患者像
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u5207\u308A\u66FF\u3048\u3092\u691C\u8A0E\u3059\u308B\u60A3\u8005\u50CF");

  var cases = [
    { situation: "\u7D4C\u53E3\u30C9\u30CD\u30DA\u30B8\u30EB\u3067\n\u6D88\u5316\u5668\u75C7\u72B6\u304C\u554F\u984C", detail: "\u60AA\u5FC3\u30FB\u98DF\u6B32\u4E0D\u632F\u30FB\u4E0B\u75E2\u304C\u6301\u7D9A\n\u6E1B\u91CF\u3067\u3082\u6539\u5584\u3057\u306A\u3044" },
    { situation: "\u5185\u670D\u30A2\u30C9\u30D2\u30A2\u30E9\u30F3\u30B9\n\u304C\u4E0D\u5B89\u5B9A", detail: "\u98F2\u307F\u5FD8\u308C\u304C\u983B\u56DE\n\u670D\u85AC\u62D2\u5426\u304C\u3042\u308B" },
    { situation: "\u5687\u4E0B\u6A5F\u80FD\u4F4E\u4E0B", detail: "\u7D4C\u53E3\u6295\u4E0E\u81EA\u4F53\u306B\u30EA\u30B9\u30AF\n\u8AA4\u56A5\u6027\u80BA\u708E\u306E\u65E2\u5F80" },
    { situation: "\u4ECB\u8B77\u8CA0\u62C5\u306E\u8EFD\u6E1B", detail: "\u670D\u85AC\u4ECB\u52A9\u306E\u8CA0\u62C5\u304C\u5927\u304D\u3044\n\u65BD\u8A2D\u5165\u6240\u3067\u7BA1\u7406\u3092\u7C21\u4FBF\u5316" },
    { situation: "\u300C\u4F55\u3068\u306A\u304F\u98DF\u304C\n\u7D30\u304F\u306A\u3063\u305F\u300D", detail: "\u660E\u78BA\u306A\u6D88\u5316\u5668\u75C7\u72B6\u306F\u306A\u3044\u304C\n\u98DF\u4E8B\u91CF\u304C\u6F38\u6E1B\u3057\u3066\u3044\u308B" },
  ];

  cases.forEach(function(c, i) {
    var yPos = 1.2 + i * 0.82;
    var bg = i % 2 === 0 ? C.light : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 4.3, h: 0.75, fill: { color: bg } });
    s.addText(c.situation, { x: 0.7, y: yPos + 0.05, w: 3.9, h: 0.65, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0, lineSpacingMultiple: 1.2 });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.8, y: yPos, w: 4.7, h: 0.75, fill: { color: bg } });
    s.addText(c.detail, { x: 5.0, y: yPos + 0.05, w: 4.3, h: 0.65, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, 8);
})();

// ============================================================
// SLIDE 9: 貼付剤の注意点
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u8CBC\u4ED8\u5264\u4F7F\u7528\u6642\u306E\u6CE8\u610F\u70B9");

  var warnings = [
    { icon: "\uD83D\uDD04", text: "\u540C\u4E00\u90E8\u4F4D\u3078\u306E\u8CBC\u4ED8\u306F 7\u65E5\u4EE5\u4E0A\u7A7A\u3051\u308B\uFF08\u30ED\u30FC\u30C6\u30FC\u30B7\u30E7\u30F3\uFF09" },
    { icon: "\u2600\uFE0F", text: "\u5265\u96E2\u5F8C 3\u9031\u9593\u306F\u76F4\u5C04\u65E5\u5149\u3092\u907F\u3051\u308B\uFF08\u5149\u7DDA\u904E\u654F\uFF09" },
    { icon: "\uD83E\uDE79", text: "\u9069\u7528\u90E8\u4F4D\u7D05\u6591 24.9%\u3001\u305D\u3046\u75D2\u611F 22.5%\u3001\u63A5\u89E6\u76AE\u819A\u708E 11.0%" },
    { icon: "\u274C", text: "\u76AE\u819A\u30C8\u30E9\u30D6\u30EB\u304C\u5F37\u3044\u60A3\u8005\u306B\u306F\u4E0D\u5411\u304D" },
  ];

  warnings.forEach(function(w, i) {
    var yPos = 1.3 + i * 1.0;
    addCard(s, 0.5, yPos, 9, 0.85);
    s.addText(w.icon, { x: 0.7, y: yPos + 0.1, w: 0.6, h: 0.6, fontSize: 28, margin: 0 });
    s.addText(w.text, { x: 1.5, y: yPos + 0.1, w: 7.7, h: 0.6, fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle" });
  });

  addPageNum(s, 9);
})();

// ============================================================
// SLIDE 10: Take-Home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take-Home Message", {
    x: 0.8, y: 0.3, w: 8.4, h: 0.7,
    fontSize: 28, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "\u30A2\u30EA\u30C9\u30CD\u30D1\u30C3\u30C1\u306F\u300C\u98DF\u6B32\u589E\u9032\u85AC\u300D\u3067\u306F\u306A\u3044" },
    { num: "2", text: "\u98DF\u6B32\u6539\u5584\u5831\u544A\u306F\u3042\u308B\u304C\u3001\u30A8\u30D3\u30C7\u30F3\u30B9\u30EC\u30D9\u30EB\u306F\u4F4E\u3044" },
    { num: "3", text: "\u8CBC\u4ED8\u5264\u3067\u3082\u6D88\u5316\u5668\u526F\u4F5C\u7528\u306F\u6B8B\u308B" },
    { num: "4", text: "\u6539\u5584\u306E\u4E3B\u56E0\u306F\u300C\u7D4C\u53E3\u5264\u306E\u4E0D\u5FEB\u8EFD\u6E1B\u300D\u3068\u8003\u3048\u308B\u306E\u304C\u59A5\u5F53" },
    { num: "5", text: "\u6D88\u5316\u5668\u75C7\u72B6\u30FB\u670D\u85AC\u56F0\u96E3\u4F8B\u3067\u306E\u5207\u308A\u66FF\u3048\u5148\u3068\u3057\u3066\u4F4D\u7F6E\u3065\u3051\u308B" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.2 + i * 0.85;
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: yPos + 0.05, w: 0.55, h: 0.55, fill: { color: C.accent } });
    s.addText(m.num, { x: 0.8, y: yPos + 0.05, w: 0.55, h: 0.55, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.6, y: yPos, w: 7.6, h: 0.65, fontSize: 20, fontFace: FONT_JP, color: C.white, margin: 0, valign: "middle" });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.5, w: 6, h: 0, line: { color: C.accent, width: 1 } });

  s.addText("\u533B\u77E5\u5275\u9020\u30E9\u30DC  \u4ECA\u6751\u4E45\u53F8", {
    x: 0.8, y: 5.1, w: 8.4, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, 10);
})();

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/alidone_patch_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Saved: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
