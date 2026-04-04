var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "パーキンソン病患者のビタ���ンB6管理";

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
var TOTAL = 12;

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

  s.addText("\u30D1\u30FC\u30AD\u30F3\u30BD\u30F3\u75C5\u60A3\u8005\u306E", {
    x: 0.8, y: 0.8, w: 8.4, h: 0.8,
    fontSize: 32, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("\u30D3\u30BF\u30DF\u30F3B6\u7BA1\u7406", {
    x: 0.8, y: 1.5, w: 8.4, h: 1.0,
    fontSize: 44, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });
  s.addText("\u30EC\u30DC\u30C9\u30D1/\u30AB\u30EB\u30D3\u30C9\u30D1\u6642\u4EE3\u306B\u300C\u907F\u3051\u308B\u300D\u3067\u306F\u306A\u304F\u300C\u8A55\u4FA1\u3057\u3066\u88DC\u3046\u300D", {
    x: 0.8, y: 2.7, w: 8.4, h: 0.7,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 3.6, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("\u533B\u77E5\u5275\u9020\u30E9\u30DC\u3000\u4ECA\u6751\u4E45\u53F8", {
    x: 0.8, y: 3.9, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("2026\u5E74FDA\u8B66\u544A\u3092\u8E0F\u307E\u3048\u305F\u5B9F\u81E8\u5E8A\u30A2\u30D7\u30ED\u30FC\u30C1", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 今日のテー���
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u4ECA\u65E5\u306E\u30C6\u30FC\u30DE");

  addCard(s, 0.5, 1.2, 9, 1.6);
  s.addText("\u30D3\u30BF\u30DF\u30F3B6\u306F\u30EC\u30DC\u30C9\u30D1\u60A3\u8005\u3067", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_JP, color: C.text, align: "center", margin: 0,
  });
  s.addText("\u300C\u907F\u3051\u308B\u3079\u304D\u300D\uFF1F\u305D\u308C\u3068\u3082\u300C\u88DC\u3046\u3079\u304D\u300D\uFF1F", {
    x: 0.8, y: 1.85, w: 8.4, h: 0.8,
    fontSize: 28, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0,
  });

  var topics = [
    { icon: "\uD83D\uDC8A", text: "\u30EC\u30DC\u30C9\u30D1\u5358\u5264 vs \u914D\u5408\u5264\u3067\u8003\u3048\u65B9\u304C\u9055\u3046" },
    { icon: "\u26A0\uFE0F", text: "2026\u5E74FDA\u8B66\u544A\uFF1AB6\u6B20\u4E4F\u3068\u767A\u4F5C" },
    { icon: "\uD83D\uDCC8", text: "B6\u6B20\u4E4F\u306E\u983B\u5EA6\u3068\u30EA\u30B9\u30AF\u56E0\u5B50" },
    { icon: "\uD83C\uDFE5", text: "\u5B9F\u81E8\u5E8A\u3067\u3044\u3064\u6E2C\u308A\u3001\u3044\u3064\u88DC\u3046\u304B" },
  ];

  topics.forEach(function(t, i) {
    var yPos = 3.2 + i * 0.55;
    addCard(s, 1.0, yPos, 8.0, 0.5);
    s.addText(t.icon + "  " + t.text, {
      x: 1.2, y: yPos + 0.05, w: 7.6, h: 0.4,
      fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  addPageNum(s, 2);
})();

// ============================================================
// SLIDE 3: なぜ混乱するのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u306A\u305C\u6DF7\u4E71\u3059\u308B\u306E\u304B\uFF1F");

  // Left card - B6過剰
  addCard(s, 0.4, 1.2, 4.4, 3.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.2, w: 4.4, h: 0.5, fill: { color: C.red } });
  s.addText("\u8A71\u2460\uFF1AB6\u304C\u300C\u591A\u3059\u304E\u308B\u300D\u3068\u56F0\u308B", {
    x: 0.4, y: 1.2, w: 4.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("\u30EC\u30DC\u30C9\u30D1\u5358\u5264\u306E\u8A71", {
    x: 0.6, y: 1.85, w: 4.0, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("B6\u304C\u672B\u68A2\u3067\u306E\u30EC\u30DC\u30C9\u30D1\u2192\u30C9\u30D1\u30DF\u30F3\n\u5909\u63DB\u3092\u4FC3\u9032\n\u2192 \u8133\u306B\u5C4A\u304F\u91CF\u304C\u6E1B\u308B", {
    x: 0.6, y: 2.3, w: 4.0, h: 1.2,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });
  s.addText("\u2192 \u53E4\u5178\u7684\u76F8\u4E92\u4F5C\u7528", {
    x: 0.6, y: 3.6, w: 4.0, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });
  s.addText("PMDA\u96FB\u5B50\u6DFB\u6587 / 2017\u5E74SR", {
    x: 0.6, y: 4.2, w: 4.0, h: 0.3,
    fontSize: 11, fontFace: FONT_EN, color: C.sub, italic: true, margin: 0,
  });

  // Right card - B6欠乏
  addCard(s, 5.2, 1.2, 4.4, 3.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.4, h: 0.5, fill: { color: C.orange } });
  s.addText("\u8A71\u2461\uFF1AB6\u304C\u300C\u8DB3\u308A\u306A\u304F\u306A\u308B\u300D\u3068\u56F0\u308B", {
    x: 5.2, y: 1.2, w: 4.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("\u30EC\u30DC\u30C9\u30D1/\u30AB\u30EB\u30D3\u30C9\u30D1\u5408\u5264\u306E\u8A71", {
    x: 5.4, y: 1.85, w: 4.0, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("\u30AB\u30EB\u30D3\u30C9\u30D1\u304CPLP\u306B\u7D50\u5408\n\u30EC\u30DC\u30C9\u30D1\u4EE3\u8B1D\u3067B6\u6D88\u8CBB\n\u2192 \u6A5F\u80FD\u7684B6\u6B20\u4E4F", {
    x: 5.4, y: 2.3, w: 4.0, h: 1.2,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });
  s.addText("\u2192 2026\u5E74FDA\u8B66\u544A", {
    x: 5.4, y: 3.6, w: 4.0, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });
  s.addText("\u672B\u68A2\u795E\u7D4C\u969C\u5BB3\u30FB\u8CA7\u8840\u30FB\u767A\u4F5C", {
    x: 5.4, y: 4.2, w: 4.0, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  addPageNum(s, 3);
})();

// ============================================================
// SLIDE 4: レボドパ単剤とB6
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u30EC\u30DC\u30C9\u30D1\u5358\u5264\u3068\u30D3\u30BF\u30DF\u30F3B6");

  addCard(s, 0.5, 1.2, 9, 1.0);
  s.addText("PMDA\u96FB\u5B50\u6DFB\u6587\uFF1A\u30D4\u30EA\u30C9\u30AD\u30B7\u30F3\u304C\u672B\u68A2\u3067\u306E\u8131\u70AD\u9178\u5316\u3092\u4FC3\u9032\u2192\u4F5C\u7528\u6E1B\u5F31", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.8,
    fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  addCard(s, 0.5, 2.5, 9, 2.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 9, h: 0.5, fill: { color: C.primary } });
  s.addText("2017\u5E74 \u7CFB\u7D71\u7684\u30EC\u30D3\u30E5\u30FC\uFF08J-STAGE\uFF09", {
    x: 0.5, y: 2.5, w: 9, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  var rows = [
    { label: "B6 \u2265 50 mg/\u65E5", value: "\u85AC\u52B9\u4F4E\u4E0B\u306E\u53EF\u80FD\u6027\u9AD8\u3044", vc: C.red },
    { label: "\u901A\u5E38\u306E\u98DF\u4E8B\u7531\u6765", value: "\u5927\u304D\u306A\u554F\u984C\u306B\u306A\u308A\u306B\u304F\u3044", vc: C.green },
    { label: "\u6CE8\u610F\u3059\u3079\u304D\u306F", value: "\u30B5\u30D7\u30EA\u30FB\u5E02\u8CA9\u85AC", vc: C.orange },
  ];

  rows.forEach(function(r, i) {
    var yPos = 3.2 + i * 0.6;
    s.addText(r.label, {
      x: 0.8, y: yPos, w: 3.5, h: 0.5,
      fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0,
    });
    s.addText(r.value, {
      x: 4.5, y: yPos, w: 4.8, h: 0.5,
      fontSize: 18, fontFace: FONT_JP, color: r.vc, bold: true, margin: 0,
    });
  });

  addPageNum(s, 4);
})();

// ============================================================
// SLIDE 5: 配合剤ではどう考えるか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u30EC\u30DC\u30C9\u30D1/\u30AB\u30EB\u30D3\u30C9\u30D1\u5408\u5264\u3067\u306F\uFF1F");

  addCard(s, 0.5, 1.2, 9, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 0.5, fill: { color: C.green } });
  s.addText("SINEMET\u6DFB\u4ED8\u6587\u66F8\uFF08\u7C73\u56FD\uFF09", {
    x: 0.5, y: 1.2, w: 9, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("\u30AB\u30EB\u30D3\u30C9\u30D1\u306FB6\u306B\u3088\u308B\u30EC\u30DC\u30C9\u30D1\u6E1B\u5F31\u3092\u6291\u3048\u308B\n\u2192 B6\u88DC\u5145\u4E2D\u306E\u60A3\u8005\u306B\u3082\u6295\u4E0E\u53EF\u80FD", {
    x: 0.8, y: 1.8, w: 8.4, h: 0.8,
    fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });

  addCard(s, 0.5, 3.0, 9, 2.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 9, h: 0.5, fill: { color: C.orange } });
  s.addText("\u65E5\u672C\u306E\u30E1\u30CD\u30B7\u30C3\u30C8\u6DFB\u4ED8\u6587\u66F8", {
    x: 0.5, y: 3.0, w: 9, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("\u76F8\u4E92\u4F5C\u7528\u6B04\u306B\u30D3\u30BF\u30DF\u30F3B6\u306E\u8A18\u8F09\u306A\u3057\nB6\u6B20\u4E4F\u30FB\u767A\u4F5C\u306E\u8A18\u8F09\u3082\u306A\u3057", {
    x: 0.8, y: 3.6, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });
  s.addText("\u2192 \u65E5\u672C\u306E\u81E8\u5E8A\u3067\u3053\u306E\u554F\u984C\u304C\u898B\u843D\u3068\u3055\u308C\u3084\u3059\u3044\u80CC\u666F", {
    x: 0.8, y: 4.4, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });

  addPageNum(s, 5);
})();

// ============================================================
// SLIDE 6: B6欠乏のエビデンス（2023年SR）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "B6\u6B20\u4E4F\u306E\u30A8\u30D3\u30C7\u30F3\u30B9");

  addCard(s, 0.5, 1.2, 9, 0.6);
  s.addText("Modica et al. J Neurol Sci 2023 \u2015 \u7CFB\u7D71\u7684\u30EC\u30D3\u30E5\u30FC", {
    x: 0.8, y: 1.25, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  // Key numbers
  var nums = [
    { val: "41.4%", desc: "PD\u60A3\u8005145\u4F8B\u4E2D\n60\u4F8B\u3067B6\u7570\u5E38", x: 0.5, bg: C.red },
    { val: "44.6%", desc: "LCIG\u3067\u306E\nB6\u4F4E\u5024\u7387", x: 3.5, bg: C.orange },
    { val: "30.1%", desc: "\u7D4C\u53E3L/C\u3067\u306E\nB6\u4F4E\u5024\u7387", x: 6.5, bg: C.primary },
  ];

  nums.forEach(function(n) {
    addCard(s, n.x, 2.1, 2.8, 2.2);
    s.addText(n.val, {
      x: n.x, y: 2.2, w: 2.8, h: 0.8,
      fontSize: 36, fontFace: FONT_EN, color: n.bg, bold: true, align: "center", margin: 0,
    });
    s.addText(n.desc, {
      x: n.x, y: 3.1, w: 2.8, h: 0.8,
      fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.3,
    });
  });

  addCard(s, 0.5, 4.5, 9, 0.7);
  s.addText("\u4F4EB6\u306E\u5831\u544A\u306E\u307B\u307C\u5168\u4F8B\u3067\u30EC\u30DC\u30C9\u30D1 1,000 mg/\u65E5\u4EE5\u4E0A", {
    x: 0.8, y: 4.55, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });

  addPageNum(s, 6);
})();

// ============================================================
// SLIDE 7: 2020年後ろ向き研究 + 2024-2025年データ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u8FFD\u52A0\u30A8\u30D3\u30C7\u30F3\u30B9\uFF1A2020\u30FB2024\u30FB2025\u5E74");

  // Rojo-Sebastian 2020
  addCard(s, 0.4, 1.2, 4.4, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.2, w: 4.4, h: 0.45, fill: { color: C.primary } });
  s.addText("Rojo-Sebasti\u00E1n 2020\uFF0824\u4F8B\uFF09", {
    x: 0.4, y: 1.2, w: 4.4, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("\u7D4C\u8178\u6295\u4E0E\uFF1A6/6\u4F8B (100%)\n\u7D4C\u53E3\u6295\u4E0E\uFF1A13/18\u4F8B (72%)\n\u2192 B6\u4F4E\u5024", {
    x: 0.6, y: 1.75, w: 4.0, h: 1.0,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });

  // Ikenaka 2024 Osaka
  addCard(s, 5.2, 1.2, 4.4, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.4, h: 0.45, fill: { color: C.orange } });
  s.addText("Ikenaka 2024 \u5927\u962A\u5927\uFF0872+31\u4F8B\uFF09", {
    x: 5.2, y: 1.2, w: 4.4, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("PD\u60A3\u8005\u3067B6\u6709\u610F\u306B\u4F4E\u5024\nAST/ALT\u4F4E\u4E0B\u304CB6\u4F4E\u5024\u3092\u53CD\u6620\n\u2192 \u30E2\u30CB\u30BF\u30EA\u30F3\u30B0\u6307\u6A19\u306B", {
    x: 5.4, y: 1.75, w: 4.0, h: 1.0,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });

  // Dery 2024
  addCard(s, 0.4, 3.5, 9.2, 1.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.5, w: 9.2, h: 0.45, fill: { color: C.green } });
  s.addText("D\u00E9ry 2024 \u6A2A\u65AD\u7684\u89B3\u5BDF\u7814\u7A76\uFF0850\u4F8B\u3001\u7D4C\u53E3LD \u2265600 mg/\u65E5\uFF09", {
    x: 0.4, y: 3.5, w: 9.2, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("7\u4F8B (14%) \u306BB6\u6B20\u4E4F\u3000/\u3000B6\u4F4E\u5024 \u2194 \u30EC\u30DC\u30C9\u30D1\u91CF\u30FB\u672B\u68A2\u795E\u7D4C\u969C\u5BB3\u91CD\u75C7\u5EA6\u304C\u76F8\u95A2\nB12\u30FB\u8449\u9178\u30FB\u30DB\u30E2\u30B7\u30B9\u30C6\u30A4\u30F3\u306F\u672B\u68A2\u795E\u7D4C\u969C\u5BB3\u3068\u7121\u95A2\u9023", {
    x: 0.6, y: 4.05, w: 8.8, h: 0.8,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });

  addPageNum(s, 7);
})();

// ============================================================
// SLIDE 8: FDA 2026年3月警告
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u26A0\uFE0F 2026\u5E743\u6708 FDA\u5B89\u5168\u6027\u8B66\u544A");

  addCard(s, 0.5, 1.2, 9, 0.8);
  s.addText("\u30AB\u30EB\u30D3\u30C9\u30D1/\u30EC\u30DC\u30C9\u30D1\u542B\u6709\u88FD\u5264\u306B\u3001B6\u6B20\u4E4F\u3068\u95A2\u9023\u767A\u4F5C\u306E\u8B66\u544A\u8FFD\u52A0\u3092\u8981\u6C42", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });

  // Data table
  var tableData = [
    { item: "\u78BA\u8A8D\u3055\u308C\u305F\u767A\u4F5C\u4F8B", val: "14\u4F8B", c: C.red },
    { item: "\u5168\u4F8B\u306E\u30EC\u30DC\u30C9\u30D1\u7528\u91CF", val: ">1,000 mg/\u65E5", c: C.red },
    { item: "\u767A\u75C7\u307E\u3067\u306E\u671F\u9593", val: "23\uFF5E132\u304B\u6708", c: C.text },
    { item: "\u767A\u4F5C\u578B", val: "\u7126\u70B9\u8D77\u59CB \u2192 \u4E8C\u6B21\u6027\u5168\u822C\u5316", c: C.text },
    { item: "B6\u88DC\u5145\u5F8C\u306E\u8EE2\u5E30", val: "9\u4F8B\u5168\u4F8B\u3067\u767A\u4F5C\u6D88\u5931", c: C.green },
    { item: "\u6B7B\u4EA1", val: "2\u4F8B\uFF08\u4F4EB6 + \u96E3\u6CBB\u767A\u4F5C\uFF09", c: C.red },
  ];

  tableData.forEach(function(r, i) {
    var yPos = 2.3 + i * 0.5;
    var bgColor = i % 2 === 0 ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 4.5, h: 0.48, fill: { color: bgColor } });
    s.addText(r.item, {
      x: 0.7, y: yPos + 0.02, w: 4.1, h: 0.44,
      fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y: yPos, w: 4.5, h: 0.48, fill: { color: bgColor } });
    s.addText(r.val, {
      x: 5.2, y: yPos + 0.02, w: 4.1, h: 0.44,
      fontSize: 15, fontFace: FONT_JP, color: r.c, bold: true, margin: 0,
    });
  });

  addPageNum(s, 8);
})();

// ============================================================
// SLIDE 9: 抗てんかん薬抵抗性とB6補充
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u6297\u3066\u3093\u304B\u3093\u85AC\u306B\u53CD\u5FDC\u3057\u306A\u3044 \u2192 B6\u3067\u6D88\u5931");

  addCard(s, 0.5, 1.2, 9, 2.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 0.5, fill: { color: C.primary } });
  s.addText("Nathan 2025 \u75C7\u4F8B\u5831\u544A (Neurohospitalist)", {
    x: 0.5, y: 1.2, w: 9, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("\u30AB\u30EB\u30D3\u30C9\u30D1/\u30EC\u30DC\u30C9\u30D1\u5897\u91CF\u5F8C\u306B\u65B0\u898F\u767A\u4F5C", {
    x: 0.8, y: 1.85, w: 8.4, h: 0.4,
    fontSize: 17, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  s.addText("\u2192 \u8907\u6570\u306E\u6297\u3066\u3093\u304B\u3093\u85AC\u306B\u53CD\u5FDC\u305B\u305A", {
    x: 0.8, y: 2.25, w: 8.4, h: 0.4,
    fontSize: 17, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });
  s.addText("\u2192 \u30D4\u30EA\u30C9\u30AD\u30B7\u30F3\u88DC\u5145\u3067\u767A\u4F5C\u6D88\u5931\u3001\u5168ASM\u96E2\u8131\u53EF\u80FD\u306B", {
    x: 0.8, y: 2.65, w: 8.4, h: 0.4,
    fontSize: 17, fontFace: FONT_JP, color: C.green, bold: true, margin: 0,
  });

  addCard(s, 0.5, 3.8, 9, 1.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.8, w: 9, h: 0.45, fill: { color: C.red } });
  s.addText("\u81E8\u5E8A\u7684\u610F\u7FA9", {
    x: 0.5, y: 3.8, w: 9, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("PD\u60A3\u8005\u306E\u65B0\u898F\u767A\u4F5C = \u69CB\u9020\u75C5\u5909\u3060\u3051\u3067\u306F\u306A\u3044\nB6\u6B20\u4E4F\u306B\u3088\u308B\u6025\u6027\u75C7\u5019\u6027\u767A\u4F5C\u3082\u9451\u5225\u306B\uFF01", {
    x: 0.8, y: 4.35, w: 8.4, h: 0.7,
    fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true, margin: 0, lineSpacingMultiple: 1.4,
  });

  addPageNum(s, 9);
})();

// ============================================================
// SLIDE 10: 誰にB6を測るべきか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u5B9F\u81E8\u5E8A\u3067\u8AB0\u306BB6\u3092\u6E2C\u308B\u3079\u304D\u304B");

  var risks = [
    { factor: "\u30EC\u30DC\u30C9\u30D1/CD \u9AD8\u7528\u91CF\uFF08>1,000 mg/\u65E5\uFF09", icon: "\uD83D\uDC8A" },
    { factor: "LCIG\uFF08\u7D4C\u8178\u6295\u4E0E\uFF09", icon: "\uD83C\uDFE5" },
    { factor: "\u4F53\u91CD\u6E1B\u5C11\u30FB\u98DF\u4E8B\u6442\u53D6\u4F4E\u4E0B", icon: "\uD83C\uDF7D\uFE0F" },
    { factor: "\u672B\u68A2\u795E\u7D4C\u969C\u5BB3", icon: "\u26A1" },
    { factor: "\u539F\u56E0\u4E0D\u660E\u306E\u8CA7\u8840", icon: "\uD83E\uDE78" },
    { factor: "\u65B0\u898F\u767A\u4F5C\u30FB\u3066\u3093\u304B\u3093\u91CD\u7A4D", icon: "\u26A0\uFE0F" },
    { factor: "\u539F\u56E0\u4E0D\u660E\u306E\u6291\u3046\u3064\u30FB\u932F\u4E71", icon: "\uD83E\uDDE0" },
  ];

  risks.forEach(function(r, i) {
    var yPos = 1.15 + i * 0.55;
    var bgColor = i % 2 === 0 ? C.light : C.white;
    addCard(s, 0.5, yPos, 9, 0.5);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9, h: 0.5, fill: { color: bgColor } });
    s.addText(r.icon + "  " + r.factor, {
      x: 0.8, y: yPos + 0.05, w: 8.4, h: 0.4,
      fontSize: 17, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  s.addText("FDA\u63A8\u5968\uFF1A\u958B\u59CB\u524D + \u6CBB\u7642\u4E2D\u5B9A\u671F\u7684\u306B\u8A55\u4FA1\u3001\u5FC5\u8981\u6642\u88DC\u5145", {
    x: 0.5, y: 5.0, w: 9, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  addPageNum(s, 10);
})();

// ============================================================
// SLIDE 11: 製剤別まとめ表
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "\u88DC\u5145\u3059\u3079\u304D\u304B\u3001\u907F\u3051\u308B\u3079\u304D\u304B \u2015 \u88FD\u5264\u3067\u5206\u3051\u308B");

  // Header row
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 3.0, h: 0.5, fill: { color: C.dark } });
  s.addText("\u88FD\u5264", { x: 0.3, y: 1.2, w: 3.0, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.3, y: 1.2, w: 3.2, h: 0.5, fill: { color: C.dark } });
  s.addText("B6\u306E\u8003\u3048\u65B9", { x: 3.3, y: 1.2, w: 3.2, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 1.2, w: 3.2, h: 0.5, fill: { color: C.dark } });
  s.addText("\u5B9F\u52D9\u5BFE\u5FDC", { x: 6.5, y: 1.2, w: 3.2, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });

  // Row 1
  var y1 = 1.7;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y1, w: 3.0, h: 1.0, fill: { color: C.light } });
  s.addText("\u30EC\u30DC\u30C9\u30D1\u5358\u5264", { x: 0.3, y: y1, w: 3.0, h: 1.0, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 4 });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.3, y: y1, w: 3.2, h: 1.0, fill: { color: C.warmBg } });
  s.addText("B6\u3067\u85AC\u52B9\u6E1B\u5F31\u306E\n\u53EF\u80FD\u6027", { x: 3.3, y: y1, w: 3.2, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: C.red, align: "center", valign: "middle", margin: 4, lineSpacingMultiple: 1.3 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: y1, w: 3.2, h: 1.0, fill: { color: C.warmBg } });
  s.addText("\u9AD8\u7528\u91CFB6\u30B5\u30D7\u30EA\u306B\u6CE8\u610F\n\u901A\u5E38\u98DF\u4E8B\u306F\u554F\u984C\u306A\u3057", { x: 6.5, y: y1, w: 3.2, h: 1.0, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 4, lineSpacingMultiple: 1.3 });

  // Row 2
  var y2 = 2.7;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y2, w: 3.0, h: 1.0, fill: { color: C.light } });
  s.addText("L/C\u5408\u5264\uFF08\u7D4C\u53E3\uFF09", { x: 0.3, y: y2, w: 3.0, h: 1.0, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 4 });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.3, y: y2, w: 3.2, h: 1.0, fill: { color: C.white } });
  s.addText("\u9577\u671F\u30FB\u9AD8\u7528\u91CF\u3067\nB6\u6B20\u4E4F\u30EA\u30B9\u30AF", { x: 3.3, y: y2, w: 3.2, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: C.orange, align: "center", valign: "middle", margin: 4, lineSpacingMultiple: 1.3 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: y2, w: 3.2, h: 1.0, fill: { color: C.white } });
  s.addText("\u98DF\u4E8B\u5236\u9650\u4E0D\u8981\n\u6B20\u4E4F\u7591\u3044\u3067\u6E2C\u5B9A\u30FB\u88DC\u5145", { x: 6.5, y: y2, w: 3.2, h: 1.0, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 4, lineSpacingMultiple: 1.3 });

  // Row 3
  var y3 = 3.7;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y3, w: 3.0, h: 1.0, fill: { color: C.light } });
  s.addText("LCIG\uFF08\u7D4C\u8178\u6295\u4E0E\uFF09", { x: 0.3, y: y3, w: 3.0, h: 1.0, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 4 });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.3, y: y3, w: 3.2, h: 1.0, fill: { color: C.warmBg } });
  s.addText("B6\u6B20\u4E4F\u30EA\u30B9\u30AF\n\u6700\u3082\u9AD8\u3044", { x: 3.3, y: y3, w: 3.2, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 4, lineSpacingMultiple: 1.3 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: y3, w: 3.2, h: 1.0, fill: { color: C.warmBg } });
  s.addText("\u5B9A\u671F\u30E2\u30CB\u30BF\u30EA\u30F3\u30B0\u63A8\u5968\n\u4E88\u9632\u7684B\u88DC\u5145\u3082\u8003\u616E", { x: 6.5, y: y3, w: 3.2, h: 1.0, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 4, lineSpacingMultiple: 1.3 });

  addPageNum(s, 11);
})();

// ============================================================
// SLIDE 12: Take Home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.8, y: 0.3, w: 8.4, h: 0.6,
    fontSize: 28, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var msgs = [
    { icon: "\u2460", text: "\u300CB6\u306F\u6575\u304B\u5473\u65B9\u304B\u300D\u3067\u306F\u306A\u304F\n\u300C\u3069\u306E\u88FD\u5264\u3067\u3001\u3069\u306E\u5834\u9762\u304B\u300D\u3092\u5206\u3051\u308B" },
    { icon: "\u2461", text: "\u30EC\u30DC\u30C9\u30D1\u5358\u5264 = B6\u76F8\u4E92\u4F5C\u7528\u306E\u76F8\u624B\n\u30EC\u30DC\u30C9\u30D1/\u30AB\u30EB\u30D3\u30C9\u30D1\u5408\u5264 = B6\u6B20\u4E4F\u3092\u898B\u843D\u3068\u3055\u306A\u3044" },
    { icon: "\u2462", text: "2026\u5E74FDA\u8B66\u544A\u3092\u8E0F\u307E\u3048\u3001B6\u7BA1\u7406\u306F\n\u300C\u7981\u6B62\u300D\u304B\u3089\u300C\u5C64\u5225\u5316\u3057\u3066\u6E2C\u5B9A\u30FB\u5FC5\u8981\u6642\u88DC\u5145\u300D\u3078" },
  ];

  msgs.forEach(function(m, i) {
    var yPos = 1.2 + i * 1.4;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.icon, {
      x: 0.8, y: yPos, w: 0.6, h: 0.6,
      fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(m.text, {
      x: 1.6, y: yPos, w: 7.6, h: 1.1,
      fontSize: 19, fontFace: FONT_JP, color: C.white, margin: 0, lineSpacingMultiple: 1.4,
    });
  });

  s.addText("\u533B\u77E5\u5275\u9020\u30E9\u30DC\u3000\u4ECA\u6751\u4E45\u53F8", {
    x: 0.8, y: 5.1, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  addPageNum(s, 12);
})();

// ============================================================
// Export
// ============================================================
var outPath = __dirname + "/PD_vitaminB6_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Saved: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
