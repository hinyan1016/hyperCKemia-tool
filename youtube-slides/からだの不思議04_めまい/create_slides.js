var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "そのめまい大丈夫？ ― 3タイプの見分け方と危険なサインを脳神経内科医が解説";

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

  s.addText("脳神経内科医が答える からだの不思議 #04", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("そのめまい大丈夫？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 3タイプの見分け方と危険なサインを脳神経内科医が解説 ―", {
    x: 0.8, y: 2.9, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.7, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.0, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 「めまい」と一言で言っても…
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "「めまい」と一言で言っても…");

  var quotes = [
    "「ぐるぐると部屋が回る感じがする」",
    "「ふわふわして地に足がつかない」",
    "「立ち上がった瞬間に目の前が暗くなる」",
  ];
  quotes.forEach(function(q, i) {
    var yPos = 1.1 + i * 1.2;
    addCard(s, 1.0, yPos, 8.0, 1.0);
    s.addText(q, {
      x: 1.3, y: yPos + 0.1, w: 7.4, h: 0.8,
      fontSize: 19, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 4.75, w: 8.0, h: 0.6, fill: { color: C.light } });
  s.addText("この3つは、まったく異なるメカニズムで起きています", {
    x: 1.0, y: 4.75, w: 8.0, h: 0.6,
    fontSize: 17, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/11");
})();

// ============================================================
// SLIDE 3: めまい3タイプの概要
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "めまいの3タイプ");

  var types = [
    { num: "①", name: "回転性めまい", en: "Vertigo", desc: "ぐるぐる・くるくる回る感覚\n内耳・脳幹・小脳の問題", color: C.red },
    { num: "②", name: "浮動性めまい", en: "Dizziness", desc: "ふわふわ・フラフラ・不安定\n多因子的（小脳・神経・薬剤など）", color: C.orange },
    { num: "③", name: "立ちくらみ", en: "Presyncope", desc: "立ったとき一瞬目が暗くなる\n一過性の脳血流低下", color: C.accent },
  ];

  types.forEach(function(t, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.1, 3.0, 3.9);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 3.0, h: 0.55, fill: { color: t.color } });
    s.addText(t.num + " " + t.name, {
      x: xPos, y: 1.1, w: 3.0, h: 0.55,
      fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(t.en, {
      x: xPos + 0.1, y: 1.75, w: 2.8, h: 0.5,
      fontSize: 14, fontFace: FONT_EN, color: C.sub, align: "center", italic: true, margin: 0,
    });
    s.addText(t.desc, {
      x: xPos + 0.1, y: 2.35, w: 2.8, h: 2.4,
      fontSize: 15, fontFace: FONT_JP, color: C.text, valign: "top", margin: 4,
      lineSpacingMultiple: 1.3,
    });
  });

  addPageNum(s, "3/11");
})();

// ============================================================
// SLIDE 4: ① 回転性めまい（Vertigo）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "① 回転性めまい（Vertigo）");

  addCard(s, 0.5, 1.1, 9.0, 1.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 0.08, h: 1.4, fill: { color: C.red } });
  s.addText("「自分や周囲がぐるぐる・くるくる回る」感覚\n内耳（ないじ）・脳幹（のうかん）・小脳（しょうのう）の異常が原因", {
    x: 0.8, y: 1.2, w: 8.5, h: 1.1,
    fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 4, lineSpacingMultiple: 1.3,
  });

  s.addText("代表疾患", {
    x: 0.5, y: 2.7, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  var diseases = [
    { name: "BPPV", sub: "良性発作性頭位めまい症", color: C.green },
    { name: "メニエール病", sub: "内リンパ水腫", color: C.orange },
    { name: "前庭神経炎", sub: "ウイルス性", color: C.accent },
    { name: "小脳・脳幹梗塞", sub: "要除外！", color: C.red },
  ];

  diseases.forEach(function(d, i) {
    var xPos = 0.3 + i * 2.4;
    addCard(s, xPos, 3.15, 2.2, 1.9);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.15, w: 2.2, h: 0.45, fill: { color: d.color } });
    s.addText(d.name, {
      x: xPos, y: 3.15, w: 2.2, h: 0.45,
      fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(d.sub, {
      x: xPos + 0.1, y: 3.65, w: 2.0, h: 1.3,
      fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 2,
      lineSpacingMultiple: 1.2,
    });
  });

  addPageNum(s, "4/11");
})();

// ============================================================
// SLIDE 5: ② 浮動性・③ 立ちくらみ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "② 浮動性めまい　③ 立ちくらみ");

  // 浮動性
  addCard(s, 0.3, 1.1, 4.6, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 4.6, h: 0.5, fill: { color: C.orange } });
  s.addText("② 浮動性めまい（Dizziness）", {
    x: 0.3, y: 1.1, w: 4.6, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("ふわふわ・フラフラ・不安定感\n回転感はなく持続的", {
    x: 0.5, y: 1.7, w: 4.2, h: 1.0,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 4, lineSpacingMultiple: 1.3,
  });
  s.addText("原因：", {
    x: 0.5, y: 2.8, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  var dizzCauses = ["小脳変性症", "多発性神経炎", "精神疾患・薬剤性", "多因子性（高齢者）"];
  dizzCauses.forEach(function(c, i) {
    s.addText("・" + c, {
      x: 0.6, y: 3.15 + i * 0.45, w: 4.0, h: 0.4,
      fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  // 立ちくらみ
  addCard(s, 5.1, 1.1, 4.6, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.6, h: 0.5, fill: { color: C.accent } });
  s.addText("③ 立ちくらみ（Presyncope）", {
    x: 5.1, y: 1.1, w: 4.6, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("立ち上がった瞬間に目の前が暗くなる\n一過性の脳血流低下", {
    x: 5.3, y: 1.7, w: 4.2, h: 1.0,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 4, lineSpacingMultiple: 1.3,
  });
  s.addText("原因：", {
    x: 5.3, y: 2.8, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  var presCauses = ["起立性低血圧（最多）", "不整脈", "脱水", "降圧薬の過剰服用"];
  presCauses.forEach(function(c, i) {
    s.addText("・" + c, {
      x: 5.4, y: 3.15 + i * 0.45, w: 4.0, h: 0.4,
      fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  addPageNum(s, "5/11");
})();

// ============================================================
// SLIDE 6: めまいタイプ判定フロー
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "めまいタイプ判定フロー");

  // --- Q1 ---
  addCard(s, 0.3, 1.1, 5.4, 0.6);
  s.addText("Q1. 立ち上がった直後に起きたか？", {
    x: 0.3, y: 1.1, w: 5.4, h: 0.6,
    fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: [0, 0, 0, 12],
  });
  // YES →
  s.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: 1.15, w: 0.9, h: 0.45, fill: { color: C.green } });
  s.addText("YES →", {
    x: 5.9, y: 1.15, w: 0.9, h: 0.45,
    fontSize: 12, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  // → 立ちくらみ結果
  s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 1.05, w: 2.8, h: 0.7, fill: { color: C.accent } });
  s.addText("③ 立ちくらみ", {
    x: 7.0, y: 1.05, w: 2.8, h: 0.7,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // ↓ NO
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.85, w: 0.65, h: 0.35, fill: { color: C.sub } });
  s.addText("↓ NO", {
    x: 0.5, y: 1.85, w: 0.65, h: 0.35,
    fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // --- Q2 ---
  addCard(s, 0.3, 2.35, 5.4, 0.6);
  s.addText("Q2. 「ぐるぐる回る」感覚があるか？", {
    x: 0.3, y: 2.35, w: 5.4, h: 0.6,
    fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: [0, 0, 0, 12],
  });
  // NO →
  s.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: 2.4, w: 0.9, h: 0.45, fill: { color: C.sub } });
  s.addText("NO →", {
    x: 5.9, y: 2.4, w: 0.9, h: 0.45,
    fontSize: 12, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  // → 浮動性めまい結果
  s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 2.25, w: 2.8, h: 0.7, fill: { color: C.orange } });
  s.addText("② 浮動性めまい", {
    x: 7.0, y: 2.25, w: 2.8, h: 0.7,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // ↓ YES
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.1, w: 0.65, h: 0.35, fill: { color: C.green } });
  s.addText("↓ YES", {
    x: 0.5, y: 3.1, w: 0.65, h: 0.35,
    fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // --- 回転性めまいラベル ---
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.6, w: 3.2, h: 0.5, fill: { color: C.red } });
  s.addText("① 回転性めまい", {
    x: 0.3, y: 3.6, w: 3.2, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // サブ分岐：3疾患
  var subDiseases = [
    { name: "BPPV", desc: "頭位で30秒以内に\n誘発・消失", color: C.green },
    { name: "メニエール病", desc: "耳鳴り・難聴を\n繰り返す", color: C.orange },
    { name: "前庭神経炎", desc: "感染後に突然発症\n数日持続", color: C.accent },
  ];
  subDiseases.forEach(function(d, i) {
    var xPos = 0.3 + i * 3.25;
    addCard(s, xPos, 4.25, 3.05, 1.0);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.25, w: 3.05, h: 0.35, fill: { color: d.color } });
    s.addText(d.name, {
      x: xPos, y: 4.25, w: 3.05, h: 0.35,
      fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(d.desc, {
      x: xPos + 0.1, y: 4.6, w: 2.85, h: 0.6,
      fontSize: 11, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 2, lineSpacingMultiple: 1.2,
    });
  });

  // 警告バー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.35, w: 9.5, h: 0.3, fill: { color: C.red } });
  s.addText("いずれのタイプでも、頭痛・複視・麻痺・構音障害・意識障害があれば脳卒中を除外 → 即受診！", {
    x: 0.3, y: 5.35, w: 9.5, h: 0.3,
    fontSize: 11, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "6/11");
})();

// ============================================================
// SLIDE 7: 代表疾患の比較表
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "代表疾患の比較");

  var rows = [
    { name: "BPPV", type: "回転性", duration: "数秒〜30秒", features: "頭位変換で誘発・消失\n耳鳴り・難聴なし", color: C.green },
    { name: "メニエール病", type: "回転性（反復）", duration: "20分〜数時間", features: "耳鳴り・難聴・耳閉感\n繰り返す", color: C.orange },
    { name: "前庭神経炎", type: "回転性（急性）", duration: "数日〜数週間", features: "感染後突然発症\n難聴なし", color: C.accent },
    { name: "小脳・脳幹梗塞", type: "回転/浮動", duration: "持続", features: "複視・麻痺・頭痛を伴う\n即救急！", color: C.red },
    { name: "起立性低血圧", type: "立ちくらみ", duration: "数秒〜1分", features: "立位直後に視野暗転\n高齢・降圧薬・脱水", color: C.primary },
  ];

  // Header row
  var hCols = [{ label: "疾患", x: 0.2, w: 2.2 }, { label: "タイプ", x: 2.45, w: 1.9 }, { label: "持続時間", x: 4.38, w: 2.0 }, { label: "特徴", x: 6.41, w: 3.45 }];
  hCols.forEach(function(h) {
    s.addShape(pres.shapes.RECTANGLE, { x: h.x, y: 0.95, w: h.w, h: 0.45, fill: { color: C.primary } });
    s.addText(h.label, {
      x: h.x, y: 0.95, w: h.w, h: 0.45,
      fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  rows.forEach(function(r, i) {
    var yPos = 1.45 + i * 0.82;
    var bg = i % 2 === 0 ? C.white : C.warmBg;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: yPos, w: 9.65, h: 0.78, fill: { color: bg } });
    // disease name with color bar
    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: yPos, w: 0.07, h: 0.78, fill: { color: r.color } });
    s.addText(r.name, {
      x: 0.32, y: yPos, w: 2.1, h: 0.78,
      fontSize: 12, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 2, lineSpacingMultiple: 1.1,
    });
    s.addText(r.type, {
      x: 2.45, y: yPos, w: 1.9, h: 0.78,
      fontSize: 12, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 2,
    });
    s.addText(r.duration, {
      x: 4.38, y: yPos, w: 2.0, h: 0.78,
      fontSize: 12, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 2,
    });
    s.addText(r.features, {
      x: 6.41, y: yPos, w: 3.45, h: 0.78,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 2, lineSpacingMultiple: 1.1,
    });
  });

  addPageNum(s, "7/11");
})();

// ============================================================
// SLIDE 8: BPPV vs メニエール病
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "BPPV vs メニエール病 ― 見分け方");

  // BPPV
  addCard(s, 0.3, 1.1, 4.5, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 4.5, h: 0.55, fill: { color: C.green } });
  s.addText("BPPV（良性発作性頭位めまい症）", {
    x: 0.3, y: 1.1, w: 4.5, h: 0.55,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var bppvPoints = [
    { key: "誘発", val: "頭を動かしたときだけ" },
    { key: "持続", val: "30秒以内に治まる" },
    { key: "耳症状", val: "耳鳴り・難聴なし" },
    { key: "原因", val: "耳石（じせき）のずれ" },
    { key: "治療", val: "頭位変換療法で改善" },
  ];
  bppvPoints.forEach(function(p, i) {
    s.addText(p.key + "：", {
      x: 0.5, y: 1.8 + i * 0.6, w: 1.1, h: 0.5,
      fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
    });
    s.addText(p.val, {
      x: 1.6, y: 1.8 + i * 0.6, w: 3.0, h: 0.5,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  // Meniere
  addCard(s, 5.2, 1.1, 4.5, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.5, h: 0.55, fill: { color: C.orange } });
  s.addText("メニエール病", {
    x: 5.2, y: 1.1, w: 4.5, h: 0.55,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var menPoints = [
    { key: "誘発", val: "頭位に関係なく発症" },
    { key: "持続", val: "20分〜数時間" },
    { key: "耳症状", val: "耳鳴り・難聴・耳閉感" },
    { key: "原因", val: "内リンパ水腫" },
    { key: "経過", val: "繰り返す（反復性）" },
  ];
  menPoints.forEach(function(p, i) {
    s.addText(p.key + "：", {
      x: 5.4, y: 1.8 + i * 0.6, w: 1.1, h: 0.5,
      fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, valign: "middle", margin: 0,
    });
    s.addText(p.val, {
      x: 6.5, y: 1.8 + i * 0.6, w: 3.0, h: 0.5,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  addPageNum(s, "8/11");
})();

// ============================================================
// SLIDE 9: 危険なめまいのサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "危険なめまいのサイン ― 今すぐ受診！");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.05, w: 9.4, h: 0.55, fill: { color: C.red } });
  s.addText("以下が1つでもあれば → 迷わず救急受診", {
    x: 0.3, y: 1.05, w: 9.4, h: 0.55,
    fontSize: 17, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var signs = [
    "突然の激しい頭痛を伴うめまい（くも膜下出血）",
    "複視（ものが二重に見える）・眼球が動きにくい",
    "顔・手足のしびれや麻痺",
    "呂律（ろれつ）が回らない・構音障害",
    "まっすぐ立てない・歩行が明らかにおかしい",
    "意識がもうろうとする",
    "眼振の向きが上下・斜め（垂直眼振）",
  ];

  signs.forEach(function(sign, i) {
    var col = i % 2 === 0 ? "FFF5F5" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.7 + i * 0.51, w: 9.4, h: 0.48, fill: { color: col } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.7 + i * 0.51, w: 0.06, h: 0.48, fill: { color: C.red } });
    s.addText("▶ " + sign, {
      x: 0.45, y: 1.7 + i * 0.51, w: 9.1, h: 0.48,
      fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 2,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.25, w: 9.4, h: 0.38, fill: { color: C.light } });
  s.addText("HINTS 検査（頭部衝動試験・眼振・スキュー偏位）で脳卒中を見分けます", {
    x: 0.3, y: 5.25, w: 9.4, h: 0.38,
    fontSize: 12, fontFace: FONT_JP, color: C.primary, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "9/11");
})();

// ============================================================
// SLIDE 10: 様子見でよいめまいの特徴
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "様子見でよいめまいの目安");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.05, w: 9.4, h: 0.5, fill: { color: C.green } });
  s.addText("以下の特徴が揃う場合 → まずかかりつけ医・耳鼻科へ", {
    x: 0.3, y: 1.05, w: 9.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var okSigns = [
    "頭を動かしたときだけ誘発され、30秒以内に自然に治まる（BPPVを示唆）",
    "以前にも同じめまいがあり、医師に診断を受けたことがある",
    "頭痛・麻痺・複視などの危険サインが一切ない",
    "立ち上がったときだけで、安静にすれば全く症状がない（立ちくらみ）",
    "発熱・胃腸炎などがあり、全身状態の悪化によるめまいと説明がつく",
  ];

  okSigns.forEach(function(sign, i) {
    var yPos = 1.7 + i * 0.63;
    addCard(s, 0.3, yPos, 9.4, 0.58);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.42, h: 0.58, fill: { color: C.green } });
    s.addText("✓", {
      x: 0.3, y: yPos, w: 0.42, h: 0.58,
      fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(sign, {
      x: 0.8, y: yPos, w: 8.8, h: 0.58,
      fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 4,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.25, w: 9.4, h: 0.38, fill: { color: "FFF3CD" } });
  s.addText("「いつもと少し違う」と感じたら → 迷わず受診を", {
    x: 0.3, y: 5.25, w: 9.4, h: 0.38,
    fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "10/11");
})();

// ============================================================
// SLIDE 11: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("まとめ ― Take Home Message", {
    x: 0.5, y: 0.15, w: 9.0, h: 0.5,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", msg: "めまいは「回転性・浮動性・立ちくらみ」の3タイプで考える", color: C.accent },
    { num: "2", msg: "最多は BPPV。頭位で30秒以内に誘発・消失、耳鳴り・難聴なし", color: C.green },
    { num: "3", msg: "耳鳴り・難聴を繰り返す回転性めまい → メニエール病を疑う", color: C.orange },
    { num: "4", msg: "突然発症＋頭痛・複視・麻痺・歩行困難が1つでも → 脳卒中を除外！", color: C.red },
    { num: "5", msg: "立ちくらみは「立った瞬間」が特徴。起立性低血圧・薬・脱水を確認", color: C.yellow },
  ];

  messages.forEach(function(m, i) {
    var yPos = 0.85 + i * 0.88;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.52, h: 0.72, fill: { color: m.color } });
    s.addText(m.num, {
      x: 0.3, y: yPos, w: 0.52, h: 0.72,
      fontSize: 22, fontFace: FONT_JP, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(m.msg, {
      x: 0.92, y: yPos + 0.05, w: 8.8, h: 0.62,
      fontSize: 16, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 4, lineSpacingMultiple: 1.2,
    });
  });

  // 受診時の一言アドバイス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.38, fill: { color: C.accent }, transparency: 30 });
  s.addText("受診時に一言（例:「ぐるぐる」「ふわふわ」「立ったときだけ暗くなる」）で診察がスムーズになります", {
    x: 0.3, y: 5.0, w: 9.4, h: 0.38,
    fontSize: 12, fontFace: FONT_JP, color: C.light, align: "center", valign: "middle", margin: 0,
  });

  s.addText("医知創造ラボ", {
    x: 0.3, y: 5.42, w: 9.4, h: 0.2,
    fontSize: 11, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });
})();

// ============================================================
// 出力
// ============================================================
var outPath = __dirname + "/めまい_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
