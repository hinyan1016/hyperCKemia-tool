var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "ヤンキー・ツッパリ・不良・やんちゃ・チンピラ・暴走族・半グレの違いとは";

// --- Color Palette (Navy series for brand consistency) ---
var C = {
  dark:    "1B3A5C",
  primary: "2C5AA0",
  accent:  "4A90D9",
  light:   "E8F0FE",
  warmBg:  "F5F7FA",
  white:   "FFFFFF",
  text:    "2D3436",
  sub:     "636E72",
  red:     "DC3545",
  orange:  "E8850C",
  yellow:  "F5C518",
  green:   "28A745",
  muted:   "95A5A6",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
}

function addCard(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: C.white }, shadow: makeShadow() });
}

function addPageNum(s, num, total) {
  s.addText(num + "/" + total, { x: 9.0, y: 5.2, w: 0.8, h: 0.4, fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "right", margin: 0 });
}

var TOTAL = "16";

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("日本語の雑学 ― 意味の違いシリーズ", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("ヤンキー・ツッパリ・不良・\nやんちゃ・チンピラ・暴走族・半グレ", {
    x: 0.5, y: 1.1, w: 9.0, h: 1.7,
    fontSize: 28, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.25,
  });

  s.addText("― 似ているようで全然違う、7つの言葉を整理する ―", {
    x: 0.5, y: 3.0, w: 9.0, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.85, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.65, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 7語一覧表（結論）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "結論：7語はこう違う");

  var words = [
    { w: "不良",     m: "素行の悪い若者の総称",         c: C.primary },
    { w: "ツッパリ", m: "反抗的な態度・虚勢",           c: C.primary },
    { w: "ヤンキー", m: "日本独自の不良文化・見た目",   c: C.primary },
    { w: "やんちゃ", m: "軽くぼかす婉曲表現",           c: C.green },
    { w: "チンピラ", m: "下っ端のならず者（侮蔑語）",   c: C.orange },
    { w: "暴走族",   m: "車両で危険走行する集団",       c: C.red },
    { w: "半グレ",   m: "暴力団に属さない犯罪集団",     c: C.red },
  ];

  words.forEach(function(it, i) {
    var yPos = 1.15 + i * 0.58;
    addCard(s, 0.6, yPos, 8.8, 0.5);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yPos, w: 0.1, h: 0.5, fill: { color: it.c } });
    s.addText(it.w, {
      x: 0.85, y: yPos, w: 2.2, h: 0.5,
      fontSize: 16, fontFace: FONT_JP, color: it.c, bold: true, valign: "middle", margin: 0,
    });
    s.addText(it.m, {
      x: 3.1, y: yPos, w: 6.2, h: 0.5,
      fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  addPageNum(s, "2", TOTAL);
})();

// ============================================================
// SLIDE 3: ニュアンスの軸
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "フォーカスする「軸」がバラバラ");

  s.addText("同じ「悪い人」に見えても、どこに注目しているかが違う", {
    x: 0.5, y: 1.1, w: 9.0, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  var axes = [
    { label: "態度", word: "ツッパリ", desc: "反抗・虚勢の見せ方" },
    { label: "文化", word: "ヤンキー", desc: "日本独自の美意識・ファッション" },
    { label: "言い換え", word: "やんちゃ", desc: "ソフトに丸める婉曲語" },
    { label: "序列", word: "チンピラ", desc: "格の低さを貶める" },
    { label: "行為", word: "暴走族", desc: "車両の危険走行" },
    { label: "犯罪性", word: "半グレ", desc: "反社会的犯罪集団" },
  ];

  axes.forEach(function(a, i) {
    var col = i % 3;
    var row = Math.floor(i / 3);
    var xPos = 0.6 + col * 3.0;
    var yPos = 1.9 + row * 1.7;
    addCard(s, xPos, yPos, 2.8, 1.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 2.8, h: 0.4, fill: { color: C.primary } });
    s.addText(a.label, {
      x: xPos, y: yPos, w: 2.8, h: 0.4,
      fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(a.word, {
      x: xPos + 0.1, y: yPos + 0.5, w: 2.6, h: 0.45,
      fontSize: 20, fontFace: FONT_JP, color: C.dark, bold: true, align: "center", margin: 0,
    });
    s.addText(a.desc, {
      x: xPos + 0.1, y: yPos + 1.0, w: 2.6, h: 0.45,
      fontSize: 11, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
    });
  });

  addPageNum(s, "3", TOTAL);
})();

// ============================================================
// SLIDE 4: 不良
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "①「不良」— 最も広い総称");

  addCard(s, 0.6, 1.15, 8.8, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 0.15, h: 1.1, fill: { color: C.primary } });
  s.addText("素行や品行が悪い若者全般を指す、\n最も広い言葉（上位概念）", {
    x: 0.95, y: 1.2, w: 8.3, h: 1.0,
    fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
  });

  addCard(s, 0.6, 2.45, 4.25, 2.2);
  s.addText("語源・初出", { x: 0.8, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("・1749年：一般義の用例\n・1905年：素行の悪い青少年の意味\n（精選版 日本国語大辞典）", {
    x: 0.8, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  addCard(s, 5.15, 2.45, 4.25, 2.2);
  s.addText("現代での位置づけ", { x: 5.35, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("・法令・統計では現役\n（非行少年・不良行為少年）\n・日常会話では硬く説明的\n・ヤンキー等の上位概念", {
    x: 5.35, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.85, w: 8.8, h: 0.5, fill: { color: C.light } });
  s.addText("例：「彼女は高校時代、不良グループと付き合っていた」", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "4", TOTAL);
})();

// ============================================================
// SLIDE 5: ツッパリ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "②「ツッパリ」— 反抗的な「態度」");

  addCard(s, 0.6, 1.15, 8.8, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 0.15, h: 1.1, fill: { color: C.primary } });
  s.addText("動詞「突っ張る（虚勢を張る）」が語源。\n行為や集団ではなく『見せ方・態度』に焦点", {
    x: 0.95, y: 1.2, w: 8.3, h: 1.0,
    fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
  });

  addCard(s, 0.6, 2.45, 4.25, 2.2);
  s.addText("語源・初出", { x: 0.8, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("・1984年：名詞用法の実例\n・動詞「突っ張る」は古い\n・70年代末〜80年代に流行", {
    x: 0.8, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  addCard(s, 5.15, 2.45, 4.25, 2.2);
  s.addText("現代での位置づけ", { x: 5.35, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("・日常会話ではほぼ使わない\n・『ろくでなしBLUES』\n　『今日から俺は!!』を語る懐古語\n・リーゼント・短ラン・ボンタン", {
    x: 5.35, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.85, w: 8.8, h: 0.5, fill: { color: C.light } });
  s.addText("例：「校内暴力全盛期のツッパリは、リーゼントで教室を闊歩していた」", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "5", TOTAL);
})();

// ============================================================
// SLIDE 6: ヤンキー
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "③「ヤンキー」— 日本独自の不良文化");

  addCard(s, 0.6, 1.15, 8.8, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 0.15, h: 1.1, fill: { color: C.primary } });
  s.addText("元々は「アメリカ人」の俗称。\n日本では文化・見た目・生活様式まで含む", {
    x: 0.95, y: 1.2, w: 8.3, h: 1.0,
    fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
  });

  addCard(s, 0.6, 2.45, 4.25, 2.2);
  s.addText("語源・初出", { x: 0.8, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("・1904年：アメリカ人の俗称\n・1973年：若者不良文化の意味\n・大阪アメリカ村から関西発祥", {
    x: 0.8, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  addCard(s, 5.15, 2.45, 4.25, 2.2);
  s.addText("現代での位置づけ", { x: 5.35, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("・文化・サブカル語として現役\n・80年代以降「ツッパリ」から\n　「ヤンキー」に中心語が移行\n・金髪・特攻服・改造車文化", {
    x: 5.35, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.85, w: 8.8, h: 0.5, fill: { color: C.light } });
  s.addText("例：「地元の祭りに集まる元ヤンキーたちが主役のドキュメンタリー」", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "6", TOTAL);
})();

// ============================================================
// SLIDE 7: やんちゃ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "④「やんちゃ」— 軽くぼかす婉曲表現");

  addCard(s, 0.6, 1.15, 8.8, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 0.15, h: 1.1, fill: { color: C.green } });
  s.addText("もっとも古い部類の日本語。\n『若い頃ちょっと悪かった』をソフトに言い換える", {
    x: 0.95, y: 1.2, w: 8.3, h: 1.0,
    fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
  });

  addCard(s, 0.6, 2.45, 4.25, 2.2);
  s.addText("語源・初出", { x: 0.8, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true, margin: 0 });
  s.addText("・1702年：子どものわがまま\n・動物のじゃれつきにも使う\n・7語中もっとも古い", {
    x: 0.8, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  addCard(s, 5.15, 2.45, 4.25, 2.2);
  s.addText("現代での位置づけ", { x: 5.35, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true, margin: 0 });
  s.addText("・芸能人の自己紹介で多用\n・喧嘩・夜遊び程度まで\n・重く受け取られない緩衝材\n・本格的犯罪は含まない", {
    x: 5.35, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.85, w: 8.8, h: 0.5, fill: { color: "EAF5EA" } });
  s.addText("例：「学生時代はやんちゃでしたが、今は真面目に働いています」", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "7", TOTAL);
})();

// ============================================================
// SLIDE 8: チンピラ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "⑤「チンピラ」— 下っ端のならず者");

  addCard(s, 0.6, 1.15, 8.8, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 0.15, h: 1.1, fill: { color: C.orange } });
  s.addText("『大物ぶる未熟者』から転じて、\n不良や暴力団の下っ端を指す侮蔑語", {
    x: 0.95, y: 1.2, w: 8.3, h: 1.0,
    fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
  });

  addCard(s, 0.6, 2.45, 4.25, 2.2);
  s.addText("語源・初出", { x: 0.8, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0 });
  s.addText("・1915〜1930年ごろ\n・元：未熟・小物\n・大正〜昭和初期に広がる", {
    x: 0.8, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  addCard(s, 5.15, 2.45, 4.25, 2.2);
  s.addText("現代での位置づけ", { x: 5.35, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0 });
  s.addText("・罵倒語・比喩として現役\n・『序列の低い悪党』\n・ビジネス・政治でも比喩\n・他6語と違い『格』が焦点", {
    x: 5.35, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.85, w: 8.8, h: 0.5, fill: { color: "FDF0E4" } });
  s.addText("例：「大物気取りだが実態はただのチンピラにすぎない」", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "8", TOTAL);
})();

// ============================================================
// SLIDE 9: 暴走族
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "⑥「暴走族」— 車両で危険走行する集団");

  addCard(s, 0.6, 1.15, 8.8, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 0.15, h: 1.1, fill: { color: C.red } });
  s.addText("行為と組織性に焦点を当てた行政用語。\n警察白書で現役", {
    x: 0.95, y: 1.2, w: 8.3, h: 1.0,
    fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
  });

  addCard(s, 0.6, 2.45, 4.25, 2.2);
  s.addText("語源・初出", { x: 0.8, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText("・1973年前後\n・昭和49年警察白書で言及\n・『共同危険行為』を繰り返す\n　二輪・四輪の集団", {
    x: 0.8, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  addCard(s, 5.15, 2.45, 4.25, 2.2);
  s.addText("現代での位置づけ", { x: 5.35, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText("・警察・報道・行政で現役\n・構成員は最盛期の1/10以下\n・『旧車會』など別形態へ移行\n・個人ではなく集団が前提", {
    x: 5.35, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.85, w: 8.8, h: 0.5, fill: { color: "FDEAEA" } });
  s.addText("例：「深夜の国道を走る暴走族の騒音が通報された」", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "9", TOTAL);
})();

// ============================================================
// SLIDE 10: 半グレ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "⑦「半グレ」— 最も重い報道語");

  addCard(s, 0.6, 1.15, 8.8, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 0.15, h: 1.1, fill: { color: C.red } });
  s.addText("暴力団に属さないが反社会的行為を\n繰り返す集団。犯罪集団性が明確", {
    x: 0.95, y: 1.2, w: 8.3, h: 1.0,
    fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
  });

  addCard(s, 0.6, 2.45, 4.25, 2.2);
  s.addText("語源・初出", { x: 0.8, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText("・2010年代に広まる\n・語源は諸説あり\n　（ぐれん隊、グレーゾーン等）\n・定説なし", {
    x: 0.8, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  addCard(s, 5.15, 2.45, 4.25, 2.2);
  s.addText("公的文書での扱い", { x: 5.35, y: 2.55, w: 4.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText("・令和2年白書「準暴力団」\n・令和5年白書で新概念\n　『匿名・流動型犯罪グループ』\n・特殊詐欺・闇バイト等と直結", {
    x: 5.35, y: 2.95, w: 4.0, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.5,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.85, w: 8.8, h: 0.5, fill: { color: "FDEAEA" } });
  s.addText("例：「特殊詐欺の摘発で半グレグループの関与が浮上した」", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "10", TOTAL);
})();

// ============================================================
// SLIDE 11: 時系列
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "時系列で見る7語の歴史");

  var events = [
    { year: "1702", word: "やんちゃ",  desc: "子どもの悪ふざけ",         c: C.green },
    { year: "1749", word: "不良",      desc: "一般義",                   c: C.primary },
    { year: "1904", word: "ヤンキー",  desc: "アメリカ人の俗称",         c: C.sub },
    { year: "1905", word: "不良",      desc: "素行の悪い青少年の意味",   c: C.primary },
    { year: "1915-30", word: "チンピラ", desc: "大正〜昭和初期に広がる", c: C.orange },
    { year: "1973", word: "暴走族 / ヤンキー", desc: "警察白書・若者語に", c: C.red },
    { year: "1984", word: "ツッパリ",  desc: "名詞用法の実例",           c: C.primary },
    { year: "2010s", word: "半グレ",   desc: "メディアで広がる",         c: C.red },
  ];

  events.forEach(function(e, i) {
    var yPos = 1.15 + i * 0.52;
    addCard(s, 0.6, yPos, 8.8, 0.45);
    s.addText(e.year, {
      x: 0.8, y: yPos, w: 1.3, h: 0.45,
      fontSize: 14, fontFace: FONT_EN, color: e.c, bold: true, valign: "middle", margin: 0,
    });
    s.addShape(pres.shapes.LINE, { x: 2.15, y: yPos + 0.05, w: 0, h: 0.35, line: { color: C.muted, width: 1 } });
    s.addText(e.word, {
      x: 2.3, y: yPos, w: 2.6, h: 0.45,
      fontSize: 14, fontFace: FONT_JP, color: C.dark, bold: true, valign: "middle", margin: 0,
    });
    s.addText(e.desc, {
      x: 5.0, y: yPos, w: 4.3, h: 0.45,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  addPageNum(s, "11", TOTAL);
})();

// ============================================================
// SLIDE 12: 混同しやすいペア
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "混同しやすい3つのペア");

  var pairs = [
    { a: "ヤンキー", b: "ツッパリ", diff: "ヤンキー＝文化、ツッパリ＝態度。ツッパリはヤンキー文化の中の『ふるまい』部分" },
    { a: "やんちゃ", b: "チンピラ", diff: "真逆の関係。やんちゃは軽くぼかす婉曲語、チンピラは序列を貶める侮蔑語" },
    { a: "暴走族",   b: "半グレ",   diff: "暴走族は車両集団の行為、半グレは暴力団以外の犯罪集団。行為の質が違う" },
  ];

  pairs.forEach(function(p, i) {
    var yPos = 1.15 + i * 1.35;
    addCard(s, 0.6, yPos, 8.8, 1.15);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yPos, w: 0.15, h: 1.15, fill: { color: C.accent } });

    s.addText(p.a, {
      x: 0.95, y: yPos + 0.1, w: 2.5, h: 0.5,
      fontSize: 20, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
    });
    s.addText("VS", {
      x: 3.5, y: yPos + 0.1, w: 0.6, h: 0.5,
      fontSize: 14, fontFace: FONT_EN, color: C.sub, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(p.b, {
      x: 4.1, y: yPos + 0.1, w: 2.5, h: 0.5,
      fontSize: 20, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
    });

    s.addText(p.diff, {
      x: 0.95, y: yPos + 0.62, w: 8.25, h: 0.48,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
    });
  });

  addPageNum(s, "12", TOTAL);
})();

// ============================================================
// SLIDE 13: 現代での使い分け
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "現代ではどう使い分ける？");

  var uses = [
    { label: "日常会話",   words: "ヤンキー／やんちゃ／不良",       c: C.primary },
    { label: "懐古・昭和語", words: "ツッパリ",                      c: C.sub },
    { label: "侮蔑・比喩",  words: "チンピラ",                       c: C.orange },
    { label: "報道語",     words: "半グレ",                          c: C.red },
    { label: "行政・警察語", words: "暴走族／半グレ（→匿名・流動型犯罪グループ）", c: C.red },
  ];

  uses.forEach(function(u, i) {
    var yPos = 1.2 + i * 0.78;
    addCard(s, 0.6, yPos, 8.8, 0.65);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yPos, w: 2.5, h: 0.65, fill: { color: u.c } });
    s.addText(u.label, {
      x: 0.6, y: yPos, w: 2.5, h: 0.65,
      fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(u.words, {
      x: 3.2, y: yPos, w: 6.1, h: 0.65,
      fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  addPageNum(s, "13", TOTAL);
})();

// ============================================================
// SLIDE 14: 現代の若者の変化
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "現代の若者像はどう変わった？");

  addCard(s, 0.6, 1.15, 4.25, 3.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 4.25, h: 0.55, fill: { color: C.primary } });
  s.addText("昭和〜平成初期", {
    x: 0.6, y: 1.15, w: 4.25, h: 0.55,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("・校内暴力・暴走族全盛\n・ツッパリ／ヤンキー文化\n・派手な見た目で自己主張\n・不良が一種の『ステータス』", {
    x: 0.85, y: 1.85, w: 3.75, h: 2.9,
    fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.7,
  });

  addCard(s, 5.15, 1.15, 4.25, 3.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.15, w: 4.25, h: 0.55, fill: { color: C.accent } });
  s.addText("令和", {
    x: 5.15, y: 1.15, w: 4.25, h: 0.55,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("・従来の不良像は激減\n・目立たず穏やかな傾向\n・不良はSNS上の\n　『匿名・流動型』集団へ\n・闇バイト・特殊詐欺が主戦場", {
    x: 5.4, y: 1.85, w: 3.75, h: 2.9,
    fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.7,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.0, w: 8.8, h: 0.4, fill: { color: C.light } });
  s.addText("出典：博報堂生活総合研究所／警察庁令和5年警察白書", {
    x: 0.8, y: 5.0, w: 8.4, h: 0.4,
    fontSize: 11, fontFace: FONT_JP, color: C.sub, align: "center", valign: "middle", italic: true, margin: 0,
  });

  addPageNum(s, "14", TOTAL);
})();

// ============================================================
// SLIDE 15: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "まとめ");

  s.addText("7語は同じ仲間ではない。焦点の軸が違う", {
    x: 0.5, y: 1.1, w: 9.0, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0,
  });

  var summary = [
    { w: "不良",     s: "広い箱（総称）" },
    { w: "ツッパリ", s: "反抗的ふるまい" },
    { w: "ヤンキー", s: "日本独自の若者文化" },
    { w: "やんちゃ", s: "軽く丸めた言い方" },
    { w: "チンピラ", s: "下っ端のならず者" },
    { w: "暴走族",   s: "乗り物集団" },
    { w: "半グレ",   s: "犯罪集団寄りの報道語" },
  ];

  summary.forEach(function(it, i) {
    var yPos = 1.8 + i * 0.45;
    s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: yPos, w: 0.08, h: 0.35, fill: { color: C.accent } });
    s.addText(it.w, {
      x: 1.3, y: yPos, w: 2.2, h: 0.35,
      fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
    });
    s.addText("＝ " + it.s, {
      x: 3.6, y: yPos, w: 5.4, h: 0.35,
      fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.0, w: 8.8, h: 0.45, fill: { color: C.light } });
  s.addText("言葉の変化を追うことは、社会の変化を追うこと", {
    x: 0.8, y: 5.0, w: 8.4, h: 0.45,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", italic: true, margin: 0,
  });

  addPageNum(s, "15", TOTAL);
})();

// ============================================================
// SLIDE 16: 出典
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "出典");

  var refs = [
    "小学館『デジタル大辞泉』『精選版 日本国語大辞典』（コトバンク経由）",
    "国立国会図書館レファレンス協同データベース「ヤンキーの語源が知りたい」",
    "警察庁『昭和49年警察白書』『令和2年警察白書』『令和5年警察白書』",
    "大阪府警察「暴走族とは」（行政定義）",
    "CiNii『今日から俺は!!にみる“80年代不良表象”の批評性』",
    "博報堂生活総合研究所「不良は消えた？新時代の不良の姿」",
    "難波功士『ヤンキー進化論』／斎藤環『ヤンキー文化論序説』／打越正行『ヤンキーと地元』",
  ];

  refs.forEach(function(r, i) {
    var yPos = 1.25 + i * 0.55;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yPos, w: 0.08, h: 0.4, fill: { color: C.accent } });
    s.addText("・" + r, {
      x: 0.85, y: yPos, w: 8.4, h: 0.4,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  s.addShape(pres.shapes.LINE, { x: 1, y: 5.2, w: 8, h: 0, line: { color: C.muted, width: 1 } });
  s.addText("医知創造ラボ  今村久司", {
    x: 0.5, y: 5.3, w: 9.0, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });

  addPageNum(s, "16", TOTAL);
})();

pres.writeFile({ fileName: "ヤンキー7語_スライド.pptx" }).then(function(name) {
  console.log("Created: " + name);
});
