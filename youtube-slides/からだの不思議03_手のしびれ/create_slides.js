var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "「手がしびれて目が覚めた」の正体 ― 朝のしびれの原因と受診のタイミング";

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

  s.addText("脳神経内科医が答える からだの不思議 #03", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("「手がしびれて\n目が覚めた」の正体", {
    x: 0.8, y: 1.1, w: 8.4, h: 1.8,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 朝のしびれの原因と受診のタイミング ―", {
    x: 0.8, y: 3.0, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.75, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.05, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: こんな経験ありませんか？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんな経験ありませんか？");

  var quotes = [
    "「夜中に手がしびれて目が覚めた。\n  手を振ったら治ったけど、また繰り返す」",
    "「朝起きると手がジンジンしている。\n  特に親指・人差し指・中指が気になる」",
    "「横向きに寝ていたら、朝に手首が下がって\n  力が入らなくなった」",
  ];
  quotes.forEach(function(q, i) {
    var yPos = 1.2 + i * 1.3;
    addCard(s, 1.0, yPos, 8.0, 1.1);
    s.addText(q, {
      x: 1.3, y: yPos + 0.1, w: 7.4, h: 0.9,
      fontSize: 17, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
      lineSpacingMultiple: 1.2,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 5.0, w: 8.0, h: 0.5, fill: { color: C.light } });
  s.addText("外来でよく相談を受ける症状 ― 原因は1つではありません", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/11");
})();

// ============================================================
// SLIDE 3: しびれとは何か
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "しびれとは何か ― 神経で何が起きているか");

  // 上部カード：しびれの種類
  addCard(s, 0.5, 1.1, 9.0, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 9.0, h: 0.5, fill: { color: C.primary } });
  s.addText("しびれの種類", { x: 0.8, y: 1.1, w: 8.4, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("ジンジン感（異常感覚） ／ 感覚が薄くなる（感覚鈍麻） ／ ピリピリ感 ／ 力が入らない", {
    x: 0.8, y: 1.65, w: 8.4, h: 0.8,
    fontSize: 16, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
  });

  // 神経の種類カード
  var nerves = [
    { name: "末梢神経", desc: "手足・体幹の\n末端を走る神経", color: C.accent },
    { name: "脊髄・神経根", desc: "背骨の中を通る\n中枢から末梢への中継点", color: C.orange },
    { name: "脳（中枢）", desc: "大脳・視床・\n感覚野の障害", color: C.red },
  ];
  nerves.forEach(function(n, i) {
    var xPos = 0.4 + i * 3.1;
    addCard(s, xPos, 2.9, 2.9, 1.7);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.9, w: 2.9, h: 0.5, fill: { color: n.color } });
    s.addText(n.name, { x: xPos, y: 2.9, w: 2.9, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(n.desc, { x: xPos, y: 3.45, w: 2.9, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.85, w: 9.0, h: 0.6, fill: { color: C.light } });
  s.addText("重要：しびれの「場所のパターン」でどの神経か特定できる ― これが鑑別の核心", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.6,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "3/11");
})();

// ============================================================
// SLIDE 4: 鑑別の最初の一歩 ― どの指か
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "鑑別の最初の一歩 ― 「どの指か」を確認する");

  // 3神経カード
  var nerveData = [
    { name: "正中神経", yomi: "せいちゅうしんけい", area: "親指・人差し指・中指・薬指の半分", route: "手根管を通る", color: C.green },
    { name: "尺骨神経", yomi: "しゃっこつしんけい", area: "小指・薬指の半分", route: "肘の内側を通る", color: C.orange },
    { name: "橈骨神経", yomi: "とうこつしんけい", area: "手背・親指〜中指の背側", route: "上腕部を走る", color: C.accent },
  ];
  nerveData.forEach(function(n, i) {
    var xPos = 0.4 + i * 3.1;
    addCard(s, xPos, 1.1, 2.9, 3.2);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 2.9, h: 0.55, fill: { color: n.color } });
    s.addText(n.name, { x: xPos, y: 1.1, w: 2.9, h: 0.55, fontSize: 17, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText("（" + n.yomi + "）", { x: xPos, y: 1.7, w: 2.9, h: 0.4, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0 });
    s.addText("担当エリア", { x: xPos + 0.2, y: 2.15, w: 2.5, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.sub, margin: 0 });
    s.addText(n.area, { x: xPos + 0.2, y: 2.5, w: 2.5, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, margin: 0, lineSpacingMultiple: 1.2 });
    s.addText("経路：" + n.route, { x: xPos + 0.2, y: 3.25, w: 2.5, h: 0.7, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0, lineSpacingMultiple: 1.2 });
  });

  // ポイントバナー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 9.0, h: 0.9, fill: { color: C.dark } });
  s.addText([
    { text: "KEY: ", options: { color: C.yellow, bold: true, fontSize: 16 } },
    { text: "「小指が入るかどうか」が最初の分岐点", options: { color: C.white, fontSize: 16 } },
  ], { x: 0.8, y: 4.5, w: 8.4, h: 0.9, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, "4/11");
})();

// ============================================================
// SLIDE 5: 手根管症候群
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "最も多い原因 ― 手根管症候群（しゅこんかんしょうこうぐん）");

  // メインカード
  addCard(s, 0.5, 1.1, 9.0, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 9.0, h: 0.5, fill: { color: C.green } });
  s.addText("手首の正中神経が「手根管」というトンネルで圧迫される疾患", {
    x: 0.8, y: 1.1, w: 8.4, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });
  s.addText("特に40〜60代の女性に多い（妊娠中・透析患者でも起こりやすい）", {
    x: 0.8, y: 1.65, w: 8.4, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
  });

  // 特徴的なサインカード
  var features = [
    { icon: "👌", title: "しびれる場所", desc: "親指・人差し指・中指\n（小指にはしびれなし！）" },
    { icon: "🌙", title: "悪化するとき", desc: "夜間〜早朝に強くなる\n手首を曲げた寝姿勢で悪化" },
    { icon: "✋", title: "Flick sign", desc: "手を振る・腕を高く上げると\n一時的に楽になる" },
  ];
  features.forEach(function(f, i) {
    var xPos = 0.4 + i * 3.1;
    addCard(s, xPos, 2.65, 2.9, 2.2);
    s.addText(f.icon, { x: xPos, y: 2.7, w: 2.9, h: 0.7, fontSize: 28, align: "center", margin: 0 });
    s.addText(f.title, { x: xPos, y: 3.4, w: 2.9, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0 });
    s.addText(f.desc, { x: xPos + 0.15, y: 3.9, w: 2.6, h: 0.85, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.2 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.05, w: 9.0, h: 0.45, fill: { color: C.light } });
  s.addText("筋萎縮が起きてからでは回復に時間がかかる ― 気になり始めたら早めに受診を", {
    x: 0.8, y: 5.05, w: 8.4, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "5/11");
})();

// ============================================================
// SLIDE 6: 肘部管症候群・胸郭出口症候群
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "小指・薬指のしびれ ― 肘と首の問題");

  // 肘部管症候群
  addCard(s, 0.4, 1.1, 4.4, 4.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 0.55, fill: { color: C.orange } });
  s.addText("肘部管症候群（ちゅうぶかんしょうこうぐん）", {
    x: 0.55, y: 1.1, w: 4.1, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });
  var elbow = [
    "障害神経: 尺骨神経（しゃっこつしんけい）",
    "部位: 肘の内側で圧迫",
    "しびれ: 小指・薬指（半分）",
    "悪化: 肘を曲げ続けると悪化\n　　　（横向き就寝時も）",
    "進行: 手の筋肉がやせる\n　　　「鷲手（わしで）変形」",
  ];
  elbow.forEach(function(t, i) {
    s.addText(t, {
      x: 0.6, y: 1.8 + i * 0.62, w: 4.0, h: 0.55,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.1,
    });
  });

  // 胸郭出口症候群
  addCard(s, 5.2, 1.1, 4.4, 4.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.55, fill: { color: C.accent } });
  s.addText("胸郭出口症候群（きょうかくでぐちしょうこうぐん）", {
    x: 5.35, y: 1.1, w: 4.1, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });
  var tos = [
    "障害神経: 腕神経叢（うでのしんけいそう）",
    "部位: 鎖骨と第一肋骨の間で圧迫",
    "しびれ: 小指・薬指側・前腕内側",
    "悪化: 腕を上げたとき\n　　　（ドライヤー・つり革）",
    "特徴: なで肩・首が長い女性に多い\n　　　肩こりを伴うことが多い",
  ];
  tos.forEach(function(t, i) {
    s.addText(t, {
      x: 5.35, y: 1.8 + i * 0.62, w: 4.1, h: 0.55,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.1,
    });
  });

  addPageNum(s, "6/11");
})();

// ============================================================
// SLIDE 7: 頸椎症・橈骨神経麻痺
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "首から来るしびれ・翌朝の麻痺");

  // 頸椎症性神経根症
  addCard(s, 0.4, 1.1, 4.4, 3.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 0.55, fill: { color: C.primary } });
  s.addText("頸椎症性神経根症（けいついしょうせいしんけいこんしょう）", {
    x: 0.55, y: 1.1, w: 4.1, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });
  var cerv = [
    "原因: 頸椎の骨・椎間板の変性\n　　　↓神経根（しんけいこん）を圧迫",
    "年齢: 50歳以降に多い\n　　　（デスクワーク・スマホ負担）",
    "しびれ: 手指〜前腕〜上腕まで広がる",
    "診断: 首の後屈で悪化\n　　　（Spurlingテスト）",
  ];
  cerv.forEach(function(t, i) {
    s.addText(t, {
      x: 0.6, y: 1.8 + i * 0.66, w: 4.0, h: 0.6,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.1,
    });
  });

  // 橈骨神経麻痺
  addCard(s, 5.2, 1.1, 4.4, 3.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.55, fill: { color: C.sub } });
  s.addText("Saturday night palsy（橈骨神経麻痺）", {
    x: 5.35, y: 1.1, w: 4.1, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });
  var radial = [
    "原因: 上腕部の橈骨神経（とうこつしんけい）を\n　　　圧迫したまま眠った翌朝に出現",
    "症状: 手首が下がる「下垂手（かすいしゅ）」\n　　　手首・指が持ち上げられない",
    "経過: 多くは数週間〜数か月で\n　　　自然回復する",
    "予防: 飲酒後に固い台で腕を圧迫\n　　　して眠らないよう注意",
  ];
  radial.forEach(function(t, i) {
    s.addText(t, {
      x: 5.35, y: 1.8 + i * 0.66, w: 4.1, h: 0.6,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.1,
    });
  });

  // ポイント
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.75, w: 9.2, h: 0.65, fill: { color: C.light } });
  s.addText("頸椎症はしびれが手だけでなく前腕・上腕まで広がるのが特徴。首・肩の痛みを伴うことも多い", {
    x: 0.6, y: 4.75, w: 8.8, h: 0.65, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "7/11");
})();

// ============================================================
// SLIDE 8: 6つの原因を比較する
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "6つの原因を比較する ― 「どの指か」と「悪化状況」が核心");

  var rows = [
    { name: "手根管症候群", area: "親指・人差し指・中指（小指は正常）", trigger: "夜間〜早朝 / 手首屈曲姿勢", note: "40〜60代女性多い", color: C.green },
    { name: "胸郭出口症候群", area: "小指・薬指側 / 前腕内側", trigger: "腕を上げたとき / 重い荷物", note: "なで肩の女性に多い", color: C.accent },
    { name: "肘部管症候群", area: "小指・薬指（半分）", trigger: "肘を曲げ続けると悪化", note: "進行で鷲手変形", color: C.orange },
    { name: "頸椎症性神経根症", area: "手指〜前腕〜上腕まで広がる", trigger: "首の後屈（Spurling）で誘発", note: "50代以降 肩の痛み伴う", color: C.primary },
    { name: "橈骨神経麻痺", area: "手背・親指〜中指の背側", trigger: "上腕圧迫翌朝に出現", note: "下垂手 多くは自然回復", color: C.sub },
    { name: "脳卒中", area: "手全体〜顔・下肢まで同側に", trigger: "突然発症（数秒〜数分）", note: "麻痺・言語障害を伴う", color: C.red },
  ];

  rows.forEach(function(r, i) {
    var yPos = 1.1 + i * 0.72;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: yPos, w: 9.6, h: 0.68, fill: { color: i % 2 === 0 ? C.white : C.warmBg }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: yPos, w: 0.15, h: 0.68, fill: { color: r.color } });
    s.addText(r.name, { x: 0.4, y: yPos, w: 2.2, h: 0.68, fontSize: 12, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(r.area, { x: 2.65, y: yPos, w: 2.7, h: 0.68, fontSize: 11, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
    s.addText(r.trigger, { x: 5.4, y: yPos, w: 2.6, h: 0.68, fontSize: 11, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
    s.addText(r.note, { x: 8.05, y: yPos, w: 1.7, h: 0.68, fontSize: 10, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0 });
  });

  // ヘッダーラベル
  s.addText("疾患名", { x: 0.4, y: 0.88, w: 2.2, h: 0.2, fontSize: 11, fontFace: FONT_JP, color: C.sub, bold: true, margin: 0 });
  s.addText("しびれる部位", { x: 2.65, y: 0.88, w: 2.7, h: 0.2, fontSize: 11, fontFace: FONT_JP, color: C.sub, bold: true, margin: 0 });
  s.addText("悪化状況", { x: 5.4, y: 0.88, w: 2.6, h: 0.2, fontSize: 11, fontFace: FONT_JP, color: C.sub, bold: true, margin: 0 });
  s.addText("その他", { x: 8.05, y: 0.88, w: 1.7, h: 0.2, fontSize: 11, fontFace: FONT_JP, color: C.sub, bold: true, margin: 0 });

  addPageNum(s, "8/11");
})();

// ============================================================
// SLIDE 9: 受診を考えるサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "受診を考えるサイン ― こんなときは専門医へ");

  var signs = [
    "しびれが毎晩・毎朝繰り返される",
    "しびれが日中もある（夜間・朝だけでなくなった）",
    "手に力が入りにくい・物を落としやすくなった",
    "手の筋肉がやせてきた（手のひらのふくらみが減った）",
    "しびれが手首から肩・首まで広がっている",
    "両手に同時にしびれがある",
    "症状が数か月以上続いている",
    "しびれとともに首や肩の痛みが強い",
  ];

  signs.forEach(function(t, i) {
    var col = i < 4 ? 0 : 1;
    var row = i < 4 ? i : i - 4;
    var xPos = 0.4 + col * 4.8;
    var yPos = 1.15 + row * 0.97;
    addCard(s, xPos, yPos, 4.5, 0.85);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 0.12, h: 0.85, fill: { color: C.red } });
    s.addText(t, {
      x: xPos + 0.25, y: yPos, w: 4.1, h: 0.85,
      fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.1, w: 9.2, h: 0.38, fill: { color: C.light } });
  s.addText("筋萎縮が起きてからでは回復に時間がかかる ― 気になり始めた段階での受診が治療の選択肢を広げる", {
    x: 0.6, y: 5.1, w: 8.8, h: 0.38, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "9/11");
})();

// ============================================================
// SLIDE 10: 脳卒中のしびれを見逃さない（救急）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.red } });
  s.addText("これは救急！ ― 脳卒中のしびれを見逃さない", {
    x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0,
  });

  // FAST解説
  var fastItems = [
    { letter: "F", word: "Face（顔のゆがみ）", desc: "顔がゆがむ・口が動きにくい" },
    { letter: "A", word: "Arm（腕の麻痺）", desc: "腕に力が入らない" },
    { letter: "S", word: "Speech（言語障害）", desc: "言葉が出ない・相手の言葉がわからない" },
    { letter: "T", word: "Time（すぐ119番！）", desc: "疑ったら今すぐ救急車を呼ぶ" },
  ];
  fastItems.forEach(function(f, i) {
    var xPos = 0.3 + i * 2.35;
    addCard(s, xPos, 1.1, 2.2, 2.3);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 2.2, h: 0.7, fill: { color: f.letter === "T" ? C.red : C.dark } });
    s.addText(f.letter, { x: xPos, y: 1.1, w: 2.2, h: 0.7, fontSize: 32, fontFace: FONT_EN, color: C.yellow, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.word, { x: xPos + 0.1, y: 1.85, w: 2.0, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0, lineSpacingMultiple: 1.1 });
    s.addText(f.desc, { x: xPos + 0.1, y: 2.4, w: 2.0, h: 0.85, fontSize: 12, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // 救急のサイン
  addCard(s, 0.4, 3.65, 9.2, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.65, w: 9.2, h: 0.45, fill: { color: C.red } });
  s.addText("すぐ119番すべきしびれのサイン", { x: 0.6, y: 3.65, w: 8.8, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  var emgSigns = "突然始まったしびれ（数秒〜数分で完成）  ／  しびれ＋手足の麻痺  ／  顔のゆがみ  ／  言葉が出ない\nひどい頭痛が同時に出た  ／  視野の一部が欠ける・二重に見える";
  s.addText(emgSigns, {
    x: 0.6, y: 4.15, w: 8.8, h: 0.95, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.25,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.2, w: 9.2, h: 0.35, fill: { color: C.dark } });
  s.addText("脳卒中治療のウィンドウは発症から4.5時間以内 ― 「様子を見よう」では間に合わない", {
    x: 0.6, y: 5.2, w: 8.8, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.yellow, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "10/11");
})();

// ============================================================
// SLIDE 11: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("まとめ ― Take Home Message", {
    x: 0.5, y: 0.15, w: 9, h: 0.7, fontSize: 22, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var msgs = [
    { num: "1", text: "「どの指か」「いつ悪化するか」が鑑別の核心。小指が入るかどうかが最初の分岐点", color: C.accent },
    { num: "2", text: "最も多いのは手根管症候群（正中神経圧迫）。小指にしびれが出ないのが特徴", color: C.green },
    { num: "3", text: "手の筋萎縮・力の低下・日中も続くしびれは早めの受診サイン", color: C.yellow },
    { num: "4", text: "突然発症のしびれ＋麻痺・言語障害・顔のゆがみは脳卒中 → 即119番！", color: C.red },
  ];

  msgs.forEach(function(m, i) {
    var yPos = 1.05 + i * 1.05;
    addCard(s, 0.4, yPos, 9.2, 0.92);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 0.55, h: 0.92, fill: { color: m.color } });
    s.addText(m.num, { x: 0.4, y: yPos, w: 0.55, h: 0.92, fontSize: 22, fontFace: FONT_EN, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, {
      x: 1.1, y: yPos, w: 8.3, h: 0.92,
      fontSize: 16, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  s.addText("チャンネル登録・コメントをお待ちしています！", {
    x: 0.5, y: 5.15, w: 9, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  addPageNum(s, "11/11");
})();

// ============================================================
// WRITE FILE
// ============================================================
var outPath = __dirname + "/手のしびれ_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
