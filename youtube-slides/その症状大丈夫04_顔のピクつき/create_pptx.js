var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#04】顔がピクピク動く";

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

function hdr(s, title, bgColor) {
  bgColor = bgColor || C.primary;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: bgColor } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 26, fontFace: FJ, color: C.white, bold: true, margin: 0 });
}

function card(s, x, y, w, h, fillColor, borderColor) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: fillColor || C.white }, shadow: shd(), rectRadius: 0.1 });
  if (borderColor) {
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.12, h: h, fill: { color: borderColor }, rectRadius: 0.05 });
  }
}

function ftr(s) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.38, fill: { color: C.dark } });
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #04", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("😊", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("顔がピクピク動く…\nこれって何の病気？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.9, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("片側顔面けいれんと\n放置していいタイプの違い", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.9,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #04", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.7, w: 8.8, h: 0.4,
    fontSize: 12, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 2: こんな経験ありませんか？
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな経験、ありませんか？");

  var items = [
    { emoji: "👁️", text: "まぶたが勝手にピクピクする" },
    { emoji: "😣", text: "片側の目の周りがけいれんする" },
    { emoji: "🗣️", text: "口の端が勝手にひきつる" },
    { emoji: "😰", text: "疲れたときに顔がピクつく" },
  ];

  items.forEach(function(item, i) {
    var y = 1.2 + i * 0.95;
    card(s, 0.6, y, 8.8, 0.8, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.1, w: 0.7, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.1, w: 7.5, h: 0.6, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.5, 5.0, 7.0, 0.2, C.accent);
  ftr(s);
})();

// ============================================================
// SLIDE 3: 顔のピクつきの種類
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "顔のピクつき ― 3つのタイプ");

  var types = [
    { name: "眼瞼ミオキミア", sub: "まぶたのピクピク", desc: "最も多い！\nほとんど心配不要", color: C.green, bg: C.lightGreen },
    { name: "片側顔面けいれん", sub: "片側の目〜口のけいれん", desc: "治療が必要\n脳神経内科へ", color: C.accent, bg: C.lightYellow },
    { name: "眼瞼けいれん", sub: "両側のまぶたの閉じ", desc: "目が開けにくい\n脳神経内科へ", color: C.red, bg: C.lightRed },
  ];

  types.forEach(function(t, i) {
    var x = 0.3 + i * 3.2;
    card(s, x, 1.1, 3.0, 3.9, t.bg, t.color);
    s.addText(t.name, { x: x + 0.15, y: 1.2, w: 2.7, h: 0.55, fontSize: 19, fontFace: FJ, color: t.color, bold: true, align: "center", margin: 0 });
    s.addText(t.sub, { x: x + 0.15, y: 1.75, w: 2.7, h: 0.4, fontSize: 14, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
    s.addShape(pres.shapes.LINE, { x: x + 0.4, y: 2.2, w: 2.2, h: 0, line: { color: t.color, width: 1 } });
    s.addText(t.desc, { x: x + 0.15, y: 2.4, w: 2.7, h: 0.8, fontSize: 18, fontFace: FJ, color: C.text, align: "center", lineSpacingMultiple: 1.2, margin: 0 });
  });

  s.addText("※ まずは「どこが」「どう」動くかで分類します", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: 眼瞼ミオキミア（心配不要タイプ）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "眼瞼ミオキミア ― 心配不要タイプ（最多！）", C.green);

  card(s, 0.5, 1.1, 9.0, 1.0, C.lightGreen, C.green);
  s.addText("まぶたの一部が細かくピクピク → ほとんどの場合、数日で自然に治る", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 21, fontFace: FJ, color: C.green, bold: true, valign: "middle", margin: 0,
  });

  var features = [
    "片方のまぶた（上か下）の一部だけがピクピク動く",
    "自分では気になるが、他人からは気づかれないことが多い",
    "疲労・睡眠不足・カフェイン・ストレスで起きやすい",
    "数日〜数週間で自然に消える",
    "休息・カフェイン減量・睡眠確保で改善",
  ];

  features.forEach(function(f, i) {
    var y = 2.35 + i * 0.55;
    s.addText("●", { x: 0.7, y: y, w: 0.4, h: 0.45, fontSize: 14, color: C.green, margin: 0 });
    s.addText(f, { x: 1.1, y: y, w: 8.2, h: 0.45, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 5: 片側顔面けいれん
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "片側顔面けいれん ― 治療が必要なタイプ");

  card(s, 0.5, 1.1, 9.0, 1.0, C.lightYellow, C.accent);
  s.addText("片側の目の周り → 頬 → 口角へ徐々に広がるけいれん", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 21, fontFace: FJ, color: C.accent, bold: true, valign: "middle", margin: 0,
  });

  card(s, 0.5, 2.4, 4.2, 2.4, C.white, C.accent);
  s.addText("特徴", { x: 0.8, y: 2.5, w: 3.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var feats = [
    "片側だけに起きる",
    "目→頬→口角の順に広がる",
    "緊張やストレスで悪化",
    "寝ている間も起きる",
    "40〜60代の女性に多い",
  ];
  feats.forEach(function(f, i) {
    s.addText("・" + f, { x: 0.8, y: 2.95 + i * 0.35, w: 3.8, h: 0.35, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.3, 2.4, 4.2, 2.4, C.white, C.primary);
  s.addText("原因と治療", { x: 5.6, y: 2.5, w: 3.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("原因: 顔面神経が血管に\n圧迫されること（多くの場合）", {
    x: 5.6, y: 3.0, w: 3.8, h: 0.7, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("治療法:", { x: 5.6, y: 3.75, w: 3.8, h: 0.3, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("① ボツリヌス毒素注射\n② 微小血管減圧術（根治）\n③ 内服薬（補助的）", {
    x: 5.6, y: 4.05, w: 3.8, h: 0.75, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 6: 眼瞼けいれん
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "眼瞼けいれん ― 目が開けにくいタイプ", C.red);

  card(s, 0.5, 1.1, 9.0, 1.0, C.lightRed, C.red);
  s.addText("両目のまぶたが不随意に閉じてしまい、目が開けにくくなる", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 21, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0,
  });

  var features = [
    "「ピクピク」よりも「目が開けにくい」が主症状",
    "両側に起きる（片側顔面けいれんとの大きな違い）",
    "明るい光・風・ストレスで悪化する",
    "重症では実質的に目が見えない状態（機能的失明）に",
    "ドライアイや眼精疲労と誤診されやすい",
  ];

  features.forEach(function(f, i) {
    var y = 2.35 + i * 0.55;
    s.addText("●", { x: 0.7, y: y, w: 0.4, h: 0.45, fontSize: 14, color: C.red, margin: 0 });
    s.addText(f, { x: 1.1, y: y, w: 8.2, h: 0.45, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 7: 3つのタイプ比較表
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "3つのタイプ ― 比較表");

  var rows = [
    ["", "眼瞼ミオキミア", "片側顔面けいれん", "眼瞼けいれん"],
    ["場所", "まぶたの一部", "片側 目→頬→口", "両側のまぶた"],
    ["見え方", "ピクピク", "ギュッと収縮", "目が閉じる"],
    ["原因", "疲労・ストレス", "血管の圧迫", "脳の機能異常"],
    ["経過", "自然に治る", "徐々に悪化", "徐々に悪化"],
    ["治療", "休息で十分", "ボトックス/手術", "ボトックス"],
    ["緊急性", "🟢 低い", "🟡 受診推奨", "🟡 受診推奨"],
  ];

  var colW = [1.8, 2.6, 2.8, 2.8];
  var startX = 0.0;

  rows.forEach(function(row, ri) {
    var y = 1.05 + ri * 0.58;
    var isHeader = ri === 0;
    row.forEach(function(cell, ci) {
      var x = startX;
      for (var k = 0; k < ci; k++) { x += colW[k]; }
      var bgColor = isHeader ? C.primary : (ri % 2 === 0 ? C.warmBg : C.white);
      var txtColor = isHeader ? C.white : C.text;
      var isBold = isHeader || ci === 0;
      s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: colW[ci], h: 0.58, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      s.addText(cell, { x: x + 0.1, y: y, w: colW[ci] - 0.2, h: 0.58, fontSize: 14, fontFace: FJ, color: txtColor, bold: isBold, valign: "middle", margin: 0 });
    });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 8: ボツリヌス毒素治療
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "ボツリヌス毒素注射（ボトックス）とは？");

  card(s, 0.5, 1.1, 9.0, 1.0, C.light, C.primary);
  s.addText("けいれんしている筋肉に少量のボツリヌス毒素を注射し、筋肉の過剰な収縮を和らげる治療", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  var points = [
    { title: "効果", desc: "注射後2〜3日で効き始め\n3〜4ヶ月間持続", icon: "✨", color: C.green, bg: C.lightGreen },
    { title: "回数", desc: "効果が切れたら再注射\n（年3〜4回が目安）", icon: "📅", color: C.primary, bg: C.light },
    { title: "保険適用", desc: "片側顔面けいれんと\n眼瞼けいれんは保険適用", icon: "💰", color: C.accent, bg: C.lightYellow },
    { title: "副作用", desc: "注射部位の腫れ・まぶた下垂\n（一時的・軽度が多い）", icon: "⚠️", color: C.red, bg: C.lightRed },
  ];

  points.forEach(function(p, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 2.35 + row * 1.5;
    card(s, x, y, 4.5, 1.3, p.bg, p.color);
    s.addText(p.icon, { x: x + 0.2, y: y + 0.1, w: 0.6, h: 0.5, fontSize: 24, margin: 0 });
    s.addText(p.title, { x: x + 0.8, y: y + 0.1, w: 3.4, h: 0.4, fontSize: 19, fontFace: FJ, color: p.color, bold: true, margin: 0 });
    s.addText(p.desc, { x: x + 0.8, y: y + 0.55, w: 3.4, h: 0.65, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 9: 危険なサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "⚠️ こんな場合は受診を ― チェックポイント", C.red);

  var signs = [
    "ピクつきが1ヶ月以上続いている",
    "目の周りから頬・口角に広がってきた",
    "目が開けにくくて日常生活に支障がある",
    "顔の片側の麻痺（口角が下がる等）を伴う",
    "急にピクつきが始まり、頭痛や視力障害もある",
  ];

  signs.forEach(function(sign, i) {
    var y = 1.15 + i * 0.78;
    card(s, 0.5, y, 9.0, 0.65, C.lightRed, C.red);
    s.addText((i + 1) + ".", { x: 0.8, y: y + 0.05, w: 0.5, h: 0.55, fontSize: 22, fontFace: FE, color: C.red, bold: true, margin: 0 });
    s.addText(sign, { x: 1.3, y: y + 0.05, w: 7.9, h: 0.55, fontSize: 20, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.5, 5.05, 7.0, 0.15, C.red);
  ftr(s);
})();

// ============================================================
// SLIDE 10: 受診の目安 ― 緊急度3段階
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");

  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "顔面の麻痺を伴う\n頭痛・視力障害がある", action: "脳卒中や腫瘍の\n可能性 → 救急へ", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "ピクつきが広がってきた\n1ヶ月以上続いている", action: "脳神経内科を受診\nボトックス治療の相談", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "まぶただけ・数日で改善\n疲労時のみ出る", action: "休息・睡眠確保\n悪化したら受診", bg: C.lightGreen, color: C.green },
  ];

  levels.forEach(function(lv, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, lv.bg, lv.color);
    s.addText(lv.emoji + " " + lv.label, { x: 0.8, y: y + 0.05, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    s.addText(lv.desc, { x: 0.8, y: y + 0.5, w: 5.0, h: 0.55, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText("→ " + lv.action, { x: 6.0, y: y + 0.15, w: 3.3, h: 0.8, fontSize: 16, fontFace: FJ, color: lv.color, bold: true, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 11: セルフケア
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まぶたのピクつき ― セルフケア");

  var tips = [
    { title: "十分な睡眠", desc: "7〜8時間の睡眠を確保。\n寝不足は最大のトリガー", icon: "😴" },
    { title: "カフェイン減量", desc: "コーヒーは1日2杯まで。\nエナジードリンクも注意", icon: "☕" },
    { title: "目の休息", desc: "PC・スマホの20分ルール。\n20分ごとに20秒遠くを見る", icon: "👁️" },
    { title: "ストレス管理", desc: "リラックス法を取り入れる。\n深呼吸・軽い運動が有効", icon: "🧘" },
  ];

  tips.forEach(function(t, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.15 + row * 2.0;
    card(s, x, y, 4.5, 1.75, C.white, C.primary);
    s.addText(t.icon, { x: x + 0.2, y: y + 0.15, w: 0.7, h: 0.6, fontSize: 28, margin: 0 });
    s.addText(t.title, { x: x + 0.9, y: y + 0.15, w: 3.3, h: 0.5, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(t.desc, { x: x + 0.9, y: y + 0.7, w: 3.3, h: 0.9, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 顔のピクつきで覚えておきたい3つ");

  var points = [
    { num: "1", text: "まぶただけのピクピクは\nほとんど心配不要", sub: "眼瞼ミオキミア（最多）\n休息で自然に治る" },
    { num: "2", text: "片側で広がるけいれんは\n脳神経内科へ", sub: "片側顔面けいれんの可能性\nボトックスで改善できる" },
    { num: "3", text: "顔面麻痺を伴う場合は\nすぐ受診", sub: "脳卒中や腫瘍の\n可能性を除外する" },
  ];

  points.forEach(function(p, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, C.primary);
    s.addText(p.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.9, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.text, { x: 1.5, y: y + 0.05, w: 3.8, h: 0.95, fontSize: 19, fontFace: FJ, color: C.text, bold: true, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText(p.sub, { x: 5.5, y: y + 0.1, w: 3.8, h: 0.9, fontSize: 15, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 13: エンディング
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("ご視聴ありがとうございました", {
    x: 0.6, y: 0.8, w: 8.8, h: 0.7,
    fontSize: 30, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 2.5, y: 1.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("チャンネル登録・高評価よろしくお願いします！", {
    x: 0.6, y: 2.0, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });

  var links = [
    "🔗 ブログ記事で詳しく読む（概要欄にリンク）",
    "🔍 顔のピクつきセルフチェックツール（概要欄にリンク）",
    "📺 次回：寝ている時に足がつる ― 危険なサインの見分け方",
  ];

  links.forEach(function(link, i) {
    s.addText(link, {
      x: 1.5, y: 2.9 + i * 0.55, w: 7.0, h: 0.45,
      fontSize: 17, fontFace: FJ, color: "B0BEC5", align: "left", margin: 0,
    });
  });

  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/facial_twitch_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
