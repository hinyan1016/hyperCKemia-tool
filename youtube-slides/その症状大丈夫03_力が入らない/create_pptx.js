var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#03】力が入らない…脳？神経？筋肉？";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #03", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("💪", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("突然手に力が入らない…\n脳？ 神経？ 筋肉？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.9, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("脳神経内科医が教える\n原因の見分け方", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.9,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #03", {
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
    { emoji: "🥢", text: "箸やペンを持つ手に力が入らない" },
    { emoji: "🧴", text: "ペットボトルのキャップが開けられない" },
    { emoji: "🚶", text: "階段を上るとき脚がガクッとする" },
    { emoji: "👕", text: "荷物を持つ腕がすぐに疲れてしまう" },
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
// SLIDE 3: 「力が入らない」とは？
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「力が入らない」の正体 ― 筋力低下");

  card(s, 0.5, 1.1, 9.0, 1.2, C.light, C.primary);
  s.addText("筋力低下（きんりょくていか）", { x: 0.8, y: 1.15, w: 6, h: 0.45, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("筋肉に力を入れようとしても十分な力が\n発揮できない状態。原因は3つの場所に分けられる", {
    x: 0.8, y: 1.55, w: 8.5, h: 0.65, fontSize: 19, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  var parts = [
    { title: "脳", desc: "命令を出す司令塔", icon: "🧠", color: C.red, bg: C.lightRed },
    { title: "神経", desc: "命令を伝える電線", icon: "⚡", color: C.accent, bg: C.lightYellow },
    { title: "筋肉", desc: "命令を実行するエンジン", icon: "💪", color: C.green, bg: C.lightGreen },
  ];

  parts.forEach(function(p, i) {
    var x = 0.4 + i * 3.15;
    card(s, x, 2.6, 2.95, 2.3, p.bg, p.color);
    s.addText(p.icon, { x: x + 0.1, y: 2.7, w: 2.75, h: 0.7, fontSize: 36, align: "center", margin: 0 });
    s.addText(p.title, { x: x + 0.1, y: 3.4, w: 2.75, h: 0.5, fontSize: 24, fontFace: FJ, color: p.color, bold: true, align: "center", margin: 0 });
    s.addText(p.desc, { x: x + 0.1, y: 3.9, w: 2.75, h: 0.5, fontSize: 17, fontFace: FJ, color: C.text, align: "center", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: 脳が原因の筋力低下
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳が原因 ― 突然の片側の脱力", C.red);

  card(s, 0.5, 1.1, 9.0, 1.0, C.lightRed, C.red);
  s.addText("片側の手足に突然力が入らなくなった = 脳卒中の可能性大", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 21, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0,
  });

  var features = [
    "右手と右足など、同じ側の手足が同時に弱くなる",
    "顔の片側の麻痺やろれつの障害を伴うことが多い",
    "脳梗塞・脳出血・TIA（一過性脳虚血発作）が原因",
    "発症時刻を記録し、すぐに119番！",
  ];

  features.forEach(function(f, i) {
    var y = 2.4 + i * 0.65;
    s.addText("●", { x: 0.7, y: y, w: 0.4, h: 0.5, fontSize: 14, color: C.red, margin: 0 });
    s.addText(f, { x: 1.1, y: y, w: 8.2, h: 0.5, fontSize: 19, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.0, 4.95, 8.0, 0.25, C.red);

  ftr(s);
})();

// ============================================================
// SLIDE 5: 神経が原因
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "神経が原因 ― しびれを伴う脱力");

  var diseases = [
    { name: "ギラン・バレー症候群", features: "風邪の1〜3週間後に両足から\n力が入らなくなり上に広がる", color: C.red, bg: C.lightRed },
    { name: "手根管症候群", features: "手首の神経の圧迫で親指側の\n力が入りにくい・しびれる", color: C.primary, bg: C.light },
    { name: "頚椎症性脊髄症", features: "首の骨の変形で脊髄が圧迫\n手のしびれ・巧緻運動障害", color: C.accent, bg: C.lightYellow },
    { name: "末梢神経障害", features: "糖尿病やビタミン欠乏で\n手足の先からしびれ・脱力", color: C.green, bg: C.lightGreen },
  ];

  diseases.forEach(function(d, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.15 + row * 2.0;
    card(s, x, y, 4.5, 1.75, d.bg, d.color);
    s.addText(d.name, { x: x + 0.25, y: y + 0.1, w: 4.0, h: 0.5, fontSize: 20, fontFace: FJ, color: d.color, bold: true, margin: 0 });
    s.addText(d.features, { x: x + 0.25, y: y + 0.65, w: 4.0, h: 0.9, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 6: 筋肉が原因
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "筋肉が原因 ― 体幹・近位筋の脱力");

  var diseases = [
    { name: "重症筋無力症", features: "夕方に悪化・まぶたが下がる\n繰り返す動作で力が落ちる", color: C.red, bg: C.lightRed },
    { name: "筋炎（多発性・皮膚筋炎）", features: "太ももや上腕の力が入らない\n筋肉痛・CK高値", color: C.primary, bg: C.light },
    { name: "甲状腺機能異常", features: "バセドウ病・甲状腺機能低下症\nでも筋力低下が起きる", color: C.accent, bg: C.lightYellow },
    { name: "低カリウム血症", features: "利尿薬・下痢・嘔吐が原因で\n急に全身の力が入らなくなる", color: C.green, bg: C.lightGreen },
  ];

  diseases.forEach(function(d, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.15 + row * 2.0;
    card(s, x, y, 4.5, 1.75, d.bg, d.color);
    s.addText(d.name, { x: x + 0.25, y: y + 0.1, w: 4.0, h: 0.5, fontSize: 20, fontFace: FJ, color: d.color, bold: true, margin: 0 });
    s.addText(d.features, { x: x + 0.25, y: y + 0.65, w: 4.0, h: 0.9, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 7: 見分けるポイント比較表
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "原因を見分けるポイント ― 比較表");

  var rows = [
    ["", "脳", "神経", "筋肉"],
    ["発症", "突然", "数日〜数週間", "数週間〜数ヶ月"],
    ["部位", "片側（手足とも）", "一部の筋肉", "両側・体幹に近い"],
    ["しびれ", "片側に伴う", "しびれが主症状", "通常なし"],
    ["腱反射", "亢進", "低下〜消失", "正常〜低下"],
    ["代表疾患", "脳卒中", "ギラン・バレー", "筋炎・MG"],
  ];

  var colW = [1.8, 2.6, 2.8, 2.8];
  var startX = 0.0;

  rows.forEach(function(row, ri) {
    var y = 1.1 + ri * 0.65;
    var isHeader = ri === 0;
    row.forEach(function(cell, ci) {
      var x = startX;
      for (var k = 0; k < ci; k++) { x += colW[k]; }
      var bgColor = isHeader ? C.primary : (ri % 2 === 0 ? C.warmBg : C.white);
      var txtColor = isHeader ? C.white : C.text;
      var isBold = isHeader || ci === 0;
      s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: colW[ci], h: 0.65, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      s.addText(cell, { x: x + 0.1, y: y, w: colW[ci] - 0.2, h: 0.65, fontSize: 15, fontFace: FJ, color: txtColor, bold: isBold, valign: "middle", margin: 0 });
    });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 8: その他の原因
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "その他の原因 ― 病気以外も");

  var causes = [
    { title: "運動不足・廃用", desc: "使わない筋肉は衰える。\n特に高齢者は急速に進行", icon: "🛋️", color: C.primary, bg: C.light },
    { title: "栄養不足", desc: "たんぱく質・ビタミンB群の\n不足で筋力低下が起きる", icon: "🥗", color: C.green, bg: C.lightGreen },
    { title: "薬の副作用", desc: "スタチン（コレステロール薬）・\nステロイドなどで筋力低下", icon: "💊", color: C.red, bg: C.lightRed },
    { title: "精神的な原因", desc: "うつ病・過労・ストレスで\n体に力が入らないと感じる", icon: "😔", color: C.accent, bg: C.lightYellow },
  ];

  causes.forEach(function(c, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.15 + row * 2.0;
    card(s, x, y, 4.5, 1.75, c.bg, c.color);
    s.addText(c.icon, { x: x + 0.2, y: y + 0.15, w: 0.7, h: 0.6, fontSize: 28, margin: 0 });
    s.addText(c.title, { x: x + 0.9, y: y + 0.15, w: 3.3, h: 0.5, fontSize: 20, fontFace: FJ, color: c.color, bold: true, margin: 0 });
    s.addText(c.desc, { x: x + 0.9, y: y + 0.7, w: 3.3, h: 0.9, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 9: 危険なサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "⚠️ こんな脱力はすぐ受診 ― 危険なサイン5つ", C.red);

  var signs = [
    "片側の手足に突然力が入らない（脳卒中の疑い）",
    "両足から急速に広がる脱力（ギラン・バレー症候群）",
    "飲み込みや呼吸がしにくくなってきた",
    "排尿・排便のコントロールが効かない",
    "数日〜数週間で明らかに悪化し続けている",
  ];

  signs.forEach(function(sign, i) {
    var y = 1.15 + i * 0.78;
    card(s, 0.5, y, 9.0, 0.65, C.lightRed, C.red);
    s.addText((i + 1) + ".", { x: 0.8, y: y + 0.05, w: 0.5, h: 0.55, fontSize: 22, fontFace: FE, color: C.red, bold: true, margin: 0 });
    s.addText(sign, { x: 1.3, y: y + 0.05, w: 7.9, h: 0.55, fontSize: 19, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
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
    { emoji: "🔴", label: "すぐ119番（救急）", desc: "突然の片側脱力\n顔の麻痺・ろれつを伴う", action: "発症時刻を記録し\n救急搬送", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "数日で進行する脱力\n飲み込み・呼吸に影響", action: "脳神経内科を\n速やかに受診", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "疲労後のみ・すぐ回復\n運動不足の自覚あり", action: "生活習慣を見直し\n悪化したら受診", bg: C.lightGreen, color: C.green },
  ];

  levels.forEach(function(lv, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, lv.bg, lv.color);
    s.addText(lv.emoji + " " + lv.label, { x: 0.8, y: y + 0.05, w: 3.5, h: 0.45, fontSize: 21, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    s.addText(lv.desc, { x: 0.8, y: y + 0.5, w: 5.0, h: 0.55, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText("→ " + lv.action, { x: 6.0, y: y + 0.15, w: 3.3, h: 0.8, fontSize: 16, fontFace: FJ, color: lv.color, bold: true, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 11: 受診前に確認すること
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診前にチェック ― 医師に伝えること");

  var checks = [
    { title: "いつから？", desc: "突然か、徐々にか\n発症の日時を思い出す", icon: "📅" },
    { title: "どこが弱い？", desc: "片側？両側？\n手？足？体幹？", icon: "🎯" },
    { title: "他の症状は？", desc: "しびれ・痛み・まぶたの下がり\n飲み込みにくさなどの有無", icon: "📝" },
    { title: "薬や既往歴は？", desc: "飲んでいる薬・持病\n最近の風邪やワクチン", icon: "💊" },
  ];

  checks.forEach(function(c, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.15 + row * 2.0;
    card(s, x, y, 4.5, 1.75, C.white, C.primary);
    s.addText(c.icon, { x: x + 0.2, y: y + 0.15, w: 0.7, h: 0.6, fontSize: 28, margin: 0 });
    s.addText(c.title, { x: x + 0.9, y: y + 0.15, w: 3.3, h: 0.5, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(c.desc, { x: x + 0.9, y: y + 0.7, w: 3.3, h: 0.9, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 力が入らない時に覚えておきたい3つ");

  var points = [
    { num: "1", text: "原因は脳・神経・筋肉\nの3か所に分けて考える", sub: "脳 = 突然・片側\n神経 = しびれを伴う\n筋肉 = 両側・近位" },
    { num: "2", text: "「突然・片側」は\n脳卒中 → 即119番", sub: "顔の麻痺やろれつを伴えば\nFASTチェックで確認" },
    { num: "3", text: "進行する脱力は\n早めに脳神経内科へ", sub: "原因の特定には\n専門的な検査が必要" },
  ];

  points.forEach(function(p, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, C.primary);
    s.addText(p.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.9, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.text, { x: 1.5, y: y + 0.05, w: 3.8, h: 0.95, fontSize: 19, fontFace: FJ, color: C.text, bold: true, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText(p.sub, { x: 5.5, y: y + 0.05, w: 3.8, h: 0.95, fontSize: 15, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.2, margin: 0 });
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
    "🔍 脱力セルフチェックツール（概要欄にリンク）",
    "📺 次回：顔がピクピク動く ― 放置していいタイプの違い",
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
var outPath = __dirname + "/weakness_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
