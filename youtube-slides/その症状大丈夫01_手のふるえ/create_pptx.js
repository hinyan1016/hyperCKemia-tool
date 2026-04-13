var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#01】手が震える…パーキンソン病？";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #01", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("✋", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("手が震える…\nこれってパーキンソン病？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.9, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("脳神経内科医が教える\n振戦の見分け方と受診の目安", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.9,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #01", {
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
    { emoji: "☕", text: "コーヒーカップを持つと手が震える" },
    { emoji: "✍️", text: "字を書くとき、線がガタガタになる" },
    { emoji: "🍽️", text: "箸で食べ物をつまみにくい" },
    { emoji: "📱", text: "スマホを持つ手が小刻みに揺れる" },
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
// SLIDE 3: 手の震え（振戦）とは
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "手の震え（振戦）とは？");

  card(s, 0.5, 1.1, 9.0, 1.2, C.light, C.primary);
  s.addText("振戦（しんせん）", { x: 0.8, y: 1.15, w: 4, h: 0.45, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("自分の意思とは関係なく、筋肉が\n規則的・リズミカルに収縮する運動", {
    x: 0.8, y: 1.55, w: 8.5, h: 0.65, fontSize: 20, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  card(s, 0.5, 2.6, 4.2, 1.5, C.white, C.accent);
  s.addText("40歳以上の約4%", { x: 0.8, y: 2.7, w: 3.8, h: 0.5, fontSize: 24, fontFace: FE, color: C.accent, bold: true, margin: 0 });
  s.addText("に何らかの振戦がある", { x: 0.8, y: 3.2, w: 3.8, h: 0.4, fontSize: 18, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("→ 珍しい症状ではない！", { x: 0.8, y: 3.6, w: 3.8, h: 0.4, fontSize: 16, fontFace: FJ, color: C.sub, margin: 0 });

  card(s, 5.3, 2.6, 4.2, 1.5, C.white, C.green);
  s.addText("原因を見分けるカギ", { x: 5.6, y: 2.7, w: 3.8, h: 0.5, fontSize: 22, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("「いつ、どんなときに\n　震えるか」を観察する", {
    x: 5.6, y: 3.2, w: 3.8, h: 0.8, fontSize: 20, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: 振戦の3つのタイプ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "振戦の3つのタイプ ― いつ震えるかで分類");

  var types = [
    { name: "安静時振戦", when: "じっとしている時", cause: "パーキンソン病", color: C.red, bg: C.lightRed },
    { name: "姿勢時・動作時振戦", when: "手を伸ばした時\n物を持つ時", cause: "本態性振戦（最多）\n甲状腺・薬剤性", color: C.primary, bg: C.light },
    { name: "企図振戦", when: "目標に手を\n近づける時", cause: "小脳疾患", color: C.accent, bg: C.lightYellow },
  ];

  types.forEach(function(t, i) {
    var x = 0.3 + i * 3.2;
    card(s, x, 1.1, 3.0, 3.9, t.bg, t.color);
    s.addText(t.name, { x: x + 0.2, y: 1.2, w: 2.6, h: 0.5, fontSize: 18, fontFace: FJ, color: t.color, bold: true, align: "center", margin: 0 });
    s.addShape(pres.shapes.LINE, { x: x + 0.4, y: 1.75, w: 2.2, h: 0, line: { color: t.color, width: 1 } });
    s.addText("いつ？", { x: x + 0.2, y: 1.9, w: 2.6, h: 0.3, fontSize: 14, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
    s.addText(t.when, { x: x + 0.2, y: 2.2, w: 2.6, h: 0.7, fontSize: 18, fontFace: FJ, color: C.text, align: "center", lineSpacingMultiple: 1.1, margin: 0 });
    s.addText("主な原因", { x: x + 0.2, y: 3.0, w: 2.6, h: 0.3, fontSize: 14, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
    s.addText(t.cause, { x: x + 0.2, y: 3.3, w: 2.6, h: 0.7, fontSize: 17, fontFace: FJ, color: C.text, bold: true, align: "center", lineSpacingMultiple: 1.1, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 5: 安静時振戦 ― パーキンソン病
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "安静時振戦 ― パーキンソン病の特徴", C.red);

  card(s, 0.5, 1.1, 9.0, 1.0, C.lightRed, C.red);
  s.addText("じっとしている時に震え、手を動かすと一時的に止まる", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 22, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0,
  });

  var features = [
    "「丸薬を丸める」ような親指と人差し指の動き（pill-rolling）",
    "片側から始まることが多い",
    "動作が遅くなる・体が硬くなる・歩幅が狭くなる等を伴う",
    "50〜70代に多い",
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
// SLIDE 6: 姿勢時・動作時振戦 ― 本態性振戦
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "姿勢時・動作時振戦 ― 本態性振戦（最多！）");

  card(s, 0.5, 1.1, 9.0, 1.0, C.light, C.primary);
  s.addText("手を伸ばした時・物を持つ時に震える ＝ 振戦で最も多い原因", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  var features = [
    "命に関わらないが、生活の質を下げることがある",
    "家族にも同じ症状がある人が約半数",
    "お酒を飲むと一時的に軽くなるのが特徴",
    "10〜20代で発症することもある",
    "薬（プロプラノロール等）で症状を軽減できる",
  ];

  features.forEach(function(f, i) {
    var y = 2.35 + i * 0.58;
    s.addText("●", { x: 0.7, y: y, w: 0.4, h: 0.5, fontSize: 14, color: C.primary, margin: 0 });
    s.addText(f, { x: 1.1, y: y, w: 8.2, h: 0.5, fontSize: 19, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 7: 比較表 パーキンソン病 vs 本態性振戦
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "パーキンソン病 vs 本態性振戦");

  var rows = [
    ["比べるポイント", "パーキンソン病", "本態性振戦"],
    ["震えるタイミング", "安静時", "動作時・姿勢時"],
    ["震え以外の症状", "あり（動作緩慢等）", "基本的にない"],
    ["頭の震え", "まれ", "よくある"],
    ["家族歴", "まれ", "約50%にあり"],
    ["飲酒の影響", "変わらない", "一時的に軽減"],
    ["食事のとき", "ゆっくりだが可能", "箸が持ちにくい"],
  ];

  var colW = [2.8, 3.1, 3.1];
  var startX = 0.5;

  rows.forEach(function(row, ri) {
    var y = 1.1 + ri * 0.55;
    var isHeader = ri === 0;
    row.forEach(function(cell, ci) {
      var x = startX;
      for (var k = 0; k < ci; k++) { x += colW[k]; }
      var bgColor = isHeader ? C.primary : (ri % 2 === 0 ? C.warmBg : C.white);
      var txtColor = isHeader ? C.white : C.text;
      var isBold = isHeader || ci === 0;
      s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: colW[ci], h: 0.55, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      s.addText(cell, { x: x + 0.1, y: y, w: colW[ci] - 0.2, h: 0.55, fontSize: 15, fontFace: FJ, color: txtColor, bold: isBold, valign: "middle", margin: 0 });
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
  hdr(s, "その他の原因 ― ストレス・薬・甲状腺");

  var causes = [
    { title: "緊張・ストレス", desc: "交感神経の活性化。\n誰にでも起こる生理的反応", icon: "😰", color: C.yellow, bg: C.lightYellow },
    { title: "薬の副作用", desc: "喘息薬・抗うつ薬・\nリチウム・抗てんかん薬等", icon: "💊", color: C.red, bg: C.lightRed },
    { title: "甲状腺の病気", desc: "バセドウ病など。動悸・\n発汗・体重減少を伴う", icon: "🦋", color: C.primary, bg: C.light },
    { title: "カフェイン過剰", desc: "コーヒー等の飲みすぎ。\n減量で改善", icon: "☕", color: C.accent, bg: C.lightYellow },
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
// SLIDE 9: 危険なサイン5つ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "⚠️ こんな震えはすぐ受診 ― 危険なサイン5つ", C.red);

  var signs = [
    "急に片側だけ震え始めた",
    "ろれつが回らない・力が入らないを伴う",
    "日常生活に支障がある（食事・字が書けない）",
    "震え以外の症状が増えている（歩きにくい等）",
    "数週間で明らかに悪化している",
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
    { emoji: "🔴", label: "すぐ受診（救急）", desc: "突然の震え＋片側の脱力・\nろれつ困難・意識の変化", action: "119番で救急搬送", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "震えが持続、生活に支障、\n他の症状もある", action: "脳神経内科を予約", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "緊張時だけ、すぐ治まる、\n生活に支障なし", action: "生活習慣見直し", bg: C.lightGreen, color: C.green },
  ];

  levels.forEach(function(lv, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, lv.bg, lv.color);
    s.addText(lv.emoji + " " + lv.label, { x: 0.8, y: y + 0.05, w: 3.5, h: 0.45, fontSize: 21, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    s.addText(lv.desc, { x: 0.8, y: y + 0.5, w: 5.5, h: 0.55, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText("→ " + lv.action, { x: 6.5, y: y + 0.25, w: 2.8, h: 0.6, fontSize: 17, fontFace: FJ, color: lv.color, bold: true, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 11: セルフチェック
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自分でできるセルフチェック");

  var checks = [
    { title: "コップテスト", desc: "水を入れたコップを持ち、口に運ぶ\n→ 震えるなら姿勢時・動作時振戦", icon: "🥤" },
    { title: "手を伸ばすテスト", desc: "両手を前に10秒間キープ\n→ 指先の震えを確認", icon: "🖐️" },
    { title: "渦巻き描画テスト", desc: "紙に渦巻きを描く\n→ 線のガタつきが振戦の指標に", icon: "🌀" },
  ];

  checks.forEach(function(c, i) {
    var y = 1.15 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, C.primary);
    s.addText(c.icon, { x: 0.8, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, margin: 0 });
    s.addText(c.title, { x: 1.6, y: y + 0.08, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(c.desc, { x: 1.6, y: y + 0.5, w: 7.5, h: 0.5, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });

  s.addText("※ セルフチェックはあくまで目安です。気になる場合は医師にご相談ください。", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 手の震えで覚えておきたい3つ");

  var points = [
    { num: "1", text: "「いつ震えるか」が\n原因を見分けるカギ", sub: "安静時 → パーキンソン病の可能性\n動作時 → 本態性振戦が多い" },
    { num: "2", text: "大半は命に関わらない", sub: "最多は本態性振戦やストレス\nただし放置すると生活の質が低下" },
    { num: "3", text: "迷ったら脳神経内科へ", sub: "振戦の原因を見分ける\nプロフェッショナルです" },
  ];

  points.forEach(function(p, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, C.primary);
    s.addText(p.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.9, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.text, { x: 1.5, y: y + 0.05, w: 3.5, h: 0.55, fontSize: 20, fontFace: FJ, color: C.text, bold: true, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText(p.sub, { x: 5.2, y: y + 0.1, w: 4.0, h: 0.9, fontSize: 15, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.2, margin: 0 });
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
    "🔍 手のふるえセルフチェックツール（概要欄にリンク）",
    "📺 次回：「ろれつが回りにくい」は脳のSOS？",
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
var outPath = __dirname + "/hand_tremor_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
