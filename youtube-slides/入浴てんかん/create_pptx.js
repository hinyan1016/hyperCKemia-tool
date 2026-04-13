var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "お風呂で決まって意識を失う — 入浴てんかんの正体と安全対策";

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
  s.addText("医知創造ラボ ｜ 入浴てんかん", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("🛁", { x: 0.6, y: 0.3, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("お風呂で決まって\n意識を失う", {
    x: 0.6, y: 1.2, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.8, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("入浴てんかんの正体と安全対策", {
    x: 0.6, y: 3.0, w: 8.8, h: 0.6,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("脳神経内科専門医が解説", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
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
    { emoji: "🛁", text: "お風呂に入ると毎回意識が飛ぶ" },
    { emoji: "⚡", text: "入浴中にけいれんを起こして倒れた" },
    { emoji: "💧", text: "湯船で意識がなくなり溺れかけた" },
    { emoji: "🤔", text: "ヒートショック？失神？原因が分からない" },
  ];

  items.forEach(function(item, i) {
    var y = 1.2 + i * 0.95;
    card(s, 0.6, y, 8.8, 0.8, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.1, w: 0.7, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.1, w: 7.5, h: 0.6, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 3: 入浴てんかんとは？
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "入浴てんかんとは？");

  card(s, 0.5, 1.1, 9.0, 1.2, C.light, C.primary);
  s.addText("反射てんかんの一種", { x: 0.8, y: 1.15, w: 4, h: 0.45, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("入浴という特定の状況でのみ\n発作が繰り返し誘発されるてんかん", {
    x: 0.8, y: 1.55, w: 8.5, h: 0.65, fontSize: 20, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  var points = [
    { icon: "🔑", label: "最大の特徴", desc: "入浴時にしか発作が起きない" },
    { icon: "📊", label: "入浴外発作", desc: "経過中62%に入浴外の自然発作も出現" },
    { icon: "🏆", label: "反射てんかんで最多", desc: "反射てんかんの中で最も頻度が高い型の一つ" },
  ];

  points.forEach(function(p, i) {
    var y = 2.6 + i * 0.85;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.accent);
    s.addText(p.icon, { x: 0.75, y: y + 0.05, w: 0.6, h: 0.6, fontSize: 24, align: "center", margin: 0 });
    s.addText(p.label, { x: 1.4, y: y + 0.05, w: 2.2, h: 0.6, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(p.desc, { x: 3.6, y: y + 0.05, w: 5.7, h: 0.6, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: 2つのタイプ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "入浴てんかんの2つのタイプ");

  // 左カード：HWE
  card(s, 0.4, 1.1, 4.4, 3.8, C.white, C.red);
  s.addText("🔥 熱水てんかん (HWE)", { x: 0.65, y: 1.2, w: 4.0, h: 0.5, fontSize: 20, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("Hot Water Epilepsy", { x: 0.65, y: 1.65, w: 4.0, h: 0.35, fontSize: 14, fontFace: FE, color: C.sub, margin: 0 });
  s.addText(
    "• 温水（37〜50℃）で誘発\n• 頭部への湯かけが主な誘因\n• 小児後期〜成人に多い\n• 男女比 約3:1（男性優位）\n• インド南部で多い報告",
    { x: 0.65, y: 2.1, w: 4.0, h: 2.6, fontSize: 17, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.4, margin: 0 }
  );

  // 右カード：BE
  card(s, 5.2, 1.1, 4.4, 3.8, C.white, C.primary);
  s.addText("💧 浸水型入浴てんかん (BE)", { x: 5.45, y: 1.2, w: 4.0, h: 0.5, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("Bathing Epilepsy", { x: 5.45, y: 1.65, w: 4.0, h: 0.35, fontSize: 14, fontFace: FE, color: C.sub, margin: 0 });
  s.addText(
    "• 水温に関係なく浸水で誘発\n• 体温に近い温度でも発症\n• 乳幼児に多い（平均15ヶ月）\n• SYN1遺伝子変異（X連鎖性）\n• 経過は良性、自然消失多い",
    { x: 5.45, y: 2.1, w: 4.0, h: 2.6, fontSize: 17, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.4, margin: 0 }
  );

  ftr(s);
})();

// ============================================================
// SLIDE 5: なぜ入浴で発作が起きるのか
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "なぜ入浴で発作が起きるのか？");

  var mechanisms = [
    { icon: "🌡️", title: "体温上昇", desc: "入浴→急速な体温上昇→脳の興奮性閾値が低下\n動物実験：直腸温41.5℃で発作出現", color: C.red },
    { icon: "🧠", title: "異常な温度調節", desc: "温水が頭皮に接触→側頭葉（海馬周辺）を異常刺激\n脳波で側頭部に異常波を認める症例が多い", color: C.primary },
    { icon: "🧬", title: "遺伝的素因", desc: "SYN1遺伝子変異（X連鎖性）：シナプス小胞制御異常\n浸水型入浴てんかんで同定", color: C.green },
  ];

  mechanisms.forEach(function(m, i) {
    var y = 1.15 + i * 1.35;
    card(s, 0.5, y, 9.0, 1.2, C.white, m.color);
    s.addText(m.icon, { x: 0.75, y: y + 0.1, w: 0.7, h: 1.0, fontSize: 30, align: "center", valign: "middle", margin: 0 });
    s.addText(m.title, { x: 1.5, y: y + 0.05, w: 2.5, h: 0.45, fontSize: 20, fontFace: FJ, color: m.color, bold: true, valign: "middle", margin: 0 });
    s.addText(m.desc, { x: 1.5, y: y + 0.5, w: 7.8, h: 0.65, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.3, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 6: どんな発作が起きる？
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "入浴てんかんの症状");

  // 発作型テーブル
  var headers = ["発作型", "頻度", "特徴"];
  headers.forEach(function(h, i) {
    var x = 0.5 + [0, 2.8, 4.8][i];
    var w = [2.8, 2.0, 4.2][i];
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.1, w: w, h: 0.5, fill: { color: C.primary } });
    s.addText(h, { x: x, y: 1.1, w: w, h: 0.5, fontSize: 16, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  var rows = [
    ["焦点性発作（意識減損）", "最も多い", "ぼーっとする、口もぐもぐ"],
    ["二次性全般化", "約38%", "全身けいれんに移行"],
    ["全般性発作のみ", "まれ", "最初から全身けいれん"],
  ];

  rows.forEach(function(row, ri) {
    var y = 1.6 + ri * 0.55;
    var bg = ri % 2 === 0 ? C.warmBg : C.white;
    row.forEach(function(cell, ci) {
      var x = 0.5 + [0, 2.8, 4.8][ci];
      var w = [2.8, 2.0, 4.2][ci];
      s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: 0.55, fill: { color: bg } });
      s.addText(cell, { x: x + 0.1, y: y, w: w - 0.2, h: 0.55, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    });
  });

  // 症状の流れ
  card(s, 0.5, 3.4, 9.0, 1.7, C.light, C.accent);
  s.addText("発作の典型的な流れ", { x: 0.8, y: 3.45, w: 4, h: 0.4, fontSize: 18, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText(
    "① 入浴後 数分〜15分 で前兆（不安感、胃からこみ上げる感覚）\n" +
    "② 意識がぼんやり → 反応がなくなる（焦点性発作）\n" +
    "③ 約4割で全身けいれん（強直間代発作）に移行\n" +
    "④ 発作後はもうろう状態が続く（記憶なし）",
    { x: 0.8, y: 3.85, w: 8.5, h: 1.15, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.3, margin: 0 }
  );

  ftr(s);
})();

// ============================================================
// SLIDE 7: ヒートショック・失神との見分け方
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "ヒートショック・失神との見分け方");

  var cols = ["鑑別ポイント", "入浴てんかん", "ヒートショック", "迷走神経反射"];
  cols.forEach(function(c, i) {
    var x = 0.3 + [0, 2.0, 4.5, 7.0][i];
    var w = [2.0, 2.5, 2.5, 2.7][i];
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.05, w: w, h: 0.5, fill: { color: C.primary } });
    s.addText(c, { x: x, y: 1.05, w: w, h: 0.5, fontSize: 14, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  var data = [
    ["再現性", "毎回ほぼ確実", "寒暖差で時々", "長湯で時々"],
    ["けいれん", "あり（多い）", "通常なし", "通常なし"],
    ["好発年齢", "小児〜若年成人", "高齢者", "全年齢"],
    ["意識回復", "もうろう状態続く", "比較的速い", "横で速やかに回復"],
    ["自動症", "口もぐもぐ等あり", "なし", "なし"],
    ["脳波異常", "40%に異常", "なし", "なし"],
  ];

  data.forEach(function(row, ri) {
    var y = 1.55 + ri * 0.55;
    var bg = ri % 2 === 0 ? C.warmBg : C.white;
    row.forEach(function(cell, ci) {
      var x = 0.3 + [0, 2.0, 4.5, 7.0][ci];
      var w = [2.0, 2.5, 2.5, 2.7][ci];
      s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: 0.55, fill: { color: bg } });
      var fontColor = ci === 1 ? C.primary : C.text;
      var isBold = ci === 1;
      s.addText(cell, { x: x + 0.08, y: y, w: w - 0.16, h: 0.55, fontSize: 13, fontFace: FJ, color: fontColor, bold: isBold, valign: "middle", margin: 0 });
    });
  });

  card(s, 0.5, 4.9, 9.0, 0.3, C.lightRed, C.red);
  s.addText("⚠ 「毎回」＋「けいれん」＋「もうろう状態」→ 入浴てんかんを疑う", {
    x: 0.8, y: 4.9, w: 8.5, h: 0.3, fontSize: 15, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 8: 診断
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "診断 — どんな検査で分かる？");

  card(s, 0.5, 1.1, 9.0, 1.0, C.light, C.primary);
  s.addText("🔑 最も重要なのは「詳細な問診」", { x: 0.8, y: 1.15, w: 5, h: 0.4, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("発作と入浴の関連、湯温、発作の様子、熱性けいれんの既往、家族歴", {
    x: 0.8, y: 1.55, w: 8.5, h: 0.45, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  var tests = [
    { name: "脳波検査（EEG）", purpose: "てんかん性異常波の検出", finding: "約40%に側頭部異常波", color: C.primary },
    { name: "頭部MRI", purpose: "器質的脳病変の除外", finding: "多くの症例で正常", color: C.green },
    { name: "ビデオ脳波モニタリング", purpose: "入浴中の発作を記録", finding: "誘発試験で発作記録可能", color: C.accent },
  ];

  tests.forEach(function(t, i) {
    var y = 2.35 + i * 0.9;
    card(s, 0.5, y, 9.0, 0.75, C.white, t.color);
    s.addText(t.name, { x: 0.75, y: y + 0.05, w: 3.0, h: 0.65, fontSize: 18, fontFace: FJ, color: t.color, bold: true, valign: "middle", margin: 0 });
    s.addText(t.purpose, { x: 3.8, y: y + 0.05, w: 2.8, h: 0.65, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText(t.finding, { x: 6.6, y: y + 0.05, w: 2.7, h: 0.65, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  });

  card(s, 0.5, 5.0, 9.0, 0.2, C.lightYellow);
  s.addText("💡 脳波正常でもてんかんは否定できない → 発作時のスマホ動画が重要", {
    x: 0.5, y: 4.65, w: 9.0, h: 0.5, fontSize: 15, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 9: 治療
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "治療 — まず湯温を下げる");

  // 非薬物療法
  card(s, 0.5, 1.1, 4.3, 3.0, C.white, C.green);
  s.addText("✅ 非薬物療法（第一歩）", { x: 0.75, y: 1.15, w: 3.8, h: 0.45, fontSize: 18, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText(
    "• 湯温を 37〜38℃以下 に設定\n  → これだけで消失する症例あり\n• 入浴時間を 10分以内 に\n• シャワー浴に切り替え\n• 頭から湯をかぶらない",
    { x: 0.75, y: 1.65, w: 3.8, h: 2.3, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.4, margin: 0 }
  );

  // 薬物療法
  card(s, 5.2, 1.1, 4.3, 3.0, C.white, C.primary);
  s.addText("💊 薬物療法", { x: 5.45, y: 1.15, w: 3.8, h: 0.45, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText(
    "• カルバマゼピン（CBZ）\n  焦点性発作に有効\n• バルプロ酸（VPA）\n  全般性発作を伴う場合\n• クロバザム（CLB）\n  入浴前1〜2時間の頓用",
    { x: 5.45, y: 1.65, w: 3.8, h: 2.3, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.4, margin: 0 }
  );

  // 予後
  card(s, 0.5, 4.3, 9.0, 0.8, C.lightGreen, C.green);
  s.addText("📈 予後は一般的に良好", { x: 0.8, y: 4.35, w: 3.5, h: 0.35, fontSize: 18, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("15年間の長期フォローアップ研究で、多くの患者が成長とともに発作を卒業（自然寛解）", {
    x: 0.8, y: 4.7, w: 8.5, h: 0.35, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 10: 7つの安全対策
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "入浴中の発作を防ぐ — 7つの安全対策", C.red);

  var items = [
    { num: "1", text: "湯温は38℃以下に設定する", icon: "🌡️" },
    { num: "2", text: "入浴時間は10分以内にする", icon: "⏱️" },
    { num: "3", text: "浴槽の湯量を減らす（半身浴程度）", icon: "💧" },
    { num: "4", text: "入浴前に家族に声をかける", icon: "🗣️" },
    { num: "5", text: "浴室の鍵をかけない", icon: "🔓" },
    { num: "6", text: "シャワー浴を検討（椅子使用）", icon: "🚿" },
    { num: "7", text: "不安定期は入浴を控える", icon: "⚠️" },
  ];

  items.forEach(function(item, i) {
    var y = 1.05 + i * 0.58;
    card(s, 0.5, y, 9.0, 0.5, C.white, C.red);
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: y + 0.05, w: 0.4, h: 0.4, fill: { color: C.red } });
    s.addText(item.num, { x: 0.7, y: y + 0.05, w: 0.4, h: 0.4, fontSize: 16, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(item.icon + " " + item.text, { x: 1.3, y: y + 0.02, w: 8.0, h: 0.46, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 11: 緊急対応
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "入浴中に発作が起きたら — 緊急対応", C.red);

  var steps = [
    { num: "STEP 1", action: "頭を水面の上に持ち上げる", desc: "まず気道を確保。顔が水に浸からないように", color: C.red },
    { num: "STEP 2", action: "浴槽の栓を抜いて排水する", desc: "本人を持ち上げるのが難しければ、先に湯を抜く", color: C.accent },
    { num: "STEP 3", action: "安全な場所に移す", desc: "意識回復後、ゆっくり浴槽から出す。無理は禁物", color: C.green },
  ];

  steps.forEach(function(step, i) {
    var y = 1.15 + i * 1.1;
    card(s, 0.5, y, 9.0, 0.95, C.white, step.color);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.75, y: y + 0.15, w: 1.2, h: 0.4, fill: { color: step.color }, rectRadius: 0.05 });
    s.addText(step.num, { x: 0.75, y: y + 0.15, w: 1.2, h: 0.4, fontSize: 14, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(step.action, { x: 2.1, y: y + 0.08, w: 7.2, h: 0.4, fontSize: 20, fontFace: FJ, color: step.color, bold: true, valign: "middle", margin: 0 });
    s.addText(step.desc, { x: 2.1, y: y + 0.5, w: 7.2, h: 0.4, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  });

  card(s, 0.5, 4.5, 9.0, 0.6, C.lightRed, C.red);
  s.addText("🚑 119番を呼ぶ：発作5分以上 / 水を飲んだ / 意識戻らず / 初めての発作 / 怪我あり", {
    x: 0.8, y: 4.52, w: 8.5, h: 0.55, fontSize: 15, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ");

  var points = [
    "入浴てんかん = 入浴時に限って発作が起きる反射てんかん",
    "2タイプ：熱水てんかん（温度）と浸水型（浸水行為）",
    "ヒートショック・失神との鑑別が重要",
    "治療の第一歩は湯温を38℃以下に下げること",
    "予後は一般的に良好 — 多くの患者が発作を卒業",
    "溺水事故の予防が最優先 — 7つの安全対策を実践",
  ];

  points.forEach(function(p, i) {
    var y = 1.1 + i * 0.65;
    card(s, 0.5, y, 9.0, 0.55, C.white, C.primary);
    s.addText("✓", { x: 0.65, y: y + 0.02, w: 0.5, h: 0.5, fontSize: 20, fontFace: FE, color: C.green, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p, { x: 1.2, y: y + 0.02, w: 8.1, h: 0.5, fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.5, 5.05, 7.0, 0.15, C.accent);

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
    fontSize: 28, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 2.5, y: 1.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });

  s.addText(
    "入浴中に繰り返し意識を失う場合は\n脳神経内科を受診してください",
    {
      x: 0.6, y: 2.0, w: 8.8, h: 0.9,
      fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.3, margin: 0,
    }
  );

  s.addText(
    "チャンネル登録・高評価をお願いします！\n🔔 通知ONで最新情報をお届け",
    {
      x: 0.6, y: 3.2, w: 8.8, h: 0.8,
      fontSize: 18, fontFace: FJ, color: "90A4AE", align: "center", lineSpacingMultiple: 1.3, margin: 0,
    }
  );

  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.4, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.9, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/bathing_epilepsy_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
