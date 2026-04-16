var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#16】スマホの使いすぎで手がしびれる？ ― デジタル時代の神経トラブル";

var C = {
  dark: "1B3A5C", primary: "2C5AA0", accent: "E8850C", light: "E8F4FD",
  warmBg: "F8F9FA", white: "FFFFFF", text: "2D3436", sub: "636E72",
  red: "DC3545", green: "28A745", yellow: "F5C518",
  lightRed: "F8D7DA", lightYellow: "FFF3CD", lightGreen: "D4EDDA",
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
  if (borderColor) { s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.12, h: h, fill: { color: borderColor }, rectRadius: 0.05 }); }
}
function ftr(s) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.38, fill: { color: C.dark } });
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #16", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("📱", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("スマホの使いすぎで\n手がしびれる？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.15, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("デジタル時代の神経トラブルを\n脳神経内科医が解説", {
    x: 0.6, y: 2.9, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #16", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.7, w: 8.8, h: 0.4,
    fontSize: 12, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// SLIDE 2: こんな経験ありませんか？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな経験、ありませんか？");
  var items = [
    { emoji: "📱", text: "スマホを長時間使っていると手がしびれる" },
    { emoji: "💻", text: "パソコン作業の後に首から肩にかけて痛い" },
    { emoji: "✋", text: "指先がピリピリ・ジンジンする" },
    { emoji: "🤚", text: "手首を曲げると指にしびれが走る" },
    { emoji: "⚡", text: "首を動かすと腕に電気が走るような感じ" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: スマホ首とは
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "スマホ首とは ― 首にかかる負担");
  // 負荷の段階
  var loads = [
    { angle: "0°", weight: "約5kg", desc: "正面を向いた状態", color: C.green },
    { angle: "15°", weight: "約12kg", desc: "少し下を向く", color: C.yellow },
    { angle: "30°", weight: "約18kg", desc: "スマホを見る角度", color: C.accent },
    { angle: "60°", weight: "約27kg", desc: "深くうつむく", color: C.red },
  ];
  loads.forEach(function(l, i) {
    var x = 0.4 + i * 2.35;
    card(s, x, 1.1, 2.15, 1.5, C.white, l.color);
    s.addText(l.angle, { x: x + 0.1, y: 1.15, w: 1.95, h: 0.4, fontSize: 24, fontFace: FE, color: l.color, bold: true, align: "center", margin: 0 });
    s.addText(l.weight, { x: x + 0.1, y: 1.55, w: 1.95, h: 0.45, fontSize: 22, fontFace: FJ, color: C.text, bold: true, align: "center", margin: 0 });
    s.addText(l.desc, { x: x + 0.1, y: 2.05, w: 1.95, h: 0.4, fontSize: 13, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  });
  // 影響の説明
  card(s, 0.5, 2.85, 9.0, 0.65, C.lightYellow, C.yellow);
  s.addText("⚠️ 60°の前傾 ＝ 小学生1人分の重さが首に！", { x: 0.8, y: 2.9, w: 8.5, h: 0.25, fontSize: 18, fontFace: FJ, color: "856404", bold: true, margin: 0 });
  s.addText("長時間の負荷 → 首の骨・椎間板・神経に影響 → ストレートネックの原因にも", { x: 0.8, y: 3.2, w: 8.5, h: 0.25, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  // ストレートネック
  card(s, 0.5, 3.75, 4.3, 0.9, C.white, C.primary);
  s.addText("正常な頸椎", { x: 0.7, y: 3.8, w: 3.8, h: 0.35, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("緩やかに前方へカーブ\n衝撃を吸収するバネの役割", { x: 0.7, y: 4.15, w: 3.8, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  card(s, 5.2, 3.75, 4.3, 0.9, C.white, C.red);
  s.addText("ストレートネック", { x: 5.4, y: 3.8, w: 3.8, h: 0.35, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("頸椎が真っすぐになった状態\n首・肩の痛み、神経症状の原因に", { x: 5.4, y: 4.15, w: 3.8, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  ftr(s);
})();

// SLIDE 4: 3つの神経トラブル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "スマホ・パソコンで起こる神経トラブル");
  var troubles = [
    { icon: "🦴", name: "頸椎症性神経根症", desc: "首の骨の変形・椎間板突出が神経の根元を圧迫\n首を動かすと腕に電気が走るような痛み・しびれ", color: C.primary },
    { icon: "🤚", name: "手根管症候群", desc: "手首のトンネルで正中神経が圧迫される\n親指〜薬指のしびれ。スマホ握り姿勢が原因に", color: C.accent },
    { icon: "💪", name: "肘部管症候群", desc: "肘の内側で尺骨神経が圧迫される\n小指と薬指のしびれ。肘を曲げた姿勢で悪化", color: "7B2D8E" },
  ];
  troubles.forEach(function(t, i) {
    var y = 1.1 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, t.color);
    s.addText(t.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.9, fontSize: 30, align: "center", valign: "middle", margin: 0 });
    s.addText(t.name, { x: 1.4, y: y + 0.05, w: 3.5, h: 0.4, fontSize: 19, fontFace: FJ, color: t.color, bold: true, margin: 0 });
    s.addText(t.desc, { x: 1.4, y: y + 0.5, w: 7.8, h: 0.55, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 頸椎症性神経根症
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "頸椎症性神経根症 ― 首から腕へのしびれ");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("首の骨の間から出る神経の根元が圧迫される病気 ― 30〜50代に多い", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
  // 症状
  s.addText("主な症状", { x: 0.5, y: 1.8, w: 9.0, h: 0.35, fontSize: 17, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var symptoms = [
    "首を後ろに反らすと腕に痛み・しびれが走る",
    "片側の腕から手にかけてのしびれ",
    "肩甲骨の周囲の痛み",
  ];
  symptoms.forEach(function(sym, i) {
    var y = 2.25 + i * 0.55;
    card(s, 0.5, y, 9.0, 0.45, C.white, C.accent);
    s.addText("• " + sym, { x: 0.8, y: y + 0.03, w: 8.5, h: 0.4, fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  // 治療
  card(s, 0.5, 4.0, 9.0, 0.85, C.lightGreen, C.green);
  s.addText("多くは安静・姿勢改善・薬物療法で改善", { x: 0.8, y: 4.05, w: 8.5, h: 0.35, fontSize: 16, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("手の筋力低下や日常生活に支障がある場合は手術が必要になることも", { x: 0.8, y: 4.4, w: 8.5, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  ftr(s);
})();

// SLIDE 6: 手根管症候群
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "手根管症候群 ― 手首のしびれ");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.accent);
  s.addText("手首の手根管で正中神経が圧迫される ― 女性に多い", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 17, fontFace: FJ, color: C.accent, bold: true, valign: "middle", margin: 0 });
  // 症状
  var symptoms = [
    { icon: "✋", text: "親指・人差し指・中指・薬指のしびれ" },
    { icon: "🌙", text: "夜間・明け方に症状が悪化する" },
    { icon: "🤏", text: "物をつまむ力が弱くなる" },
  ];
  symptoms.forEach(function(sym, i) {
    var y = 1.8 + i * 0.7;
    card(s, 0.5, y, 4.3, 0.6, C.white, C.accent);
    s.addText(sym.icon, { x: 0.7, y: y + 0.05, w: 0.5, h: 0.5, fontSize: 22, align: "center", margin: 0 });
    s.addText(sym.text, { x: 1.3, y: y + 0.05, w: 3.3, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  // 原因と対策
  card(s, 5.2, 1.8, 4.3, 1.0, C.white, C.primary);
  s.addText("原因", { x: 5.4, y: 1.85, w: 3.8, h: 0.3, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("スマホを握る姿勢\nマウス操作で手首が圧迫\n妊娠中・更年期に発症しやすい", { x: 5.4, y: 2.15, w: 3.8, h: 0.55, fontSize: 13, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  card(s, 5.2, 3.0, 4.3, 0.7, C.lightGreen, C.green);
  s.addText("軽症 → 手首の安静・サポーター\n進行 → 手術が必要な場合も", { x: 5.4, y: 3.05, w: 3.8, h: 0.6, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 7: 肘部管症候群と胸郭出口症候群
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "肘部管症候群と胸郭出口症候群");
  // 肘部管
  card(s, 0.5, 1.1, 9.0, 1.6, C.white, "7B2D8E");
  s.addText("💪 肘部管症候群", { x: 0.7, y: 1.15, w: 8.5, h: 0.4, fontSize: 19, fontFace: FJ, color: "7B2D8E", bold: true, margin: 0 });
  s.addText("肘の内側で尺骨神経が圧迫される", { x: 0.7, y: 1.55, w: 4.5, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("小指と薬指の小指側がしびれる", { x: 0.7, y: 1.85, w: 4.5, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("悪化する姿勢：スマホを耳に当てる\n肘をついてスマホを見る", { x: 5.5, y: 1.55, w: 3.8, h: 0.5, fontSize: 14, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.15, margin: 0 });
  s.addText("⚠️ 進行すると手の筋肉がやせて箸が使いにくくなることも", { x: 0.7, y: 2.2, w: 8.5, h: 0.3, fontSize: 14, fontFace: FJ, color: C.red, margin: 0 });
  // 胸郭出口
  card(s, 0.5, 2.95, 9.0, 1.6, C.white, C.primary);
  s.addText("🫁 胸郭出口症候群", { x: 0.7, y: 3.0, w: 8.5, h: 0.4, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("首から腕に向かう神経・血管が鎖骨の周囲で圧迫される", { x: 0.7, y: 3.4, w: 4.5, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("腕を上げると手がしびれる・冷たくなる", { x: 0.7, y: 3.7, w: 4.5, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("なりやすい方：なで肩\nデスクワークで猫背の方", { x: 5.5, y: 3.4, w: 3.8, h: 0.5, fontSize: 14, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.15, margin: 0 });
  ftr(s);
})();

// SLIDE 8: 受診先の選び方
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "整形外科？脳神経内科？ ― 受診先の選び方");
  // 整形外科
  card(s, 0.5, 1.1, 4.3, 2.5, C.white, C.accent);
  s.addText("🏥 整形外科", { x: 0.7, y: 1.15, w: 3.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var ortho = [
    "首を動かすと痛み・しびれが出る",
    "手首や肘の特定の動作で症状が出る",
    "首や手首に原因がありそう",
  ];
  ortho.forEach(function(item, i) {
    var y = 1.65 + i * 0.6;
    s.addText("• " + item, { x: 0.7, y: y, w: 3.8, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  // 脳神経内科
  card(s, 5.2, 1.1, 4.3, 2.5, C.white, C.primary);
  s.addText("🧠 脳神経内科", { x: 5.4, y: 1.15, w: 3.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var neuro = [
    "しびれの原因がはっきりしない",
    "両手両足など広い範囲にしびれ",
    "ろれつ困難・歩行障害を伴う",
  ];
  neuro.forEach(function(item, i) {
    var y = 1.65 + i * 0.6;
    s.addText("• " + item, { x: 5.4, y: y, w: 3.8, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  // まとめ
  card(s, 0.5, 3.85, 9.0, 0.6, C.light, C.primary);
  s.addText("迷ったらまずどちらかを受診 → 適切な科に紹介してもらえます", { x: 0.8, y: 3.9, w: 8.5, h: 0.5, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", align: "center", margin: 0 });
  ftr(s);
})();

// SLIDE 9: 危険なサイン
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "危険なサイン ― 単なるスマホ首ではない", C.red);
  var signs = [
    { text: "突然の片側の手足のしびれ＋ろれつ困難\n（脳卒中の可能性 → 119番）", icon: "🚑" },
    { text: "両手がしびれてボタンがかけにくい・歩行ふらつき\n（頸椎症性脊髄症の可能性）", icon: "🚶" },
    { text: "手の筋肉がやせてきた\n（進行した神経圧迫・運動ニューロン疾患）", icon: "✋" },
    { text: "しびれが日に日に広がっている\n（末梢神経障害・脊髄の病気の精査が必要）", icon: "📈" },
  ];
  signs.forEach(function(sign, i) {
    var y = 1.1 + i * 0.95;
    card(s, 0.5, y, 9.0, 0.8, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.1, w: 0.6, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.7, fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 10: 受診の目安
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "しびれ＋ろれつ困難・顔面麻痺\nしびれ＋歩行障害", action: "脳卒中の可能性\n→119番", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "しびれが2週間以上続く\n手の筋力低下・ボタンがかけにくい\n首を動かすたびに腕に強い痛み", action: "整形外科・\n脳神経内科を受診", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "スマホ・PC使用後だけ一時的にしびれる\n姿勢を変えると改善する", action: "使用時間見直し\nストレッチから", bg: C.lightGreen, color: C.green },
  ];
  levels.forEach(function(lv, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, lv.bg, lv.color);
    s.addText(lv.emoji + " " + lv.label, { x: 0.8, y: y + 0.05, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    s.addText(lv.desc, { x: 0.8, y: y + 0.5, w: 5.0, h: 0.55, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText(lv.action, { x: 6.0, y: y + 0.15, w: 3.3, h: 0.8, fontSize: 16, fontFace: FJ, color: lv.color, bold: true, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 11: 予防と対策
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "予防と対策 ― デジタルデバイスとの付き合い方");
  var tips = [
    { num: "1", title: "30分に1回は休憩", desc: "連続使用は30分を目安に\n首や手首のストレッチを行う", icon: "⏰" },
    { num: "2", title: "スマホの位置を目の高さに", desc: "下を向く角度が小さいほど\n首への負担は軽減される", icon: "📱" },
    { num: "3", title: "正しい姿勢でデスクワーク", desc: "画面は目の高さ・肘は90°\n足は床につく姿勢が理想的", icon: "🪑" },
    { num: "4", title: "首と手首のストレッチ", desc: "首をゆっくり左右に傾ける\n手首を回す・指を開閉する", icon: "🤸" },
  ];
  tips.forEach(function(t, i) {
    var y = 1.05 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, C.primary);
    s.addText(t.num, { x: 0.7, y: y + 0.05, w: 0.5, h: 0.75, fontSize: 30, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.icon, { x: 1.3, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 26, align: "center", margin: 0 });
    s.addText(t.title, { x: 2.0, y: y + 0.05, w: 3.0, h: 0.35, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(t.desc, { x: 2.0, y: y + 0.42, w: 7.3, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 12: まとめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― デジタル時代の神経を守る");
  var points = [
    { num: "1", text: "スマホ首は首に\n最大27kgの負荷", sub: "頸椎症性神経根症・手根管症候群\n肘部管症候群などの原因に" },
    { num: "2", text: "30分に1回の休憩と\nストレッチが最大の予防", sub: "スマホの位置を高く\n正しい姿勢を心がける" },
    { num: "3", text: "2週間以上のしびれや\n筋力低下は受診を", sub: "整形外科・脳神経内科で\n原因を特定し適切な治療へ" },
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

// SLIDE 13: エンディング
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("ご視聴ありがとうございました", {
    x: 0.6, y: 0.6, w: 8.8, h: 0.7,
    fontSize: 30, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 1.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("チャンネル登録・高評価よろしくお願いします！", {
    x: 0.6, y: 1.7, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("🔗 ブログ記事で詳しく読む（概要欄にリンク）", {
    x: 1.5, y: 2.5, w: 7.0, h: 0.4, fontSize: 16, fontFace: FJ, color: "B0BEC5", margin: 0,
  });
  s.addText("次回予告", {
    x: 0.6, y: 3.3, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: C.accent, bold: true, align: "center", margin: 0,
  });
  s.addText("#17 その首の痛み、実は神経が原因？ ― 整形外科と脳神経内科の境界線", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 17, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/smartphone_nerve_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
