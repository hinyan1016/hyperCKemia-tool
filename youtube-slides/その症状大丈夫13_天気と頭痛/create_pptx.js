var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#13】雨の日に頭が痛くなる科学的理由";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #13", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🌧️", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("雨の日に頭が痛くなる\n科学的理由", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.15, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("気象病と片頭痛の関係を\n脳神経内科医が解説", {
    x: 0.6, y: 2.9, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #13", {
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
    { emoji: "🌧️", text: "雨が降る前に頭が痛くなる" },
    { emoji: "📉", text: "天気予報を見ると頭痛が予測できる" },
    { emoji: "🌀", text: "台風が近づくと体調が悪くなる" },
    { emoji: "🤷", text: "「気のせい」と言われてつらい" },
    { emoji: "💊", text: "痛み止めが手放せない" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 気象病のメカニズム
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "なぜ天気で頭痛が起こるか ― 科学的メカニズム");
  var steps = [
    { icon: "📉", name: "気圧が低下", desc: "低気圧が近づくと大気圧が下がる", color: C.primary },
    { icon: "👂", name: "内耳が感知", desc: "内耳のセンサーが気圧変化をキャッチ\n脳に「環境変化」のシグナルを送る", color: C.accent },
    { icon: "🧠", name: "自律神経が乱れる", desc: "交感神経が過剰に活性化\n血管の収縮・拡張のバランスが崩れる", color: "7B2D8E" },
    { icon: "💥", name: "頭痛が発生", desc: "脳の血管が拡張 → 周囲の神経を刺激\n炎症物質が放出されて痛みが起こる", color: C.red },
  ];
  steps.forEach(function(st, i) {
    var y = 1.1 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, st.color);
    s.addText(st.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 28, align: "center", margin: 0 });
    s.addText(st.name, { x: 1.4, y: y + 0.05, w: 2.5, h: 0.4, fontSize: 19, fontFace: FJ, color: st.color, bold: true, margin: 0 });
    s.addText(st.desc, { x: 4.0, y: y + 0.05, w: 5.3, h: 0.75, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 4: 片頭痛の特徴
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "天気で悪化しやすい頭痛 ― 片頭痛の特徴");
  var features = [
    { icon: "💓", name: "ズキンズキンと脈打つ痛み", desc: "心臓の拍動に合わせた痛み。片側が多いが両側のことも", color: C.red },
    { icon: "🔆", name: "光・音・においに敏感になる", desc: "明るい場所や騒がしい場所がつらい。暗い静かな部屋で休みたい", color: C.accent },
    { icon: "🤢", name: "吐き気・嘔吐を伴う", desc: "頭痛と同時に気持ちが悪くなる。食事がとれないことも", color: C.primary },
    { icon: "⏰", name: "4〜72時間続く", desc: "頭痛の発作は数時間〜3日間。日常生活に支障が出る", color: "7B2D8E" },
  ];
  features.forEach(function(f, i) {
    var y = 1.1 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, f.color);
    s.addText(f.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 28, align: "center", margin: 0 });
    s.addText(f.name, { x: 1.4, y: y + 0.05, w: 3.5, h: 0.4, fontSize: 18, fontFace: FJ, color: f.color, bold: true, margin: 0 });
    s.addText(f.desc, { x: 1.4, y: y + 0.45, w: 7.8, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 片頭痛 vs 緊張型頭痛
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "片頭痛と緊張型頭痛の見分け方");
  var rows = [
    { label: "痛みの質", migraine: "ズキンズキン（拍動性）", tension: "締めつけられる感じ" },
    { label: "場所", migraine: "片側が多い", tension: "両側・頭全体" },
    { label: "動くと", migraine: "悪化する（寝込む）", tension: "変わらない・軽減" },
    { label: "光・音", migraine: "敏感になる", tension: "通常影響なし" },
    { label: "吐き気", migraine: "伴うことが多い", tension: "通常なし" },
    { label: "天気の影響", migraine: "悪化しやすい", tension: "影響は少ない" },
  ];
  // ヘッダー
  card(s, 0.5, 1.05, 9.0, 0.5, C.primary);
  s.addText("項目", { x: 0.7, y: 1.08, w: 2.0, h: 0.45, fontSize: 16, fontFace: FJ, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("片頭痛", { x: 2.8, y: 1.08, w: 3.2, h: 0.45, fontSize: 16, fontFace: FJ, color: C.white, bold: true, valign: "middle", align: "center", margin: 0 });
  s.addText("緊張型頭痛", { x: 6.1, y: 1.08, w: 3.2, h: 0.45, fontSize: 16, fontFace: FJ, color: C.white, bold: true, valign: "middle", align: "center", margin: 0 });
  rows.forEach(function(r, i) {
    var y = 1.6 + i * 0.58;
    var bg = i % 2 === 0 ? C.light : C.white;
    card(s, 0.5, y, 9.0, 0.5, bg);
    s.addText(r.label, { x: 0.7, y: y + 0.05, w: 2.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(r.migraine, { x: 2.8, y: y + 0.05, w: 3.2, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", align: "center", margin: 0 });
    s.addText(r.tension, { x: 6.1, y: y + 0.05, w: 3.2, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", align: "center", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 6: 片頭痛の治療
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "片頭痛の治療 ― 発作時と予防");
  // 発作時
  s.addText("発作時の治療", { x: 0.5, y: 1.05, w: 4.3, h: 0.35, fontSize: 17, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var acute = [
    { name: "トリプタン", desc: "片頭痛の特効薬。痛みが出たら早めに服用" },
    { name: "NSAIDs", desc: "市販の鎮痛薬。軽い発作に。月10日以上は要注意" },
  ];
  acute.forEach(function(a, i) {
    var y = 1.5 + i * 0.7;
    card(s, 0.5, y, 4.3, 0.6, C.white, C.red);
    s.addText(a.name, { x: 0.7, y: y + 0.05, w: 1.8, h: 0.25, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
    s.addText(a.desc, { x: 0.7, y: y + 0.3, w: 3.8, h: 0.25, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 予防
  s.addText("予防の治療", { x: 5.2, y: 1.05, w: 4.3, h: 0.35, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var prev = [
    { name: "予防薬", desc: "月4回以上の発作に。毎日服用して頻度を減らす" },
    { name: "CGRP関連抗体", desc: "注射の新薬。月1回で片頭痛日数を大幅に減少" },
  ];
  prev.forEach(function(p, i) {
    var y = 1.5 + i * 0.7;
    card(s, 5.2, y, 4.3, 0.6, C.white, C.primary);
    s.addText(p.name, { x: 5.4, y: y + 0.05, w: 2.0, h: 0.25, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(p.desc, { x: 5.4, y: y + 0.3, w: 3.8, h: 0.25, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 注意
  card(s, 0.5, 3.1, 9.0, 0.8, C.lightYellow, C.yellow);
  s.addText("⚠️ 薬物乱用頭痛に注意", { x: 0.8, y: 3.15, w: 8.5, h: 0.35, fontSize: 17, fontFace: FJ, color: "856404", bold: true, margin: 0 });
  s.addText("鎮痛薬を月10日以上使用すると、かえって頭痛が悪化する「薬物乱用頭痛」になることがあります。\n頻繁に薬が必要な方は脳神経内科・頭痛外来を受診してください。", {
    x: 0.8, y: 3.5, w: 8.5, h: 0.35, fontSize: 13, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });
  // セルフケア
  s.addText("セルフケア", { x: 0.5, y: 4.15, w: 9.0, h: 0.3, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var self = ["頭痛ダイアリーをつける", "規則正しい睡眠", "気圧予報アプリの活用", "カフェインの適度な摂取"];
  self.forEach(function(item, i) {
    var x = 0.5 + i * 2.3;
    card(s, x, 4.5, 2.1, 0.55, C.white, C.accent);
    s.addText(item, { x: x + 0.15, y: 4.53, w: 1.85, h: 0.5, fontSize: 13, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 7: 頭痛ダイアリー
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "頭痛ダイアリー ― 記録が治療の第一歩");
  var records = [
    { num: "1", title: "いつ・どのくらい続いたか", desc: "日時・持続時間を記録\n天気（気圧）もメモすると傾向が見える", icon: "📅" },
    { num: "2", title: "どんな痛みか", desc: "ズキンズキン？締めつけ？強さは10段階で\n場所（片側・両側・後頭部）も記録", icon: "📝" },
    { num: "3", title: "何をしたか・何が効いたか", desc: "服用した薬と効果・寝込んだか・トリガー\n（食事・睡眠・ストレス・月経周期）", icon: "💊" },
  ];
  records.forEach(function(r, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, C.primary);
    s.addText(r.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.85, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(r.icon, { x: 1.5, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addText(r.title, { x: 2.3, y: y + 0.1, w: 3.5, h: 0.45, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(r.desc, { x: 2.3, y: y + 0.55, w: 6.5, h: 0.45, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
  });
  s.addText("※ 1〜2ヶ月分の記録があると、医師が的確に診断できます", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });
  ftr(s);
})();

// SLIDE 8: 危険な頭痛の見分け方
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "気象病ではない ― 危険な頭痛のサイン", C.red);
  var signs = [
    { text: "人生最悪の突然の激しい頭痛（くも膜下出血）", icon: "💥" },
    { text: "発熱＋頭痛＋首の硬さ（髄膜炎）", icon: "🤒" },
    { text: "日に日に悪化する頭痛（脳腫瘍・慢性硬膜下血腫）", icon: "📈" },
    { text: "50歳以降に初めて起きた頭痛", icon: "⚠️" },
    { text: "視力低下や意識障害を伴う頭痛", icon: "👁️" },
  ];
  signs.forEach(function(sign, i) {
    var y = 1.15 + i * 0.78;
    card(s, 0.5, y, 9.0, 0.65, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.05, w: 0.6, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.55, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  card(s, 1.0, 5.0, 8.0, 0.2, C.red);
  ftr(s);
})();

// SLIDE 9: 受診の目安
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "突然の激しい頭痛（人生最悪）\n発熱＋頭痛＋首の硬さ\n意識障害・けいれんを伴う", action: "くも膜下出血・髄膜炎\n→119番", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "月4回以上の頭痛がある\n鎮痛薬を月10日以上使用している\n日常生活に支障が出ている", action: "脳神経内科・頭痛外来\n→ 予防治療の相談", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "天気の変わり目の軽い頭痛\n市販薬で数時間で改善\n月に数回程度", action: "頭痛ダイアリーで\n記録を開始", bg: C.lightGreen, color: C.green },
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

// SLIDE 10: セルフケア
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "天気頭痛のセルフケア");
  var tips = [
    { num: "1", title: "気圧予報アプリを活用", desc: "気圧の低下を事前に把握 → 予兆段階で薬を服用\n「頭痛ーる」などのアプリが便利", icon: "📱" },
    { num: "2", title: "規則正しい生活リズム", desc: "睡眠の過不足・空腹・脱水は片頭痛のトリガー\n週末の寝だめも避ける", icon: "🛏️" },
    { num: "3", title: "耳のマッサージ", desc: "耳を引っ張る・回す → 内耳の血流を改善\n気圧変化による不調の緩和に効果が期待", icon: "👂" },
  ];
  tips.forEach(function(t, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, C.primary);
    s.addText(t.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.85, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.icon, { x: 1.5, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addText(t.title, { x: 2.3, y: y + 0.1, w: 3.5, h: 0.45, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.desc, { x: 2.3, y: y + 0.55, w: 6.5, h: 0.45, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 11: Q&A
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある質問 Q&A");
  var qa = [
    { q: "気象病は正式な病名ですか？", a: "「気象病」は正式な医学用語ではありませんが、気圧変化が片頭痛やめまいを悪化させることは医学的に認められています。近年は「気象関連頭痛」として研究が進んでいます" },
    { q: "片頭痛は遺伝しますか？", a: "片頭痛は遺伝的要因が大きく、親が片頭痛の場合は子どもの発症リスクが約2〜4倍高くなります。特に母親からの遺伝が多いとされています" },
    { q: "市販の頭痛薬だけで大丈夫ですか？", a: "軽い発作には有効ですが、月に10日以上使用すると薬物乱用頭痛のリスクがあります。頻繁に薬が必要な方はトリプタンや予防薬の処方を相談してください" },
  ];
  qa.forEach(function(item, i) {
    var y = 1.1 + i * 1.35;
    card(s, 0.5, y, 9.0, 1.15, C.white, C.accent);
    s.addText("Q.", { x: 0.8, y: y + 0.05, w: 0.5, h: 0.45, fontSize: 22, fontFace: FE, color: C.accent, bold: true, margin: 0 });
    s.addText(item.q, { x: 1.3, y: y + 0.05, w: 8.0, h: 0.45, fontSize: 18, fontFace: FJ, color: C.accent, bold: true, valign: "middle", margin: 0 });
    s.addText("A.", { x: 0.8, y: y + 0.55, w: 0.5, h: 0.5, fontSize: 22, fontFace: FE, color: C.primary, bold: true, margin: 0 });
    s.addText(item.a, { x: 1.3, y: y + 0.55, w: 8.0, h: 0.55, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.1, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 12: まとめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 天気と頭痛の付き合い方");
  var points = [
    { num: "1", text: "天気頭痛は「気のせい」\nではない", sub: "内耳が気圧変化を感知し\n自律神経を介して頭痛を引き起こす" },
    { num: "2", text: "片頭痛には効果的な\n治療がある", sub: "トリプタン・予防薬・CGRP関連抗体\n我慢せず脳神経内科へ" },
    { num: "3", text: "頭痛ダイアリーが\n治療の第一歩", sub: "記録があれば医師が的確に\n診断・治療方針を立てられる" },
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
  s.addText("#14 味がしない・味が変わった ― 味覚障害の原因", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/weather_headache_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
