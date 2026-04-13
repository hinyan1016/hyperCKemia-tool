var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#11】歩き方がおかしい ― 歩行パターンでわかる隠れた病気";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #11", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🚶", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("歩き方がおかしい", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.0,
    fontSize: 42, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("歩行パターンでわかる隠れた病気を\n脳神経内科医が解説", {
    x: 0.6, y: 2.7, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #11", {
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
    { emoji: "🚶", text: "家族に「歩き方が変」と言われた" },
    { emoji: "🦶", text: "つまずきやすくなった" },
    { emoji: "🌀", text: "まっすぐ歩いているつもりがふらつく" },
    { emoji: "👣", text: "歩幅が小さくなった気がする" },
    { emoji: "🦵", text: "長く歩くと足が重くなって休みたくなる" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 歩行のメカニズム ― なぜ歩き方で病気がわかるか
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "歩行のメカニズム ― なぜ歩き方で病気がわかる？");
  var parts = [
    { icon: "🧠", name: "大脳・基底核", role: "歩行の開始・速度・リズムを制御", disease: "パーキンソン病→小刻み歩行", color: C.primary },
    { icon: "🔵", name: "小脳", role: "バランス・協調運動を調整", disease: "小脳疾患→ふらつき歩行", color: C.accent },
    { icon: "🦴", name: "脊髄", role: "脳からの運動指令を手足に伝える", disease: "脊髄障害→つっぱり歩行", color: "7B2D8E" },
    { icon: "⚡", name: "末梢神経・筋肉", role: "筋肉を動かし関節を制御", disease: "神経障害→つま先が上がらない", color: C.red },
  ];
  parts.forEach(function(p, i) {
    var y = 1.1 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, p.color);
    s.addText(p.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 28, align: "center", margin: 0 });
    s.addText(p.name, { x: 1.4, y: y + 0.05, w: 2.0, h: 0.4, fontSize: 19, fontFace: FJ, color: p.color, bold: true, margin: 0 });
    s.addText(p.role, { x: 3.5, y: y + 0.05, w: 3.3, h: 0.4, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText(p.disease, { x: 3.5, y: y + 0.45, w: 5.5, h: 0.35, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  });
  s.addText("※ 歩行は脳から筋肉まで全身の協調で成り立つ → どこが障害されたかで歩き方が変わる", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.2, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0,
  });
  ftr(s);
})();

// SLIDE 4: 代表的な歩行パターン（前半）
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "歩行パターンでわかる病気 ― 前半");
  var gaits = [
    { icon: "👣", name: "小刻み歩行", feature: "歩幅が狭く前かがみ\nすくみ足・腕振りが減る", cause: "パーキンソン病\n正常圧水頭症", color: C.primary },
    { icon: "🦵", name: "痙性歩行", feature: "足がつっぱり円を描くように\n振り出す（ぶん回し歩行）", cause: "脳卒中後遺症\n脊髄障害", color: C.red },
    { icon: "🌀", name: "失調性歩行", feature: "ふらつきが大きく酔っ払いのよう\n足を広げてバランスをとる", cause: "小脳疾患\n脊髄小脳変性症", color: C.accent },
  ];
  gaits.forEach(function(g, i) {
    var y = 1.1 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, g.color);
    s.addText(g.icon, { x: 0.7, y: y + 0.15, w: 0.6, h: 0.8, fontSize: 30, align: "center", margin: 0 });
    s.addText(g.name, { x: 1.4, y: y + 0.05, w: 2.3, h: 0.5, fontSize: 22, fontFace: FJ, color: g.color, bold: true, valign: "middle", margin: 0 });
    s.addText(g.feature, { x: 3.8, y: y + 0.05, w: 3.0, h: 0.6, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText(g.cause, { x: 3.8, y: y + 0.65, w: 3.0, h: 0.4, fontSize: 14, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText("原因", { x: 7.0, y: y + 0.05, w: 0.7, h: 0.4, fontSize: 12, fontFace: FJ, color: g.color, bold: true, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 代表的な歩行パターン（後半）
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "歩行パターンでわかる病気 ― 後半");
  var gaits = [
    { icon: "🦶", name: "鶏歩", feature: "つま先が上がらず膝を高く上げて\nペタペタと歩く（下垂足）", cause: "腓骨神経麻痺\n腰椎椎間板ヘルニア", color: "7B2D8E" },
    { icon: "🦆", name: "動揺歩行", feature: "体を左右に揺らしながら歩く\n（あひる歩き）", cause: "筋ジストロフィー\n近位筋の筋力低下", color: C.accent },
    { icon: "🛑", name: "間欠性跛行", feature: "歩くと足が重く痛くなり\n休むと回復して歩ける", cause: "脊柱管狭窄症\n末梢動脈疾患", color: C.red },
  ];
  gaits.forEach(function(g, i) {
    var y = 1.1 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, g.color);
    s.addText(g.icon, { x: 0.7, y: y + 0.15, w: 0.6, h: 0.8, fontSize: 30, align: "center", margin: 0 });
    s.addText(g.name, { x: 1.4, y: y + 0.05, w: 2.3, h: 0.5, fontSize: 22, fontFace: FJ, color: g.color, bold: true, valign: "middle", margin: 0 });
    s.addText(g.feature, { x: 3.8, y: y + 0.05, w: 3.0, h: 0.6, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText(g.cause, { x: 3.8, y: y + 0.65, w: 3.0, h: 0.4, fontSize: 14, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText("原因", { x: 7.0, y: y + 0.05, w: 0.7, h: 0.4, fontSize: 12, fontFace: FJ, color: g.color, bold: true, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 6: よくある原因 ― 心配度で分類
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "歩行障害の原因 ― 心配度で分類");
  var causes = [
    { level: "🟢 軽度", name: "加齢による筋力低下", desc: "歩幅の減少・歩行速度の低下。筋トレで改善可能", bg: C.lightGreen, color: C.green },
    { level: "🟡 要注意", name: "パーキンソン病", desc: "小刻み歩行・すくみ足。早期治療で症状コントロール可能", bg: C.lightYellow, color: "856404" },
    { level: "🟡 要注意", name: "脊柱管狭窄症", desc: "長く歩くと足が重い。前かがみや休憩で回復", bg: C.lightYellow, color: "856404" },
    { level: "🟡 要注意", name: "正常圧水頭症", desc: "小刻み歩行＋認知症＋尿失禁。手術で改善可能", bg: C.lightYellow, color: "856404" },
    { level: "🔴 緊急", name: "脳卒中", desc: "突然の歩行困難＋片側の脱力・ろれつ障害", bg: C.lightRed, color: C.red },
  ];
  causes.forEach(function(c, i) {
    var y = 1.1 + i * 0.8;
    card(s, 0.5, y, 9.0, 0.67, c.bg, c.color);
    s.addText(c.level, { x: 0.7, y: y + 0.05, w: 1.3, h: 0.55, fontSize: 15, fontFace: FJ, color: c.color, bold: true, valign: "middle", margin: 0 });
    s.addText(c.name, { x: 2.0, y: y + 0.05, w: 3.0, h: 0.55, fontSize: 17, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(c.desc, { x: 5.1, y: y + 0.05, w: 4.2, h: 0.55, fontSize: 14, fontFace: FJ, color: C.sub, valign: "middle", lineSpacingMultiple: 1.1, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 7: 見逃されやすい原因 ― 正常圧水頭症
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "見逃されやすい原因 ― 正常圧水頭症（iNPH）");
  // 三徴カード
  s.addText("治る認知症とも呼ばれる ― 手術で改善可能な歩行障害", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.35, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0,
  });
  var triad = [
    { icon: "🚶", name: "歩行障害", desc: "小刻み歩行・すくみ足\n足が床に張り付く感覚", note: "最も早期に出現" },
    { icon: "🧠", name: "認知症", desc: "物忘れ・注意力低下\n意欲の減退", note: "アルツハイマーと誤診されやすい" },
    { icon: "🚽", name: "尿失禁", desc: "頻尿・切迫性尿失禁\n間に合わない", note: "進行すると出現" },
  ];
  triad.forEach(function(t, i) {
    var y = 1.55 + i * 1.1;
    card(s, 0.5, y, 9.0, 0.95, C.white, C.primary);
    s.addText(t.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.75, fontSize: 28, align: "center", margin: 0 });
    s.addText(t.name, { x: 1.4, y: y + 0.05, w: 1.8, h: 0.45, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.desc, { x: 3.3, y: y + 0.05, w: 3.5, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText(t.note, { x: 3.3, y: y + 0.55, w: 5.5, h: 0.35, fontSize: 13, fontFace: FJ, color: C.accent, margin: 0 });
  });
  s.addText("※ 髄液シャント術で歩行障害の約8割が改善 ― 早期発見が重要", {
    x: 0.5, y: 4.95, w: 9.0, h: 0.25, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });
  ftr(s);
})();

// SLIDE 8: セルフチェック
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自分でできるチェック ― 歩行の変化");
  var checks = [
    { num: "1", title: "歩幅と速度を確認", desc: "横断歩道を青信号で渡りきれるか？\n以前より歩幅が小さくなっていないか", icon: "📏" },
    { num: "2", title: "バランスを確認", desc: "まっすぐ歩けるか？ふらつかないか\n方向転換でよろめかないか", icon: "⚖️" },
    { num: "3", title: "左右差を確認", desc: "片足だけ引きずっていないか\n靴底の減り方に左右差がないか", icon: "👟" },
  ];
  checks.forEach(function(ch, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, C.primary);
    s.addText(ch.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.85, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(ch.icon, { x: 1.5, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addText(ch.title, { x: 2.3, y: y + 0.1, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(ch.desc, { x: 2.3, y: y + 0.55, w: 6.5, h: 0.45, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 9: 危険なサイン
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな場合はすぐ受診 ― 危険なサイン", C.red);
  var signs = [
    { text: "突然歩けなくなった＋片側の脱力・ろれつ障害（脳卒中）", icon: "🧠" },
    { text: "数日〜数週間で急速に歩行が悪化（脊髄障害・ギラン・バレー）", icon: "⚡" },
    { text: "歩行障害＋排尿障害がある（脊髄圧迫の可能性）", icon: "⚠️" },
    { text: "転倒を繰り返すようになった（骨折・頭部外傷のリスク）", icon: "🩹" },
    { text: "歩行障害＋認知機能低下＋尿失禁（正常圧水頭症の三徴）", icon: "🚶" },
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

// SLIDE 10: 受診の目安
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "突然の歩行困難＋片側脱力・ろれつ障害\n急速に歩行が悪化している", action: "脳卒中・脊髄障害\n→119番", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "数週間〜数ヶ月で徐々に歩きにくくなった\n転倒が増えた・歩幅が小さくなった\n認知機能低下や尿失禁を伴う", action: "脳神経内科で精査\nMRI・歩行分析", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "疲労や体調不良時の一時的なふらつき\n長時間の歩行後の足の重さ", action: "休息で改善すれば\n心配不要", bg: C.lightGreen, color: C.green },
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

// SLIDE 11: Q&A
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある質問 Q&A");
  var qa = [
    { q: "歩き方がおかしいのは何科に行けばいい？", a: "まず脳神経内科をおすすめします。歩行障害の原因は脳・脊髄・末梢神経・筋肉と多岐にわたりますが、脳神経内科はこれらを総合的に診断できます。必要に応じて整形外科等に紹介します" },
    { q: "パーキンソン病の歩行障害は治りますか？", a: "薬物療法で歩行を含む症状を大幅に改善できます。特に早期から適切な治療とリハビリを行うことで、長期間にわたって良好な歩行機能を維持できます" },
    { q: "正常圧水頭症の手術は安全ですか？", a: "髄液シャント術は確立された手術で、歩行障害の改善率は約80%と報告されています。高齢者でも比較的安全に行える手術です" },
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
  hdr(s, "まとめ ― 歩き方がおかしいと感じたら");
  var points = [
    { num: "1", text: "歩行パターンが\n原因疾患を示唆する", sub: "小刻み→パーキンソン病\nふらつき→小脳 / つっぱり→脊髄" },
    { num: "2", text: "正常圧水頭症は\n手術で改善可能", sub: "歩行障害＋認知症＋尿失禁の\n三徴に注目" },
    { num: "3", text: "歩行障害は脳神経内科で\n総合的に診断できる", sub: "脳・脊髄・末梢神経・筋肉の\nどこが問題かを判断" },
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
  s.addText("🔍 歩行セルフチェックツール（概要欄にリンク）", {
    x: 1.5, y: 2.95, w: 7.0, h: 0.4, fontSize: 16, fontFace: FJ, color: "B0BEC5", margin: 0,
  });
  s.addText("次回予告", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: C.accent, bold: true, align: "center", margin: 0,
  });
  s.addText("#12 頭をぶつけた後、病院に行くべき？", {
    x: 0.6, y: 4.1, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/gait_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
