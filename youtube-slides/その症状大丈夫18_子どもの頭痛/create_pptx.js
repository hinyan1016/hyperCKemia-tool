var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#18】子どもが「頭が痛い」と言ったら ― 親が知るべき判断基準";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #18", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("👧", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("子どもが「頭が痛い」\nと言ったら", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.15, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("親が知るべき判断基準を\n脳神経内科医が解説", {
    x: 0.6, y: 2.9, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #18", {
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
    { emoji: "😢", text: "子どもが「頭が痛い」と訴えるが、どう判断すれば？" },
    { emoji: "🏫", text: "学校で頭痛を訴えて保健室に行くことが増えた" },
    { emoji: "🤮", text: "頭痛と一緒に吐くことがある" },
    { emoji: "😴", text: "朝起きたときに頭が痛いと言う" },
    { emoji: "❓", text: "病院に連れて行くべきか迷う" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 子どもの頭痛は珍しくない
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "子どもの頭痛は珍しくない");
  card(s, 0.5, 1.1, 9.0, 0.6, C.light, C.primary);
  s.addText("小中学生の約3〜5割が頭痛を経験 ― 多くは心配のないタイプ", { x: 0.8, y: 1.15, w: 8.5, h: 0.5, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
  // 2つのタイプ
  card(s, 0.5, 1.95, 4.3, 2.0, C.white, C.primary);
  s.addText("🔵 一次性頭痛（頭痛そのものが病気）", { x: 0.7, y: 2.0, w: 3.8, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var primary = ["片頭痛（子どもにも多い）", "緊張型頭痛（最多）", "頭痛の原因となる病気はない", "→ 多くは適切な対応で改善"];
  primary.forEach(function(item, i) {
    s.addText("• " + item, { x: 0.7, y: 2.5 + i * 0.32, w: 3.8, h: 0.32, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });
  card(s, 5.2, 1.95, 4.3, 2.0, C.white, C.red);
  s.addText("🔴 二次性頭痛（別の病気が原因）", { x: 5.4, y: 2.0, w: 3.8, h: 0.4, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var secondary = ["髄膜炎・脳炎", "脳腫瘍", "水頭症", "→ 見逃してはいけない"];
  secondary.forEach(function(item, i) {
    s.addText("• " + item, { x: 5.4, y: 2.5 + i * 0.32, w: 3.8, h: 0.32, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });
  // ポイント
  card(s, 0.5, 4.2, 9.0, 0.65, C.lightYellow, C.yellow);
  s.addText("💡 ほとんどは一次性頭痛ですが、「いつもと違う頭痛」を見逃さないことが大切", { x: 0.8, y: 4.25, w: 8.5, h: 0.55, fontSize: 17, fontFace: FJ, color: "856404", bold: true, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 4: 子どもの片頭痛の特徴
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "子どもの片頭痛 ― 大人とここが違う");
  // 大人との違い
  var diffs = [
    { child: "両側が痛い（おでこ全体）", adult: "片側が痛い", icon: "🧒" },
    { child: "持続時間が短い（1〜2時間）", adult: "4〜72時間", icon: "⏱️" },
    { child: "吐き気・嘔吐が目立つ", adult: "頭痛が主体", icon: "🤢" },
    { child: "顔色が悪くなる", adult: "光・音に敏感", icon: "😶" },
    { child: "寝ると治ることが多い", adult: "薬が必要なことが多い", icon: "😴" },
  ];
  diffs.forEach(function(d, i) {
    var y = 1.1 + i * 0.75;
    card(s, 0.5, y, 4.3, 0.6, C.white, C.primary);
    s.addText(d.icon, { x: 0.65, y: y + 0.05, w: 0.5, h: 0.5, fontSize: 22, align: "center", margin: 0 });
    s.addText("子ども: " + d.child, { x: 1.2, y: y + 0.05, w: 3.4, h: 0.5, fontSize: 14, fontFace: FJ, color: C.primary, valign: "middle", margin: 0 });
    card(s, 5.2, y, 4.3, 0.6, C.white, C.sub);
    s.addText("大人: " + d.adult, { x: 5.5, y: y + 0.05, w: 3.7, h: 0.5, fontSize: 14, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  });
  card(s, 0.5, 4.95, 9.0, 0.2, C.lightGreen);
  s.addText("※ 子どもの片頭痛は「お腹が痛い」と表現されることもあります", { x: 0.8, y: 4.95, w: 8.5, h: 0.2, fontSize: 12, fontFace: FJ, color: C.green, margin: 0 });
  ftr(s);
})();

// SLIDE 5: 緊張型頭痛 ― 最も多い子どもの頭痛
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "緊張型頭痛 ― 最も多い子どもの頭痛");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("頭全体が締めつけられるような痛み ― ストレス・姿勢・睡眠不足が原因", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
  // 特徴
  s.addText("特徴", { x: 0.5, y: 1.8, w: 4.3, h: 0.35, fontSize: 17, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var features = [
    "頭全体が「ぎゅっ」と痛い",
    "我慢できる程度の痛み",
    "吐き気は少ない",
    "動いても悪化しない",
  ];
  features.forEach(function(f, i) {
    var y = 2.25 + i * 0.45;
    card(s, 0.5, y, 4.3, 0.38, C.white, C.accent);
    s.addText("• " + f, { x: 0.8, y: y + 0.02, w: 3.8, h: 0.34, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  // 原因
  card(s, 5.2, 1.8, 4.3, 1.2, C.white, C.primary);
  s.addText("よくある原因", { x: 5.4, y: 1.85, w: 3.8, h: 0.3, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("学校でのストレス・友人関係\nスマホ・ゲームの長時間使用\n睡眠不足・不規則な生活\n姿勢不良（猫背）", { x: 5.4, y: 2.2, w: 3.8, h: 0.7, fontSize: 13, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  // 対処
  card(s, 5.2, 3.15, 4.3, 0.7, C.lightGreen, C.green);
  s.addText("生活習慣の見直しで\n多くは改善します", { x: 5.4, y: 3.2, w: 3.8, h: 0.6, fontSize: 15, fontFace: FJ, color: C.green, bold: true, lineSpacingMultiple: 1.2, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 6: 頭痛ダイアリーのすすめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "頭痛ダイアリー ― 受診前にやっておくと便利");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("頭痛の記録をつけることで、医師が正確に診断しやすくなります", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
  var items = [
    { icon: "📅", title: "いつ？", desc: "日付・曜日・時間帯（朝？夕方？）" },
    { icon: "⏱️", title: "どのくらい？", desc: "痛みの持続時間（分〜時間）" },
    { icon: "📍", title: "どこが？", desc: "痛む場所（おでこ？後頭部？片側？）" },
    { icon: "😣", title: "どんな痛み？", desc: "ズキズキ？ぎゅっと？ガンガン？" },
    { icon: "🤢", title: "他の症状は？", desc: "吐き気・嘔吐・光や音が嫌など" },
    { icon: "🎯", title: "きっかけは？", desc: "寝不足？テスト前？天気？運動後？" },
  ];
  items.forEach(function(item, i) {
    var y = 1.8 + i * 0.55;
    card(s, 0.5, y, 9.0, 0.45, C.white, C.accent);
    s.addText(item.icon, { x: 0.7, y: y + 0.02, w: 0.5, h: 0.4, fontSize: 20, align: "center", margin: 0 });
    s.addText(item.title, { x: 1.3, y: y + 0.02, w: 1.8, h: 0.4, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, valign: "middle", margin: 0 });
    s.addText(item.desc, { x: 3.2, y: y + 0.02, w: 6.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 7: 危険な頭痛のサイン
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "危険な頭痛のサイン ― これがあればすぐ受診", C.red);
  var signs = [
    { text: "突然の激しい頭痛（今まで経験したことがない）\n（くも膜下出血・脳出血の可能性）", icon: "🚑" },
    { text: "頭痛＋高熱＋首のこわばり＋嘔吐\n（髄膜炎・脳炎の可能性 → 救急受診）", icon: "🔥" },
    { text: "頭痛が日に日に悪化している\n（脳腫瘍・水頭症の可能性）", icon: "📈" },
    { text: "朝の頭痛＋噴射性の嘔吐（噴水のように吐く）\n（頭蓋内圧亢進の可能性）", icon: "🤮" },
  ];
  signs.forEach(function(sign, i) {
    var y = 1.1 + i * 0.95;
    card(s, 0.5, y, 9.0, 0.8, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.1, w: 0.6, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.7, fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 8: 「いつもと違う」を見逃さない
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「いつもと違う」を見逃さない");
  // いつもの頭痛
  card(s, 0.5, 1.1, 4.3, 2.2, C.white, C.green);
  s.addText("🟢 いつもの頭痛パターン", { x: 0.7, y: 1.15, w: 3.8, h: 0.4, fontSize: 17, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  var usual = [
    "痛み方が毎回似ている",
    "寝ると治る",
    "食欲があり元気",
    "遊ぶと忘れている",
    "学校に行ける程度",
  ];
  usual.forEach(function(item, i) {
    s.addText("• " + item, { x: 0.7, y: 1.65 + i * 0.3, w: 3.8, h: 0.3, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });
  // いつもと違う
  card(s, 5.2, 1.1, 4.3, 2.2, C.white, C.red);
  s.addText("🔴 いつもと違う ← 要注意", { x: 5.4, y: 1.15, w: 3.8, h: 0.4, fontSize: 17, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var unusual = [
    "今までにない激しい痛み",
    "寝ても治らない・悪化する",
    "元気がなくぐったりしている",
    "痛みで目が覚める",
    "学校に行けないほど強い",
  ];
  unusual.forEach(function(item, i) {
    s.addText("• " + item, { x: 5.4, y: 1.65 + i * 0.3, w: 3.8, h: 0.3, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });
  // まとめ
  card(s, 0.5, 3.55, 9.0, 0.65, C.lightYellow, C.yellow);
  s.addText("💡 子どもの頭痛は「普段のパターン」を知っておくことが重要です", { x: 0.8, y: 3.6, w: 8.5, h: 0.55, fontSize: 18, fontFace: FJ, color: "856404", bold: true, valign: "middle", margin: 0 });
  // 行動の変化
  card(s, 0.5, 4.4, 9.0, 0.55, C.white, C.accent);
  s.addText("※ 小さなお子さんは「頭が痛い」と言えません。ぐずる・食欲低下・頭を抱えるなどの行動に注目", { x: 0.8, y: 4.45, w: 8.5, h: 0.45, fontSize: 13, fontFace: FJ, color: C.accent, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 9: 受診先の選び方
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "何科を受診する？ ― 子どもの頭痛");
  // 小児科
  card(s, 0.5, 1.1, 4.3, 1.6, C.white, C.accent);
  s.addText("👶 まずは小児科", { x: 0.7, y: 1.15, w: 3.8, h: 0.4, fontSize: 19, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var peds = [
    "子どもの頭痛の第一選択",
    "全身状態の評価ができる",
    "感染症の除外ができる",
    "必要に応じて専門科へ紹介",
  ];
  peds.forEach(function(item, i) {
    s.addText("✓ " + item, { x: 0.7, y: 1.65 + i * 0.25, w: 3.8, h: 0.25, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 脳神経内科・小児神経科
  card(s, 5.2, 1.1, 4.3, 1.6, C.white, C.primary);
  s.addText("🧠 脳神経内科・小児神経科", { x: 5.4, y: 1.15, w: 3.8, h: 0.4, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var neuro = [
    "頭痛が繰り返す場合",
    "片頭痛が疑われる場合",
    "神経症状がある場合",
    "MRI検査が必要な場合",
  ];
  neuro.forEach(function(item, i) {
    s.addText("✓ " + item, { x: 5.4, y: 1.65 + i * 0.25, w: 3.8, h: 0.25, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 眼科
  card(s, 0.5, 2.95, 4.3, 0.8, C.white, C.sub);
  s.addText("👁️ 眼科も忘れずに", { x: 0.7, y: 3.0, w: 3.8, h: 0.3, fontSize: 16, fontFace: FJ, color: C.sub, bold: true, margin: 0 });
  s.addText("視力低下・斜視による眼精疲労が\n頭痛の原因になることも", { x: 0.7, y: 3.3, w: 3.8, h: 0.4, fontSize: 13, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  // 救急
  card(s, 5.2, 2.95, 4.3, 0.8, C.lightRed, C.red);
  s.addText("🚑 救急（119番）", { x: 5.4, y: 3.0, w: 3.8, h: 0.3, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("意識がおかしい・けいれん\n突然の激しい頭痛＋嘔吐", { x: 5.4, y: 3.3, w: 3.8, h: 0.4, fontSize: 13, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  ftr(s);
})();

// SLIDE 10: 受診の目安 ― 緊急度3段階
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "突然の激しい頭痛＋嘔吐\n高熱＋首のこわばり\n意識がおかしい・けいれん", action: "119番 or\n緊急受診", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "頭痛が1週間以上続く\n朝の頭痛が繰り返す\n頭痛が日に日に強くなっている\n頭痛で学校を休むことが増えた", action: "小児科 or\n脳神経内科", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "一時的で寝ると治る\n元気で食欲もある\nいつもと同じパターンの頭痛", action: "生活習慣の\n見直しから", bg: C.lightGreen, color: C.green },
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

// SLIDE 11: 家庭でできる対処法
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "家庭でできる対処法");
  var tips = [
    { num: "1", title: "十分な睡眠", desc: "小学生は9〜11時間、中学生は\n8〜10時間の睡眠を確保", icon: "😴" },
    { num: "2", title: "規則正しい生活", desc: "朝食を抜かない、就寝時間を\n一定にする", icon: "🕐" },
    { num: "3", title: "画面時間の制限", desc: "スマホ・ゲームは1日1〜2時間\nまでを目安に", icon: "📱" },
    { num: "4", title: "水分をしっかり", desc: "脱水は頭痛の原因に。こまめに\n水分を取らせる", icon: "💧" },
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
  hdr(s, "まとめ ― 子どもの頭痛への正しい対応");
  var points = [
    { num: "1", text: "子どもの頭痛の多くは\n片頭痛か緊張型頭痛", sub: "小中学生の3〜5割が経験\n生活習慣の見直しで改善" },
    { num: "2", text: "「いつもと違う頭痛」\nが受診の判断基準", sub: "普段のパターンを知っておく\n頭痛ダイアリーが役立つ" },
    { num: "3", text: "高熱＋嘔吐＋首のこわばり\nは救急受診を", sub: "髄膜炎・脳腫瘍の可能性\n朝の頭痛の繰り返しにも注意" },
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
  s.addText("#19 認知症と運転免許 ― 家族が知っておくべき制度と説得のコツ", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 17, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/child_headache_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
