var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#17】その首の痛み、実は神経が原因？ ― 整形外科と脳神経内科の境界線";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #17", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🦴", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("その首の痛み、\n実は神経が原因？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.15, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("整形外科と脳神経内科の境界線を\n脳神経内科医が解説", {
    x: 0.6, y: 2.9, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #17", {
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
    { emoji: "😣", text: "首の後ろがズキズキ痛い" },
    { emoji: "⚡", text: "首を動かすと腕に電気が走る" },
    { emoji: "🤚", text: "肩から手にかけてしびれる" },
    { emoji: "🏥", text: "整形外科に行ったが「異常なし」と言われた" },
    { emoji: "❓", text: "首の痛みで何科に行けばいいかわからない" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 首の痛みの原因は大きく2つ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "首の痛み ― 原因は大きく2つ");
  // 筋肉・骨の問題
  card(s, 0.5, 1.1, 4.3, 2.0, C.white, C.accent);
  s.addText("🏥 筋肉・骨の問題", { x: 0.7, y: 1.15, w: 3.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("整形外科の領域", { x: 0.7, y: 1.55, w: 3.8, h: 0.3, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  var ortho = ["筋肉のこり・緊張", "椎間板ヘルニア", "変形性頸椎症", "寝違え・むち打ち"];
  ortho.forEach(function(item, i) {
    s.addText("• " + item, { x: 0.7, y: 1.95 + i * 0.28, w: 3.8, h: 0.28, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 神経の問題
  card(s, 5.2, 1.1, 4.3, 2.0, C.white, C.primary);
  s.addText("🧠 神経の問題", { x: 5.4, y: 1.15, w: 3.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("脳神経内科の領域", { x: 5.4, y: 1.55, w: 3.8, h: 0.3, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  var neuro = ["頸椎症性神経根症", "頸椎症性脊髄症", "後頭神経痛", "髄膜炎・くも膜下出血"];
  neuro.forEach(function(item, i) {
    s.addText("• " + item, { x: 5.4, y: 1.95 + i * 0.28, w: 3.8, h: 0.28, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  });
  // ポイント
  card(s, 0.5, 3.35, 9.0, 0.65, C.light, C.primary);
  s.addText("💡 同じ「首が痛い」でも、原因によって受診すべき科が違います", { x: 0.8, y: 3.4, w: 8.5, h: 0.55, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
  // 重なる領域
  card(s, 0.5, 4.2, 9.0, 0.75, C.lightYellow, C.yellow);
  s.addText("⚠️ 実際にはこの2つが重なることも多い", { x: 0.8, y: 4.25, w: 8.5, h: 0.3, fontSize: 16, fontFace: FJ, color: "856404", bold: true, margin: 0 });
  s.addText("例: 頸椎の変形（骨の問題）→ 神経が圧迫される（神経の問題）", { x: 0.8, y: 4.55, w: 8.5, h: 0.3, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  ftr(s);
})();

// SLIDE 4: 筋肉・骨が原因の首の痛み
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "筋肉・骨が原因の首の痛み（整形外科領域）");
  var conditions = [
    { name: "頸肩腕症候群（肩こり）", desc: "デスクワーク・ストレス・姿勢不良で首から肩の筋肉が緊張\n首〜肩〜腕の痛み・だるさ。最も多い原因", icon: "💻" },
    { name: "頸椎椎間板ヘルニア", desc: "椎間板が飛び出して神経を圧迫する\n首の痛み＋片方の腕へのしびれ・痛み", icon: "🦴" },
    { name: "変形性頸椎症", desc: "加齢で首の骨が変形し、関節がすり減る\n首の動きが悪くなり、動かすとゴリゴリ音がする", icon: "👤" },
  ];
  conditions.forEach(function(c, i) {
    var y = 1.1 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, C.accent);
    s.addText(c.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.9, fontSize: 30, align: "center", valign: "middle", margin: 0 });
    s.addText(c.name, { x: 1.4, y: y + 0.05, w: 7.8, h: 0.4, fontSize: 19, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
    s.addText(c.desc, { x: 1.4, y: y + 0.5, w: 7.8, h: 0.55, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 神経が原因の首の痛み
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "神経が原因の首の痛み（脳神経内科領域）");
  var conditions = [
    { name: "頸椎症性神経根症", desc: "首の骨の変形が神経の根元を圧迫\n首を後ろに反らすと腕に電気が走る痛み", icon: "⚡" },
    { name: "頸椎症性脊髄症", desc: "脊髄そのものが圧迫される重症タイプ\n手のしびれ・細かい動作がしにくい・歩行障害", icon: "🚶" },
    { name: "後頭神経痛", desc: "後頭部〜首の付け根にズキッとする鋭い痛み\n片側に多く、触ると痛みが誘発される", icon: "💥" },
  ];
  conditions.forEach(function(c, i) {
    var y = 1.1 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, C.primary);
    s.addText(c.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.9, fontSize: 30, align: "center", valign: "middle", margin: 0 });
    s.addText(c.name, { x: 1.4, y: y + 0.05, w: 7.8, h: 0.4, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(c.desc, { x: 1.4, y: y + 0.5, w: 7.8, h: 0.55, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 6: 見分けるポイント ― 痛みだけ？しびれもある？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "見分けるポイント ― 痛みだけ？しびれもある？");
  // 痛みだけ
  card(s, 0.5, 1.1, 4.3, 2.2, C.white, C.accent);
  s.addText("😣 痛みだけ → 筋肉・骨の可能性大", { x: 0.7, y: 1.15, w: 3.8, h: 0.4, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var painOnly = [
    "首を動かすと痛い",
    "肩がこる・張る",
    "首の付け根が重い",
    "温めると楽になる",
    "マッサージで改善する",
  ];
  painOnly.forEach(function(item, i) {
    s.addText("• " + item, { x: 0.7, y: 1.65 + i * 0.3, w: 3.8, h: 0.3, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 痛み＋しびれ
  card(s, 5.2, 1.1, 4.3, 2.2, C.white, C.primary);
  s.addText("⚡ 痛み＋しびれ → 神経の可能性あり", { x: 5.4, y: 1.15, w: 3.8, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var painNumb = [
    "腕や手にしびれが走る",
    "指先がピリピリする",
    "手に力が入りにくい",
    "首を反らすとしびれが悪化",
    "細かい作業がしにくい",
  ];
  painNumb.forEach(function(item, i) {
    s.addText("• " + item, { x: 5.4, y: 1.65 + i * 0.3, w: 3.8, h: 0.3, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });
  // まとめ
  card(s, 0.5, 3.55, 9.0, 0.65, C.lightYellow, C.yellow);
  s.addText("💡 「しびれ」「力が入らない」「細かい動作がしにくい」は神経が原因のサイン", { x: 0.8, y: 3.6, w: 8.5, h: 0.55, fontSize: 17, fontFace: FJ, color: "856404", bold: true, valign: "middle", margin: 0 });
  // 注意
  card(s, 0.5, 4.4, 9.0, 0.55, C.white, C.sub);
  s.addText("※ 筋肉の問題でも腕のだるさを感じることがあり、完全に分けられない場合もあります", { x: 0.8, y: 4.45, w: 8.5, h: 0.45, fontSize: 13, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 7: 後頭神経痛 ― 意外と多い首の痛みの正体
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "後頭神経痛 ― 意外と多い首の痛みの正体");
  // 概要
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("後頭部〜首の付け根にズキッ・ビリッとする鋭い痛み。「頭痛」として受診する方も多い", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
  // 特徴
  s.addText("特徴", { x: 0.5, y: 1.8, w: 4.3, h: 0.35, fontSize: 17, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var features = [
    "片側の後頭部〜耳の後ろに鋭い痛み",
    "数秒〜数十秒の電撃的な痛みが繰り返す",
    "痛い部位を押すと痛みが誘発される",
    "首の動きで悪化することがある",
  ];
  features.forEach(function(f, i) {
    var y = 2.25 + i * 0.45;
    card(s, 0.5, y, 4.3, 0.38, C.white, C.accent);
    s.addText("• " + f, { x: 0.8, y: y + 0.02, w: 3.8, h: 0.34, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  // 原因と対処
  card(s, 5.2, 1.8, 4.3, 1.2, C.white, C.primary);
  s.addText("原因", { x: 5.4, y: 1.85, w: 3.8, h: 0.3, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("ストレス・姿勢不良・頸椎の異常\nが後頭部の神経を刺激する\n長時間のデスクワークがきっかけに\nなることが多い", { x: 5.4, y: 2.15, w: 3.8, h: 0.75, fontSize: 13, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  card(s, 5.2, 3.15, 4.3, 0.7, C.lightGreen, C.green);
  s.addText("多くは鎮痛薬で改善します\n頸椎に原因がある場合はMRI検査も", { x: 5.4, y: 3.2, w: 3.8, h: 0.55, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 8: 頸椎症性脊髄症 ― 見逃してはいけない重症タイプ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "頸椎症性脊髄症 ― 見逃せない重症タイプ", C.red);
  card(s, 0.5, 1.1, 9.0, 0.5, C.lightRed, C.red);
  s.addText("脊髄そのものが圧迫される ― 放置すると回復が難しくなることも", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 17, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0 });
  // 3つのサイン
  s.addText("見逃してはいけない3つのサイン", { x: 0.5, y: 1.8, w: 9.0, h: 0.35, fontSize: 17, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var signs = [
    { icon: "✋", title: "手の巧緻運動障害", desc: "ボタンがかけにくい\n箸がうまく使えない\n字が書きにくくなった" },
    { icon: "🚶", title: "歩行障害", desc: "歩くときにふらつく\n階段が怖い\n足がもつれる" },
    { icon: "🔋", title: "膀胱直腸障害", desc: "尿が出にくい\n頻尿になった\n残尿感がある" },
  ];
  signs.forEach(function(sign, i) {
    var x = 0.5 + i * 3.1;
    card(s, x, 2.3, 2.9, 2.0, C.white, C.red);
    s.addText(sign.icon, { x: x + 0.2, y: 2.35, w: 0.6, h: 0.5, fontSize: 28, align: "center", margin: 0 });
    s.addText(sign.title, { x: x + 0.8, y: 2.35, w: 2.0, h: 0.45, fontSize: 16, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0 });
    s.addText(sign.desc, { x: x + 0.2, y: 2.9, w: 2.5, h: 1.2, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });
  // 注意
  card(s, 0.5, 4.5, 9.0, 0.55, C.lightYellow, C.yellow);
  s.addText("⚠️ これらの症状がある場合は早めに脳神経内科または整形外科を受診してください", { x: 0.8, y: 4.55, w: 8.5, h: 0.45, fontSize: 16, fontFace: FJ, color: "856404", bold: true, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 9: 整形外科と脳神経内科 ― どちらに行くべき？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "整形外科 vs 脳神経内科 ― どちらに行く？");
  // 整形外科
  card(s, 0.5, 1.1, 4.3, 1.8, C.white, C.accent);
  s.addText("🏥 整形外科を受診", { x: 0.7, y: 1.15, w: 3.8, h: 0.4, fontSize: 19, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var orthoItems = [
    "首・肩の痛みが中心",
    "首を動かすと痛い",
    "事故やスポーツ後の痛み",
    "肩こりがひどい",
  ];
  orthoItems.forEach(function(item, i) {
    s.addText("✓ " + item, { x: 0.7, y: 1.65 + i * 0.3, w: 3.8, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 脳神経内科
  card(s, 5.2, 1.1, 4.3, 1.8, C.white, C.primary);
  s.addText("🧠 脳神経内科を受診", { x: 5.4, y: 1.15, w: 3.8, h: 0.4, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var neuroItems = [
    "手足にしびれがある",
    "手に力が入りにくい",
    "細かい動作がしにくい",
    "歩行がふらつく",
  ];
  neuroItems.forEach(function(item, i) {
    s.addText("✓ " + item, { x: 5.4, y: 1.65 + i * 0.3, w: 3.8, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 両方の可能性
  card(s, 0.5, 3.15, 9.0, 0.65, C.lightYellow, C.yellow);
  s.addText("🤝 両方にまたがる場合も多い ― どちらを先に受診しても、必要に応じて紹介してもらえます", { x: 0.8, y: 3.2, w: 8.5, h: 0.55, fontSize: 16, fontFace: FJ, color: "856404", bold: true, valign: "middle", margin: 0 });
  // かかりつけ
  card(s, 0.5, 4.05, 9.0, 0.65, C.lightGreen, C.green);
  s.addText("💡 迷ったらまず「かかりつけ医」に相談 → 適切な専門科を紹介してもらうのも良い方法です", { x: 0.8, y: 4.1, w: 8.5, h: 0.55, fontSize: 16, fontFace: FJ, color: C.green, bold: true, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 10: 危険なサイン ― すぐに受診すべき首の痛み
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "危険なサイン ― すぐに受診すべき首の痛み", C.red);
  var signs = [
    { text: "突然の激しい首の痛み＋高熱\n（髄膜炎の可能性 → 119番）", icon: "🚑" },
    { text: "首の痛み＋手足の麻痺・感覚障害\n（脊髄圧迫の可能性 → 緊急受診）", icon: "⚡" },
    { text: "頭痛を伴う首のこわばり（首が前に曲がらない）\n（くも膜下出血・髄膜炎の可能性）", icon: "🧠" },
    { text: "事故・転倒後の首の痛み\n（骨折・脱臼の可能性 → 動かさず救急へ）", icon: "🦴" },
  ];
  signs.forEach(function(sign, i) {
    var y = 1.1 + i * 0.95;
    card(s, 0.5, y, 9.0, 0.8, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.1, w: 0.6, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.7, fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 11: 受診の目安 ― 緊急度3段階
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "首の痛み＋高熱・意識障害\n手足の麻痺・歩行困難\n事故後の首の痛み", action: "119番 or\n緊急受診", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "首の痛みが2週間以上続く\n腕や手にしびれが広がる\n手の筋力低下・巧緻運動障害\nだんだん悪くなっている", action: "整形外科 or\n脳神経内科", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "寝違え程度の一時的な痛み\n姿勢を変えると楽になる\n温めると改善する", action: "ストレッチ・\n姿勢改善から", bg: C.lightGreen, color: C.green },
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

// SLIDE 12: まとめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 首の痛みの正しい判断");
  var points = [
    { num: "1", text: "首の痛みの原因は\n「筋肉・骨」と「神経」", sub: "痛みだけなら整形外科\nしびれ・脱力があれば脳神経内科" },
    { num: "2", text: "後頭神経痛は意外と多い\n「頭痛」の正体", sub: "首の付け根の鋭い痛みが特徴\nストレス・姿勢不良が誘因に" },
    { num: "3", text: "手の巧緻運動障害・\n歩行障害は早めに受診", sub: "頸椎症性脊髄症は放置すると\n回復が難しくなる可能性あり" },
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
  s.addText("#18 子どもが「頭が痛い」と言ったら ― 親が知るべき判断基準", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 17, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/neck_pain_nerve_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
