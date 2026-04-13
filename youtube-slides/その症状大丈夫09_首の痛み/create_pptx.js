var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#09】首が痛い ― 原因と見分け方";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #09", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🦴", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("首が痛い", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.0,
    fontSize: 42, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("原因と見分け方を\n脳神経内科医が解説", {
    x: 0.6, y: 2.7, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #09", {
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
    { emoji: "😣", text: "朝起きたら首が痛くて動かせない" },
    { emoji: "💻", text: "デスクワークの後に首から肩がガチガチ" },
    { emoji: "⚡", text: "首を回すと腕や手にしびれが走る" },
    { emoji: "🤕", text: "首の痛みとともに頭痛がする" },
    { emoji: "🚶", text: "つまずきやすい・箸が使いにくくなった" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 首の痛み ― 2つのタイプ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "首の痛み ― 2つのタイプ");

  card(s, 0.4, 1.2, 4.4, 2.8, C.lightGreen, C.green);
  s.addText("💪 筋骨格系の痛み", { x: 0.7, y: 1.3, w: 3.8, h: 0.5, fontSize: 21, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("首や肩の筋肉・関節が原因", { x: 0.7, y: 1.8, w: 3.8, h: 0.35, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  var left = ["肩こり・寝違え", "筋筋膜性疼痛", "変形性頚椎症（軽度）", "→ 整形外科が主な受診先"];
  left.forEach(function(t, i) {
    s.addText("・" + t, { x: 0.8, y: 2.3 + i * 0.38, w: 3.8, h: 0.33, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.2, 1.2, 4.4, 2.8, C.lightRed, C.red);
  s.addText("🧠 神経系の痛み", { x: 5.5, y: 1.3, w: 3.8, h: 0.5, fontSize: 21, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("頚椎の神経や脊髄・脳の病気", { x: 5.5, y: 1.8, w: 3.8, h: 0.35, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  var right = ["頚椎症性脊髄症", "椎間板ヘルニア（神経圧迫）", "髄膜炎・椎骨動脈解離", "→ 脳神経内科の受診が必要"];
  right.forEach(function(t, i) {
    s.addText("・" + t, { x: 5.6, y: 2.3 + i * 0.38, w: 3.8, h: 0.33, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  });

  s.addText("⚠️ 手のしびれ・脱力・歩行障害を伴う首の痛みは神経の問題を疑いましょう", {
    x: 0.5, y: 4.95, w: 9.0, h: 0.3, fontSize: 14, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });
  ftr(s);
})();

// SLIDE 4: 原因 ― 心配度で分類
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "首の痛みの原因 ― 心配度で分類");
  var causes = [
    { level: "🟢 軽度", name: "肩こり・寝違え", desc: "筋肉の緊張やこり。ストレッチや姿勢改善で軽快", bg: C.lightGreen, color: C.green },
    { level: "🟡 要注意", name: "頚椎症（変形性）", desc: "加齢で頚椎が変形。50歳以上に非常に多い", bg: C.lightYellow, color: "856404" },
    { level: "🟡 要注意", name: "頚椎椎間板ヘルニア", desc: "椎間板の突出で神経を圧迫。腕のしびれが特徴", bg: C.lightYellow, color: "856404" },
    { level: "🔴 緊急", name: "頚椎症性脊髄症", desc: "脊髄の圧迫。手の巧緻運動障害＋歩行障害", bg: C.lightRed, color: C.red },
    { level: "🔴 緊急", name: "髄膜炎・椎骨動脈解離", desc: "発熱＋首の硬直、突然の後頭部痛。命に関わる", bg: C.lightRed, color: C.red },
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

// SLIDE 5: 見逃されやすい原因 ― 脳神経の病気
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "見逃されやすい原因 ― 脳神経の病気");
  var diseases = [
    { icon: "🦴", name: "頚椎症性脊髄症", desc: "脊髄の圧迫で手の巧緻運動障害＋歩行障害。\n首の痛みが軽い・ないこともあり見逃されやすい", color: C.primary },
    { icon: "🦠", name: "髄膜炎", desc: "脳と脊髄を包む膜の感染症。\n発熱＋首の硬直（前に曲げると痛い）が特徴", color: C.red },
    { icon: "💥", name: "椎骨動脈解離", desc: "首の後ろの動脈壁が裂ける。突然の後頭部痛。\n脳卒中の原因になる重大な病気", color: C.accent },
    { icon: "🚨", name: "くも膜下出血", desc: "突然の激しい頭痛＋首の痛み・硬直。\n一刻を争う緊急事態", color: "7B2D8E" },
  ];
  diseases.forEach(function(d, i) {
    var y = 1.1 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, d.color);
    s.addText(d.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 26, align: "center", margin: 0 });
    s.addText(d.name, { x: 1.4, y: y + 0.05, w: 2.8, h: 0.4, fontSize: 19, fontFace: FJ, color: d.color, bold: true, margin: 0 });
    s.addText(d.desc, { x: 1.4, y: y + 0.4, w: 7.8, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 6: 頚椎症性脊髄症
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "首の痛みと頚椎症性脊髄症");

  card(s, 0.5, 1.1, 9.0, 0.8, C.light, C.primary);
  s.addText("頚椎の変形で脊髄が圧迫される病気 ― 脳神経内科で最も多い首の疾患の一つ", {
    x: 0.8, y: 1.15, w: 8.5, h: 0.35, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, margin: 0,
  });
  s.addText("首の痛みが軽い・全くないこともあり、見逃されやすい", {
    x: 0.8, y: 1.55, w: 8.5, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  var features = [
    { icon: "✋", title: "手の巧緻運動障害", desc: "ボタンが留めにくい\n箸が使いにくい・字が下手に" },
    { icon: "🚶", title: "歩行障害", desc: "階段の下りが怖い\n小走りができなくなる" },
    { icon: "🔢", title: "しびれ", desc: "両手のしびれが多い\n足のしびれを伴うことも" },
    { icon: "⚠️", title: "放置すると進行", desc: "不可逆的な障害が残る\n早期発見・早期治療が重要" },
  ];
  features.forEach(function(f, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 2.15 + row * 1.4;
    card(s, x, y, 4.5, 1.2, C.white, C.primary);
    s.addText(f.icon, { x: x + 0.2, y: y + 0.08, w: 0.6, h: 0.5, fontSize: 22, align: "center", margin: 0 });
    s.addText(f.title, { x: x + 0.8, y: y + 0.08, w: 3.4, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(f.desc, { x: x + 0.8, y: y + 0.5, w: 3.4, h: 0.6, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 7: セルフチェック
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自分でできるチェック ― 3つのテスト");
  var checks = [
    { num: "1", title: "10秒テスト", desc: "両手をパーに開いて10秒間グーパーを繰り返す\n20回以上→正常 / 20回未満→巧緻運動障害の疑い", icon: "✊" },
    { num: "2", title: "スパーリングテスト", desc: "首をやや後ろに反らして痛い側に傾ける\n腕に痛み・しびれが放散→神経根の圧迫の疑い", icon: "🔍" },
    { num: "3", title: "歩行チェック", desc: "一直線上を歩く / つま先歩き・かかと歩き\nふらつきがあれば脊髄症の可能性", icon: "🚶" },
  ];
  checks.forEach(function(ch, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, C.primary);
    s.addText(ch.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.85, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(ch.icon, { x: 1.5, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addText(ch.title, { x: 2.3, y: y + 0.1, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(ch.desc, { x: 2.3, y: y + 0.55, w: 6.5, h: 0.45, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
  });
  s.addText("※ 痛みが強い場合は無理をせず中止してください", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });
  ftr(s);
})();

// SLIDE 8: 危険なサイン
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな場合はすぐ受診 ― 危険なサイン", C.red);
  var signs = [
    { text: "発熱＋首の硬直（髄膜炎の可能性）", icon: "🦠" },
    { text: "突然の激しい後頭部の痛み（椎骨動脈解離・くも膜下出血）", icon: "💥" },
    { text: "手の巧緻運動障害・歩行のふらつきが進行（脊髄症）", icon: "✋" },
    { text: "首の痛み＋両手足のしびれ（脊髄の圧迫）", icon: "⚡" },
    { text: "外傷後の首の痛み＋手足のしびれ（頚椎骨折・脊髄損傷）", icon: "🚑" },
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

// SLIDE 9: 受診の目安 ― 緊急度3段階
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "発熱＋首の硬直\n突然の激しい後頭部の痛み\n外傷後の首の痛み＋手足のしびれ", action: "髄膜炎・くも膜下出血\n脊髄損傷→119番", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "手のしびれ・脱力が続く\n歩行のふらつきが進行\n2週間以上改善しない", action: "脳神経内科・整形外科\nで原因精査", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "寝違え・肩こり程度\nストレッチで改善する", action: "悪化したら受診\n姿勢改善で予防", bg: C.lightGreen, color: C.green },
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

// SLIDE 10: 検査
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "首の痛みを調べる検査");
  var tests = [
    { icon: "🦴", name: "頚椎MRI", targets: "脊髄の圧迫・椎間板ヘルニア・脊髄症", note: "しびれ・歩行障害では必須" },
    { icon: "📷", name: "頚椎レントゲン", targets: "骨の変形・配列の異常", note: "まず最初のスクリーニング" },
    { icon: "🧠", name: "頭部MRI・MRA", targets: "椎骨動脈解離・くも膜下出血", note: "突然の痛みでは必須" },
    { icon: "🩸", name: "血液検査・腰椎穿刺", targets: "髄膜炎の診断", note: "発熱＋首の硬直がある場合" },
    { icon: "⚡", name: "神経伝導検査・筋電図", targets: "末梢神経の障害の評価", note: "しびれの原因鑑別" },
  ];
  tests.forEach(function(t, i) {
    var y = 1.05 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(t.icon, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.55, fontSize: 22, align: "center", margin: 0 });
    s.addText(t.name, { x: 1.4, y: y + 0.08, w: 2.8, h: 0.55, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.targets, { x: 4.3, y: y + 0.08, w: 4.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText(t.note, { x: 4.3, y: y + 0.4, w: 4.0, h: 0.25, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 11: Q&A
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある質問 Q&A");
  var qa = [
    { q: "ストレートネックが首の痛みの原因？", a: "ストレートネック自体が直接の痛みの原因とは限りませんが、首の筋肉に負担がかかりやすくなります。姿勢の改善とストレッチが基本です" },
    { q: "首をボキボキ鳴らすのは危険ですか？", a: "自分で鳴らす程度では問題ないことが多いですが、他人に強く首をひねるのは危険です。まれに椎骨動脈解離を起こし脳卒中のリスクがあります" },
    { q: "枕を変えれば首の痛みは治りますか？", a: "枕の高さは首の負担に影響しますが、万人に合う枕はありません。2週間以上続く場合は枕の問題ではなく別の原因を疑って受診しましょう" },
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
  hdr(s, "まとめ ― 首が痛いとき");
  var points = [
    { num: "1", text: "手のしびれ・脱力を\n伴う首の痛みに注意", sub: "ボタンが留めにくい等の\n巧緻運動障害は脊髄症のサイン" },
    { num: "2", text: "発熱＋首の硬直、\n突然の後頭部痛は緊急", sub: "髄膜炎・椎骨動脈解離・\nくも膜下出血→119番" },
    { num: "3", text: "2週間以上続く首の\n痛みは「肩こり」で終わらせない", sub: "脳神経内科・整形外科で\n原因を精査しましょう" },
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
    "🔍 首の痛みセルフチェックツール（概要欄にリンク）",
    "📺 次回：手足がしびれる ― 受診すべき科",
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

var outPath = __dirname + "/neck_pain_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
