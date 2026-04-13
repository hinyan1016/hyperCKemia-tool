var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#10】手足がしびれる ― 受診すべき科";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #10", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🖐️", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("手足がしびれる", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.0,
    fontSize: 42, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("受診すべき科と原因の見分け方を\n脳神経内科医が解説", {
    x: 0.6, y: 2.7, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #10（最終回）", {
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
    { emoji: "🖐️", text: "朝起きたら手がしびれている" },
    { emoji: "🦶", text: "足の裏にジンジンする感覚が続く" },
    { emoji: "🤏", text: "指先の感覚が鈍い" },
    { emoji: "❓", text: "正座をしていないのに足がしびれる" },
    { emoji: "📱", text: "手がしびれて物を落としやすくなった" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: しびれの正体 ― 3つのタイプ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "しびれの正体 ― 3つのタイプ");
  var types = [
    { icon: "🤚", name: "感覚低下", desc: "触った感覚が鈍い\n温度がわかりにくい", color: C.primary },
    { icon: "⚡", name: "異常感覚", desc: "ジンジン・ピリピリ\nチクチクした不快な感覚", color: C.accent },
    { icon: "💪", name: "運動麻痺", desc: "力が入りにくい・動かしにくい\n「しびれる」と表現することがある", color: C.red },
  ];
  types.forEach(function(t, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, t.color);
    s.addText(t.icon, { x: 0.7, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 32, align: "center", margin: 0 });
    s.addText(t.name, { x: 1.5, y: y + 0.1, w: 2.5, h: 0.45, fontSize: 24, fontFace: FJ, color: t.color, bold: true, valign: "middle", margin: 0 });
    s.addText(t.desc, { x: 4.2, y: y + 0.1, w: 5.0, h: 0.85, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });
  s.addText("※ 受診時に「どのタイプのしびれか」を伝えると原因特定の大きな手がかりに", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });
  ftr(s);
})();

// SLIDE 4: しびれの原因 ― どこが問題？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "しびれの原因 ― どこが問題？");
  var locs = [
    { icon: "🧠", name: "脳", pattern: "片側の手足がしびれる", example: "脳卒中・脳腫瘍", dept: "脳神経内科", color: C.red },
    { icon: "🦴", name: "脊髄", pattern: "両手や両足がしびれる\n＋歩行障害・巧緻運動障害", example: "頚椎症性脊髄症・脊髄腫瘍", dept: "脳神経内科・整形外科", color: C.primary },
    { icon: "⚡", name: "末梢神経", pattern: "手袋・靴下の範囲\n特定の指だけ", example: "糖尿病性神経障害・手根管症候群", dept: "脳神経内科", color: C.accent },
  ];
  locs.forEach(function(l, i) {
    var y = 1.1 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, l.color);
    s.addText(l.icon, { x: 0.7, y: y + 0.15, w: 0.7, h: 0.8, fontSize: 30, align: "center", margin: 0 });
    s.addText(l.name, { x: 1.5, y: y + 0.05, w: 1.2, h: 0.5, fontSize: 24, fontFace: FJ, color: l.color, bold: true, valign: "middle", margin: 0 });
    s.addText(l.pattern, { x: 2.8, y: y + 0.05, w: 3.2, h: 0.55, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText(l.example, { x: 2.8, y: y + 0.6, w: 3.2, h: 0.4, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
    s.addText(l.dept, { x: 6.5, y: y + 0.2, w: 2.8, h: 0.7, fontSize: 15, fontFace: FJ, color: l.color, bold: true, valign: "middle", align: "center", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 心配度で分類
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "しびれの原因 ― 心配度で分類");
  var causes = [
    { level: "🟢 軽度", name: "一時的な血行不良", desc: "正座や腕枕。姿勢を変えればすぐ改善", bg: C.lightGreen, color: C.green },
    { level: "🟡 要注意", name: "手根管症候群", desc: "親指〜薬指のしびれ。夜間〜朝に悪化", bg: C.lightYellow, color: "856404" },
    { level: "🟡 要注意", name: "糖尿病性神経障害", desc: "足先から手袋靴下型に広がる。患者の約半数に合併", bg: C.lightYellow, color: "856404" },
    { level: "🟡 要注意", name: "ビタミンB12欠乏症", desc: "両足のしびれ＋歩行障害。菜食・胃手術後に多い", bg: C.lightYellow, color: "856404" },
    { level: "🔴 緊急", name: "脳卒中", desc: "突然の片側のしびれ。ろれつ障害・脱力を伴う", bg: C.lightRed, color: C.red },
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

// SLIDE 6: 見逃されやすい原因 ― 末梢神経障害
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "見逃されやすい原因 ― 末梢神経障害");
  var diseases = [
    { icon: "🩸", name: "糖尿病性神経障害", desc: "足先から始まり徐々に上がる「手袋靴下型」。\n糖尿病患者の約半数に合併。早期の血糖管理が重要", color: C.accent },
    { icon: "🤚", name: "手根管症候群", desc: "夜間〜朝に手がしびれ、振ると楽になる。\n親指〜薬指の半分がしびれるのが典型的", color: C.primary },
    { icon: "🦠", name: "ギラン・バレー症候群", desc: "感染症の後に急速に四肢のしびれと脱力が進行。\n入院治療が必要な緊急疾患", color: C.red },
    { icon: "💊", name: "ビタミンB12欠乏症", desc: "菜食主義者や胃の手術後に多い。\n両足のしびれから始まり歩行障害に進行", color: "7B2D8E" },
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

// SLIDE 7: セルフチェック
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自分でできるチェック ― しびれの分布");
  var checks = [
    { num: "1", title: "しびれの分布を確認", desc: "片側だけ→脳 / 手袋靴下型→末梢神経\n両手＋歩行障害→脊髄 / 特定の指→絞扼性", icon: "🗺️" },
    { num: "2", title: "いつ悪化するかを確認", desc: "夜間〜朝→手根管症候群\n長く歩くと悪化→脊柱管狭窄症", icon: "🕐" },
    { num: "3", title: "しびれ以外の症状を確認", desc: "脱力→神経障害の進行 / 歩行障害→脊髄\n足の傷が治りにくい→糖尿病性", icon: "📋" },
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

// SLIDE 8: 危険なサイン
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな場合はすぐ受診 ― 危険なサイン", C.red);
  var signs = [
    { text: "突然の片側のしびれ＋ろれつ障害・脱力（脳卒中の可能性）", icon: "🧠" },
    { text: "両手のしびれ＋歩行障害の進行（頚椎症性脊髄症・脊髄腫瘍）", icon: "🦴" },
    { text: "感染後に急速にしびれと脱力が広がる（ギラン・バレー症候群）", icon: "🦠" },
    { text: "しびれ＋排尿障害がある（脊髄の圧迫）", icon: "⚠️" },
    { text: "足のしびれ＋傷が治りにくい（糖尿病性神経障害の進行）", icon: "🩹" },
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
    { emoji: "🔴", label: "すぐ受診", desc: "突然の片側のしびれ＋ろれつ・脱力\n急速にしびれと脱力が広がっている", action: "脳卒中・ギラン・バレー\n→119番", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "数週間以上続くしびれ\n巧緻運動障害・歩行障害を伴う\nしびれの範囲が広がっている", action: "脳神経内科で原因精査\n神経伝導検査・MRI", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "正座や腕枕の後の一時的なしびれ\n姿勢を変えればすぐ改善", action: "繰り返さなければ\n心配不要", bg: C.lightGreen, color: C.green },
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
  hdr(s, "しびれを調べる検査");
  var tests = [
    { icon: "⚡", name: "神経伝導検査", targets: "末梢神経の伝導速度を測定", note: "手根管症候群・末梢神経障害の診断に最重要" },
    { icon: "🦴", name: "頚椎MRI", targets: "脊髄の圧迫を評価", note: "両手のしびれ・歩行障害に必須" },
    { icon: "🧠", name: "頭部MRI", targets: "脳卒中・脳腫瘍の評価", note: "突然の片側のしびれに必須" },
    { icon: "🩸", name: "血液検査", targets: "糖尿病・ビタミンB12・甲状腺", note: "末梢神経障害の原因検索" },
    { icon: "📊", name: "筋電図", targets: "筋肉の電気活動を記録", note: "末梢神経障害の程度評価" },
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
    { q: "しびれは何科に行けばいいですか？", a: "まず脳神経内科をおすすめします。脳・脊髄・末梢神経のすべてを診る科で、原因がどこにあるか総合的に判断できます。必要に応じて整形外科等に紹介します" },
    { q: "糖尿病のしびれは治りますか？", a: "早期なら血糖コントロール改善で進行を抑えられます。進行した神経障害は回復しにくいため、早期発見・早期治療が重要です" },
    { q: "手根管症候群は手術が必要ですか？", a: "軽度なら手首の安静やスプリント装着で改善することがあります。筋萎縮が進んでいる場合は手術を検討します" },
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
  hdr(s, "まとめ ― 手足がしびれたら");
  var points = [
    { num: "1", text: "しびれの分布パターンが\n原因を示唆する", sub: "片側→脳 / 手袋靴下型→末梢\n両手＋歩行障害→脊髄" },
    { num: "2", text: "突然の片側のしびれは\n脳卒中の可能性", sub: "ろれつ障害・脱力を伴う場合\nすぐに119番" },
    { num: "3", text: "しびれは脳神経内科が\n総合的に診断できる", sub: "脳・脊髄・末梢神経の\nどこが問題かを判断" },
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

// SLIDE 13: エンディング（シリーズ最終回）
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("シリーズ全10回\nご視聴ありがとうございました", {
    x: 0.6, y: 0.5, w: 8.8, h: 1.0,
    fontSize: 28, fontFace: FJ, color: C.white, bold: true, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 1.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("チャンネル登録・高評価よろしくお願いします！", {
    x: 0.6, y: 1.9, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  var topics = [
    "#01 手のふるえ  #02 ろれつ  #03 脱力  #04 顔のピクつき  #05 足がつる",
    "#06 自律神経  #07 複視  #08 嚥下障害  #09 首の痛み  #10 しびれ",
  ];
  topics.forEach(function(t, i) {
    s.addText(t, {
      x: 0.6, y: 2.5 + i * 0.4, w: 8.8, h: 0.35,
      fontSize: 13, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
    });
  });
  s.addText("🔗 ブログ記事で詳しく読む（概要欄にリンク）", {
    x: 1.5, y: 3.5, w: 7.0, h: 0.4, fontSize: 16, fontFace: FJ, color: "B0BEC5", margin: 0,
  });
  s.addText("🔍 セルフチェックツール（概要欄にリンク）", {
    x: 1.5, y: 3.95, w: 7.0, h: 0.4, fontSize: 16, fontFace: FJ, color: "B0BEC5", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/numbness_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
