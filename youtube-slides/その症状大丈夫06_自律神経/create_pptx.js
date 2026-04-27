var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#06】自律神経失調症と言われたら";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #06", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("🧠", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("「自律神経失調症」と\n言われたら？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.9, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("本当の原因と正しい受診先を\n脳神経内科医が解説", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.9,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #06", {
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
    { emoji: "😵‍💫", text: "めまいやふらつきが続く" },
    { emoji: "💓", text: "急に動悸がして不安になる" },
    { emoji: "😰", text: "理由のない倦怠感・だるさ" },
    { emoji: "🏥", text: "「自律神経失調症ですね」と言われた" },
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
// SLIDE 3: 自律神経とは？
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自律神経の役割 ― 体の自動操縦システム");

  card(s, 0.5, 1.1, 9.0, 0.8, C.light, C.primary);
  s.addText("自律神経 ＝ 自分の意思に関係なく体を自動調節する神経", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.6, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  // 交感神経 vs 副交感神経
  card(s, 0.4, 2.15, 4.4, 2.6, C.lightRed, C.red);
  s.addText("⚡ 交感神経（アクセル）", { x: 0.7, y: 2.25, w: 3.8, h: 0.5, fontSize: 20, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var sympathetic = [
    "心拍数を上げる",
    "血圧を上げる",
    "瞳孔を広げる",
    "消化を抑える",
    "緊張・興奮時に活発",
  ];
  sympathetic.forEach(function(t, i) {
    s.addText("・" + t, { x: 0.8, y: 2.85 + i * 0.35, w: 3.8, h: 0.3, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.2, 2.15, 4.4, 2.6, C.lightGreen, C.green);
  s.addText("🌿 副交感神経（ブレーキ）", { x: 5.5, y: 2.25, w: 3.8, h: 0.5, fontSize: 20, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  var parasympathetic = [
    "心拍数を下げる",
    "血圧を下げる",
    "消化を促進する",
    "体を修復・回復させる",
    "リラックス時に活発",
  ];
  parasympathetic.forEach(function(t, i) {
    s.addText("・" + t, { x: 5.6, y: 2.85 + i * 0.35, w: 3.8, h: 0.3, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  });

  s.addText("この2つのバランスが崩れると、さまざまな不調が現れます", {
    x: 0.5, y: 4.95, w: 9.0, h: 0.3, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: 自律神経失調症の正体
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「自律神経失調症」の正体");

  card(s, 0.5, 1.1, 9.0, 1.0, C.lightYellow, C.accent);
  s.addText("「自律神経失調症」は正式な病名ではない", {
    x: 0.8, y: 1.15, w: 8.5, h: 0.5, fontSize: 22, fontFace: FJ, color: C.accent, bold: true, margin: 0,
  });
  s.addText("検査で異常が見つからないときに使われる「仮の診断名」です", {
    x: 0.8, y: 1.65, w: 8.5, h: 0.35, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  var points = [
    { icon: "📋", title: "正式な疾患分類に存在しない", desc: "ICD（国際疾病分類）やDSM\n（精神疾患診断マニュアル）にない" },
    { icon: "🔍", title: "「原因不明」の仮ラベル", desc: "検査で異常なし→とりあえず\n自律神経失調症と説明される" },
    { icon: "⚠️", title: "本当の原因を見逃す危険", desc: "背後に治療が必要な病気が\n隠れている可能性がある" },
  ];

  points.forEach(function(p, i) {
    var y = 2.4 + i * 0.9;
    card(s, 0.5, y, 9.0, 0.75, C.white, C.primary);
    s.addText(p.icon, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.6, fontSize: 24, align: "center", margin: 0 });
    s.addText(p.title, { x: 1.5, y: y + 0.05, w: 3.5, h: 0.65, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(p.desc, { x: 5.2, y: y + 0.05, w: 4.2, h: 0.65, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 5: よくある症状
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自律神経の不調でよくある症状");

  var symptoms = [
    { icon: "😵‍💫", text: "めまい・立ちくらみ" },
    { icon: "💓", text: "動悸・胸のドキドキ" },
    { icon: "🤢", text: "吐き気・胃もたれ" },
    { icon: "🥵", text: "発汗異常（多汗・無汗）" },
    { icon: "😴", text: "不眠・過眠" },
    { icon: "🫠", text: "全身の倦怠感・だるさ" },
    { icon: "🌡️", text: "微熱・体温調節がうまくいかない" },
    { icon: "💩", text: "便秘・下痢の繰り返し" },
  ];

  symptoms.forEach(function(sym, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.1 + row * 0.95;
    card(s, x, y, 4.5, 0.78, C.white, C.primary);
    s.addText(sym.icon, { x: x + 0.2, y: y + 0.1, w: 0.6, h: 0.55, fontSize: 22, align: "center", margin: 0 });
    s.addText(sym.text, { x: x + 0.8, y: y + 0.1, w: 3.4, h: 0.55, fontSize: 19, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  s.addText("※ これらの症状は他の病気でも起こるため、「自律神経のせい」と決めつけないことが大切", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.25, fontSize: 13, fontFace: FJ, color: C.red, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 6: 本当の原因（見逃される病気）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「自律神経失調症」の裏に隠れた本当の原因", C.red);

  var diseases = [
    { name: "甲状腺疾患", desc: "動悸・発汗・体重変化\n血液検査で判明", color: C.accent, bg: C.lightYellow },
    { name: "貧血", desc: "めまい・だるさ・息切れ\n鉄欠乏が多い", color: C.accent, bg: C.lightYellow },
    { name: "うつ病・不安障害", desc: "倦怠感・不眠・食欲低下\n適切な治療で改善可能", color: C.primary, bg: C.light },
    { name: "起立性調節障害", desc: "立ち上がりでめまい\n若年者に多い", color: C.primary, bg: C.light },
    { name: "不整脈", desc: "動悸の原因が心臓にある\n心電図検査で確認", color: C.red, bg: C.lightRed },
    { name: "副腎不全", desc: "慢性的なだるさ・低血圧\nホルモン検査で判明", color: C.red, bg: C.lightRed },
  ];

  diseases.forEach(function(d, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.1 + row * 1.3;
    card(s, x, y, 4.5, 1.1, d.bg, d.color);
    s.addText(d.name, { x: x + 0.25, y: y + 0.08, w: 4.0, h: 0.4, fontSize: 19, fontFace: FJ, color: d.color, bold: true, margin: 0 });
    s.addText(d.desc, { x: x + 0.25, y: y + 0.5, w: 4.0, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 7: まず受けるべき検査
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まず受けるべき5つの検査");

  var tests = [
    { icon: "🩸", name: "血液検査", targets: "貧血・甲状腺機能・血糖値・肝腎機能・電解質", note: "基本中の基本" },
    { icon: "💓", name: "心電図", targets: "不整脈・虚血性心疾患", note: "動悸があれば必須" },
    { icon: "🩺", name: "起立試験（シェロンテスト）", targets: "起立性低血圧・起立性調節障害", note: "めまいがあれば必須" },
    { icon: "🧪", name: "尿検査", targets: "腎臓・副腎の異常", note: "スクリーニング" },
    { icon: "🧠", name: "頭部MRI（必要時）", targets: "脳幹・小脳の異常", note: "神経症状があれば" },
  ];

  tests.forEach(function(t, i) {
    var y = 1.05 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(t.icon, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.55, fontSize: 22, align: "center", margin: 0 });
    s.addText(t.name, { x: 1.4, y: y + 0.08, w: 3.0, h: 0.55, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.targets, { x: 4.5, y: y + 0.08, w: 4.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText(t.note, { x: 4.5, y: y + 0.4, w: 4.0, h: 0.25, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 8: 正しい受診先
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "正しい受診先 ― 症状で選ぶ診療科");

  var routes = [
    { symptom: "めまい・しびれ・歩行障害", dept: "脳神経内科", color: C.primary, bg: C.light },
    { symptom: "動悸・胸痛・息切れ", dept: "循環器内科", color: C.red, bg: C.lightRed },
    { symptom: "不眠・気分の落ち込み・不安", dept: "心療内科 / 精神科", color: C.accent, bg: C.lightYellow },
    { symptom: "倦怠感・体重変化・発汗異常", dept: "内分泌内科", color: C.green, bg: C.lightGreen },
    { symptom: "どこに行けばいいかわからない", dept: "まずは内科（かかりつけ医）", color: C.primary, bg: C.light },
  ];

  routes.forEach(function(r, i) {
    var y = 1.1 + i * 0.8;
    card(s, 0.5, y, 9.0, 0.65, r.bg, r.color);
    s.addText(r.symptom, { x: 0.8, y: y + 0.05, w: 5.0, h: 0.55, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText("→ " + r.dept, { x: 5.8, y: y + 0.05, w: 3.5, h: 0.55, fontSize: 18, fontFace: FJ, color: r.color, bold: true, valign: "middle", margin: 0 });
  });

  s.addText("※ 「自律神経失調症」と言われたら、他の原因を除外するための検査を受けたか確認しましょう", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.25, fontSize: 13, fontFace: FJ, color: C.red, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 9: 危険なサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな症状は見逃さないで ― 危険なサイン", C.red);

  var signs = [
    { text: "急に始まった激しいめまい＋嘔吐（脳卒中の可能性）", icon: "🚨" },
    { text: "失神した・意識を失った（不整脈・てんかんなど）", icon: "⚡" },
    { text: "体重が急に減っている（甲状腺・悪性腫瘍）", icon: "📉" },
    { text: "夜間に大量の汗をかく（感染症・リンパ腫）", icon: "🌡️" },
    { text: "動悸＋胸痛＋息苦しさ（心臓の病気）", icon: "💔" },
  ];

  signs.forEach(function(sign, i) {
    var y = 1.15 + i * 0.78;
    card(s, 0.5, y, 9.0, 0.65, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.05, w: 0.6, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.55, fontSize: 19, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.0, 5.0, 8.0, 0.2, C.red);
  ftr(s);
})();

// ============================================================
// SLIDE 10: セルフケア ― 自律神経を整える5つの習慣
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自律神経を整える5つの習慣");

  var tips = [
    { icon: "🌅", title: "規則正しい睡眠", desc: "毎日同じ時間に起きる。\n朝の光を浴びて体内時計をリセット" },
    { icon: "🏃", title: "適度な有酸素運動", desc: "ウォーキング・ヨガなど。\n週3回30分で自律神経のバランス改善" },
    { icon: "🛁", title: "ぬるめの入浴", desc: "38〜40℃で15分程度。\n副交感神経を優位にしてリラックス" },
    { icon: "🫁", title: "腹式呼吸", desc: "4秒吸って・7秒止めて・8秒吐く\n（4-7-8呼吸法）" },
    { icon: "📱", title: "デジタルデトックス", desc: "就寝1時間前からスマホOFF。\nブルーライトが交感神経を刺激する" },
  ];

  tips.forEach(function(t, i) {
    var y = 1.05 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(t.icon, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(t.title, { x: 1.4, y: y + 0.08, w: 2.6, h: 0.55, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.desc, { x: 4.1, y: y + 0.05, w: 5.2, h: 0.6, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.1, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 11: よくある質問
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある質問 Q&A");

  var qa = [
    { q: "自律神経失調症は治りますか？", a: "「自律神経失調症」自体は病名ではないため、原因を見つけて治療することが大切。背後の病気が治れば症状も改善します" },
    { q: "自律神経を鍛える方法はありますか？", a: "有酸素運動・腹式呼吸・規則的な生活で副交感神経の活性化が期待できます。即効性はなく、数週間の継続が必要です" },
    { q: "サプリメントは効きますか？", a: "ビタミンB群やマグネシウムが神経機能に関与しますが、まずは食事とバランスの良い生活が基本。効果を過信しないようにしましょう" },
  ];

  qa.forEach(function(item, i) {
    var y = 1.1 + i * 1.35;
    card(s, 0.5, y, 9.0, 1.15, C.white, C.accent);
    s.addText("Q.", { x: 0.8, y: y + 0.05, w: 0.5, h: 0.45, fontSize: 22, fontFace: FE, color: C.accent, bold: true, margin: 0 });
    s.addText(item.q, { x: 1.3, y: y + 0.05, w: 8.0, h: 0.45, fontSize: 18, fontFace: FJ, color: C.accent, bold: true, valign: "middle", margin: 0 });
    s.addText("A.", { x: 0.8, y: y + 0.55, w: 0.5, h: 0.5, fontSize: 22, fontFace: FE, color: C.primary, bold: true, margin: 0 });
    s.addText(item.a, { x: 1.3, y: y + 0.55, w: 8.0, h: 0.55, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.1, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 「自律神経失調症」と言われたら");

  var points = [
    { num: "1", text: "「自律神経失調症」は\n正式な病名ではない", sub: "検査で異常がないときの\n仮の診断名にすぎない" },
    { num: "2", text: "背後に隠れた病気を\n見逃さない", sub: "甲状腺・貧血・うつ病・不整脈\nなど、治療可能な原因を精査" },
    { num: "3", text: "症状に合った診療科を\n受診しよう", sub: "「自律神経だから仕方ない」\nで終わらせないことが大切" },
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
    "🔍 自律神経セルフチェックツール（概要欄にリンク）",
    "📺 次回：物が二重に見えたら要注意 ― 複視の原因と受診すべき科",
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
var outPath = __dirname + "/autonomic_nerves_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
