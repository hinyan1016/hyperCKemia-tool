var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#08】むせやすくなったは老化のせい？";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #08", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("🍵", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("「むせやすくなった」は\n老化のせい？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.9, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("嚥下障害の早期サインと受診の目安を\n脳神経内科医が解説", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.9,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #08", {
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
    { emoji: "😳", text: "食事中にむせることが増えた" },
    { emoji: "🍵", text: "お茶や水でむせやすい" },
    { emoji: "💊", text: "薬の錠剤が飲みにくくなった" },
    { emoji: "🗣️", text: "食後に声がかすれる" },
    { emoji: "⏱️", text: "食事に時間がかかるようになった" },
  ];

  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 3: 嚥下のしくみ ― 飲み込みの5段階
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「嚥下」のしくみ ― 飲み込みの5段階");

  var stages = [
    { num: "1", name: "先行期", desc: "食べ物を目で見て認識し、硬さや量を判断", icon: "👀" },
    { num: "2", name: "準備期", desc: "かんで唾液と混ぜ、飲み込みやすい塊にする", icon: "🦷" },
    { num: "3", name: "口腔期", desc: "舌を使って食べ物をのどに送り込む", icon: "👅" },
    { num: "4", name: "咽頭期", desc: "のどの反射で食道に送る。喉頭蓋がフタをする", icon: "🔒" },
    { num: "5", name: "食道期", desc: "食道のぜん動運動で胃に運ぶ", icon: "⬇️" },
  ];

  stages.forEach(function(st, i) {
    var y = 1.05 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(st.num, { x: 0.7, y: y + 0.08, w: 0.5, h: 0.55, fontSize: 28, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.icon, { x: 1.3, y: y + 0.08, w: 0.5, h: 0.55, fontSize: 22, align: "center", margin: 0 });
    s.addText(st.name, { x: 1.9, y: y + 0.08, w: 1.6, h: 0.55, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(st.desc, { x: 3.6, y: y + 0.08, w: 5.5, h: 0.55, fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: 嚥下障害の原因 ― 心配度で分類
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "嚥下障害の原因 ― 心配度で分類");

  var causes = [
    { level: "🟢 軽度", name: "加齢による筋力低下", desc: "のどの筋肉や反射の衰え。ゆっくり食べることで対応", bg: C.lightGreen, color: C.green },
    { level: "🟡 要注意", name: "逆流性食道炎", desc: "胃酸の逆流でのどに違和感。消化器内科で治療", bg: C.lightYellow, color: "856404" },
    { level: "🟡 要注意", name: "薬の副作用", desc: "抗精神病薬・抗コリン薬などが嚥下機能を低下", bg: C.lightYellow, color: "856404" },
    { level: "🔴 要精査", name: "脳卒中", desc: "脳幹の脳卒中で嚥下障害が主症状になることがある", bg: C.lightRed, color: C.red },
    { level: "🔴 要精査", name: "パーキンソン病・MG・ALS", desc: "神経変性疾患や自己免疫疾患。初期症状のことも", bg: C.lightRed, color: C.red },
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

// ============================================================
// SLIDE 5: 見逃されやすい原因 ― 脳神経の病気
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "見逃されやすい原因 ― 脳神経の病気");

  var diseases = [
    { icon: "🧠", name: "脳卒中", desc: "脳幹の脳卒中で嚥下障害が主症状に。\nワレンベルグ症候群ではめまい＋声がれ＋嚥下障害の3徴", color: C.red },
    { icon: "🤲", name: "パーキンソン病", desc: "初期から嚥下機能が低下していることが多い。\n誤嚥性肺炎はパーキンソン病の主な死因の一つ", color: C.accent },
    { icon: "💪", name: "重症筋無力症", desc: "食事の後半にむせが増えるのが特徴。\n眼症状に次いで多い症状", color: C.primary },
    { icon: "⚡", name: "ALS（筋萎縮性側索硬化症）", desc: "球麻痺型では嚥下障害が初期症状。\n飲み込みにくさや声のかすれで発症", color: "7B2D8E" },
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

// ============================================================
// SLIDE 6: 高齢者に多い誤嚥性肺炎
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "高齢者に多い誤嚥性肺炎");

  card(s, 0.5, 1.1, 9.0, 0.8, C.lightRed, C.red);
  s.addText("誤嚥 ＝ 食べ物や唾液が気管に入ってしまうこと", {
    x: 0.8, y: 1.15, w: 8.5, h: 0.35, fontSize: 19, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });
  s.addText("高齢者の肺炎死亡の多くは誤嚥性肺炎が原因", {
    x: 0.8, y: 1.55, w: 8.5, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  var points = [
    { icon: "🤫", title: "不顕性誤嚥に注意", desc: "むせずに誤嚥する「サイレント\nアスピレーション」が最も危険" },
    { icon: "🔥", title: "繰り返す発熱はサイン", desc: "原因不明の肺炎や発熱が\n繰り返すなら嚥下障害を疑う" },
    { icon: "🪥", title: "口腔ケアで予防", desc: "口の中の細菌を減らすことで\n誤嚥しても肺炎リスクを軽減" },
    { icon: "🍽️", title: "食事の工夫", desc: "食べやすい形態の食事と\n適切なとろみで予防できる" },
  ];

  points.forEach(function(p, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 2.15 + row * 1.4;
    card(s, x, y, 4.5, 1.2, C.white, C.primary);
    s.addText(p.icon, { x: x + 0.2, y: y + 0.08, w: 0.6, h: 0.5, fontSize: 22, align: "center", margin: 0 });
    s.addText(p.title, { x: x + 0.8, y: y + 0.08, w: 3.4, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(p.desc, { x: x + 0.8, y: y + 0.5, w: 3.4, h: 0.6, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 7: セルフチェック ― 3つのサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自分でできるチェック ― 3つのサイン");

  var checks = [
    { num: "1", title: "反復唾液嚥下テスト", desc: "30秒間にできるだけ何回もつばを飲み込む\n3回以上→正常 / 2回以下→要注意", icon: "💧" },
    { num: "2", title: "水飲みテスト", desc: "30ml（大さじ2杯）の水をコップで飲む\nむせたり声がかすれたりしたら要注意", icon: "🥛" },
    { num: "3", title: "食事の観察サイン", desc: "むせが増えた / 食後に声が湿った感じ\n食事に時間がかかるようになった", icon: "🍽️" },
  ];

  checks.forEach(function(ch, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, C.primary);
    s.addText(ch.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.85, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(ch.icon, { x: 1.5, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addText(ch.title, { x: 2.3, y: y + 0.1, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(ch.desc, { x: 2.3, y: y + 0.55, w: 6.5, h: 0.45, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
  });

  s.addText("※ あくまで参考です。気になる場合は医師にご相談ください", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 8: 危険なサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな場合はすぐ受診 ― 危険なサイン", C.red);

  var signs = [
    { text: "突然の嚥下障害＋めまい・ろれつ障害（脳卒中の可能性）", icon: "🚨" },
    { text: "繰り返す発熱・原因不明の肺炎（不顕性誤嚥の可能性）", icon: "🔥" },
    { text: "嚥下障害＋体重が急に減っている（食道がん等の可能性）", icon: "📉" },
    { text: "嚥下障害＋しゃべりにくさ・筋力低下の進行（ALS等の可能性）", icon: "⚡" },
    { text: "食べ物がのどに完全に詰まって呼吸が苦しい（窒息の危険）", icon: "🫁" },
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

// ============================================================
// SLIDE 9: 受診の目安 ― 緊急度3段階
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");

  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "突然の嚥下障害＋めまい・ろれつ障害\n食べ物がのどに詰まり呼吸困難", action: "脳卒中・窒息の可能性\n→119番", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "数週間以上続く / 体重減少\n繰り返す発熱 / まぶた下垂・筋力低下", action: "脳神経内科・耳鼻科を\n受診して精査", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "たまにむせる程度\nゆっくり食べると問題ない", action: "悪化したら受診\n口腔ケアで予防", bg: C.lightGreen, color: C.green },
  ];

  levels.forEach(function(lv, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, lv.bg, lv.color);
    s.addText(lv.emoji + " " + lv.label, { x: 0.8, y: y + 0.05, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    s.addText(lv.desc, { x: 0.8, y: y + 0.5, w: 5.0, h: 0.55, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText(lv.action, { x: 6.0, y: y + 0.15, w: 3.3, h: 0.8, fontSize: 16, fontFace: FJ, color: lv.color, bold: true, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 10: 嚥下障害の検査
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "嚥下障害の検査");

  var tests = [
    { icon: "📹", name: "嚥下内視鏡検査（VE）", targets: "鼻から内視鏡を入れて飲み込みを観察", note: "最も重要な検査の一つ" },
    { icon: "🔬", name: "嚥下造影検査（VF）", targets: "バリウム入り食品をX線で撮影", note: "嚥下の全過程を評価" },
    { icon: "🧠", name: "頭部MRI", targets: "脳卒中・脳腫瘍の評価", note: "突然発症では必須" },
    { icon: "🩸", name: "血液検査", targets: "重症筋無力症の抗体・甲状腺・栄養状態", note: "原因の絞り込み" },
    { icon: "🔍", name: "上部消化管内視鏡", targets: "食道の狭窄・腫瘍の評価", note: "食べ物が引っかかる場合" },
  ];

  tests.forEach(function(t, i) {
    var y = 1.05 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(t.icon, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.55, fontSize: 22, align: "center", margin: 0 });
    s.addText(t.name, { x: 1.4, y: y + 0.08, w: 3.0, h: 0.55, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.targets, { x: 4.5, y: y + 0.08, w: 3.8, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText(t.note, { x: 4.5, y: y + 0.4, w: 3.8, h: 0.25, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
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
    { q: "むせやすいのは年のせいですか？", a: "加齢で嚥下機能は低下しますが「年のせい」で片付けるのは危険です。脳卒中やパーキンソン病が隠れている可能性があり、悪化傾向があれば受診をおすすめします" },
    { q: "嚥下障害は治りますか？", a: "原因によります。脳卒中後はリハビリで改善、重症筋無力症は薬で改善。パーキンソン病は進行を遅らせる治療があります。原因の特定と早期介入が重要です" },
    { q: "とろみをつければ安心ですか？", a: "水分にとろみをつけることは誤嚥予防に有効ですが、全員に同じとろみが適しているわけではありません。嚥下評価で適切な程度を確認することが大切です" },
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

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― むせやすくなったら");

  var points = [
    { num: "1", text: "「むせ」を年のせいに\nしない", sub: "脳卒中・パーキンソン病・\nMG・ALSが隠れている可能性" },
    { num: "2", text: "突然の嚥下障害＋\nめまい・ろれつ障害は緊急", sub: "脳卒中の可能性\nすぐに119番" },
    { num: "3", text: "繰り返す発熱は\n不顕性誤嚥のサイン", sub: "嚥下内視鏡検査（VE）で\n嚥下機能を評価してもらおう" },
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
    "🔍 嚥下セルフチェックツール（概要欄にリンク）",
    "📺 次回：首が痛い ― 原因と見分け方",
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
var outPath = __dirname + "/dysphagia_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
