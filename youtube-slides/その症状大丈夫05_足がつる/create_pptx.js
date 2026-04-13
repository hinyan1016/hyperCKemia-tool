var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#05】寝ている時に足がつる";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #05", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("🦵", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("寝ている時に足がつる…\nこれって危ないサイン？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.9, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("脳神経内科医が教える\n原因と危険なサイン", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.9,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #05", {
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
    { emoji: "🌙", text: "寝ている最中にふくらはぎが突然つる" },
    { emoji: "😖", text: "痛みで目が覚めて動けない" },
    { emoji: "🏃", text: "運動の後に足がつりやすい" },
    { emoji: "🔁", text: "最近、頻繁に足がつるようになった" },
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
// SLIDE 3: 足がつる仕組み
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "足がつる（こむら返り）とは？");

  card(s, 0.5, 1.1, 9.0, 1.0, C.light, C.primary);
  s.addText("筋肉が自分の意志とは関係なく、突然強く収縮して痛みが出る現象", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  var steps = [
    { num: "1", title: "神経の異常興奮", desc: "筋肉を動かす神経が\n何らかの原因で過剰に興奮" },
    { num: "2", title: "筋肉の強直性収縮", desc: "筋肉が一気に縮み\nリラックスできなくなる" },
    { num: "3", title: "激しい痛み", desc: "数秒〜数分間\n強い痛みが続く" },
  ];

  steps.forEach(function(st, i) {
    var x = 0.3 + i * 3.2;
    card(s, x, 2.35, 3.0, 2.5, C.white, C.primary);
    s.addText(st.num, { x: x + 0.2, y: 2.45, w: 0.6, h: 0.55, fontSize: 28, fontFace: FE, color: C.primary, bold: true, align: "center", margin: 0 });
    s.addText(st.title, { x: x + 0.8, y: 2.45, w: 2.0, h: 0.55, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addShape(pres.shapes.LINE, { x: x + 0.3, y: 3.1, w: 2.4, h: 0, line: { color: C.primary, width: 1 } });
    s.addText(st.desc, { x: x + 0.2, y: 3.25, w: 2.6, h: 0.8, fontSize: 17, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  s.addText("※ 医学用語では「有痛性筋けいれん（muscle cramp）」と呼びます", {
    x: 0.5, y: 4.95, w: 9.0, h: 0.3, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: よくある原因（心配不要タイプ）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある原因 ― ほとんどは心配不要", C.green);

  var causes = [
    { icon: "💧", title: "脱水・水分不足", desc: "汗をかいた後や就寝中に\n体内の水分が減る", bg: C.lightGreen, color: C.green },
    { icon: "⚡", title: "電解質バランスの乱れ", desc: "マグネシウム・カリウム・\nカルシウムの不足", bg: C.lightGreen, color: C.green },
    { icon: "🏋️", title: "筋肉の疲労", desc: "運動後や長時間の立ち仕事\n筋肉への負荷が原因", bg: C.lightGreen, color: C.green },
    { icon: "🥶", title: "冷え", desc: "冬場やクーラーの効いた部屋\n血行不良で筋肉が緊張", bg: C.lightGreen, color: C.green },
    { icon: "👴", title: "加齢", desc: "50歳以上で頻度が増加\n筋肉量の減少が一因", bg: C.lightGreen, color: C.green },
    { icon: "🤰", title: "妊娠", desc: "妊娠後期に多い\nホルモン・体重変化の影響", bg: C.lightGreen, color: C.green },
  ];

  causes.forEach(function(c, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.1 + row * 1.35;
    card(s, x, y, 4.5, 1.15, c.bg, c.color);
    s.addText(c.icon, { x: x + 0.2, y: y + 0.08, w: 0.6, h: 0.5, fontSize: 24, margin: 0 });
    s.addText(c.title, { x: x + 0.8, y: y + 0.08, w: 3.4, h: 0.4, fontSize: 18, fontFace: FJ, color: c.color, bold: true, margin: 0 });
    s.addText(c.desc, { x: x + 0.8, y: y + 0.5, w: 3.4, h: 0.55, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 5: 薬の副作用
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "意外と多い ― 薬の副作用による足のつり");

  card(s, 0.5, 1.1, 9.0, 0.8, C.lightYellow, C.accent);
  s.addText("普段飲んでいる薬が原因で足がつりやすくなることがあります", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.6, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, valign: "middle", margin: 0,
  });

  var drugs = [
    { name: "利尿薬", desc: "高血圧・むくみの薬", reason: "カリウム・マグネシウムが\n尿と一緒に排出される" },
    { name: "スタチン", desc: "コレステロールの薬", reason: "筋肉へのダメージ\n（横紋筋融解症に注意）" },
    { name: "降圧薬（ARB/ACE）", desc: "血圧を下げる薬", reason: "電解質バランスの変化" },
    { name: "吸入ステロイド", desc: "喘息・COPDの薬", reason: "カリウム低下の可能性" },
  ];

  drugs.forEach(function(d, i) {
    var y = 2.15 + i * 0.73;
    card(s, 0.5, y, 9.0, 0.6, C.white, C.primary);
    s.addText(d.name, { x: 0.8, y: y + 0.05, w: 2.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(d.desc, { x: 2.8, y: y + 0.05, w: 2.5, h: 0.5, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
    s.addText(d.reason, { x: 5.5, y: y + 0.02, w: 3.8, h: 0.55, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.1, margin: 0 });
  });

  s.addText("※ 自己判断で薬を中止せず、主治医にご相談ください", {
    x: 0.5, y: 4.95, w: 9.0, h: 0.3, fontSize: 14, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 6: 脳神経内科医が注目する危険な原因
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳神経内科医が注目する ― 隠れた病気", C.red);

  var diseases = [
    { name: "末梢神経障害", sub: "（糖尿病性など）", desc: "しびれや感覚低下を伴う\n糖尿病の合併症で最多", color: C.accent, bg: C.lightYellow },
    { name: "腰部脊柱管狭窄症", sub: "（せきちゅうかんきょうさくしょう）", desc: "歩くと足がしびれ・つる\n休むと楽になる", color: C.accent, bg: C.lightYellow },
    { name: "ALS（筋萎縮性側索硬化症）", sub: "", desc: "筋力低下・筋萎縮が進行\n足のつりが初期症状のことも", color: C.red, bg: C.lightRed },
    { name: "甲状腺機能低下症", sub: "", desc: "むくみ・疲労感・冷え\n代謝低下による筋けいれん", color: C.primary, bg: C.light },
  ];

  diseases.forEach(function(d, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.1 + row * 2.0;
    card(s, x, y, 4.5, 1.75, d.bg, d.color);
    s.addText(d.name, { x: x + 0.25, y: y + 0.1, w: 4.0, h: 0.45, fontSize: 18, fontFace: FJ, color: d.color, bold: true, margin: 0 });
    if (d.sub) {
      s.addText(d.sub, { x: x + 0.25, y: y + 0.55, w: 4.0, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
    }
    var descY = d.sub ? y + 0.85 : y + 0.6;
    s.addText(d.desc, { x: x + 0.25, y: descY, w: 4.0, h: 0.7, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 7: 病気が隠れている場合の特徴
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな「つり方」は要注意");

  var signs = [
    { text: "週に2回以上、繰り返し足がつる", icon: "🔁" },
    { text: "足のつりと一緒に、しびれや感覚の鈍さがある", icon: "🦶" },
    { text: "筋肉が痩せてきた（筋萎縮）", icon: "📉" },
    { text: "ふくらはぎ以外（太もも・手・全身）もつる", icon: "🫳" },
    { text: "歩行中につる＋休むと治る（間欠性跛行）", icon: "🚶" },
  ];

  signs.forEach(function(sign, i) {
    var y = 1.15 + i * 0.78;
    card(s, 0.5, y, 9.0, 0.65, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.05, w: 0.6, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.55, fontSize: 20, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.0, 5.0, 8.0, 0.2, C.red);
  ftr(s);
})();

// ============================================================
// SLIDE 8: 受診の目安 ― 緊急度3段階
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");

  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "足のつり＋筋力低下が急速に進行\n足がパンパンに腫れて熱を持つ", action: "神経疾患・深部静脈\n血栓症の可能性→救急", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "週2回以上つる / しびれを伴う\n筋肉が痩せてきた気がする", action: "脳神経内科を受診\n原因精査が必要", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "たまにつる程度（月1〜2回）\n疲労・運動後に限定される", action: "セルフケアで対応\n悪化したら受診", bg: C.lightGreen, color: C.green },
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
// SLIDE 9: つった時の対処法
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "足がつった！ ― その場でできる対処法");

  var steps = [
    { num: "1", title: "つま先を手前に引く", desc: "つった筋肉をゆっくり\nストレッチする", icon: "🦶", color: C.primary },
    { num: "2", title: "ゆっくりマッサージ", desc: "痛みが和らいだら\n筋肉を軽くほぐす", icon: "👐", color: C.primary },
    { num: "3", title: "温める", desc: "蒸しタオルや温水で\n血行を促進する", icon: "🔥", color: C.accent },
  ];

  steps.forEach(function(st, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, st.color);
    s.addText(st.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.85, fontSize: 36, fontFace: FE, color: st.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.icon, { x: 1.5, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addText(st.title, { x: 2.3, y: y + 0.1, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: st.color, bold: true, valign: "middle", margin: 0 });
    s.addText(st.desc, { x: 2.3, y: y + 0.55, w: 6.5, h: 0.45, fontSize: 17, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });

  s.addText("⚠️ 無理に伸ばすと肉離れの危険！ ゆっくり丁寧に", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 10: 予防法
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "足のつりを予防する5つの習慣");

  var tips = [
    { icon: "💧", title: "水分をこまめに摂る", desc: "就寝前にコップ1杯の水。\n1日1.5〜2Lが目安" },
    { icon: "🥬", title: "ミネラルを意識した食事", desc: "マグネシウム（ナッツ・海藻）\nカリウム（バナナ・ほうれん草）" },
    { icon: "🧘", title: "就寝前のストレッチ", desc: "ふくらはぎを30秒ずつ伸ばす。\n毎晩の習慣にする" },
    { icon: "🧦", title: "足を冷やさない", desc: "レッグウォーマーや靴下で保温。\n夏のクーラーにも注意" },
    { icon: "🚶", title: "適度な運動", desc: "ウォーキング30分/日。\n筋力維持が予防の基本" },
  ];

  tips.forEach(function(t, i) {
    var y = 1.05 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(t.icon, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(t.title, { x: 1.4, y: y + 0.08, w: 2.8, h: 0.55, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.desc, { x: 4.3, y: y + 0.05, w: 5.0, h: 0.6, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.1, margin: 0 });
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
    { q: "芍薬甘草湯（しゃくやくかんぞうとう）は効きますか？", a: "即効性があり有効ですが、長期連用は低カリウム血症のリスク。頓用で使いましょう" },
    { q: "足がつるのは脳梗塞の前兆ですか？", a: "足のつり単独では脳梗塞を疑う根拠にはなりません。片側の麻痺やろれつ障害を伴う場合は別です" },
    { q: "サプリメントでマグネシウムを摂ってもいい？", a: "食事からの摂取が基本ですが、不足気味なら補助的に有効です。腎臓の病気がある方は医師に相談を" },
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
  hdr(s, "まとめ ― 足がつる時に覚えておきたい3つ");

  var points = [
    { num: "1", text: "たまにつる程度は\nほとんど心配不要", sub: "脱水・疲労・冷えが主な原因\n生活習慣の見直しで改善" },
    { num: "2", text: "頻繁につる＋しびれは\n病気のサインかも", sub: "末梢神経障害・脊柱管狭窄症\n脳神経内科で精査を" },
    { num: "3", text: "薬の副作用も\nチェックしよう", sub: "利尿薬・スタチンなど\n主治医に相談を" },
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
    "🔍 足のつりセルフチェックツール（概要欄にリンク）",
    "📺 次回：「自律神経失調症」と言われたら？ ― 本当の原因と受診先",
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
var outPath = __dirname + "/leg_cramps_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
