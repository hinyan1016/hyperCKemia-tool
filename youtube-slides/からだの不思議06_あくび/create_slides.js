var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "なぜあくびはうつるのか？ ― ミラーニューロンと共感の脳科学";

// --- Color Palette ---
var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "4A90D9",
  light: "E8F0FE",
  warmBg: "F5F7FA",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "DC3545",
  orange: "E8850C",
  yellow: "F5C518",
  green: "28A745",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var FONT_H = FONT_JP;
var FONT_B = FONT_JP;

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
}

function addCard(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: C.white }, shadow: makeShadow() });
}

function addPageNum(s, num) {
  s.addText(num, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_B, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("脳神経内科医が答える からだの不思議 #06", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("なぜあくびはうつるのか？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.2,
    fontSize: 38, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― ミラーニューロンと共感の脳科学 ―", {
    x: 0.8, y: 2.5, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.3, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.6, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: こんな経験ありませんか？（もらいあくびの不思議）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんな経験ありませんか？");

  var quotes = [
    "会議中に誰かがあくびをした瞬間、\n自分もあくびが出てしまった",
    "「あくび」という文字を読んだだけで\n口が開きかけた",
    "目が見えていなくても、あくびの音を聞いただけで\nうつることがある",
  ];
  quotes.forEach(function(q, i) {
    var yPos = 1.2 + i * 1.35;
    addCard(s, 1.0, yPos, 8.0, 1.15);
    s.addText(q, {
      x: 1.3, y: yPos + 0.1, w: 7.4, h: 0.95,
      fontSize: 18, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
      lineSpacingMultiple: 1.2,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 5.0, w: 8.0, h: 0.5, fill: { color: C.light } });
  s.addText("成人の約50〜60%が他者のあくびを見て60秒以内にもらいあくびをする", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: ミラーニューロン系の役割
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳の中で何が起きているか ― ミラーニューロン");

  // 定義カード
  addCard(s, 0.5, 1.1, 9.0, 1.4);
  s.addText("ミラーニューロン（Mirror Neuron）", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_EN, color: C.primary, bold: true, margin: 0,
  });
  s.addText("「自分が動作をするときだけでなく、他者が同じ動作をするのを\n観察しただけでも活動する神経細胞」（1990年代にサルの研究から発見）", {
    x: 0.8, y: 1.65, w: 8.4, h: 0.7,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.25,
  });

  // 3つの関連脳領域カード
  var regions = [
    { name: "下前頭回", role: "動作の模倣\n観察・実行", icon: "🔁" },
    { name: "頭頂葉", role: "空間的な\n動作処理", icon: "📍" },
    { name: "補足運動野", role: "運動プログラム\nの活性化", icon: "⚡" },
  ];
  regions.forEach(function(r, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 2.75, 3.0, 2.0);
    s.addText(r.icon, { x: xPos, y: 2.85, w: 3.0, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(r.name, { x: xPos, y: 3.4, w: 3.0, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0 });
    s.addText(r.role, { x: xPos, y: 3.9, w: 3.0, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
  });

  s.addText("→", { x: 3.1, y: 3.2, w: 0.5, h: 0.5, fontSize: 26, color: C.accent, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 6.3, y: 3.2, w: 0.5, h: 0.5, fontSize: 26, color: C.accent, align: "center", valign: "middle", margin: 0 });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: 共感との深いつながり
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "共感との深いつながり");

  // 共感回路の2領域カード
  var regions = [
    { name: "前帯状皮質（ACC）", role: "他者の感情状態を\n自分のものとして処理\n（情動的共感の中枢）", color: C.primary, icon: "💙" },
    { name: "島皮質（insula）", role: "身体感覚と感情を統合\n社会的・共感的処理の\n中心的役割", color: C.accent, icon: "🌐" },
  ];
  regions.forEach(function(r, i) {
    var xPos = 0.5 + i * 4.8;
    addCard(s, xPos, 1.2, 4.3, 2.6);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 4.3, h: 0.55, fill: { color: r.color } });
    s.addText(r.name, { x: xPos, y: 1.2, w: 4.3, h: 0.55, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(r.icon, { x: xPos, y: 1.8, w: 4.3, h: 0.5, fontSize: 26, align: "center", margin: 0 });
    s.addText(r.role, { x: xPos + 0.2, y: 2.35, w: 3.9, h: 1.2, fontSize: 15, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
  });

  // 社会的絆ポイント
  addCard(s, 0.5, 4.1, 9.0, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.1, w: 0.15, h: 1.3, fill: { color: C.green } });
  s.addText([
    { text: "もらいあくびは社会的絆の強さを反映する", options: { bold: true, color: C.primary, breakLine: true } },
    { text: "親しい人のあくびはうつりやすく、見知らぬ人・写真・アニメーションはうつりにくい", options: { color: C.text, breakLine: true } },
    { text: "→ これらの領域の活動の強さがもらいあくびの起こりやすさと相関する", options: { color: C.green, bold: true } },
  ], { x: 0.9, y: 4.2, w: 8.4, h: 1.1, fontSize: 14, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: 動物もあくびはうつる
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "動物にもうつるあくび ― 社会性をもつ生き物の共通反応");

  var animals = [
    { name: "チンパンジー\nボノボ", note: "ヒトに最も近い\n霊長類での報告あり", color: C.primary, icon: "🦧" },
    { name: "イヌ", note: "ヒトのあくびを\n見てうつることが\n報告されている", color: C.green, icon: "🐕" },
    { name: "オウム", note: "鳥類でも\n類似反応の\n報告あり", color: C.orange, icon: "🦜" },
    { name: "ゾウ", note: "群れ内での\n同調行動として\n観察される", color: C.accent, icon: "🐘" },
  ];
  animals.forEach(function(a, i) {
    var xPos = 0.2 + i * 2.4;
    addCard(s, xPos, 1.15, 2.2, 2.8);
    s.addText(a.icon, { x: xPos, y: 1.25, w: 2.2, h: 0.7, fontSize: 32, align: "center", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.95, w: 2.2, h: 0.5, fill: { color: a.color } });
    s.addText(a.name, { x: xPos, y: 1.95, w: 2.2, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(a.note, { x: xPos + 0.1, y: 2.5, w: 2.0, h: 1.3, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // 底部メッセージ
  addCard(s, 0.5, 4.2, 9.0, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.2, w: 9.0, h: 0.45, fill: { color: C.dark } });
  s.addText("なぜ社会性のある動物に広くみられるのか？", { x: 0.8, y: 4.2, w: 8.4, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("→ もらいあくびは「社会的同調・絆の形成」に関わる進化的に保存されたシステムと考えられる", {
    x: 0.8, y: 4.7, w: 8.4, h: 0.6,
    fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
  });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: うつりやすい人・うつりにくい人
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "うつりやすい人・うつりにくい人");

  var rows = [
    { factor: "共感性", easy: "共感スコアが高い", hard: "共感スコアが低い" },
    { factor: "年齢", easy: "成人（特に20〜40代）", hard: "乳幼児（4歳未満）・高齢者" },
    { factor: "相手との関係", easy: "家族・友人・親しい同僚", hard: "見知らぬ人・写真・動画" },
    { factor: "神経発達特性", easy: "定型発達", hard: "自閉スペクトラム症の一部" },
    { factor: "注意の向き", easy: "相手の顔に注目している", hard: "目をそらしている・スマホ操作中" },
  ];

  // テーブルヘッダ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 2.0, h: 0.5, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.3, y: 1.1, w: 3.7, h: 0.5, fill: { color: C.green } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 1.1, w: 3.7, h: 0.5, fill: { color: C.red } });
  s.addText("特徴", { x: 0.3, y: 1.1, w: 2.0, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("うつりやすい傾向", { x: 2.3, y: 1.1, w: 3.7, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("うつりにくい傾向", { x: 6.0, y: 1.1, w: 3.7, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.6 + i * 0.72;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.72, fill: { color: bgColor } });
    s.addText(r.factor, { x: 0.3, y: yPos, w: 2.0, h: 0.72, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.easy, { x: 2.3, y: yPos, w: 3.7, h: 0.72, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 8, 0, 8] });
    s.addText(r.hard, { x: 6.0, y: yPos, w: 3.7, h: 0.72, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: [0, 8, 0, 8] });
  });

  // 底部注意書き
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.15, w: 9.4, h: 0.4, fill: { color: C.light } });
  s.addText("「もらいあくびしない＝共感力がない」は誤り。疲労・睡眠不足・注意の向き方など状況要因も大きく影響する", {
    x: 0.5, y: 5.15, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: 病的なあくびと神経疾患
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなときは要注意 ― 病的なあくびと神経疾患");

  var items = [
    {
      title: "脳幹・視床下部病変",
      desc: "延髄・橋・視床下部が制御を担う\n梗塞・腫瘍・脱髄（多発性硬化症など）で過剰なあくびが出現\nめまい・構音障害・嚥下障害・顔面感覚障害を伴う場合は緊急性あり",
      color: C.red,
      icon: "🧠",
    },
    {
      title: "てんかん発作（側頭葉てんかんなど）",
      desc: "発作の前兆（前駆症状）としてあくびが繰り返し出ることがある\nデジャブ・腹部不快感などの前兆と合わせて診断の手がかりに",
      color: C.orange,
      icon: "⚡",
    },
    {
      title: "薬剤性あくび（SSRI・SNRI）",
      desc: "セロトニン作動性薬の副作用として過剰なあくびが出現することがある\n薬剤開始・増量後にあくびが増えた場合は主治医に相談を",
      color: C.accent,
      icon: "💊",
    },
  ];

  items.forEach(function(item, i) {
    var yPos = 1.15 + i * 1.35;
    addCard(s, 0.5, yPos, 9.0, 1.15);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 0.15, h: 1.15, fill: { color: item.color } });
    s.addText(item.icon, { x: 0.7, y: yPos + 0.1, w: 0.7, h: 0.9, fontSize: 22, align: "center", valign: "middle", margin: 0 });
    s.addText(item.title, { x: 1.5, y: yPos + 0.05, w: 7.7, h: 0.38, fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true, margin: 0 });
    s.addText(item.desc, { x: 1.5, y: yPos + 0.44, w: 7.7, h: 0.66, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: 緊急受診サイン
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなあくびはすぐ受診を");

  var signs = [
    "急に止まらないあくびが出現し、めまい・ろれつが回らない・物が二重に見える などを伴う",
    "あくびの直前後に意識がぼんやりする・自分でも気づかない行動をしている",
    "抗うつ薬など（SSRI・SNRI）を開始・増量した後から、著しくあくびが増えた",
    "あくびと同時に頭痛・嘔吐（おうと）・複視（ふくし）が伴う",
  ];

  signs.forEach(function(sign, i) {
    var yPos = 1.1 + i * 0.98;
    addCard(s, 0.5, yPos, 9.0, 0.82);
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.16, w: 0.5, h: 0.5, fill: { color: C.red } });
    s.addText("!" , { x: 0.7, y: yPos + 0.16, w: 0.5, h: 0.5, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(sign, { x: 1.35, y: yPos, w: 7.9, h: 0.82, fontSize: 15, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6] });
  });

  // 底部メッセージ
  addCard(s, 0.5, 5.0, 9.0, 0.55);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9.0, h: 0.55, fill: { color: C.red } });
  s.addText("「ふだんと違う・急に増えた・他の症状を伴う」あくびは脳神経内科へ", {
    x: 0.8, y: 5.0, w: 8.4, h: 0.55,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "8/9");
})();

// ============================================================
// SLIDE 9: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.3, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "もらいあくびはミラーニューロン系と共感回路（前帯状皮質・島皮質）が関与する社会的現象" },
    { num: "2", text: "視覚だけでなく聴覚・言語でもうつる多感覚現象で、社会的絆の強さを反映する" },
    { num: "3", text: "うつりやすさには個人差があり、共感性・年齢・相手との関係が影響する" },
    { num: "4", text: "止まらない・急に増えたあくびは脳幹病変・てんかん・薬剤性のサインの可能性あり\nめまい・構音障害・意識障害などを伴う場合は緊急受診を" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.1 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.2, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.1, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.2, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "9/9");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/あくび_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
