var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "目がピクピクする ― 放置していいのか？ まぶたの痙攣を脳神経内科医が解説";

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

  s.addText("脳神経内科医が答える からだの不思議 #08", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("目がピクピクする\n― 放置していいのか？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― まぶたのピクピクの正体とその見分け方 ―", {
    x: 0.8, y: 2.9, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.7, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.0, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: こんな経験ありませんか？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんな経験ありませんか？");

  var quotes = [
    "「残業続きの週に、下まぶたがピクピクし始めた」",
    "「コーヒーをたくさん飲んだ日に症状が出やすい」",
    "「まぶたがピクピクする ― 脳に異常があるのでは？」",
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
  s.addText("ほぼ全員が一生に一度は経験するありふれた現象 ― でも心配なケースも", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/10");
})();

// ============================================================
// SLIDE 3: 眼瞼ミオキミアとは
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "眼瞼ミオキミア（Eyelid Myokymia）とは");

  addCard(s, 0.5, 1.2, 9.0, 1.5);
  s.addText("Eyelid Myokymia（眼瞼ミオキミア）", {
    x: 0.8, y: 1.25, w: 8.4, h: 0.5,
    fontSize: 22, fontFace: FONT_EN, color: C.primary, bold: true, margin: 0,
  });
  s.addText("眼輪筋の一部が自発的に不規則な電気信号を発して収縮する現象", {
    x: 0.8, y: 1.8, w: 8.4, h: 0.7,
    fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  var triggers = [
    { label: "睡眠不足\n疲労", color: C.red },
    { label: "カフェイン\n過多", color: C.orange },
    { label: "ストレス\n緊張", color: C.yellow },
    { label: "スクリーン\n疲労", color: C.accent },
    { label: "ドライアイ\nアルコール", color: C.green },
  ];
  triggers.forEach(function(t, i) {
    var xPos = 0.3 + i * 1.88;
    addCard(s, xPos, 3.1, 1.7, 1.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.1, w: 1.7, h: 0.45, fill: { color: t.color } });
    s.addText(t.label, {
      x: xPos, y: 3.1, w: 1.7, h: 1.5,
      fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
      lineSpacingMultiple: 1.3,
    });
  });

  s.addText("主な引き金", {
    x: 0.3, y: 2.75, w: 9.4, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.8, w: 9.0, h: 0.6, fill: { color: C.light } });
  s.addText("下まぶた・片目のみ・細かいピクピク・引き金解消で数日以内に自然消失", {
    x: 0.8, y: 4.8, w: 8.4, h: 0.6,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "3/10");
})();

// ============================================================
// SLIDE 4: メカニズム ― なぜ疲れると起きるのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "メカニズム ― なぜ疲れるとピクピクするのか");

  var steps = [
    { icon: "😴", label: "疲労・睡眠不足\nカフェイン・Mg不足", color: C.orange },
    { icon: "⚡", label: "神経・筋肉の\n静止電位が不安定に", color: C.yellow },
    { icon: "🔀", label: "わずかな刺激で\n自発放電が起きやすくなる", color: C.red },
    { icon: "👁️", label: "指令なしに\n眼輪筋がピクッと収縮", color: C.primary },
  ];

  steps.forEach(function(st, i) {
    var xPos = 0.2 + i * 2.45;
    addCard(s, xPos, 1.2, 2.2, 2.5);
    s.addText(st.icon, { x: xPos, y: 1.35, w: 2.2, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.1, y: 2.05, w: 2.0, h: 0.4, fill: { color: st.color } });
    s.addText(st.label, {
      x: xPos, y: 2.5, w: 2.2, h: 1.0,
      fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 4,
      lineSpacingMultiple: 1.2,
    });
  });

  s.addText("→", { x: 2.35, y: 2.0, w: 0.4, h: 0.5, fontSize: 24, color: C.accent, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 4.8, y: 2.0, w: 0.4, h: 0.5, fontSize: 24, color: C.accent, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 7.25, y: 2.0, w: 0.4, h: 0.5, fontSize: 24, color: C.accent, align: "center", valign: "middle", margin: 0 });

  addCard(s, 0.5, 4.0, 9.0, 1.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 9.0, h: 0.45, fill: { color: C.primary } });
  s.addText("核心メッセージ", { x: 0.8, y: 4.0, w: 8.4, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "ピクピクは「筋肉の誤作動」 ― 疲労などで静止電位が不安定になっているだけ", options: { fontSize: 15, bold: true, color: C.red, breakLine: true } },
    { text: "→ 原因を取り除けば自然に落ち着く", options: { fontSize: 14, color: C.text } },
  ], { x: 0.8, y: 4.5, w: 8.4, h: 0.85, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "4/10");
})();

// ============================================================
// SLIDE 5: 良性 vs 病的 ― 比較表
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "良性と病的なピクピクの比較");

  var rows = [
    { item: "左右", normal: "片目のみ", warning: "片目→顔半側に広がる\nまたは両目" },
    { item: "部位", normal: "下まぶた中心\n限局的", warning: "頬・口角・まぶた全体" },
    { item: "動き方", normal: "細かくピクピク\n（不規則）", warning: "顔が引きつる\n目がギュッと閉じる" },
    { item: "持続時間", normal: "数秒〜数分\n断続的", warning: "持続的・発作的に繰り返す" },
    { item: "自然経過", normal: "誘因解消で\n数日以内に消失", warning: "徐々に悪化\n自然消失しない" },
  ];

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 1.9, h: 0.55, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.2, y: 1.1, w: 3.5, h: 0.55, fill: { color: C.green } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.7, y: 1.1, w: 4.0, h: 0.55, fill: { color: C.red } });
  s.addText("項目", { x: 0.3, y: 1.1, w: 1.9, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("良性（眼瞼ミオキミア）", { x: 2.2, y: 1.1, w: 3.5, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("要注意（病的なもの）", { x: 5.7, y: 1.1, w: 4.0, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.65 + i * 0.76;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.76, fill: { color: bgColor } });
    s.addText(r.item, { x: 0.3, y: yPos, w: 1.9, h: 0.76, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.normal, { x: 2.2, y: yPos, w: 3.5, h: 0.76, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.1 });
    s.addText(r.warning, { x: 5.7, y: yPos, w: 4.0, h: 0.76, fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.1 });
  });

  addPageNum(s, "5/10");
})();

// ============================================================
// SLIDE 6: 代表的な3つの病態
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "放置してはいけない3つの病態");

  var conditions = [
    {
      num: "1",
      name: "片側顔面痙攣",
      eng: "Hemifacial spasm",
      desc: "まぶたから始まり同側の頬・口角へ広がる。顔面神経が動脈で圧迫されることで発症。自然消失しない。",
      treatment: "ボツリヌス毒素注射 / 微小血管減圧術",
      color: C.red,
    },
    {
      num: "2",
      name: "眼瞼痙攣",
      eng: "Blepharospasm",
      desc: "両目が意思に反して閉じてしまう。大脳基底核のジストニア。強い光・疲労・ストレスで悪化。",
      treatment: "ボツリヌス毒素注射（第一選択）",
      color: C.orange,
    },
    {
      num: "3",
      name: "マイジュ症候群",
      eng: "Meige syndrome",
      desc: "眼瞼痙攣＋口・顎・舌の不随意運動を伴う顔面ジストニア。中年以降の女性に多い。",
      treatment: "ボツリヌス毒素注射＋薬物療法",
      color: C.primary,
    },
  ];

  conditions.forEach(function(c, i) {
    var yPos = 1.1 + i * 1.45;
    addCard(s, 0.3, yPos, 9.4, 1.3);
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: yPos + 0.2, w: 0.9, h: 0.9, fill: { color: c.color } });
    s.addText(c.num, { x: 0.5, y: yPos + 0.2, w: 0.9, h: 0.9, fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.name + "  " + c.eng, {
      x: 1.6, y: yPos + 0.05, w: 7.8, h: 0.45,
      fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
    });
    s.addText(c.desc, {
      x: 1.6, y: yPos + 0.5, w: 5.5, h: 0.7,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2,
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 7.3, y: yPos + 0.45, w: 2.2, h: 0.65, fill: { color: c.color } });
    s.addText(c.treatment, {
      x: 7.3, y: yPos + 0.45, w: 2.2, h: 0.65,
      fontSize: 11, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 2,
      lineSpacingMultiple: 1.1,
    });
  });

  addPageNum(s, "6/10");
})();

// ============================================================
// SLIDE 7: セルフケア（良性ミオキミアの場合）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "セルフケア ― 良性のピクピクへの対処法");

  var cares = [
    { icon: "🛏️", title: "睡眠を確保する", desc: "7〜8時間を目安に。睡眠不足は最多の引き金" },
    { icon: "☕", title: "カフェインを控える", desc: "コーヒー・エナジードリンクを一時的に減らす" },
    { icon: "👁️", title: "画面から目を休める", desc: "1時間に1回、20秒以上遠くを見る（20-20-20ルール）" },
    { icon: "🥗", title: "Mg含む食品を意識する", desc: "ナッツ・豆腐・ほうれん草。電解質バランスを整える" },
    { icon: "💧", title: "目薬でドライアイを和らげる", desc: "人工涙液を使い、目の表面の刺激を減らす" },
  ];

  cares.forEach(function(c, i) {
    var col = (i < 3) ? 0 : 1;
    var row = (i < 3) ? i : i - 3;
    var xPos = (col === 0) ? 0.3 : 5.2;
    var w = (col === 0) ? 4.6 : 4.5;
    var yPos = 1.15 + row * 1.22;
    addCard(s, xPos, yPos, w, 1.05);
    s.addText(c.icon, { x: xPos + 0.1, y: yPos + 0.1, w: 0.8, h: 0.85, fontSize: 26, align: "center", valign: "middle", margin: 0 });
    s.addText(c.title, { x: xPos + 1.0, y: yPos + 0.05, w: w - 1.1, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
    s.addText(c.desc, { x: xPos + 1.0, y: yPos + 0.45, w: w - 1.1, h: 0.55, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.9, w: 9.4, h: 0.5, fill: { color: C.light } });
  s.addText("1〜2週間試しても改善しない場合や頬・口角への広がりがある場合は受診を", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "7/10");
})();

// ============================================================
// SLIDE 8: 受診基準 ― こんなときは要注意
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなときは脳神経内科へ ― 受診のサイン");

  var signs = [
    "頬や口角まで引きつりが広がる（片側顔面痙攣の可能性）",
    "両目が意思に反して閉じてしまう（眼瞼痙攣の可能性）",
    "痙攣が2週間以上続いて止まらない",
    "口・舌の不随意運動も一緒に出現している",
    "顔の麻痺（動かしにくさ）やしびれを伴う",
    "耳鳴りや難聴を伴う（顔面神経周囲の病変の可能性）",
  ];

  signs.forEach(function(sg, i) {
    var yPos = 1.1 + i * 0.7;
    addCard(s, 0.5, yPos, 9.0, 0.6);
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.12, w: 0.36, h: 0.36, fill: { color: C.red } });
    s.addText("" + (i + 1), { x: 0.7, y: yPos + 0.12, w: 0.36, h: 0.36, fontSize: 12, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(sg, {
      x: 1.2, y: yPos, w: 8.0, h: 0.6,
      fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.15, w: 9.0, h: 0.4, fill: { color: C.light } });
  s.addText("「下まぶただけ・疲れたとき限定」なら過度な心配は不要。広がる・止まらない・閉じるは要受診", {
    x: 0.7, y: 5.15, w: 8.6, h: 0.4,
    fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "8/10");
})();

// ============================================================
// SLIDE 9: 脳神経内科とは
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "「顔の動きの異常」は脳神経内科の専門領域");

  addCard(s, 0.5, 1.2, 9.0, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9.0, h: 0.5, fill: { color: C.primary } });
  s.addText("まぶたのピクピクに気づいたら、まず脳神経内科へ", {
    x: 0.8, y: 1.2, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });
  s.addText("「精神科では？」と思われがちですが、顔面の不随意運動・痙攣は脳神経内科の得意領域です。\n顔面神経・大脳基底核・脳幹などの評価から、ボツリヌス治療まで一貫して対応できます。", {
    x: 0.8, y: 1.78, w: 8.4, h: 0.65,
    fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.3,
  });

  var tips = [
    { label: "こんな一言で受診しやすい", items: ["「最近まぶたがずっとピクピクしている」", "「頬まで広がってきた気がする」", "「目が自然に閉じてしまうことがある」"] },
    { label: "診察で行われる検査", items: ["問診（症状のパターン・誘因）", "神経学的診察（顔面神経の評価）", "必要に応じてMRI・筋電図"] },
  ];

  tips.forEach(function(tp, i) {
    var xPos = 0.5 + i * 4.8;
    addCard(s, xPos, 2.7, 4.5, 2.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.7, w: 4.5, h: 0.45, fill: { color: (i === 0) ? C.green : C.accent } });
    s.addText(tp.label, { x: xPos, y: 2.7, w: 4.5, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    tp.items.forEach(function(it, j) {
      s.addText("●  " + it, {
        x: xPos + 0.2, y: 3.25 + j * 0.6, w: 4.1, h: 0.55,
        fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
      });
    });
  });

  addPageNum(s, "9/10");
})();

// ============================================================
// SLIDE 10: まとめ（Take Home Message）
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
    { num: "1", text: "日常的な「まぶたのピクピク」の大多数は眼瞼ミオキミア\n疲労・カフェイン・ストレスが引き金で自然に治る" },
    { num: "2", text: "頬や口角へ広がる → 片側顔面痙攣を疑う\n自然消失しないため脳神経内科へ" },
    { num: "3", text: "両目が意思に反して閉じる → 眼瞼痙攣\nボツリヌス毒素注射が有効" },
    { num: "4", text: "口・舌の不随意運動も伴う → マイジュ症候群\n「広がる・止まらない・閉じる」が受診のサイン" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.1 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.2, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.1, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.2, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "10/10");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/まぶたピクピク_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
