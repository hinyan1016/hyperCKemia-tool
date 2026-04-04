var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "この頭痛、病院に行くべき？ ― 危険な頭痛の見分け方";

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

  s.addText("脳神経内科医が答える からだの不思議 #13", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("この頭痛、病院に行くべき？\n― 危険な頭痛の見分け方 ―", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 34, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("一次性 vs 二次性・雷鳴頭痛・SNNOOP10", {
    x: 0.8, y: 2.9, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
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
    "「バットで殴られたような頭痛が突然きて、救急車を呼んだ」",
    "「月に数回、目の奥がズキズキして光がまぶしくなる」",
    "「毎日のように頭全体が締め付けられる感じがある」",
    "「高熱と一緒に頭痛が来て、首が硬い気がした」",
  ];
  quotes.forEach(function(q, i) {
    var yPos = 1.1 + i * 1.0;
    addCard(s, 0.8, yPos, 8.4, 0.8);
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.2, w: 0.4, h: 0.4, fill: { color: C.red } });
    s.addText(q, {
      x: 1.5, y: yPos, w: 7.5, h: 0.8,
      fontSize: 16, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 5.0, w: 8.4, h: 0.5, fill: { color: C.light } });
  s.addText("すべて「頭痛」でも、緊急度はまったく異なります", {
    x: 0.8, y: 5.0, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: 一次性 vs 二次性頭痛
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "頭痛の2大分類 ― 一次性 vs 二次性");

  var rows = [
    { item: "定義", primary: "頭痛自体が疾患\n（構造的異常なし）", secondary: "他の病気が原因" },
    { item: "頻度", primary: "全頭痛の約90%以上", secondary: "全頭痛の約10%未満" },
    { item: "代表例", primary: "緊張型・片頭痛・群発頭痛", secondary: "くも膜下出血・髄膜炎\n脳腫瘍・高血圧性頭痛" },
    { item: "命の危険", primary: "基本的になし", secondary: "場合により生命に関わる" },
    { item: "発症パターン", primary: "繰り返す・慢性・パターンあり", secondary: "突然の発症・今まで経験なし" },
  ];

  // ヘッダ行
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.0, w: 2.0, h: 0.55, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.3, y: 1.0, w: 3.7, h: 0.55, fill: { color: C.green } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 1.0, w: 3.7, h: 0.55, fill: { color: C.red } });
  s.addText("項目", { x: 0.3, y: 1.0, w: 2.0, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("一次性頭痛", { x: 2.3, y: 1.0, w: 3.7, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("二次性頭痛（要注意）", { x: 6.0, y: 1.0, w: 3.7, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.55 + i * 0.78;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.78, fill: { color: bgColor } });
    s.addText(r.item, { x: 0.3, y: yPos, w: 2.0, h: 0.78, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.primary, { x: 2.3, y: yPos, w: 3.7, h: 0.78, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 8, 0, 8] });
    s.addText(r.secondary, { x: 6.0, y: yPos, w: 3.7, h: 0.78, fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: [0, 8, 0, 8] });
  });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: 雷鳴頭痛 ― くも膜下出血のサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "「人生最悪の頭痛」― くも膜下出血のサイン");

  // 雷鳴頭痛定義
  addCard(s, 0.5, 1.0, 9.0, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 0.15, h: 1.1, fill: { color: C.red } });
  s.addText([
    { text: "雷鳴頭痛（サンダークラップヘッドエイク）", options: { bold: true, color: C.red, fontSize: 18, breakLine: true } },
    { text: "雷が落ちたように突然（1秒以内）に始まる、人生最強の頭痛", options: { color: C.text, fontSize: 15 } },
  ], { x: 0.85, y: 1.05, w: 8.4, h: 1.0, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  // 即119番の警告
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.25, w: 9.0, h: 0.45, fill: { color: C.red } });
  s.addText("今すぐ119番　― 以下のどれか1つでも当てはまれば", {
    x: 0.5, y: 2.25, w: 9.0, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var sings = [
    "突然始まった「今まで生きてきた中で最もひどい頭痛」",
    "頭痛と同時に意識を失いかけた・失った",
    "頭痛と同時に首が硬くなった（項部硬直）",
    "頭痛に加えて手足の麻痺・ろれつが回らない",
    "頭痛と同時に高熱・意識の混濁",
  ];
  sings.forEach(function(sg, i) {
    var yPos = 2.8 + i * 0.46;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.4, fill: { color: (i % 2 === 0) ? "FFF5F5" : C.white } });
    s.addShape(pres.shapes.OVAL, { x: 0.65, y: yPos + 0.08, w: 0.24, h: 0.24, fill: { color: C.red } });
    s.addText(sg, { x: 1.05, y: yPos, w: 8.2, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: SNNOOP10 レッドフラッグ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "危険な頭痛のレッドフラッグ ― SNNOOP10");

  // 上部説明
  s.addText("国際頭痛学会が提唱。1つでも当てはまれば医療機関での精査が必要です", {
    x: 0.5, y: 0.95, w: 9.0, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  var flags = [
    { letter: "S", word: "Systemic symptoms", desc: "全身症状：発熱・体重減少・悪性腫瘍の既往" },
    { letter: "N", word: "Neuropathy / Neurodeficit", desc: "神経症状：麻痺・言語障害・視力障害・意識変容" },
    { letter: "N", word: "New / different headache", desc: "新しい頭痛・いつもと違う頭痛" },
    { letter: "O", word: "Onset thunderclap", desc: "突然発症：雷鳴頭痛・1分以内に最強になる" },
    { letter: "O", word: "Older age (>50)", desc: "高齢発症：50歳以降に初めて起こった頭痛" },
    { letter: "P", word: "Progressive", desc: "進行性：徐々に悪化している頭痛パターン" },
    { letter: "+4", word: "追加サイン", desc: "体位変換で増悪 / 咳・いきみで誘発 / 免疫不全 / 妊娠・産褥期" },
  ];

  flags.forEach(function(f, i) {
    var col = (i < 4) ? 0 : 1;
    var row = (i < 4) ? i : i - 4;
    var xPos = 0.3 + col * 4.9;
    var yPos = 1.55 + row * 0.97;
    addCard(s, xPos, yPos, 4.6, 0.82);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 0.6, h: 0.82, fill: { color: C.red } });
    s.addText(f.letter, { x: xPos, y: yPos, w: 0.6, h: 0.82, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.word, { x: xPos + 0.7, y: yPos + 0.02, w: 3.8, h: 0.35, fontSize: 12, fontFace: FONT_EN, color: C.primary, bold: true, margin: 0 });
    s.addText(f.desc, { x: xPos + 0.7, y: yPos + 0.35, w: 3.8, h: 0.45, fontSize: 11, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: 救急車 vs 専門外来の判断
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "救急車 vs 専門外来 ― どちらに行くべきか");

  // 左: 今すぐ119番
  addCard(s, 0.3, 1.0, 4.3, 4.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.0, w: 4.3, h: 0.55, fill: { color: C.red } });
  s.addText("今すぐ119番（救急車）", { x: 0.3, y: 1.0, w: 4.3, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var emergency = [
    "今まで経験なし・突然始まった最強の頭痛",
    "意識がおかしい・呼びかけに反応しない",
    "手足の麻痺・顔の歪み・ろれつが回らない",
    "高熱＋頭痛＋首が前に曲がらない",
    "けいれん発作を起こした",
  ];
  emergency.forEach(function(e, i) {
    var yPos = 1.65 + i * 0.68;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: yPos + 0.14, w: 0.3, h: 0.3, fill: { color: C.red } });
    s.addText(e, { x: 0.95, y: yPos, w: 3.5, h: 0.6, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  // 右: 数日以内に専門外来
  addCard(s, 5.0, 1.0, 4.7, 4.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y: 1.0, w: 4.7, h: 0.55, fill: { color: C.orange } });
  s.addText("数日以内に専門外来（脳神経内科）", { x: 5.0, y: 1.0, w: 4.7, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var outpatient = [
    "いつもの頭痛と何か違う気がする",
    "頭痛の頻度・強さが最近増してきた",
    "市販薬を月に10日以上飲んでいる",
    "50歳以降に初めて頭痛が出現した",
    "日常生活・仕事に支障が出ている",
  ];
  outpatient.forEach(function(o, i) {
    var yPos = 1.65 + i * 0.68;
    s.addShape(pres.shapes.OVAL, { x: 5.15, y: yPos + 0.14, w: 0.3, h: 0.3, fill: { color: C.orange } });
    s.addText(o, { x: 5.6, y: yPos, w: 3.9, h: 0.6, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: 代表的な一次性頭痛の概要
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "代表的な一次性頭痛 ― 3種類を押さえよう");

  var types = [
    {
      name: "緊張型頭痛",
      freq: "最も頻度が高い",
      color: C.accent,
      features: [
        "頭全体の締め付け感・重い圧迫感",
        "日本人の約30〜40%が経験",
        "市販鎮痛薬が有効（使い過ぎ注意）",
        "長時間のPC・スマホ作業で増加",
      ],
    },
    {
      name: "片頭痛",
      freq: "脳神経内科で最もよく診る",
      color: C.primary,
      features: [
        "片側のズキンズキンする拍動性頭痛",
        "吐き気・光過敏・音過敏を伴う",
        "日本では約840万人（20〜40代女性に多い）",
        "トリプタン系薬が特効薬",
      ],
    },
    {
      name: "群発頭痛",
      freq: "最もつらい頭痛のひとつ",
      color: C.red,
      features: [
        "目の奥をえぐるような片側激痛",
        "1〜2時間続き数週間毎日同時刻に発作",
        "涙・鼻水・まぶたの腫れを伴う",
        "トリプタン注射・酸素吸入が有効",
      ],
    },
  ];

  types.forEach(function(t, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.0, 3.0, 4.4);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.0, w: 3.0, h: 0.65, fill: { color: t.color } });
    s.addText(t.name, { x: xPos, y: 1.0, w: 3.0, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.freq, { x: xPos, y: 1.38, w: 3.0, h: 0.28, fontSize: 10, fontFace: FONT_JP, color: C.white, align: "center", italic: true, margin: 0 });
    t.features.forEach(function(f, j) {
      var yPos = 1.75 + j * 0.85;
      s.addShape(pres.shapes.OVAL, { x: xPos + 0.15, y: yPos + 0.25, w: 0.2, h: 0.2, fill: { color: t.color } });
      s.addText(f, { x: xPos + 0.45, y: yPos, w: 2.4, h: 0.85, fontSize: 11, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    });
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: 警告頭痛（sentinel headache）と予防
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "「警告頭痛」を見逃さない ― SAH発症の前兆");

  // 警告頭痛の説明
  addCard(s, 0.5, 1.1, 9.0, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 0.15, h: 1.5, fill: { color: C.orange } });
  s.addText([
    { text: "警告頭痛（sentinel headache）", options: { bold: true, color: C.orange, fontSize: 17, breakLine: true } },
    { text: "くも膜下出血の約30〜50%は、大出血の数日〜数週間前に一過性の頭痛が起こる", options: { color: C.text, fontSize: 14, breakLine: true } },
    { text: "「突然ひどい頭痛がして、数時間で治まった」 → この段階で受診すれば動脈瘤破裂を防げる場合がある", options: { color: C.primary, fontSize: 13, bold: true } },
  ], { x: 0.85, y: 1.15, w: 8.4, h: 1.4, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  // 50歳以降の新規頭痛
  addCard(s, 0.5, 2.85, 9.0, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.85, w: 0.15, h: 1.1, fill: { color: C.yellow } });
  s.addText([
    { text: "50歳以降の新規頭痛は特に注意", options: { bold: true, color: C.orange, fontSize: 16, breakLine: true } },
    { text: "脳腫瘍・側頭動脈炎（巨細胞性動脈炎）など見逃せない疾患を念頭に。眼科・神経内科への早期受診を", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.85, y: 2.9, w: 8.4, h: 1.0, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  // 薬物乱用頭痛への注意
  addCard(s, 0.5, 4.1, 9.0, 1.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.1, w: 0.15, h: 1.0, fill: { color: C.accent } });
  s.addText([
    { text: "薬物乱用頭痛（MOH）にも注意", options: { bold: true, color: C.primary, fontSize: 15, breakLine: true } },
    { text: "市販鎮痛薬を月10日以上服用 → 慢性頭痛へ移行するリスク。脳神経内科への相談を", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.85, y: 4.15, w: 8.4, h: 0.9, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "8/9");
})();

// ============================================================
// SLIDE 9: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("まとめ ― Take Home Message", {
    x: 0.5, y: 0.15, w: 9.0, h: 0.55,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var msgs = [
    { num: "1", text: "頭痛は「一次性（命に関わらない）」と「二次性（要注意）」に大別される", color: C.accent },
    { num: "2", text: "「突然始まった人生最悪の頭痛」はくも膜下出血の可能性 → 即119番", color: C.red },
    { num: "3", text: "SNNOOP10のサインが1つでもあれば必ず医療機関で精査を", color: C.orange },
    { num: "4", text: "「いつもと何か違う」という感覚を大切に。その直感が命を救うことがある", color: C.yellow },
    { num: "5", text: "片頭痛・緊張型は脳神経内科で管理できる。我慢せず相談を", color: C.green },
  ];

  msgs.forEach(function(m, i) {
    var yPos = 0.9 + i * 0.92;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.78, fill: { color: "1F3D5C" } });
    s.addShape(pres.shapes.OVAL, { x: 0.65, y: yPos + 0.14, w: 0.5, h: 0.5, fill: { color: m.color } });
    s.addText(m.num, { x: 0.65, y: yPos + 0.14, w: 0.5, h: 0.5, fontSize: 16, fontFace: FONT_EN, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.35, y: yPos, w: 7.9, h: 0.78, fontSize: 14, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  s.addText("医知創造ラボ", {
    x: 0.5, y: 5.4, w: 9.0, h: 0.3,
    fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
})();

// ============================================================
// 出力
// ============================================================
var outPath = __dirname + "/危険な頭痛_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
