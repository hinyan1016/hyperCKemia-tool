var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "ミネブロ（エサキセレノン）徹底解説｜MRA4剤比較と臨床での使い分け";

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
  purple: "6F42C1",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";

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
  s.addText(num, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("降圧薬シリーズ", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("ミネブロ（エサキセレノン）", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.0,
    fontSize: 38, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addText("MRA4剤比較と臨床での使い分け", {
    x: 0.8, y: 2.3, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.15, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText([
    { text: "非ステロイド型MRA  |  ", options: { color: C.light, fontSize: 14 } },
    { text: "ESAX-HTN / ESAX-DN", options: { color: C.accent, fontSize: 14, italic: true } },
  ], { x: 0.8, y: 3.4, w: 8.4, h: 0.4, fontFace: FONT_EN, align: "center", margin: 0 });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: アルドステロンの多面的作用
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "アルドステロンの多面的作用");

  // 左側: アルドステロンの基本作用
  addCard(s, 0.4, 1.1, 4.4, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 0.45, fill: { color: C.primary } });
  s.addText("アルドステロンとは", { x: 0.6, y: 1.1, w: 4, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "副腎皮質から分泌される\nミネラルコルチコイドホルモン\n\n", options: { fontSize: 14, color: C.text } },
    { text: "腎臓での作用\n", options: { fontSize: 14, color: C.primary, bold: true } },
    { text: "MR受容体に結合\n", options: { fontSize: 13, color: C.text } },
    { text: "  Na再吸収 ", options: { fontSize: 13, color: C.text } },
    { text: "↑", options: { fontSize: 14, color: C.red, bold: true } },
    { text: " + K排泄 ", options: { fontSize: 13, color: C.text } },
    { text: "↑", options: { fontSize: 14, color: C.red, bold: true } },
    { text: "\n  → 循環血液量↑ → ", options: { fontSize: 13, color: C.text } },
    { text: "血圧上昇", options: { fontSize: 14, color: C.red, bold: true } },
  ], { x: 0.6, y: 1.7, w: 4, h: 3.2, fontFace: FONT_JP, valign: "top", margin: 4 });

  // 右側: 臓器障害
  addCard(s, 5.2, 1.1, 4.4, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.45, fill: { color: C.red } });
  s.addText("臓器障害への関与", { x: 5.4, y: 1.1, w: 4, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var organs = [
    { icon: "心臓", text: "心筋線維化 → 心肥大・拡張障害", y: 1.75 },
    { icon: "血管", text: "血管炎症・リモデリング → 動脈硬化", y: 2.55 },
    { icon: "腎臓", text: "メサンギウム増殖・硬化 → CKD進行", y: 3.35 },
  ];
  for (var i = 0; i < organs.length; i++) {
    var o = organs[i];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.5, y: o.y, w: 1.0, h: 0.55, fill: { color: C.red }, rectRadius: 0.08 });
    s.addText(o.icon, { x: 5.5, y: o.y, w: 1.0, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(o.text, { x: 6.65, y: o.y, w: 2.8, h: 0.55, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.5, y: 4.15, w: 3.9, h: 0.7, fill: { color: C.light }, rectRadius: 0.08 });
  s.addText([
    { text: "アルドステロンブレイクスルー\n", options: { fontSize: 12, color: C.primary, bold: true } },
    { text: "RAS阻害下でもアルドステロンが再上昇", options: { fontSize: 11, color: C.text } },
  ], { x: 5.6, y: 4.15, w: 3.7, h: 0.7, fontFace: FONT_JP, valign: "middle", margin: 2 });

  addPageNum(s, "2");
})();

// ============================================================
// SLIDE 3: MRAの世代分類
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "MRA（ミネラルコルチコイド受容体拮抗薬）の世代分類");

  // 世代テーブル
  var tableRows = [
    [
      { text: "世代", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
      { text: "商品名", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
      { text: "一般名", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
      { text: "構造", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
      { text: "承認年", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
      { text: "適応", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
    ],
    [
      { text: "第1世代", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "アルダクトンA", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "スピロノラクトン", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "ステロイド型", options: { fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true } },
      { text: "1978年", options: { fontSize: 12, fontFace: FONT_EN, color: C.text } },
      { text: "高血圧・心不全\n原発性アルドステロン症", options: { fontSize: 11, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "第2世代", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "セララ", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "エプレレノン", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "ステロイド型", options: { fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true } },
      { text: "2007年", options: { fontSize: 12, fontFace: FONT_EN, color: C.text } },
      { text: "高血圧・心不全", options: { fontSize: 11, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "第3世代", options: { fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true } },
      { text: "ミネブロ", options: { fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true } },
      { text: "エサキセレノン", options: { fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true } },
      { text: "非ステロイド型", options: { fontSize: 12, fontFace: FONT_JP, color: C.green, bold: true } },
      { text: "2019年", options: { fontSize: 12, fontFace: FONT_EN, color: C.primary, bold: true } },
      { text: "高血圧", options: { fontSize: 11, fontFace: FONT_JP, color: C.primary, bold: true } },
    ],
    [
      { text: "第3世代", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "ケレンディア", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "フィネレノン", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "非ステロイド型", options: { fontSize: 12, fontFace: FONT_JP, color: C.green, bold: true } },
      { text: "2022年", options: { fontSize: 12, fontFace: FONT_EN, color: C.text } },
      { text: "CKD(2型DM)\n心不全(2025年追加)", options: { fontSize: 11, fontFace: FONT_JP, color: C.text } },
    ],
  ];

  s.addTable(tableRows, {
    x: 0.4, y: 1.15, w: 9.2,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    rowH: [0.45, 0.65, 0.65, 0.65, 0.65],
    colW: [1.0, 1.5, 1.7, 1.4, 1.0, 2.6],
    margin: [4, 6, 4, 6],
  });

  // 下部注釈
  addCard(s, 0.4, 4.0, 9.2, 1.0);
  s.addText([
    { text: "非ステロイド型MRAの利点：", options: { bold: true, color: C.primary, fontSize: 14 } },
    { text: "ステロイド骨格を持たないため、スピロノラクトンで問題となった", options: { color: C.text, fontSize: 13 } },
    { text: "女性化乳房・月経異常", options: { color: C.red, bold: true, fontSize: 13 } },
    { text: "などの性ホルモン関連副作用を回避できる", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.6, y: 4.1, w: 8.8, h: 0.8, fontFace: FONT_JP, valign: "middle", margin: 4 });

  addPageNum(s, "3");
})();

// ============================================================
// SLIDE 4: ミネブロの薬理プロファイル
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ミネブロ（エサキセレノン）の薬理プロファイル");

  // 左カード: 薬物動態
  addCard(s, 0.4, 1.1, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 0.45, fill: { color: C.primary } });
  s.addText("薬物動態", { x: 0.6, y: 1.1, w: 4, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var pkItems = [
    ["半減期", "約18.6時間（1日1回投与）"],
    ["Tmax", "2〜3時間"],
    ["BA", "89%"],
    ["代謝", "主にCYP3A（酸化）"],
    ["食事影響", "なし"],
  ];
  for (var i = 0; i < pkItems.length; i++) {
    var yPos = 1.7 + i * 0.4;
    s.addText(pkItems[i][0], { x: 0.6, y: yPos, w: 1.2, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(pkItems[i][1], { x: 1.9, y: yPos, w: 2.7, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  }

  // 右カード: 臨床的強み
  addCard(s, 5.2, 1.1, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.45, fill: { color: C.green } });
  s.addText("臨床上の強み", { x: 5.4, y: 1.1, w: 4, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var strengths = [
    "非ステロイド型 → 性ホルモン副作用なし",
    "中等度腎機能障害にも使用可",
    "DM+アルブミン尿の高血圧にも使用可",
    "CYP3A4強力阻害薬との併用禁忌なし",
    "2.5mg ≒ エプレレノン50mg",
  ];
  for (var i = 0; i < strengths.length; i++) {
    var yPos = 1.7 + i * 0.4;
    s.addText("✓", { x: 5.4, y: yPos, w: 0.3, h: 0.35, fontSize: 13, fontFace: FONT_EN, color: C.green, bold: true, valign: "middle", margin: 0 });
    s.addText(strengths[i], { x: 5.7, y: yPos, w: 3.7, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  }

  // 下部: 用法用量テーブル
  var doseRows = [
    [
      { text: "患者群", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 12, fontFace: FONT_JP } },
      { text: "開始用量", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 12, fontFace: FONT_JP } },
      { text: "最大用量", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 12, fontFace: FONT_JP } },
      { text: "備考", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 12, fontFace: FONT_JP } },
    ],
    [
      { text: "通常", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "2.5mg 1日1回", options: { fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true } },
      { text: "5mg", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "効果不十分時に増量", options: { fontSize: 11, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "中等度腎機能障害\n(eGFR 30-60)", options: { fontSize: 11, fontFace: FONT_JP, color: C.text } },
      { text: "1.25mg 1日1回", options: { fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true } },
      { text: "2.5mg", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "4週以降に増量可", options: { fontSize: 11, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "DM+アルブミン尿", options: { fontSize: 11, fontFace: FONT_JP, color: C.text } },
      { text: "1.25mg 1日1回", options: { fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true } },
      { text: "2.5mg", options: { fontSize: 12, fontFace: FONT_JP, color: C.text } },
      { text: "K値頻回モニタリング", options: { fontSize: 11, fontFace: FONT_JP, color: C.red } },
    ],
  ];

  s.addTable(doseRows, {
    x: 0.4, y: 3.85, w: 9.2,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    rowH: [0.38, 0.38, 0.45, 0.38],
    colW: [2.2, 2.2, 1.3, 3.5],
    margin: [3, 6, 3, 6],
  });

  addPageNum(s, "4");
})();

// ============================================================
// SLIDE 5: MRA4剤 完全比較表
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "MRA4剤 完全比較表");

  var compRows = [
    [
      { text: "項目", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, fontFace: FONT_JP } },
      { text: "アルダクトンA\nスピロノラクトン", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, fontFace: FONT_JP } },
      { text: "セララ\nエプレレノン", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, fontFace: FONT_JP } },
      { text: "ミネブロ\nエサキセレノン", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, fontFace: FONT_JP } },
      { text: "ケレンディア\nフィネレノン", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, fontFace: FONT_JP } },
    ],
    [
      { text: "構造", options: { fontSize: 10, fontFace: FONT_JP, color: C.text, bold: true } },
      { text: "ステロイド型", options: { fontSize: 10, fontFace: FONT_JP, color: C.red } },
      { text: "ステロイド型", options: { fontSize: 10, fontFace: FONT_JP, color: C.red } },
      { text: "非ステロイド型", options: { fontSize: 10, fontFace: FONT_JP, color: C.green, bold: true } },
      { text: "非ステロイド型", options: { fontSize: 10, fontFace: FONT_JP, color: C.green, bold: true } },
    ],
    [
      { text: "適応症", options: { fontSize: 10, fontFace: FONT_JP, color: C.text, bold: true } },
      { text: "高血圧\n心不全\n原発性アルド症", options: { fontSize: 9, fontFace: FONT_JP, color: C.text } },
      { text: "高血圧\n心不全", options: { fontSize: 9, fontFace: FONT_JP, color: C.text } },
      { text: "高血圧", options: { fontSize: 10, fontFace: FONT_JP, color: C.primary, bold: true } },
      { text: "CKD(2型DM)\n心不全", options: { fontSize: 9, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "降圧力", options: { fontSize: 10, fontFace: FONT_JP, color: C.text, bold: true } },
      { text: "強い", options: { fontSize: 10, fontFace: FONT_JP, color: C.text } },
      { text: "中等度", options: { fontSize: 10, fontFace: FONT_JP, color: C.text } },
      { text: "強い", options: { fontSize: 10, fontFace: FONT_JP, color: C.primary, bold: true } },
      { text: "弱い\n(3-7mmHg)", options: { fontSize: 9, fontFace: FONT_JP, color: C.sub } },
    ],
    [
      { text: "半減期", options: { fontSize: 10, fontFace: FONT_JP, color: C.text, bold: true } },
      { text: "活性代謝物\n16.5h", options: { fontSize: 9, fontFace: FONT_JP, color: C.text } },
      { text: "4-6h", options: { fontSize: 10, fontFace: FONT_EN, color: C.text } },
      { text: "18.6h", options: { fontSize: 10, fontFace: FONT_EN, color: C.primary, bold: true } },
      { text: "2-3h", options: { fontSize: 10, fontFace: FONT_EN, color: C.text } },
    ],
    [
      { text: "腎機能制限", options: { fontSize: 10, fontFace: FONT_JP, color: C.text, bold: true } },
      { text: "重度で\n慎重投与", options: { fontSize: 9, fontFace: FONT_JP, color: C.text } },
      { text: "CCr<50\n禁忌(HT)", options: { fontSize: 9, fontFace: FONT_JP, color: C.red, bold: true } },
      { text: "eGFR<30\n禁忌", options: { fontSize: 9, fontFace: FONT_JP, color: C.orange } },
      { text: "CKD制限なし\nHF:eGFR<25禁忌", options: { fontSize: 8, fontFace: FONT_JP, color: C.text } },
    ],
    [
      { text: "DM+腎症", options: { fontSize: 10, fontFace: FONT_JP, color: C.text, bold: true } },
      { text: "使用可", options: { fontSize: 10, fontFace: FONT_JP, color: C.text } },
      { text: "禁忌", options: { fontSize: 10, fontFace: FONT_JP, color: C.red, bold: true } },
      { text: "使用可\n(1.25mg開始)", options: { fontSize: 9, fontFace: FONT_JP, color: C.green, bold: true } },
      { text: "適応", options: { fontSize: 10, fontFace: FONT_JP, color: C.green, bold: true } },
    ],
    [
      { text: "性ホルモン\n副作用", options: { fontSize: 9, fontFace: FONT_JP, color: C.text, bold: true } },
      { text: "あり", options: { fontSize: 10, fontFace: FONT_JP, color: C.red, bold: true } },
      { text: "まれ", options: { fontSize: 10, fontFace: FONT_JP, color: C.text } },
      { text: "なし", options: { fontSize: 10, fontFace: FONT_JP, color: C.green, bold: true } },
      { text: "なし", options: { fontSize: 10, fontFace: FONT_JP, color: C.green, bold: true } },
    ],
  ];

  s.addTable(compRows, {
    x: 0.2, y: 1.1, w: 9.6,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    rowH: [0.55, 0.4, 0.6, 0.4, 0.45, 0.5, 0.45, 0.45],
    colW: [1.2, 2.0, 2.0, 2.2, 2.2],
    margin: [2, 4, 2, 4],
  });

  addPageNum(s, "5");
})();

// ============================================================
// SLIDE 6: ESAX-HTN試験
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ESAX-HTN試験 ― 本態性高血圧における降圧効果");

  // 試験デザイン
  addCard(s, 0.4, 1.1, 9.2, 1.2);
  s.addText([
    { text: "試験デザイン：", options: { bold: true, color: C.primary, fontSize: 13 } },
    { text: "無作為化二重盲検実薬対照第3相試験  |  ", options: { color: C.text, fontSize: 12 } },
    { text: "本態性高血圧 1,001例  |  12週間\n", options: { color: C.text, fontSize: 12 } },
    { text: "3群：エサキセレノン2.5mg / 5mg / エプレレノン50mg", options: { color: C.sub, fontSize: 12 } },
  ], { x: 0.6, y: 1.2, w: 8.8, h: 1.0, fontFace: FONT_JP, valign: "middle", margin: 4 });

  // 結果テーブル
  var resRows = [
    [
      { text: "群", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
      { text: "座位収縮期血圧変化", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
      { text: "判定", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 13, fontFace: FONT_JP } },
    ],
    [
      { text: "エサキセレノン 2.5mg", options: { fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true } },
      { text: "-13.7 mmHg", options: { fontSize: 16, fontFace: FONT_EN, color: C.primary, bold: true } },
      { text: "非劣性達成 ✓", options: { fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true } },
    ],
    [
      { text: "エサキセレノン 5mg", options: { fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true } },
      { text: "-14.3 mmHg", options: { fontSize: 16, fontFace: FONT_EN, color: C.primary, bold: true } },
      { text: "優越性達成 ✓", options: { fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true } },
    ],
    [
      { text: "エプレレノン 50mg", options: { fontSize: 13, fontFace: FONT_JP, color: C.sub } },
      { text: "-12.4 mmHg", options: { fontSize: 16, fontFace: FONT_EN, color: C.sub } },
      { text: "対照", options: { fontSize: 13, fontFace: FONT_JP, color: C.sub } },
    ],
  ];

  s.addTable(resRows, {
    x: 0.4, y: 2.55, w: 9.2,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    rowH: [0.45, 0.55, 0.55, 0.55],
    colW: [3.0, 3.2, 3.0],
    margin: [4, 8, 4, 8],
  });

  // ポイント
  addCard(s, 0.4, 4.4, 9.2, 0.7);
  s.addText([
    { text: "Point：", options: { bold: true, color: C.primary, fontSize: 13 } },
    { text: "ミネブロ2.5mgはエプレレノン50mgと同等以上の降圧効果。24時間持続的な降圧を実現。", options: { color: C.text, fontSize: 12 } },
  ], { x: 0.6, y: 4.45, w: 8.8, h: 0.6, fontFace: FONT_JP, valign: "middle", margin: 4 });

  s.addText("Ito S, et al. Hypertension. 2020;75(1):51-58", { x: 0.4, y: 5.15, w: 5, h: 0.3, fontSize: 9, fontFace: FONT_EN, color: C.sub, italic: true, margin: 0 });
  addPageNum(s, "6");
})();

// ============================================================
// SLIDE 7: ESAX-DN試験
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ESAX-DN試験 ― 糖尿病性腎症におけるアルブミン尿改善");

  // 試験デザイン
  addCard(s, 0.4, 1.1, 9.2, 0.9);
  s.addText([
    { text: "試験デザイン：", options: { bold: true, color: C.primary, fontSize: 13 } },
    { text: "プラセボ対照二重盲検比較試験  |  2型DM+微量アルブミン尿 455例  |  52週間", options: { color: C.text, fontSize: 12 } },
  ], { x: 0.6, y: 1.15, w: 8.8, h: 0.8, fontFace: FONT_JP, valign: "middle", margin: 4 });

  // 結果カード左
  addCard(s, 0.4, 2.2, 4.4, 2.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.2, w: 4.4, h: 0.45, fill: { color: C.green } });
  s.addText("主要結果", { x: 0.6, y: 2.2, w: 4, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText("UACR寛解率（主要評価項目）", { x: 0.6, y: 2.8, w: 4, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.text, bold: true, margin: 0 });
  s.addText([
    { text: "エサキセレノン ", options: { fontSize: 13, color: C.text } },
    { text: "22%", options: { fontSize: 22, color: C.primary, bold: true } },
    { text: "  vs  プラセボ ", options: { fontSize: 13, color: C.text } },
    { text: "4%", options: { fontSize: 22, color: C.sub, bold: true } },
  ], { x: 0.6, y: 3.15, w: 4, h: 0.5, fontFace: FONT_EN, valign: "middle", margin: 0 });
  s.addText("P < 0.001  |  絶対差 18%", { x: 0.6, y: 3.65, w: 4, h: 0.3, fontSize: 11, fontFace: FONT_EN, color: C.red, bold: true, margin: 0 });

  s.addText("UACR変化率", { x: 0.6, y: 4.05, w: 4, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.text, bold: true, margin: 0 });
  s.addText([
    { text: "エサキセレノン ", options: { fontSize: 13, color: C.text } },
    { text: "-58%", options: { fontSize: 18, color: C.green, bold: true } },
    { text: "  vs  プラセボ ", options: { fontSize: 13, color: C.text } },
    { text: "+8%", options: { fontSize: 18, color: C.red, bold: true } },
  ], { x: 0.6, y: 4.35, w: 4, h: 0.45, fontFace: FONT_EN, valign: "middle", margin: 0 });

  // 結果カード右
  addCard(s, 5.2, 2.2, 4.4, 2.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.2, w: 4.4, h: 0.45, fill: { color: C.orange } });
  s.addText("安全性", { x: 5.4, y: 2.2, w: 4, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText("高カリウム血症発現率", { x: 5.4, y: 2.8, w: 4, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.text, bold: true, margin: 0 });
  s.addText([
    { text: "エサキセレノン ", options: { fontSize: 13, color: C.text } },
    { text: "9%", options: { fontSize: 22, color: C.red, bold: true } },
    { text: "  vs  プラセボ ", options: { fontSize: 13, color: C.text } },
    { text: "2%", options: { fontSize: 22, color: C.sub, bold: true } },
  ], { x: 5.4, y: 3.15, w: 4, h: 0.5, fontFace: FONT_EN, valign: "middle", margin: 0 });

  s.addText([
    { text: "顕性アルブミン尿への進行率\n", options: { fontSize: 12, color: C.text, bold: true } },
    { text: "エサキセレノン ", options: { fontSize: 13, color: C.text } },
    { text: "1.4%", options: { fontSize: 16, color: C.green, bold: true } },
    { text: " vs プラセボ ", options: { fontSize: 13, color: C.text } },
    { text: "7.5%", options: { fontSize: 16, color: C.red, bold: true } },
  ], { x: 5.4, y: 3.9, w: 4, h: 0.9, fontFace: FONT_EN, valign: "middle", margin: 0 });

  s.addText("Ito S, et al. Clin J Am Soc Nephrol. 2020;15(12):1715-1727", { x: 0.4, y: 5.15, w: 6, h: 0.3, fontSize: 9, fontFace: FONT_EN, color: C.sub, italic: true, margin: 0 });
  addPageNum(s, "7");
})();

// ============================================================
// SLIDE 8: 臨床での使い分けフローチャート
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "MRA 臨床での使い分けフローチャート");

  // 3列レイアウト
  // 左列: 降圧目的
  addCard(s, 0.3, 1.15, 3.0, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.15, w: 3.0, h: 0.45, fill: { color: C.primary } });
  s.addText("降圧が主目的", { x: 0.3, y: 1.15, w: 3.0, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.8, w: 2.6, h: 0.5, fill: { color: C.light }, rectRadius: 0.08 });
  s.addText("eGFR ≥ 60\nDM腎症なし", { x: 0.5, y: 1.8, w: 2.6, h: 0.5, fontSize: 11, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0 });
  s.addText("↓", { x: 0.5, y: 2.3, w: 2.6, h: 0.3, fontSize: 16, fontFace: FONT_EN, color: C.primary, align: "center", margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 2.6, h: 0.45, fill: { color: C.primary }, rectRadius: 0.08 });
  s.addText("ミネブロ 2.5mg or セララ", { x: 0.5, y: 2.6, w: 2.6, h: 0.45, fontSize: 11, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.25, w: 2.6, h: 0.5, fill: { color: "FFF3CD" }, rectRadius: 0.08 });
  s.addText("eGFR 30-60 or\nDM+アルブミン尿", { x: 0.5, y: 3.25, w: 2.6, h: 0.5, fontSize: 11, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0 });
  s.addText("↓", { x: 0.5, y: 3.75, w: 2.6, h: 0.3, fontSize: 16, fontFace: FONT_EN, color: C.primary, align: "center", margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.05, w: 2.6, h: 0.55, fill: { color: C.primary }, rectRadius: 0.08 });
  s.addText([
    { text: "ミネブロ 1.25mg\n", options: { fontSize: 12, color: C.white, bold: true } },
    { text: "（セララは禁忌）", options: { fontSize: 10, color: C.accent } },
  ], { x: 0.5, y: 4.05, w: 2.6, h: 0.55, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  // 中列: 心不全
  addCard(s, 3.5, 1.15, 3.0, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 1.15, w: 3.0, h: 0.45, fill: { color: C.green } });
  s.addText("心不全", { x: 3.5, y: 1.15, w: 3.0, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.7, y: 1.8, w: 2.6, h: 0.5, fill: { color: C.light }, rectRadius: 0.08 });
  s.addText("HFrEF / HFpEF", { x: 3.7, y: 1.8, w: 2.6, h: 0.5, fontSize: 12, fontFace: FONT_EN, color: C.text, align: "center", valign: "middle", margin: 0 });
  s.addText("↓", { x: 3.7, y: 2.3, w: 2.6, h: 0.3, fontSize: 16, fontFace: FONT_EN, color: C.green, align: "center", margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.7, y: 2.6, w: 2.6, h: 0.55, fill: { color: C.green }, rectRadius: 0.08 });
  s.addText([
    { text: "セララ\n", options: { fontSize: 12, color: C.white, bold: true } },
    { text: "(Fantastic Four)", options: { fontSize: 10, color: "D4EDDA" } },
  ], { x: 3.7, y: 2.6, w: 2.6, h: 0.55, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.7, y: 3.4, w: 2.6, h: 0.55, fill: { color: C.green }, rectRadius: 0.08 });
  s.addText([
    { text: "ケレンディア\n", options: { fontSize: 12, color: C.white, bold: true } },
    { text: "(FINEARTS-HF)", options: { fontSize: 10, color: "D4EDDA" } },
  ], { x: 3.7, y: 3.4, w: 2.6, h: 0.55, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  s.addText("※ ミネブロは心不全適応なし", { x: 3.7, y: 4.2, w: 2.6, h: 0.3, fontSize: 10, fontFace: FONT_JP, color: C.red, align: "center", margin: 0 });

  // 右列: 腎保護
  addCard(s, 6.7, 1.15, 3.0, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 6.7, y: 1.15, w: 3.0, h: 0.45, fill: { color: C.purple } });
  s.addText("腎保護が主目的", { x: 6.7, y: 1.15, w: 3.0, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.9, y: 1.8, w: 2.6, h: 0.5, fill: { color: C.light }, rectRadius: 0.08 });
  s.addText("2型DM + CKD", { x: 6.9, y: 1.8, w: 2.6, h: 0.5, fontSize: 12, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0 });
  s.addText("↓", { x: 6.9, y: 2.3, w: 2.6, h: 0.3, fontSize: 16, fontFace: FONT_EN, color: C.purple, align: "center", margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.9, y: 2.6, w: 2.6, h: 0.55, fill: { color: C.purple }, rectRadius: 0.08 });
  s.addText([
    { text: "ケレンディア\n", options: { fontSize: 12, color: C.white, bold: true } },
    { text: "(FIDELIO/FIGARO)", options: { fontSize: 10, color: "E8D5F5" } },
  ], { x: 6.9, y: 2.6, w: 2.6, h: 0.55, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  s.addText([
    { text: "降圧効果は弱い\n", options: { fontSize: 10, color: C.text } },
    { text: "(3-7mmHg)", options: { fontSize: 10, color: C.sub } },
  ], { x: 6.9, y: 3.35, w: 2.6, h: 0.5, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  addPageNum(s, "8");
})();

// ============================================================
// SLIDE 9: 処方時の実務チェックポイント
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ミネブロ 処方時の実務チェックポイント");

  // 左カード: Kモニタリング
  addCard(s, 0.4, 1.1, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 0.45, fill: { color: C.red } });
  s.addText("K値モニタリングスケジュール", { x: 0.6, y: 1.1, w: 4, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var kSchedule = [
    { timing: "投与開始前", action: "K > 5.0 → 投与不可", color: C.red },
    { timing: "2週以内", action: "K値測定（必須）", color: C.red },
    { timing: "約1ヵ月", action: "K値測定（必須）", color: C.red },
    { timing: "以後定期的", action: "増量時は増量後2週以内にも", color: C.primary },
  ];
  for (var i = 0; i < kSchedule.length; i++) {
    var yPos = 1.7 + i * 0.45;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 1.4, h: 0.35, fill: { color: kSchedule[i].color }, rectRadius: 0.06 });
    s.addText(kSchedule[i].timing, { x: 0.6, y: yPos, w: 1.4, h: 0.35, fontSize: 10, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(kSchedule[i].action, { x: 2.1, y: yPos, w: 2.5, h: 0.35, fontSize: 11, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  }

  // 右カード: 禁忌チェック
  addCard(s, 5.2, 1.1, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.45, fill: { color: C.orange } });
  s.addText("処方前 5項目チェック", { x: 5.4, y: 1.1, w: 4, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var checks = [
    "eGFR確認（<30 → 禁忌）",
    "血清K値（>5.0 → 禁忌）",
    "DM+アルブミン尿 → 1.25mg開始",
    "他MRA・K製剤 → 併用禁忌",
    "RAS阻害薬 → 高K血症に注意",
  ];
  for (var i = 0; i < checks.length; i++) {
    var yPos = 1.7 + i * 0.42;
    s.addText((i + 1) + ".", { x: 5.4, y: yPos, w: 0.4, h: 0.35, fontSize: 13, fontFace: FONT_EN, color: C.orange, bold: true, valign: "middle", margin: 0 });
    s.addText(checks[i], { x: 5.8, y: yPos, w: 3.6, h: 0.35, fontSize: 11, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  }

  // 下部: K値異常時の対応
  addCard(s, 0.4, 3.85, 9.2, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.85, w: 9.2, h: 0.45, fill: { color: C.red } });
  s.addText("K値上昇時の対応", { x: 0.6, y: 3.85, w: 8.8, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.45, w: 4.0, h: 0.55, fill: { color: "FFF3CD" }, rectRadius: 0.08 });
  s.addText([
    { text: "K ≥ 5.5 mEq/L\n", options: { fontSize: 13, color: C.orange, bold: true } },
    { text: "→ 減量 or 休薬を検討", options: { fontSize: 12, color: C.text } },
  ], { x: 0.8, y: 4.45, w: 4.0, h: 0.55, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.4, y: 4.45, w: 4.0, h: 0.55, fill: { color: "F8D7DA" }, rectRadius: 0.08 });
  s.addText([
    { text: "K > 6.0 mEq/L\n", options: { fontSize: 13, color: C.red, bold: true } },
    { text: "→ ただちに中止", options: { fontSize: 12, color: C.red, bold: true } },
  ], { x: 5.4, y: 4.45, w: 4.0, h: 0.55, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  addPageNum(s, "9");
})();

// ============================================================
// SLIDE 10: Take Home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "ミネブロは非ステロイド型MRAで、2.5mgでエプレレノン50mgと同等の降圧効果", color: C.accent },
    { num: "2", text: "セララで禁忌の中等度腎機能障害・DM+腎症にもMRAが使える唯一の選択肢", color: C.green },
    { num: "3", text: "ESAX-DN試験でアルブミン尿の寛解率22% vs 4%、UACR 58%低下を達成", color: C.accent },
    { num: "4", text: "最重要の副作用は高カリウム血症。投与前・2週・1ヵ月のK値測定が必須", color: C.yellow },
    { num: "5", text: "MRA選択のコツ：降圧→ミネブロ、心不全→セララ、腎保護→ケレンディア", color: C.accent },
  ];

  for (var i = 0; i < messages.length; i++) {
    var yPos = 1.15 + i * 0.8;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: yPos + 0.05, w: 0.5, h: 0.5, fill: { color: messages[i].color } });
    s.addText(messages[i].num, { x: 0.6, y: yPos + 0.05, w: 0.5, h: 0.5, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(messages[i].text, { x: 1.3, y: yPos, w: 8.2, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0 });
  }

  s.addShape(pres.shapes.LINE, { x: 1, y: 5.1, w: 8, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  |  脳神経内科専門医", {
    x: 0.5, y: 5.2, w: 9, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "10");
})();

// === OUTPUT ===
var outPath = __dirname + "/ミネブロ_MRA比較_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
