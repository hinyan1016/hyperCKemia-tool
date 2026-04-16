var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "高血圧の新基準と減塩対策 ― 日本人の3人に1人が該当する沈黙の殺し屋";

// --- Color Palette (循環器・高血圧テーマ) ---
var C = {
  dark: "8B1A1A",
  primary: "C0392B",
  accent: "E74C3C",
  light: "FDEDEC",
  warmBg: "FDF5F4",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "DC3545",
  orange: "E8850C",
  yellow: "F5C518",
  green: "28A745",
  teal: "17A2B8",
  blue: "2C5AA0",
  navy: "1B3A5C",
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

  s.addText("一般向け医学解説", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("高血圧の新基準と減塩対策", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.2,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 日本人の3人に1人が該当する「沈黙の殺し屋」 ―", {
    x: 0.8, y: 2.5, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.3, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  var keys = [
    { text: "4,300万人の衝撃", color: C.primary },
    { text: "JSH2025 新基準", color: C.orange },
    { text: "今日からの減塩術", color: C.green },
  ];
  keys.forEach(function(k, i) {
    var xPos = 0.5 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.5, w: 2.8, h: 0.55, fill: { color: k.color } });
    s.addText(k.text, {
      x: xPos, y: 3.5, w: 2.8, h: 0.55,
      fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.3, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 日本の高血圧の現状
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "日本の高血圧の現状 ― 衝撃の数字");

  // 大きな数字カード
  addCard(s, 0.4, 1.1, 4.4, 2.0);
  s.addText("高血圧有病者数", {
    x: 0.6, y: 1.2, w: 4.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
  s.addText("約4,300万人", {
    x: 0.6, y: 1.6, w: 4.0, h: 0.8,
    fontSize: 36, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0,
  });
  s.addText("日本人の約3人に1人", {
    x: 0.6, y: 2.5, w: 4.0, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.text, align: "center", bold: true, margin: 0,
  });

  addCard(s, 5.2, 1.1, 4.4, 2.0);
  s.addText("高血圧に起因する死亡", {
    x: 5.4, y: 1.2, w: 4.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
  s.addText("年間 約10万人", {
    x: 5.4, y: 1.6, w: 4.0, h: 0.8,
    fontSize: 36, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0,
  });
  s.addText("脳卒中・心臓病含む", {
    x: 5.4, y: 2.5, w: 4.0, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.text, align: "center", bold: true, margin: 0,
  });

  // 下部: 寄与割合
  addCard(s, 0.4, 3.4, 9.2, 1.9);
  s.addText("高血圧は血管病死亡の最大のリスク因子", {
    x: 0.6, y: 3.5, w: 8.8, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });

  var risks = [
    { label: "全血管病死亡", pct: "50%", color: C.primary },
    { label: "脳卒中死亡", pct: "52%", color: C.accent },
    { label: "冠動脈疾患死亡", pct: "59%", color: C.red },
  ];
  risks.forEach(function(r, i) {
    var xPos = 0.8 + i * 3.0;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.15, w: 2.6, h: 0.9, fill: { color: r.color } });
    s.addText(r.label, {
      x: xPos, y: 4.15, w: 2.6, h: 0.4,
      fontSize: 13, fontFace: FONT_JP, color: C.white, align: "center", valign: "middle", margin: 0,
    });
    s.addText(r.pct, {
      x: xPos, y: 4.5, w: 2.6, h: 0.55,
      fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  addPageNum(s, "2");
})();

// ============================================================
// SLIDE 3: コントロール不良の3,100万人
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "コントロール不良の3,100万人");

  // 円グラフ的表示（テキストベース）
  var segments = [
    { label: "高血圧と認識していない", num: "1,400万人", pct: "33%", color: C.red },
    { label: "自覚はあるが未治療", num: "450万人", pct: "10%", color: C.orange },
    { label: "治療中だがコントロール不良", num: "1,250万人", pct: "29%", color: C.yellow },
    { label: "適切にコントロール", num: "1,200万人", pct: "28%", color: C.green },
  ];

  segments.forEach(function(seg, i) {
    var yPos = 1.15 + i * 1.05;
    addCard(s, 0.4, yPos, 9.2, 0.9);

    // 色バー
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 0.15, h: 0.9, fill: { color: seg.color } });

    s.addText(seg.label, {
      x: 0.8, y: yPos + 0.05, w: 4.5, h: 0.4,
      fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0,
    });
    s.addText(seg.num, {
      x: 0.8, y: yPos + 0.45, w: 4.5, h: 0.35,
      fontSize: 14, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0,
    });

    // パーセンテージバー
    var barMaxW = 3.5;
    var barW = barMaxW * (parseInt(seg.pct) / 100);
    s.addShape(pres.shapes.RECTANGLE, { x: 5.6, y: yPos + 0.25, w: barMaxW, h: 0.4, fill: { color: "EEEEEE" } });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.6, y: yPos + 0.25, w: barW, h: 0.4, fill: { color: seg.color } });
    s.addText(seg.pct, {
      x: 5.6, y: yPos + 0.25, w: barMaxW, h: 0.4,
      fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.15, w: 9.2, h: 0.4, fill: { color: C.light } });
  s.addText("→ 高血圧の方の3人に1人は自分が高血圧であることを知らない", {
    x: 0.5, y: 5.15, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "3");
})();

// ============================================================
// SLIDE 4: 2025年の新基準（JSH2025）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "2025年の新基準 ― JSH2025ガイドライン");

  // メインメッセージ
  addCard(s, 0.4, 1.1, 9.2, 1.2);
  s.addText("最大の変更点", {
    x: 0.6, y: 1.15, w: 2.0, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });
  s.addText("全年齢で降圧目標を統一", {
    x: 0.6, y: 1.5, w: 8.8, h: 0.7,
    fontSize: 26, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  // 比較テーブル風
  // ヘッダー行
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.5, w: 9.2, h: 0.5, fill: { color: C.primary } });
  s.addText("項目", { x: 0.5, y: 2.5, w: 3.0, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: [0, 0, 0, 8] });
  s.addText("JSH2019（旧）", { x: 3.5, y: 2.5, w: 2.8, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("JSH2025（新）", { x: 6.5, y: 2.5, w: 3.1, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var rows = [
    { label: "75歳未満の降圧目標", old: "130/80 mmHg未満", nw: "130/80 mmHg未満" },
    { label: "75歳以上の降圧目標", old: "140/90 mmHg未満", nw: "130/80 mmHg未満 ↓" },
    { label: "家庭血圧の目標", old: "125/75（75歳未満のみ）", nw: "全年齢 125/75 未満" },
  ];

  rows.forEach(function(r, i) {
    var yPos = 3.0 + i * 0.7;
    var bgColor = i % 2 === 0 ? "F8F9FA" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 9.2, h: 0.7, fill: { color: bgColor } });
    s.addText(r.label, { x: 0.5, y: yPos, w: 3.0, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: [0, 0, 0, 8] });
    s.addText(r.old, { x: 3.5, y: yPos, w: 2.8, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", valign: "middle", margin: 0 });

    var isChanged = r.nw.indexOf("↓") >= 0;
    s.addText(r.nw, {
      x: 6.5, y: yPos, w: 3.1, h: 0.7,
      fontSize: 14, fontFace: FONT_JP, color: isChanged ? C.red : C.text, bold: isChanged, align: "center", valign: "middle", margin: 0,
    });
  });

  // 下部メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.1, w: 9.2, h: 0.5, fill: { color: C.light } });
  s.addText("診察室血圧 130/80 mmHg未満 ／ 家庭血圧 125/75 mmHg未満 を全年齢で目指す", {
    x: 0.5, y: 5.1, w: 9.0, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "4");
})();

// ============================================================
// SLIDE 5: 血圧の分類とリスク
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "血圧の分類とリスク");

  // ヘッダー行
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 9.2, h: 0.5, fill: { color: C.primary } });
  s.addText("分類", { x: 0.5, y: 1.1, w: 2.4, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: [0, 0, 0, 8] });
  s.addText("収縮期（上）", { x: 2.9, y: 1.1, w: 2.0, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("拡張期（下）", { x: 5.0, y: 1.1, w: 2.0, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("心血管リスク", { x: 7.1, y: 1.1, w: 2.4, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var levels = [
    { cls: "正常血圧", sys: "< 120", dia: "< 80", risk: "基準（1倍）", barColor: C.green },
    { cls: "正常高値血圧", sys: "120 - 129", dia: "< 80", risk: "やや注意", barColor: C.teal },
    { cls: "高値血圧", sys: "130 - 139", dia: "80 - 89", risk: "約1.7倍", barColor: C.yellow },
    { cls: "I度高血圧", sys: "140 - 159", dia: "90 - 99", risk: "約3.3倍", barColor: C.orange },
    { cls: "II度高血圧", sys: "160 - 179", dia: "100 - 109", risk: "高リスク", barColor: C.accent },
    { cls: "III度高血圧", sys: "180 以上", dia: "110 以上", risk: "約8.5倍", barColor: C.red },
  ];

  levels.forEach(function(lv, i) {
    var yPos = 1.6 + i * 0.6;
    var bgColor = i % 2 === 0 ? "F8F9FA" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 9.2, h: 0.6, fill: { color: bgColor } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 0.12, h: 0.6, fill: { color: lv.barColor } });
    s.addText(lv.cls, { x: 0.65, y: yPos, w: 2.2, h: 0.6, fontSize: 13, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(lv.sys, { x: 2.9, y: yPos, w: 2.0, h: 0.6, fontSize: 13, fontFace: FONT_EN, color: C.text, align: "center", valign: "middle", margin: 0 });
    s.addText(lv.dia, { x: 5.0, y: yPos, w: 2.0, h: 0.6, fontSize: 13, fontFace: FONT_EN, color: C.text, align: "center", valign: "middle", margin: 0 });
    s.addText(lv.risk, { x: 7.1, y: yPos, w: 2.4, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: lv.barColor, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.2, w: 9.2, h: 0.35, fill: { color: C.light } });
  s.addText("血圧が上がるほどリスクは指数関数的に増加する", {
    x: 0.5, y: 5.2, w: 9.0, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "5");
})();

// ============================================================
// SLIDE 6: なぜ日本は高血圧大国なのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "なぜ日本は高血圧大国なのか");

  // 食塩摂取量の大きな数字
  addCard(s, 0.4, 1.1, 4.4, 2.3);
  s.addText("日本人の平均食塩摂取量（2023年）", {
    x: 0.6, y: 1.2, w: 4.0, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
  s.addText("9.8 g/日", {
    x: 0.6, y: 1.7, w: 4.0, h: 0.8,
    fontSize: 40, fontFace: FONT_EN, color: C.red, bold: true, align: "center", margin: 0,
  });
  s.addText("男性 10.7g ／ 女性 9.1g", {
    x: 0.6, y: 2.5, w: 4.0, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 2.9, w: 3.4, h: 0.35, fill: { color: C.green } });
  s.addText("目標: 6g未満（日本高血圧学会）", {
    x: 1.0, y: 2.9, w: 3.4, h: 0.35,
    fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 右側: 血圧の現状
  addCard(s, 5.2, 1.1, 4.4, 2.3);
  s.addText("日本人男性の平均収縮期血圧", {
    x: 5.4, y: 1.2, w: 4.0, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
  s.addText("131.6 mmHg", {
    x: 5.4, y: 1.7, w: 4.0, h: 0.8,
    fontSize: 36, fontFace: FONT_EN, color: C.orange, bold: true, align: "center", margin: 0,
  });
  s.addText("すでに目標130を超過", {
    x: 5.4, y: 2.5, w: 4.0, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.red, align: "center", bold: true, margin: 0,
  });
  s.addText("140以上: 男性27.5% 女性22.5%", {
    x: 5.4, y: 2.9, w: 4.0, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", margin: 0,
  });

  // 下部: 塩分の多い食文化
  addCard(s, 0.4, 3.6, 9.2, 1.8);
  s.addText("塩分の多い日本の食文化", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });

  var foods = [
    { name: "味噌汁1杯", salt: "約1.5g" },
    { name: "ラーメン（汁込み）", salt: "約5-7g" },
    { name: "梅干し1個", salt: "約2g" },
    { name: "たくあん3切れ", salt: "約1.3g" },
    { name: "醤油 大さじ1", salt: "約2.6g" },
  ];

  foods.forEach(function(f, i) {
    var xPos = 0.6 + i * 1.8;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.2, w: 1.6, h: 1.0, fill: { color: C.light } });
    s.addText(f.name, {
      x: xPos, y: 4.25, w: 1.6, h: 0.45,
      fontSize: 11, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0,
    });
    s.addText(f.salt, {
      x: xPos, y: 4.7, w: 1.6, h: 0.45,
      fontSize: 14, fontFace: FONT_EN, color: C.red, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  addPageNum(s, "6");
})();

// ============================================================
// SLIDE 7: 沈黙の殺し屋 ― 合併症
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "沈黙の殺し屋 ― 高血圧が引き起こす病気");

  var organs = [
    { organ: "脳", disease: "脳卒中\n（脳梗塞・脳出血）", mech: "血管が詰まる・破れる", color: C.red },
    { organ: "心臓", disease: "心筋梗塞\n心不全", mech: "冠動脈の動脈硬化", color: C.accent },
    { organ: "腎臓", disease: "慢性腎臓病\n（CKD）", mech: "糸球体の血管障害", color: C.orange },
    { organ: "目", disease: "高血圧性\n網膜症", mech: "網膜の血管障害", color: C.yellow },
    { organ: "血管", disease: "大動脈解離\n大動脈瘤", mech: "血管壁への持続的負荷", color: C.primary },
  ];

  organs.forEach(function(o, i) {
    var xPos = 0.3 + i * 1.9;
    addCard(s, xPos, 1.1, 1.7, 3.5);

    // 臓器名ヘッダー
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 1.7, h: 0.6, fill: { color: o.color } });
    s.addText(o.organ, {
      x: xPos, y: 1.1, w: 1.7, h: 0.6,
      fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });

    // 疾患名
    s.addText(o.disease, {
      x: xPos + 0.1, y: 1.9, w: 1.5, h: 1.2,
      fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0,
      lineSpacingMultiple: 1.3,
    });

    // メカニズム
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.1, y: 3.3, w: 1.5, h: 0.05, fill: { color: "DDDDDD" } });
    s.addText(o.mech, {
      x: xPos + 0.1, y: 3.5, w: 1.5, h: 0.9,
      fontSize: 11, fontFace: FONT_JP, color: C.sub, align: "center", valign: "top", margin: 0,
      lineSpacingMultiple: 1.3,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.8, w: 9.2, h: 0.7, fill: { color: C.light } });
  s.addText("高血圧は「自覚症状がないまま」全身の血管を傷つけ続ける\n→ だからこそ定期的な血圧測定が重要", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.7,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.4,
  });

  addPageNum(s, "7");
})();

// ============================================================
// SLIDE 8: 家庭血圧の正しい測り方
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "家庭血圧を測ろう ― 正しい測定法");

  // 左側: 測定ポイント
  addCard(s, 0.4, 1.1, 5.0, 4.3);
  s.addText("正しい測り方 7つのポイント", {
    x: 0.6, y: 1.2, w: 4.6, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  var points = [
    "上腕式の血圧計を使う",
    "朝: 起床後1時間以内・排尿後・朝食前",
    "夜: 就寝前のリラックス時",
    "1-2分安静にしてから測る",
    "カフを心臓の高さに合わせる",
    "各回2回測定し平均値を記録",
    "結果を手帳やアプリに記録",
  ];

  points.forEach(function(p, i) {
    var yPos = 1.75 + i * 0.5;
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.07, w: 0.3, h: 0.3, fill: { color: C.primary } });
    s.addText(String(i + 1), {
      x: 0.7, y: yPos + 0.07, w: 0.3, h: 0.3,
      fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(p, {
      x: 1.15, y: yPos, w: 4.1, h: 0.45,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  // 右側: 白衣高血圧と仮面高血圧
  addCard(s, 5.6, 1.1, 4.0, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.6, y: 1.1, w: 4.0, h: 0.5, fill: { color: C.teal } });
  s.addText("白衣高血圧", {
    x: 5.7, y: 1.1, w: 3.8, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });
  s.addText("病院では高いが\n自宅では正常\n→ 治療不要なことが多い", {
    x: 5.8, y: 1.7, w: 3.6, h: 1.2,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
    lineSpacingMultiple: 1.4,
  });

  addCard(s, 5.6, 3.3, 4.0, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.6, y: 3.3, w: 4.0, h: 0.5, fill: { color: C.red } });
  s.addText("仮面高血圧", {
    x: 5.7, y: 3.3, w: 3.8, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });
  s.addText("病院では正常なのに\n自宅では高い\n→ 見逃されやすく危険！", {
    x: 5.8, y: 3.9, w: 3.6, h: 1.2,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
    lineSpacingMultiple: 1.4,
  });

  addPageNum(s, "8");
})();

// ============================================================
// SLIDE 9: 減塩のコツ + DASH食
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "減塩のコツとDASH食 ― 今日からできる対策");

  // 左側: 減塩のコツ
  addCard(s, 0.4, 1.1, 4.6, 3.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.6, h: 0.5, fill: { color: C.primary } });
  s.addText("すぐに始められる減塩のコツ", {
    x: 0.5, y: 1.1, w: 4.4, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });

  var tips = [
    "味噌汁は1日1杯まで",
    "醤油はかけずにつける",
    "麺類の汁は残す",
    "だし・酢・レモンで味付け",
    "加工食品の食塩相当量を確認",
    "外食時は「薄味で」と注文",
  ];

  tips.forEach(function(t, i) {
    var yPos = 1.75 + i * 0.38;
    s.addText("✓ " + t, {
      x: 0.7, y: yPos, w: 4.1, h: 0.35,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  // 右側: DASH食
  addCard(s, 5.2, 1.1, 4.4, 3.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.5, fill: { color: C.green } });
  s.addText("DASH食とは？", {
    x: 5.3, y: 1.1, w: 4.2, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });

  s.addText("積極的に摂るもの", {
    x: 5.4, y: 1.75, w: 4.0, h: 0.3,
    fontSize: 12, fontFace: FONT_JP, color: C.green, bold: true, margin: 0,
  });
  s.addText("  野菜・果物・低脂肪乳製品\n  全粒穀物・ナッツ・魚", {
    x: 5.4, y: 2.05, w: 4.0, h: 0.7,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  s.addText("控えるもの", {
    x: 5.4, y: 2.8, w: 4.0, h: 0.3,
    fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });
  s.addText("  赤肉・砂糖・清涼飲料水\n  飽和脂肪酸・加工食品", {
    x: 5.4, y: 3.1, w: 4.0, h: 0.7,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // 下部: DASH食+減塩の効果
  addCard(s, 0.4, 4.3, 9.2, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.3, w: 0.15, h: 1.2, fill: { color: C.green } });
  s.addText("DASH食 + 減塩 = 最強の組み合わせ", {
    x: 0.8, y: 4.35, w: 5.0, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });
  s.addText("収縮期血圧 最大11 mmHg以上低下", {
    x: 6.0, y: 4.35, w: 3.4, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });
  s.addText("食塩を1g減らすだけで血圧は約1 mmHg下がる。カリウムが豊富なバナナ・ほうれん草・納豆も効果的。", {
    x: 0.8, y: 4.8, w: 8.6, h: 0.6,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  addPageNum(s, "9");
})();

// ============================================================
// SLIDE 10: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.25, w: 9.0, h: 0.6,
    fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, margin: 0,
  });

  var messages = [
    { num: "1", text: "日本は高血圧大国", detail: "4,300万人が該当、3,100万人がコントロール不良", color: C.accent },
    { num: "2", text: "2025年の新基準", detail: "全年齢で診察室130/80、家庭125/75 mmHg未満", color: C.orange },
    { num: "3", text: "高血圧は最大のリスク因子", detail: "脳卒中・心臓病死亡の約50%に関与", color: C.yellow },
    { num: "4", text: "家庭血圧の毎日測定", detail: "朝・晩2回ずつ、記録して受診時に持参", color: C.teal },
    { num: "5", text: "減塩が最重要", detail: "1日6g未満を目標に。DASH食で効果倍増", color: C.green },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.1 + i * 0.85;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.7, fill: { color: "3D1515" } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 0.1, h: 0.7, fill: { color: m.color } });

    s.addShape(pres.shapes.OVAL, { x: 0.8, y: yPos + 0.12, w: 0.45, h: 0.45, fill: { color: m.color } });
    s.addText(m.num, {
      x: 0.8, y: yPos + 0.12, w: 0.45, h: 0.45,
      fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });

    s.addText(m.text, {
      x: 1.5, y: yPos + 0.02, w: 4.0, h: 0.35,
      fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
    });
    s.addText(m.detail, {
      x: 1.5, y: yPos + 0.35, w: 7.8, h: 0.3,
      fontSize: 12, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0,
    });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.4, w: 6, h: 0, line: { color: C.accent, width: 1.5 } });

  s.addText("まずは家庭血圧計を1台用意して、毎日測ることから始めましょう", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });

  addPageNum(s, "10");
})();

// ============================================================
// 保存
// ============================================================
var outputPath = "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/高血圧/高血圧_スライド.pptx";
pres.writeFile({ fileName: outputPath }).then(function() {
  console.log("PPTX saved: " + outputPath);
}).catch(function(err) {
  console.error("Error:", err);
});
