var pptxgen = require("pptxgenjs");
var pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "医知創造ラボ";
pres.title = "甲状腺機能異常の鑑別診断";

var C = {
  navy: "1B3A5C", blue: "2C5AA0", ltBlue: "E8F0FA", bg: "F0F4F8",
  white: "FFFFFF", text: "333333", red: "DC3545", yellow: "F5C518",
  green: "28A745", orange: "E8850C", ltGray: "F8F9FA", gray: "555555", subText: "666666",
  dkRed: "8B1A1A", ltRed: "FDE8E8", ltGreen: "E6F7EC", ltYellow: "FFF8E1",
  ltOrange: "FFF3E0", purple: "6F42C1", ltPurple: "F0E6FF",
};
var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var mkShadow = function() { return { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 }; };

// ── フッター ──
function addFooter(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.15, w: 10, h: 0.475, fill: { color: C.navy } });
  slide.addText("甲状腺機能異常の鑑別診断", { x: 0.5, y: 5.15, w: 6, h: 0.475, fontSize: 18, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
}

// ── セクションタイトル ──
function addSectionTitle(slide, title) {
  slide.addText(title, { x: 0.6, y: 0.35, w: 8.8, h: 0.65, fontSize: 28, fontFace: FONT_JP, color: C.navy, bold: true, margin: 0 });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.05, w: 1.2, h: 0.06, fill: { color: C.blue } });
}

// ── カード描画ヘルパー ──
function addCard(slide, x, y, w, h, bgColor, borderColor) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: w, h: h,
    fill: { color: bgColor },
    line: { color: borderColor, width: 1.5 },
    rectRadius: 0.1,
    shadow: mkShadow()
  });
}

// ── 番号付きバッジ ──
function addBadge(slide, x, y, label, bgColor) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 0.45, h: 0.4,
    fill: { color: bgColor },
    rectRadius: 0.08
  });
  slide.addText(label, { x: x, y: y, w: 0.45, h: 0.4, fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
}

// ══════════════════════════════════════════════════════
// Slide 1: タイトルスライド
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.navy };

  // 装飾ライン上部
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.blue } });

  // メインタイトル
  s.addText("甲状腺機能異常の鑑別診断", {
    x: 0.5, y: 1.5, w: 9, h: 1.3,
    fontSize: 48, fontFace: FONT_JP, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });

  // 装飾ライン
  s.addShape(pres.shapes.RECTANGLE, { x: 3, y: 2.95, w: 4, h: 0.05, fill: { color: C.yellow } });

  // サブタイトル
  s.addText("― 7つのステップで系統的に考える ―", {
    x: 0.5, y: 3.2, w: 9, h: 0.8,
    fontSize: 26, fontFace: FONT_JP, color: C.ltBlue,
    align: "center", valign: "middle", margin: 0
  });

  // チャンネル名
  s.addText("医知創造ラボ", {
    x: 0.5, y: 4.5, w: 9, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.ltBlue,
    align: "center", valign: "middle", margin: 0
  });

  // 底部ライン
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.06, fill: { color: C.blue } });
})();

// ══════════════════════════════════════════════════════
// Slide 2: 本日の内容（7ステップ）
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "本日の内容");
  addFooter(s);

  var steps = [
    { id: "1", label: "TSHの解釈", color: C.blue },
    { id: "2", label: "TSH×FT4/FT3\nパターン分類", color: C.blue },
    { id: "3", label: "破壊性 vs\n産生過剰", color: C.orange },
    { id: "4", label: "中毒症の\n主要疾患", color: C.red },
    { id: "5", label: "機能低下症の\n鑑別", color: C.green },
    { id: "6", label: "潜在性\n甲状腺異常", color: C.purple },
    { id: "7", label: "Red Flags", color: C.dkRed }
  ];

  var startX = 0.3;
  var stepW = 1.2;
  var gap = 0.1;

  for (var i = 0; i < steps.length; i++) {
    var sx = startX + i * (stepW + gap);
    var sy = 1.5;

    // ステップカード
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: sx, y: sy, w: stepW, h: 2.8,
      fill: { color: C.white },
      line: { color: steps[i].color, width: 2 },
      rectRadius: 0.1,
      shadow: mkShadow()
    });

    // バッジ
    s.addShape(pres.shapes.OVAL, {
      x: sx + (stepW - 0.45) / 2, y: sy + 0.2, w: 0.45, h: 0.45,
      fill: { color: steps[i].color }
    });
    s.addText(steps[i].id, {
      x: sx + (stepW - 0.45) / 2, y: sy + 0.2, w: 0.45, h: 0.45,
      fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0
    });

    // ステップ名
    s.addText(steps[i].label, {
      x: sx + 0.05, y: sy + 0.85, w: stepW - 0.1, h: 1.6,
      fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3
    });

    // 矢印（最後以外）
    if (i < steps.length - 1) {
      s.addText("▶", {
        x: sx + stepW, y: sy + 1.0, w: gap, h: 0.6,
        fontSize: 10, fontFace: FONT_EN, color: C.gray,
        align: "center", valign: "middle", margin: 0
      });
    }
  }
})();

// ══════════════════════════════════════════════════════
// Slide 3: Step 1 — TSHの解釈
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 1：TSHの解釈");
  addFooter(s);

  // 中央フロー: TSH
  addCard(s, 3.5, 1.3, 3, 0.7, C.navy, C.navy);
  s.addText("TSH測定", {
    x: 3.5, y: 1.3, w: 3, h: 0.7,
    fontSize: 22, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // 矢印（左）
  s.addText("▼", {
    x: 1.8, y: 2.1, w: 1.5, h: 0.4,
    fontSize: 18, fontFace: FONT_EN, color: C.red,
    align: "center", valign: "middle", margin: 0
  });

  // 矢印（右）
  s.addText("▼", {
    x: 6.7, y: 2.1, w: 1.5, h: 0.4,
    fontSize: 18, fontFace: FONT_EN, color: C.blue,
    align: "center", valign: "middle", margin: 0
  });

  // TSH低値ボックス
  addCard(s, 0.5, 2.5, 4, 1.0, C.ltRed, C.red);
  s.addText("TSH 低値", {
    x: 0.5, y: 2.5, w: 4, h: 0.45,
    fontSize: 20, fontFace: FONT_JP, color: C.dkRed, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("→ 甲状腺中毒症を疑う", {
    x: 0.5, y: 2.95, w: 4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0
  });

  // TSH高値ボックス
  addCard(s, 5.5, 2.5, 4, 1.0, C.ltBlue, C.blue);
  s.addText("TSH 高値", {
    x: 5.5, y: 2.5, w: 4, h: 0.45,
    fontSize: 20, fontFace: FONT_JP, color: C.navy, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("→ 甲状腺機能低下症を疑う", {
    x: 5.5, y: 2.95, w: 4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0
  });

  // 例外ボックス
  addCard(s, 0.6, 3.8, 8.8, 1.1, C.ltYellow, C.orange);
  s.addText("⚠ 例外に注意", {
    x: 0.8, y: 3.85, w: 3, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0
  });
  s.addText("・中枢性甲状腺機能低下症：TSH低値〜正常 + FT4低値\n・TSH産生腫瘍：TSH高値 + FT4/FT3高値（不適切TSH分泌）", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.65,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4
  });
})();

// ══════════════════════════════════════════════════════
// Slide 4: Step 2 — TSH × FT4/FT3パターン分類（表）
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 2：TSH × FT4/FT3 パターン分類");
  addFooter(s);

  var hOpt = function(text) {
    return { text: text, options: { fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } };
  };
  var cOpt = function(text, bg) {
    return { text: text, options: { fontSize: 12, fontFace: FONT_JP, color: C.text, fill: { color: bg }, align: "center", valign: "middle" } };
  };

  var rows = [
    [hOpt("パターン"), hOpt("TSH"), hOpt("FT4"), hOpt("FT3"), hOpt("考える疾患")],
    [cOpt("顕性中毒症", C.ltRed), cOpt("↓", C.ltRed), cOpt("↑", C.ltRed), cOpt("↑", C.ltRed), cOpt("バセドウ病, 中毒性結節性甲状腺腫", C.ltRed)],
    [cOpt("T3中毒症", C.ltOrange), cOpt("↓", C.ltOrange), cOpt("正常", C.ltOrange), cOpt("↑", C.ltOrange), cOpt("バセドウ病初期, 自律性結節", C.ltOrange)],
    [cOpt("潜在性中毒症", C.ltYellow), cOpt("↓", C.ltYellow), cOpt("正常", C.ltYellow), cOpt("正常", C.ltYellow), cOpt("軽症バセドウ病, 自律性結節", C.ltYellow)],
    [cOpt("顕性低下症", C.ltBlue), cOpt("↑", C.ltBlue), cOpt("↓", C.ltBlue), cOpt("↓〜正常", C.ltBlue), cOpt("橋本病, ヨード過剰, 薬剤性", C.ltBlue)],
    [cOpt("潜在性低下症", C.white), cOpt("↑", C.white), cOpt("正常", C.white), cOpt("正常", C.white), cOpt("橋本病（初期〜軽症）", C.white)],
    [cOpt("中枢性低下症", C.ltPurple), cOpt("低〜正常", C.ltPurple), cOpt("↓", C.ltPurple), cOpt("↓", C.ltPurple), cOpt("下垂体腫瘍, シーハン症候群", C.ltPurple)],
    [cOpt("不適切TSH分泌", C.ltGreen), cOpt("正常〜↑", C.ltGreen), cOpt("↑", C.ltGreen), cOpt("↑", C.ltGreen), cOpt("TSH産生腫瘍, 甲状腺ホルモン不応症", C.ltGreen)]
  ];

  s.addTable(rows, {
    x: 0.3, y: 1.2, w: 9.4,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [1.5, 0.9, 0.9, 1.0, 5.1],
    rowH: [0.4, 0.46, 0.46, 0.46, 0.46, 0.46, 0.46, 0.46],
    margin: [2, 4, 2, 4],
    autoPage: false
  });
})();

// ══════════════════════════════════════════════════════
// Slide 5: 甲状腺中毒症の分類
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "甲状腺中毒症の分類");
  addFooter(s);

  // 上部: 甲状腺中毒症
  addCard(s, 2.5, 1.3, 5, 0.65, C.navy, C.navy);
  s.addText("甲状腺中毒症（Thyrotoxicosis）", {
    x: 2.5, y: 1.3, w: 5, h: 0.65,
    fontSize: 20, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // 分岐矢印
  s.addText("▼", {
    x: 2.0, y: 2.05, w: 1.5, h: 0.35,
    fontSize: 16, fontFace: FONT_EN, color: C.red,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("▼", {
    x: 6.5, y: 2.05, w: 1.5, h: 0.35,
    fontSize: 16, fontFace: FONT_EN, color: C.orange,
    align: "center", valign: "middle", margin: 0
  });

  // 左: 産生過剰型
  var lx = 0.4, ly = 2.45, lw = 4.4, lh = 2.0;
  addCard(s, lx, ly, lw, lh, C.white, C.red);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: lx, y: ly, w: lw, h: 0.5,
    fill: { color: C.red }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: lx, y: ly + 0.3, w: lw, h: 0.2,
    fill: { color: C.red }
  });
  s.addText("産生過剰型（甲状腺機能亢進症）", {
    x: lx + 0.1, y: ly, w: lw - 0.2, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  var prodDx = [
    "バセドウ病",
    "中毒性多結節性甲状腺腫",
    "中毒性腺腫（Plummer病）",
    "TSH産生腫瘍"
  ];
  for (var i = 0; i < prodDx.length; i++) {
    s.addText("・" + prodDx[i], {
      x: lx + 0.3, y: ly + 0.55 + i * 0.34, w: lw - 0.5, h: 0.32,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
    });
  }

  // 右: 破壊性
  var rx = 5.2, rw = 4.4;
  addCard(s, rx, ly, rw, lh, C.white, C.orange);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: rx, y: ly, w: rw, h: 0.5,
    fill: { color: C.orange }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: rx, y: ly + 0.3, w: rw, h: 0.2,
    fill: { color: C.orange }
  });
  s.addText("破壊性甲状腺炎", {
    x: rx + 0.1, y: ly, w: rw - 0.2, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  var destDx = [
    "亜急性甲状腺炎",
    "無痛性甲状腺炎",
    "産後甲状腺炎",
    "薬剤性（アミオダロン AIT2）"
  ];
  for (var i = 0; i < destDx.length; i++) {
    s.addText("・" + destDx[i], {
      x: rx + 0.3, y: ly + 0.55 + i * 0.34, w: rw - 0.5, h: 0.32,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
    });
  }

  // ポイント
  addCard(s, 0.4, 4.6, 4.4, 0.35, C.ltRed, C.red);
  s.addText("→ 抗甲状腺薬が有効", {
    x: 0.4, y: 4.6, w: 4.4, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.dkRed, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  addCard(s, 5.2, 4.6, 4.4, 0.35, C.ltOrange, C.orange);
  s.addText("→ 抗甲状腺薬は無効（対症療法）", {
    x: 5.2, y: 4.6, w: 4.4, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true,
    align: "center", valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 6: Step 3 — 破壊性 vs 産生過剰の鑑別表
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 3：破壊性 vs 産生過剰の鑑別");
  addFooter(s);

  var hOpt = function(text) {
    return { text: text, options: { fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } };
  };
  var cOpt = function(text, bg) {
    return { text: text, options: { fontSize: 13, fontFace: FONT_JP, color: C.text, fill: { color: bg }, align: "center", valign: "middle" } };
  };

  var rows = [
    [hOpt("検査"), hOpt("産生過剰型"), hOpt("破壊性")],
    [cOpt("TRAb", C.ltBlue), cOpt("陽性（バセドウ病）", C.ltRed), cOpt("陰性", C.ltGreen)],
    [cOpt("RI摂取率\n（シンチ）", C.white), cOpt("上昇\n（びまん性 or 結節性）", C.ltRed), cOpt("著明に低下", C.ltGreen)],
    [cOpt("エコー血流", C.ltBlue), cOpt("増加\n（thyroid inferno）", C.ltRed), cOpt("正常〜低下", C.ltGreen)],
    [cOpt("サイログロブリン", C.white), cOpt("軽度上昇", C.ltYellow), cOpt("著明に上昇", C.ltOrange)],
    [cOpt("経過", C.ltBlue), cOpt("持続性", C.ltRed), cOpt("一過性\n（自然軽快）", C.ltGreen)]
  ];

  s.addTable(rows, {
    x: 0.6, y: 1.25, w: 8.8,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [2.0, 3.4, 3.4],
    rowH: [0.5, 0.55, 0.65, 0.65, 0.55, 0.55],
    margin: [4, 6, 4, 6]
  });

  // 補足
  addCard(s, 0.6, 4.65, 8.8, 0.4, C.ltYellow, C.orange);
  s.addText("💡 実臨床ではTRAb＋エコー血流で多くの場合鑑別可能。迷う場合はシンチが決め手。", {
    x: 0.8, y: 4.65, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 7: バセドウ病のポイント
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "バセドウ病のポイント");
  addFooter(s);

  // 病態
  addCard(s, 0.6, 1.25, 4.1, 1.5, C.white, C.blue);
  s.addText("病態", {
    x: 0.8, y: 1.3, w: 3, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("・TRAbがTSH受容体を持続刺激\n・びまん性甲状腺腫大\n・眼球突出（Graves眼症）\n・前脛骨粘液水腫", {
    x: 0.8, y: 1.65, w: 3.7, h: 1.0,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.35
  });

  // 診断
  addCard(s, 5.3, 1.25, 4.1, 1.5, C.white, C.green);
  s.addText("診断", {
    x: 5.5, y: 1.3, w: 3, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.green, bold: true, margin: 0
  });
  s.addText("・甲状腺中毒症の存在\n・TRAb陽性 or TSAb陽性\n・シンチでびまん性取り込み増加\n  （上記いずれかで確定）", {
    x: 5.5, y: 1.65, w: 3.7, h: 1.0,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.35
  });

  // 治療
  addCard(s, 0.6, 2.9, 8.8, 1.2, C.white, C.navy);
  s.addText("治療 3つのオプション", {
    x: 0.8, y: 2.92, w: 4, h: 0.32,
    fontSize: 15, fontFace: FONT_JP, color: C.navy, bold: true, margin: 0
  });

  var treatments = [
    { num: "1", text: "抗甲状腺薬（MMI第一選択, PTUは妊娠初期・MMI不耐時）", color: C.blue },
    { num: "2", text: "放射性ヨード治療（131I）", color: C.green },
    { num: "3", text: "手術（甲状腺全摘/亜全摘）", color: C.purple }
  ];
  for (var i = 0; i < treatments.length; i++) {
    addBadge(s, 0.85, 3.28 + i * 0.28, treatments[i].num, treatments[i].color);
    s.addText(treatments[i].text, {
      x: 1.4, y: 3.28 + i * 0.28, w: 7.8, h: 0.28,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
    });
  }

  // 注意点
  addCard(s, 0.6, 4.25, 8.8, 0.7, C.ltRed, C.red);
  s.addText("⚠ 無顆粒球症に注意（0.1〜0.5%）— 発熱・咽頭痛→直ちに白血球分画を確認", {
    x: 0.8, y: 4.28, w: 8.4, h: 0.32,
    fontSize: 13, fontFace: FONT_JP, color: C.dkRed, bold: true, margin: 0
  });
  s.addText("投与開始3ヶ月以内に多い。患者への事前説明が必須。", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.28,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 8: 亜急性甲状腺炎のポイント
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "亜急性甲状腺炎のポイント");
  addFooter(s);

  // 特徴
  addCard(s, 0.6, 1.25, 4.1, 1.6, C.white, C.orange);
  s.addText("臨床的特徴", {
    x: 0.8, y: 1.3, w: 3, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0
  });
  s.addText("・ウイルス感染後に発症\n・有痛性甲状腺腫大（圧痛著明）\n・CRP・赤沈の著明な上昇\n・エコーで低エコー領域\n・シンチで取り込み低下", {
    x: 0.8, y: 1.65, w: 3.7, h: 1.1,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.35
  });

  // 3フェーズ
  addCard(s, 5.3, 1.25, 4.1, 1.6, C.white, C.blue);
  s.addText("臨床経過（3フェーズ）", {
    x: 5.5, y: 1.3, w: 3.5, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });

  var phases = [
    { label: "中毒期", desc: "疼痛＋甲状腺ホルモン↑", color: C.red },
    { label: "低下期", desc: "一過性機能低下", color: C.blue },
    { label: "回復期", desc: "正常に回復", color: C.green }
  ];
  for (var i = 0; i < phases.length; i++) {
    addBadge(s, 5.5, 1.75 + i * 0.42, String(i + 1), phases[i].color);
    s.addText(phases[i].label + "：" + phases[i].desc, {
      x: 6.05, y: 1.75 + i * 0.42, w: 3.2, h: 0.38,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
    });
  }

  // 治療
  addCard(s, 0.6, 3.05, 8.8, 1.0, C.ltGreen, C.green);
  s.addText("治療", {
    x: 0.8, y: 3.08, w: 2, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.green, bold: true, margin: 0
  });
  s.addText("・NSAIDs（軽症〜中等症）  ・プレドニゾロン 15〜30mg（痛みが強い場合、漸減）  ・抗甲状腺薬は不要", {
    x: 0.8, y: 3.4, w: 8.4, h: 0.55,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // 鑑別ポイント
  addCard(s, 0.6, 4.2, 8.8, 0.65, C.ltYellow, C.orange);
  s.addText("💡 亜急性 vs 無痛性の鑑別：圧痛の有無＋CRP上昇の有無が決め手", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.65,
    fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true,
    valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 9: 無痛性甲状腺炎・産後甲状腺炎
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "無痛性甲状腺炎・産後甲状腺炎");
  addFooter(s);

  // 無痛性甲状腺炎
  addCard(s, 0.6, 1.25, 4.1, 2.2, C.white, C.blue);
  s.addText("無痛性甲状腺炎", {
    x: 0.8, y: 1.3, w: 3.5, h: 0.35,
    fontSize: 18, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("・痛みを伴わない破壊性甲状腺炎\n・橋本病ベースの患者に多い\n・中毒期→低下期→回復期\n・CRP・赤沈は正常〜軽度上昇\n・TRAb陰性\n・多くは自然軽快", {
    x: 0.8, y: 1.7, w: 3.7, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.35
  });

  // 産後甲状腺炎
  addCard(s, 5.3, 1.25, 4.1, 2.2, C.white, C.purple);
  s.addText("産後甲状腺炎", {
    x: 5.5, y: 1.3, w: 3.5, h: 0.35,
    fontSize: 18, fontFace: FONT_JP, color: C.purple, bold: true, margin: 0
  });
  s.addText("・産後1年以内に発症\n・無痛性甲状腺炎の亜型\n・頻度：産後女性の5〜10%\n・産後うつとの鑑別が重要\n・一部は永続的低下症へ移行\n・抗TPO抗体陽性者にリスク高", {
    x: 5.5, y: 1.7, w: 3.7, h: 1.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.35
  });

  // 注意点
  addCard(s, 0.6, 3.7, 8.8, 1.2, C.ltYellow, C.orange);
  s.addText("⚠ フォローアップの重要性", {
    x: 0.8, y: 3.75, w: 5, h: 0.35,
    fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0
  });
  s.addText("・無痛性甲状腺炎は再発することがある\n・産後甲状腺炎の約20〜30%が永続的甲状腺機能低下症に移行\n・次回妊娠時にも再発リスクあり → 妊娠初期にTSH測定を推奨", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.7,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.35
  });
})();

// ══════════════════════════════════════════════════════
// Slide 10: 薬剤性（アミオダロン AIT1 vs AIT2）
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "薬剤性甲状腺中毒症（アミオダロン）");
  addFooter(s);

  var hOpt = function(text) {
    return { text: text, options: { fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } };
  };
  var cOpt = function(text, bg) {
    return { text: text, options: { fontSize: 12, fontFace: FONT_JP, color: C.text, fill: { color: bg }, align: "center", valign: "middle" } };
  };

  var rows = [
    [hOpt(""), hOpt("AIT Type 1"), hOpt("AIT Type 2")],
    [cOpt("機序", C.ltBlue), cOpt("ヨード過剰\n（潜在的甲状腺疾患あり）", C.ltRed), cOpt("甲状腺への直接毒性\n（破壊性甲状腺炎）", C.ltOrange)],
    [cOpt("基礎疾患", C.white), cOpt("バセドウ病, 多結節性甲状腺腫", C.ltRed), cOpt("正常甲状腺", C.ltOrange)],
    [cOpt("エコー血流", C.ltBlue), cOpt("増加", C.ltRed), cOpt("正常〜低下", C.ltOrange)],
    [cOpt("IL-6", C.white), cOpt("正常〜軽度上昇", C.ltRed), cOpt("著明上昇", C.ltOrange)],
    [cOpt("治療", C.ltBlue), cOpt("抗甲状腺薬", C.ltRed), cOpt("ステロイド", C.ltOrange)]
  ];

  s.addTable(rows, {
    x: 0.6, y: 1.25, w: 8.8,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [1.6, 3.6, 3.6],
    rowH: [0.45, 0.6, 0.5, 0.45, 0.45, 0.45],
    margin: [3, 5, 3, 5]
  });

  // その他薬剤
  addCard(s, 0.6, 4.2, 8.8, 0.8, C.ltPurple, C.purple);
  s.addText("その他の薬剤性甲状腺機能異常", {
    x: 0.8, y: 4.22, w: 5, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.purple, bold: true, margin: 0
  });
  s.addText("・免疫チェックポイント阻害薬（ICI）→ 甲状腺炎（破壊性）\n・リチウム → 甲状腺機能低下症（ヨード取り込み阻害）   ・インターフェロン → 甲状腺炎", {
    x: 0.8, y: 4.55, w: 8.4, h: 0.4,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });
})();

// ══════════════════════════════════════════════════════
// Slide 11: Step 5 — 甲状腺機能低下症の鑑別
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 5：甲状腺機能低下症の鑑別");
  addFooter(s);

  // フロー: TSH高値 + FT4低値
  addCard(s, 2.5, 1.25, 5, 0.55, C.navy, C.navy);
  s.addText("TSH高値 + FT4低値 → 甲状腺機能低下症", {
    x: 2.5, y: 1.25, w: 5, h: 0.55,
    fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // 矢印
  s.addText("▼", {
    x: 4.25, y: 1.85, w: 1.5, h: 0.3,
    fontSize: 14, fontFace: FONT_EN, color: C.navy,
    align: "center", valign: "middle", margin: 0
  });

  // 原因リスト
  var causes = [
    { label: "橋本病（慢性甲状腺炎）", pct: "最多", color: C.blue, bgColor: C.ltBlue },
    { label: "ヨード過剰摂取", pct: "昆布・ヨード造影剤", color: C.green, bgColor: C.ltGreen },
    { label: "薬剤性（リチウム、アミオダロン等）", pct: "", color: C.purple, bgColor: C.ltPurple },
    { label: "RI治療後・手術後", pct: "", color: C.orange, bgColor: C.ltOrange },
    { label: "中枢性（下垂体・視床下部障害）", pct: "TSH低〜正常", color: C.dkRed, bgColor: C.ltRed }
  ];

  for (var i = 0; i < causes.length; i++) {
    var cy = 2.2 + i * 0.55;
    addCard(s, 0.6, cy, 8.8, 0.48, causes[i].bgColor, causes[i].color);
    addBadge(s, 0.75, cy + 0.04, String(i + 1), causes[i].color);
    s.addText(causes[i].label, {
      x: 1.35, y: cy, w: 5, h: 0.48,
      fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true,
      valign: "middle", margin: 0
    });
    if (causes[i].pct) {
      s.addText(causes[i].pct, {
        x: 6.5, y: cy, w: 2.7, h: 0.48,
        fontSize: 12, fontFace: FONT_JP, color: C.gray,
        valign: "middle", align: "right", margin: [0, 10, 0, 0]
      });
    }
  }

  // 鑑別の進め方
  addCard(s, 0.6, 4.95, 8.8, 0.12, C.navy, C.navy);
})();

// ══════════════════════════════════════════════════════
// Slide 12: 橋本病のポイント
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "橋本病（慢性甲状腺炎）のポイント");
  addFooter(s);

  // 診断
  addCard(s, 0.6, 1.2, 4.1, 1.4, C.white, C.blue);
  s.addText("診断", {
    x: 0.8, y: 1.22, w: 3, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("・抗TPO抗体陽性（90%以上）\n・抗サイログロブリン抗体\n・エコーで不均一な低エコー\n・全員が低下症ではない\n  （euthyroidが初期は多い）", {
    x: 0.8, y: 1.52, w: 3.7, h: 1.0,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // LT4補充の適応
  addCard(s, 5.3, 1.2, 4.1, 1.4, C.white, C.green);
  s.addText("LT4補充の適応", {
    x: 5.5, y: 1.22, w: 3.5, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.green, bold: true, margin: 0
  });
  s.addText("【顕性低下症】全例で適応\n【潜在性低下症】以下で考慮：\n  ・TSH ≥ 10  ・症状あり\n  ・妊娠希望  ・脂質異常症", {
    x: 5.5, y: 1.52, w: 3.7, h: 1.0,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // 投与量・モニタリング
  addCard(s, 0.6, 2.8, 8.8, 1.2, C.ltBlue, C.blue);
  s.addText("LT4投与の実際", {
    x: 0.8, y: 2.82, w: 4, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("・目安量：1.6 μg/kg（成人で50〜150 μg/日）\n・高齢者・冠動脈疾患：少量（12.5〜25 μg）から緩やかに増量\n・TSH測定は6〜8週後 → 目標TSH 0.5〜4.0 mIU/L\n・空腹時内服推奨。鉄剤・Ca剤とは4時間以上あける", {
    x: 0.8, y: 3.15, w: 8.4, h: 0.8,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // 注意点
  addCard(s, 0.6, 4.15, 8.8, 0.75, C.ltYellow, C.orange);
  s.addText("⚠ 橋本病 ≠ 必ず治療が必要", {
    x: 0.8, y: 4.18, w: 5, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0
  });
  s.addText("euthyroidの橋本病は経過観察（年1回TSH測定）", {
    x: 0.8, y: 4.48, w: 8.4, h: 0.3,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 13: 中枢性甲状腺機能低下症
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "中枢性甲状腺機能低下症");
  addFooter(s);

  // 特徴
  addCard(s, 0.6, 1.2, 4.1, 1.7, C.white, C.purple);
  s.addText("診断のポイント", {
    x: 0.8, y: 1.22, w: 3, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.purple, bold: true, margin: 0
  });
  s.addText("・FT4低値なのにTSHが\n  適切に上昇していない\n・TSHは低値〜正常〜軽度高値\n  と多彩（bioactivity低下）\n・TSH単独スクリーニングで\n  見逃される", {
    x: 0.8, y: 1.52, w: 3.7, h: 1.3,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // 原因疾患
  addCard(s, 5.3, 1.2, 4.1, 1.7, C.white, C.blue);
  s.addText("疑うべき臨床状況", {
    x: 5.5, y: 1.22, w: 3.5, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("・下垂体腫瘍（最多）\n・下垂体手術後・放射線後\n・シーハン症候群（産後出血）\n・下垂体炎（ICI関連含む）\n・頭部外傷後", {
    x: 5.5, y: 1.52, w: 3.7, h: 1.3,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // 危険な注意点
  addCard(s, 0.6, 3.1, 8.8, 1.8, C.ltRed, C.red);
  s.addText("🚨 副腎不全合併に注意", {
    x: 0.8, y: 3.15, w: 5, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.dkRed, bold: true, margin: 0
  });
  s.addText("・中枢性低下症では他の下垂体ホルモン欠損の合併が多い\n・副腎不全合併時：甲状腺ホルモン補充の前にコルチゾール補充\n  → 順序を誤ると副腎クリーゼを誘発する危険\n・精査：下垂体MRI＋下垂体前葉ホルモン一式測定", {
    x: 0.8, y: 3.55, w: 8.4, h: 1.2,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4
  });
})();

// ══════════════════════════════════════════════════════
// Slide 14: Step 6 — 潜在性甲状腺機能異常
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 6：潜在性甲状腺機能異常の対応");
  addFooter(s);

  // 潜在性低下症
  addCard(s, 0.6, 1.25, 4.1, 3.5, C.white, C.blue);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.25, w: 4.1, h: 0.5,
    fill: { color: C.blue }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.55, w: 4.1, h: 0.2,
    fill: { color: C.blue }
  });
  s.addText("潜在性甲状腺機能低下症", {
    x: 0.7, y: 1.25, w: 3.9, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("TSH↑ + FT4正常", {
    x: 0.8, y: 1.85, w: 3.7, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("① 4〜12週後に再検\n  （一過性でないことを確認）\n\n② TSH ≥ 10 → LT4補充を考慮\n\n③ TSH 4.5〜10 で治療検討：\n  ・抗TPO抗体陽性\n  ・症状あり\n  ・妊娠希望\n  ・脂質異常症合併", {
    x: 0.8, y: 2.15, w: 3.7, h: 2.4,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // 潜在性中毒症
  addCard(s, 5.3, 1.25, 4.1, 3.5, C.white, C.red);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.3, y: 1.25, w: 4.1, h: 0.5,
    fill: { color: C.red }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.55, w: 4.1, h: 0.2,
    fill: { color: C.red }
  });
  s.addText("潜在性甲状腺中毒症", {
    x: 5.4, y: 1.25, w: 3.9, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("TSH↓ + FT4/FT3正常", {
    x: 5.5, y: 1.85, w: 3.7, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, margin: 0
  });
  s.addText("① TSH < 0.1 が持続\n  → 心房細動・骨粗鬆症リスク↑\n  → 原因検索＋治療を積極検討\n\n② TSH 0.1〜0.4\n  → 基本は経過観察\n  → 以下で治療考慮：\n    ・65歳以上\n    ・心疾患合併\n    ・閉経後女性", {
    x: 5.5, y: 2.15, w: 3.7, h: 2.4,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });
})();

// ══════════════════════════════════════════════════════
// Slide 15: Red Flags — 甲状腺クリーゼ
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Red Flags：甲状腺クリーゼ");
  addFooter(s);

  // 概要
  addCard(s, 0.6, 1.25, 8.8, 0.6, C.ltRed, C.red);
  s.addText("🚨 甲状腺中毒症の急激な増悪 → 多臓器不全 → 死亡率 10〜30%", {
    x: 0.8, y: 1.25, w: 8.4, h: 0.6,
    fontSize: 16, fontFace: FONT_JP, color: C.dkRed, bold: true,
    valign: "middle", margin: 0
  });

  // Burch-Wartofsky スコア（テーブル）
  var hOpt = function(text) {
    return { text: text, options: { fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.dkRed }, align: "center", valign: "middle" } };
  };
  var cOpt = function(text, bg) {
    return { text: text, options: { fontSize: 11, fontFace: FONT_JP, color: C.text, fill: { color: bg }, align: "center", valign: "middle" } };
  };

  var rows = [
    [hOpt("項目"), hOpt("所見"), hOpt("スコア")],
    [cOpt("体温", C.ltRed), cOpt("38℃以上（段階的に加点）", C.ltRed), cOpt("5〜30", C.ltRed)],
    [cOpt("中枢神経", C.white), cOpt("不穏/せん妄/昏睡", C.white), cOpt("10〜30", C.white)],
    [cOpt("消化器", C.ltRed), cOpt("嘔吐/下痢/黄疸", C.ltRed), cOpt("10〜20", C.ltRed)],
    [cOpt("心血管", C.white), cOpt("頻脈/心房細動/心不全", C.white), cOpt("5〜25", C.white)],
    [cOpt("トリガー", C.ltRed), cOpt("あり", C.ltRed), cOpt("10", C.ltRed)]
  ];

  s.addTable(rows, {
    x: 0.6, y: 2.0, w: 4.2,
    border: { type: "solid", color: C.red, pt: 0.8 },
    colW: [1.2, 2.0, 1.0],
    rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
    margin: [2, 4, 2, 4]
  });

  s.addText("≥45点：甲状腺クリーゼ", {
    x: 0.6, y: 4.15, w: 4.2, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.dkRed, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // 治療
  addCard(s, 5.3, 2.0, 4.1, 2.8, C.white, C.red);
  s.addText("緊急治療（ICU管理）", {
    x: 5.5, y: 2.05, w: 3.5, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.dkRed, bold: true, margin: 0
  });

  var txList = [
    { num: "1", text: "PTU大量投与\n（T4→T3変換阻害もあり）" },
    { num: "2", text: "無機ヨード\n（PTU投与1時間後に開始）" },
    { num: "3", text: "β遮断薬\n（プロプラノロール）" },
    { num: "4", text: "ステロイド\n（ヒドロコルチゾン）" },
    { num: "5", text: "冷却（解熱）" }
  ];
  for (var i = 0; i < txList.length; i++) {
    addBadge(s, 5.5, 2.5 + i * 0.45, txList[i].num, C.red);
    s.addText(txList[i].text, {
      x: 6.05, y: 2.5 + i * 0.45, w: 3.2, h: 0.42,
      fontSize: 11, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.1
    });
  }
})();

// ══════════════════════════════════════════════════════
// Slide 16: Red Flags — 粘液水腫性昏睡
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Red Flags：粘液水腫性昏睡");
  addFooter(s);

  // 概要
  addCard(s, 0.6, 1.25, 8.8, 0.6, C.ltRed, C.red);
  s.addText("🚨 甲状腺機能低下症の究極の重症型 — 高齢女性に多い", {
    x: 0.8, y: 1.25, w: 8.4, h: 0.6,
    fontSize: 16, fontFace: FONT_JP, color: C.dkRed, bold: true,
    valign: "middle", margin: 0
  });

  // 臨床像
  addCard(s, 0.6, 2.0, 4.1, 2.3, C.white, C.blue);
  s.addText("臨床像", {
    x: 0.8, y: 2.02, w: 3, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });

  var features = [
    { text: "意識障害", icon: "●" },
    { text: "低体温（<35℃）", icon: "●" },
    { text: "徐脈・低血圧", icon: "●" },
    { text: "低換気・CO2ナルコーシス", icon: "●" },
    { text: "低Na血症", icon: "●" },
    { text: "低血糖", icon: "●" }
  ];
  for (var i = 0; i < features.length; i++) {
    s.addText(features[i].icon + " " + features[i].text, {
      x: 0.8, y: 2.35 + i * 0.3, w: 3.7, h: 0.28,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
    });
  }

  // 治療
  addCard(s, 5.3, 2.0, 4.1, 2.3, C.white, C.green);
  s.addText("治療", {
    x: 5.5, y: 2.02, w: 3, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.green, bold: true, margin: 0
  });
  s.addText("・LT4静注（経口吸収が不確実）\n  初日200〜500μg→50〜100μg/日\n・ヒドロコルチゾン併用\n  （副腎不全合併を想定）\n・支持療法：保温（急速復温避ける）\n  低Na→水分制限、呼吸管理", {
    x: 5.5, y: 2.35, w: 3.7, h: 1.8,
    fontSize: 11, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // トリガー
  addCard(s, 0.6, 4.5, 8.8, 0.4, C.ltOrange, C.orange);
  s.addText("トリガー：感染症、寒冷曝露、鎮静薬、手術、甲状腺薬の中断", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.4,
    fontSize: 12, fontFace: FONT_JP, color: C.orange, bold: true,
    valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 17: 鑑別フローまとめ
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "鑑別フローまとめ");
  addFooter(s);

  // TSH起点
  addCard(s, 3.5, 1.25, 3, 0.55, C.navy, C.navy);
  s.addText("TSH測定", {
    x: 3.5, y: 1.25, w: 3, h: 0.55,
    fontSize: 20, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // 左矢印
  s.addText("TSH↓", {
    x: 0.8, y: 1.85, w: 2.5, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // 右矢印
  s.addText("TSH↑", {
    x: 6.7, y: 1.85, w: 2.5, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.blue, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // 左ブランチ: 中毒症
  addCard(s, 0.3, 2.2, 4.4, 0.45, C.red, C.red);
  s.addText("甲状腺中毒症", {
    x: 0.3, y: 2.2, w: 4.4, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  s.addText("FT4/FT3で\n顕性 vs 潜在性", {
    x: 0.3, y: 2.7, w: 4.4, h: 0.45,
    fontSize: 12, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2
  });

  addCard(s, 0.3, 3.2, 4.4, 0.45, C.ltRed, C.red);
  s.addText("TRAb / エコー血流 / シンチ", {
    x: 0.3, y: 3.2, w: 4.4, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.dkRed, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  addCard(s, 0.3, 3.75, 2.05, 0.7, C.ltOrange, C.orange);
  s.addText("産生過剰\nバセドウ病\n自律性結節", {
    x: 0.3, y: 3.75, w: 2.05, h: 0.7,
    fontSize: 11, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2
  });

  addCard(s, 2.65, 3.75, 2.05, 0.7, C.ltYellow, C.orange);
  s.addText("破壊性\n亜急性甲状腺炎\n無痛性甲状腺炎", {
    x: 2.65, y: 3.75, w: 2.05, h: 0.7,
    fontSize: 11, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2
  });

  // 右ブランチ: 低下症
  addCard(s, 5.3, 2.2, 4.4, 0.45, C.blue, C.blue);
  s.addText("甲状腺機能低下症", {
    x: 5.3, y: 2.2, w: 4.4, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  s.addText("FT4で\n顕性 vs 潜在性", {
    x: 5.3, y: 2.7, w: 4.4, h: 0.45,
    fontSize: 12, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2
  });

  addCard(s, 5.3, 3.2, 4.4, 0.45, C.ltBlue, C.blue);
  s.addText("抗TPO抗体 / 薬剤歴 / ヨード", {
    x: 5.3, y: 3.2, w: 4.4, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.navy, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  addCard(s, 5.3, 3.75, 2.05, 0.7, C.ltBlue, C.blue);
  s.addText("橋本病\n（最多）", {
    x: 5.3, y: 3.75, w: 2.05, h: 0.7,
    fontSize: 12, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2
  });

  addCard(s, 7.65, 3.75, 2.05, 0.7, C.ltPurple, C.purple);
  s.addText("薬剤性\nヨード過剰\n中枢性", {
    x: 7.65, y: 3.75, w: 2.05, h: 0.7,
    fontSize: 11, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2
  });

  // 例外パターン
  addCard(s, 0.6, 4.65, 8.8, 0.4, C.ltYellow, C.orange);
  s.addText("例外：TSH正常〜低値+FT4↓→中枢性 / TSH↑+FT4↑→不適切TSH分泌", {
    x: 0.8, y: 4.65, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true,
    valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 18: Take Home Message
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Take Home Message");
  addFooter(s);

  var messages = [
    { num: "1", text: "TSHは最も感度の高い検査だが、必ずFT4/FT3とセットで解釈する" },
    { num: "2", text: "甲状腺中毒症では「破壊性 vs 産生過剰」が治療方針を決める（TRAb・エコー・シンチ）" },
    { num: "3", text: "甲状腺機能低下症の最多原因は橋本病 — 抗TPO抗体が診断の鍵" },
    { num: "4", text: "中枢性低下症はTSH単独では見逃される — FT4低値+TSH非上昇で疑う" },
    { num: "5", text: "甲状腺クリーゼ・粘液水腫性昏睡は致死的緊急疾患 — 疑ったら直ちに治療" }
  ];

  for (var i = 0; i < messages.length; i++) {
    var my = 1.4 + i * 0.72;

    // カード
    addCard(s, 0.6, my, 8.8, 0.62, C.white, C.blue);

    // 番号バッジ
    s.addShape(pres.shapes.OVAL, {
      x: 0.8, y: my + 0.11, w: 0.4, h: 0.4,
      fill: { color: C.blue }
    });
    s.addText(messages[i].num, {
      x: 0.8, y: my + 0.11, w: 0.4, h: 0.4,
      fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0
    });

    // メッセージ
    s.addText(messages[i].text, {
      x: 1.4, y: my, w: 7.8, h: 0.62,
      fontSize: 15, fontFace: FONT_JP, color: C.text,
      valign: "middle", margin: 0
    });
  }

  // 底部アクセント
  addCard(s, 0.6, 4.95, 8.8, 0.1, C.navy, C.navy);
})();

// ── 出力 ──
pres.writeFile({ fileName: "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/甲状腺機能異常/甲状腺機能異常_鑑別診断.pptx" })
  .then(function() { console.log("Done: 甲状腺機能異常_鑑別診断.pptx"); })
  .catch(function(e) { console.error(e); });
