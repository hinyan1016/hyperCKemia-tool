var pptxgen = require("pptxgenjs");
var pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "医知創造ラボ";
pres.title = "貧血の鑑別診断";

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
  slide.addText("貧血の鑑別診断", { x: 0.5, y: 5.15, w: 6, h: 0.475, fontSize: 18, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
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
  s.addText("貧血の鑑別診断", {
    x: 0.5, y: 1.5, w: 9, h: 1.3,
    fontSize: 52, fontFace: FONT_JP, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });

  // 装飾ライン
  s.addShape(pres.shapes.RECTANGLE, { x: 3, y: 2.95, w: 4, h: 0.05, fill: { color: C.yellow } });

  // サブタイトル
  s.addText("― MCV分類×8ステップで系統的に迫る ―", {
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
// Slide 2: 貧血の定義とWHO基準
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "貧血の定義とWHO基準（2024）");
  addFooter(s);

  // テーブルデータ
  var headerRow = [
    { text: "対象", options: { fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } },
    { text: "正常 (g/dL)", options: { fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } },
    { text: "軽度", options: { fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } },
    { text: "中等度", options: { fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } },
    { text: "重度", options: { fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } }
  ];

  var cellOpt = function(bg) {
    return { fontSize: 13, fontFace: FONT_JP, color: C.text, fill: { color: bg }, align: "center", valign: "middle" };
  };

  var rows = [
    headerRow,
    [
      { text: "男性 15-65歳", options: cellOpt(C.ltBlue) },
      { text: "≥13.0", options: cellOpt(C.ltGreen) },
      { text: "11.0-12.9", options: cellOpt(C.ltYellow) },
      { text: "8.0-10.9", options: cellOpt(C.ltOrange) },
      { text: "<8.0", options: cellOpt(C.ltRed) }
    ],
    [
      { text: "女性 15-65歳", options: cellOpt(C.white) },
      { text: "≥12.0", options: cellOpt(C.ltGreen) },
      { text: "11.0-11.9", options: cellOpt(C.ltYellow) },
      { text: "8.0-10.9", options: cellOpt(C.ltOrange) },
      { text: "<8.0", options: cellOpt(C.ltRed) }
    ],
    [
      { text: "妊娠 1/3期", options: cellOpt(C.ltBlue) },
      { text: "≥11.0", options: cellOpt(C.ltGreen) },
      { text: "10.0-10.9", options: cellOpt(C.ltYellow) },
      { text: "7.0-9.9", options: cellOpt(C.ltOrange) },
      { text: "<7.0", options: cellOpt(C.ltRed) }
    ]
  ];

  s.addTable(rows, {
    x: 0.6, y: 1.4, w: 8.8,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [2.2, 1.6, 1.6, 1.8, 1.6],
    rowH: [0.5, 0.55, 0.55, 0.55],
    margin: [4, 6, 4, 6]
  });

  // 補足ボックス
  addCard(s, 0.6, 3.7, 8.8, 1.1, C.white, C.blue);
  s.addText("💡 ポイント", {
    x: 0.8, y: 3.75, w: 3, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("・高地居住・喫煙・脱水で相対的に高値\n・Hb値だけでなく症状（動悸・息切れ・倦怠感）も重要", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.65,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });
})();

// ══════════════════════════════════════════════════════
// Slide 3: 鑑別アルゴリズムの全体像（8ステップ）
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "鑑別アルゴリズムの全体像");
  addFooter(s);

  var steps = [
    { id: "A", label: "初期トリアージ", color: C.red },
    { id: "B", label: "MCV分類", color: C.blue },
    { id: "C", label: "網状赤血球", color: C.blue },
    { id: "D", label: "溶血所見", color: C.orange },
    { id: "E", label: "鉄代謝", color: C.green },
    { id: "F", label: "末梢血塗抹", color: C.purple },
    { id: "G", label: "臨床背景", color: C.navy },
    { id: "H", label: "確定診断", color: C.navy }
  ];

  var startX = 0.3;
  var stepW = 1.05;
  var gap = 0.12;

  for (var i = 0; i < steps.length; i++) {
    var sx = startX + i * (stepW + gap);
    var sy = 1.5;

    // ステップカード
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: sx, y: sy, w: stepW, h: 2.4,
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
      x: sx + 0.05, y: sy + 0.85, w: stepW - 0.1, h: 1.4,
      fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true,
      align: "center", valign: "top", margin: 0
    });

    // 矢印（最後以外）
    if (i < steps.length - 1) {
      s.addText("→", {
        x: sx + stepW - 0.05, y: sy + 0.9, w: 0.3, h: 0.5,
        fontSize: 20, fontFace: FONT_EN, color: C.blue, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    }
  }

  // 下部メッセージ
  addCard(s, 0.6, 4.2, 8.8, 0.7, C.ltBlue, C.blue);
  s.addText("💡 8つのステップを順序立てて進めることで、効率的に鑑別を絞り込む", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.7,
    fontSize: 15, fontFace: FONT_JP, color: C.navy, bold: true,
    align: "center", valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 4: Step A 初期トリアージ
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step A  初期トリアージ");
  addFooter(s);

  // 上部警告バー
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.25, w: 8.8, h: 0.55,
    fill: { color: C.red }, rectRadius: 0.08
  });
  s.addText("⚠️ まず緊急性を判断 — 以下があれば即時対応", {
    x: 0.8, y: 1.25, w: 8.4, h: 0.55,
    fontSize: 17, fontFace: FONT_JP, color: C.white, bold: true,
    valign: "middle", margin: 0
  });

  var items = [
    { emoji: "🆘", title: "血行動態不安定", desc: "頻脈・低血圧・ショック → 輸血・輸液を優先" },
    { emoji: "🩸", title: "Hb低値＋症候性", desc: "Hb<7 g/dL＋動悸/呼吸困難 → 輸血検討" },
    { emoji: "⚡", title: "MAHA疑い", desc: "血小板減少＋破砕赤血球 → TTP → 血漿交換を遅らせない" },
    { emoji: "🧠", title: "神経症状＋巨赤芽球性", desc: "B12欠乏 → 不可逆性あり → 早期補充" }
  ];

  for (var i = 0; i < items.length; i++) {
    var iy = 2.05 + i * 0.75;
    addCard(s, 0.6, iy, 8.8, 0.65, C.white, C.red);

    s.addText(items[i].emoji, {
      x: 0.75, y: iy, w: 0.5, h: 0.65,
      fontSize: 22, align: "center", valign: "middle", margin: 0
    });
    s.addText(items[i].title, {
      x: 1.35, y: iy, w: 2.5, h: 0.65,
      fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true,
      valign: "middle", margin: 0
    });
    s.addText(items[i].desc, {
      x: 3.9, y: iy, w: 5.3, h: 0.65,
      fontSize: 13, fontFace: FONT_JP, color: C.text,
      valign: "middle", margin: 0
    });
  }
})();

// ══════════════════════════════════════════════════════
// Slide 5: Step B MCV分類
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step B  MCV分類");
  addFooter(s);

  // 3つのカード
  var cards = [
    {
      emoji: "🔴", title: "小球性 (MCV<80)",
      color: C.red, bgColor: C.ltRed,
      items: "IDA, サラセミア,\n鉄芽球性, 鉛中毒, IRIDA"
    },
    {
      emoji: "🟡", title: "正球性 (80-100)",
      color: C.orange, bgColor: C.ltOrange,
      items: "出血, 溶血, ACD,\nCKD, 骨髄不全, PRCA"
    },
    {
      emoji: "🔵", title: "大球性 (MCV>100)",
      color: C.blue, bgColor: C.ltBlue,
      items: "B12欠乏, 葉酸欠乏,\nアルコール/肝, MDS"
    }
  ];

  var cardW = 2.7;
  var cardGap = 0.3;
  var baseX = 0.55;

  for (var i = 0; i < cards.length; i++) {
    var cx = baseX + i * (cardW + cardGap);
    var cy = 1.3;

    addCard(s, cx, cy, cardW, 2.6, cards[i].bgColor, cards[i].color);

    // ヘッダー帯
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx, y: cy, w: cardW, h: 0.55,
      fill: { color: cards[i].color },
      rectRadius: 0.1
    });
    // 下角を埋める
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy + 0.35, w: cardW, h: 0.2,
      fill: { color: cards[i].color }
    });

    s.addText(cards[i].emoji + " " + cards[i].title, {
      x: cx + 0.1, y: cy, w: cardW - 0.2, h: 0.55,
      fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0
    });

    s.addText(cards[i].items, {
      x: cx + 0.15, y: cy + 0.7, w: cardW - 0.3, h: 1.8,
      fontSize: 14, fontFace: FONT_JP, color: C.text,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.5
    });
  }

  // 警告ボックス
  addCard(s, 0.55, 4.15, 8.9, 0.7, C.ltYellow, C.orange);
  s.addText("⚠️ MCVが正常でも鉄欠乏は否定できない（初期や混合欠乏の場合）", {
    x: 0.75, y: 4.15, w: 8.5, h: 0.7,
    fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true,
    align: "center", valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 6: Step C 網状赤血球（RPI）
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step C  網状赤血球（RPI）");
  addFooter(s);

  // RPI計算式ボックス
  addCard(s, 0.6, 1.3, 8.8, 0.9, C.ltBlue, C.blue);
  s.addText("RPI = 網赤(%) × (Hct / 0.45) ÷ 成熟補正係数", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.55,
    fontSize: 20, fontFace: FONT_JP, color: C.navy, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("成熟補正: Hct 45%=1.0, 35%=1.5, 25%=2.0, 15%=2.5", {
    x: 0.8, y: 1.85, w: 8.4, h: 0.35,
    fontSize: 12, fontFace: FONT_JP, color: C.blue,
    align: "center", valign: "middle", margin: 0
  });

  // 左カード: RPI>3
  addCard(s, 0.6, 2.5, 4.1, 2.2, C.ltGreen, C.green);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 2.5, w: 4.1, h: 0.5,
    fill: { color: C.green }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.8, w: 4.1, h: 0.2,
    fill: { color: C.green }
  });
  s.addText("RPI > 3  反応良好（赤↑）", {
    x: 0.7, y: 2.5, w: 3.9, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("・急性出血\n・溶血性貧血\n  → 骨髄の代償反応あり\n  → 産生は保たれている", {
    x: 0.8, y: 3.1, w: 3.7, h: 1.5,
    fontSize: 14, fontFace: FONT_JP, color: C.text,
    valign: "top", margin: 0, lineSpacingMultiple: 1.4
  });

  // 右カード: RPI<2
  addCard(s, 5.3, 2.5, 4.1, 2.2, C.ltRed, C.red);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.3, y: 2.5, w: 4.1, h: 0.5,
    fill: { color: C.red }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 2.8, w: 4.1, h: 0.2,
    fill: { color: C.red }
  });
  s.addText("RPI < 2  産生低下（赤↓）", {
    x: 5.4, y: 2.5, w: 3.9, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("・骨髄不全\n・栄養欠乏（Fe, B12, 葉酸）\n・慢性疾患（ACD/AI）\n  → 骨髄の産生障害", {
    x: 5.5, y: 3.1, w: 3.7, h: 1.5,
    fontSize: 14, fontFace: FONT_JP, color: C.text,
    valign: "top", margin: 0, lineSpacingMultiple: 1.4
  });

  // 注意書き
  s.addText("※ AIHAでもDAT+網赤不十分の場合あり（赤芽球への自己抗体）", {
    x: 0.6, y: 4.85, w: 8.8, h: 0.3,
    fontSize: 12, fontFace: FONT_JP, color: C.gray, margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 7: 小球性貧血の鑑別（5疾患）
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "小球性貧血の鑑別（5疾患）");
  addFooter(s);

  var diseases = [
    { name: "IDA（鉄欠乏）", key: "フェリチン↓, pencil cells", detail: "出血源検索が重要", color: C.red },
    { name: "サラセミア", key: "RBC保たれる, target cells", detail: "HbA2↑（β型）", color: C.orange },
    { name: "鉄芽球性貧血", key: "環状鉄芽球", detail: "Pappenheimer bodies", color: C.purple },
    { name: "鉛中毒", key: "塩基性斑点赤血球", detail: "血中鉛測定, FEP↑", color: C.navy },
    { name: "IRIDA", key: "経口鉄不応", detail: "ヘプシジン高値, TMPRSS6変異", color: C.green }
  ];

  for (var i = 0; i < diseases.length; i++) {
    var dy = 1.25 + i * 0.72;
    addCard(s, 0.6, dy, 8.8, 0.62, C.white, diseases[i].color);

    // 左側カラーバー
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: dy, w: 0.12, h: 0.62,
      fill: { color: diseases[i].color }
    });

    s.addText(diseases[i].name, {
      x: 0.85, y: dy, w: 2.3, h: 0.62,
      fontSize: 15, fontFace: FONT_JP, color: diseases[i].color, bold: true,
      valign: "middle", margin: 0
    });
    s.addText(diseases[i].key, {
      x: 3.2, y: dy, w: 3.2, h: 0.62,
      fontSize: 13, fontFace: FONT_JP, color: C.text,
      valign: "middle", margin: 0
    });
    s.addText(diseases[i].detail, {
      x: 6.5, y: dy, w: 2.7, h: 0.62,
      fontSize: 12, fontFace: FONT_JP, color: C.gray,
      valign: "middle", margin: 0
    });
  }

  // 補足
  addCard(s, 0.6, 4.95, 4.2, 0.15, C.red, C.red);
  // 空の細いバーで視覚的アクセント
})();

// ══════════════════════════════════════════════════════
// Slide 8: 正球性貧血の鑑別（11疾患）
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "正球性貧血の鑑別（11疾患）");
  addFooter(s);

  // 3グループに分割

  // 出血グループ
  var gx1 = 0.4, gy = 1.25, gw = 2.8;
  addCard(s, gx1, gy, gw, 3.6, C.white, C.red);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: gx1, y: gy, w: gw, h: 0.45,
    fill: { color: C.red }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: gx1, y: gy + 0.3, w: gw, h: 0.15,
    fill: { color: C.red }
  });
  s.addText("🩸 出血", {
    x: gx1 + 0.1, y: gy, w: gw - 0.2, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("・急性出血性貧血", {
    x: gx1 + 0.15, y: gy + 0.55, w: gw - 0.3, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0
  });

  // 溶血グループ
  var gx2 = 3.4;
  addCard(s, gx2, gy, gw, 3.6, C.white, C.orange);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: gx2, y: gy, w: gw, h: 0.45,
    fill: { color: C.orange }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: gx2, y: gy + 0.3, w: gw, h: 0.15,
    fill: { color: C.orange }
  });
  s.addText("💥 溶血", {
    x: gx2 + 0.1, y: gy, w: gw - 0.2, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("・AIHA（DAT陽性）\n・TMA（破砕赤血球）\n・機械的溶血\n・遺伝性球状赤血球症\n・G6PD欠損症\n・PNH", {
    x: gx2 + 0.15, y: gy + 0.55, w: gw - 0.3, h: 2.8,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.4
  });

  // 低産生グループ
  var gx3 = 6.4;
  addCard(s, gx3, gy, gw + 0.2, 3.6, C.white, C.blue);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: gx3, y: gy, w: gw + 0.2, h: 0.45,
    fill: { color: C.blue }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: gx3, y: gy + 0.3, w: gw + 0.2, h: 0.15,
    fill: { color: C.blue }
  });
  s.addText("⬇️ 低産生", {
    x: gx3 + 0.1, y: gy, w: gw, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("・ACD/AI（慢性疾患）\n・CKD（EPO低下）\n・再生不良性貧血\n・骨髄浸潤\n・PRCA（赤芽球癆）", {
    x: gx3 + 0.15, y: gy + 0.55, w: gw - 0.1, h: 2.8,
    fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.4
  });
})();

// ══════════════════════════════════════════════════════
// Slide 9: 大球性貧血の鑑別（5疾患）
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "大球性貧血の鑑別（5疾患）");
  addFooter(s);

  // 左: 巨赤芽球性
  var lx = 0.5, ly = 1.3, lw = 4.3, lh = 3.4;
  addCard(s, lx, ly, lw, lh, C.white, C.purple);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: lx, y: ly, w: lw, h: 0.5,
    fill: { color: C.purple }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: lx, y: ly + 0.3, w: lw, h: 0.2,
    fill: { color: C.purple }
  });
  s.addText("巨赤芽球性", {
    x: lx + 0.1, y: ly, w: lw - 0.2, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // B12
  addCard(s, lx + 0.15, ly + 0.7, lw - 0.3, 1.1, C.ltPurple, C.purple);
  s.addText("ビタミンB12欠乏", {
    x: lx + 0.3, y: ly + 0.75, w: lw - 0.5, h: 0.35,
    fontSize: 15, fontFace: FONT_JP, color: C.purple, bold: true, margin: 0
  });
  s.addText("過分葉好中球, 神経症状（亜急性脊髄変性症）\nMMA↑, ホモシステイン↑", {
    x: lx + 0.3, y: ly + 1.1, w: lw - 0.5, h: 0.6,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // 葉酸
  addCard(s, lx + 0.15, ly + 1.95, lw - 0.3, 1.0, C.ltPurple, C.purple);
  s.addText("葉酸欠乏", {
    x: lx + 0.3, y: ly + 2.0, w: lw - 0.5, h: 0.35,
    fontSize: 15, fontFace: FONT_JP, color: C.purple, bold: true, margin: 0
  });
  s.addText("MMA正常（B12との鑑別点）\n妊娠・アルコール・薬剤（MTX等）で多い", {
    x: lx + 0.3, y: ly + 2.35, w: lw - 0.5, h: 0.55,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // 右: 非巨赤芽球性
  var rx = 5.2, rw = 4.3;
  addCard(s, rx, ly, rw, lh, C.white, C.blue);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: rx, y: ly, w: rw, h: 0.5,
    fill: { color: C.blue }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: rx, y: ly + 0.3, w: rw, h: 0.2,
    fill: { color: C.blue }
  });
  s.addText("非巨赤芽球性", {
    x: rx + 0.1, y: ly, w: rw - 0.2, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  var nonMega = [
    { name: "アルコール/肝疾患", desc: "直接毒性, 栄養障害" },
    { name: "甲状腺機能低下症", desc: "EPO産生低下" },
    { name: "MDS（骨髄異形成症候群）", desc: "芽球%, 染色体異常" }
  ];

  for (var i = 0; i < nonMega.length; i++) {
    var ny = ly + 0.7 + i * 0.85;
    addCard(s, rx + 0.15, ny, rw - 0.3, 0.75, C.ltBlue, C.blue);
    s.addText(nonMega[i].name, {
      x: rx + 0.3, y: ny + 0.05, w: rw - 0.5, h: 0.3,
      fontSize: 14, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
    });
    s.addText(nonMega[i].desc, {
      x: rx + 0.3, y: ny + 0.35, w: rw - 0.5, h: 0.3,
      fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0
    });
  }
})();

// ══════════════════════════════════════════════════════
// Slide 10: 鉄代謝指標の読み方
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "鉄代謝指標の読み方");
  addFooter(s);

  // テーブル
  var hdrOpt = { fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" };
  var cellC = function(bg) { return { fontSize: 13, fontFace: FONT_JP, color: C.text, fill: { color: bg }, align: "center", valign: "middle" }; };
  var cellL = function(bg) { return { fontSize: 13, fontFace: FONT_JP, color: C.text, fill: { color: bg }, valign: "middle" }; };

  var rows = [
    [
      { text: "指標", options: hdrOpt },
      { text: "IDA", options: { fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.red }, align: "center", valign: "middle" } },
      { text: "ACD/AI", options: { fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.orange }, align: "center", valign: "middle" } }
    ],
    [
      { text: "血清鉄", options: cellL(C.ltGray) },
      { text: "低 ↓", options: cellC(C.ltRed) },
      { text: "低 ↓", options: cellC(C.ltOrange) }
    ],
    [
      { text: "フェリチン", options: cellL(C.white) },
      { text: "<15-30 ↓↓", options: cellC(C.ltRed) },
      { text: "正常～高 ↑", options: cellC(C.ltOrange) }
    ],
    [
      { text: "TIBC", options: cellL(C.ltGray) },
      { text: "高 ↑", options: cellC(C.ltRed) },
      { text: "低～正常 ↓", options: cellC(C.ltOrange) }
    ],
    [
      { text: "TSAT", options: cellL(C.white) },
      { text: "低 ↓", options: cellC(C.ltRed) },
      { text: "低 ↓", options: cellC(C.ltOrange) }
    ],
    [
      { text: "sTfR", options: cellL(C.ltGray) },
      { text: "↑ 上昇", options: cellC(C.ltRed) },
      { text: "正常", options: cellC(C.ltOrange) }
    ]
  ];

  s.addTable(rows, {
    x: 0.6, y: 1.3, w: 5.5,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [1.8, 1.85, 1.85],
    rowH: [0.45, 0.45, 0.45, 0.45, 0.45, 0.45],
    margin: [4, 8, 4, 8]
  });

  // 右側: フェリチンカットオフ
  addCard(s, 6.4, 1.3, 3.2, 3.6, C.white, C.green);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.4, y: 1.3, w: 3.2, h: 0.45,
    fill: { color: C.green }, rectRadius: 0.1
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.4, y: 1.55, w: 3.2, h: 0.2,
    fill: { color: C.green }
  });
  s.addText("フェリチン カットオフ", {
    x: 6.5, y: 1.3, w: 3.0, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  var cutoffs = [
    { val: "<15 ng/mL", desc: "IDA確定的\n（特異度99%）" },
    { val: "<30 ng/mL", desc: "IDA疑い\n（感度92%）" },
    { val: "<100 ng/mL", desc: "炎症合併時の\nIDA判定閾値" }
  ];

  for (var i = 0; i < cutoffs.length; i++) {
    var cy = 1.95 + i * 0.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 6.6, y: cy, w: 2.8, h: 0.8,
      fill: { color: C.ltGreen },
      line: { color: C.green, width: 1 },
      rectRadius: 0.08
    });
    s.addText(cutoffs[i].val, {
      x: 6.7, y: cy + 0.05, w: 2.6, h: 0.3,
      fontSize: 14, fontFace: FONT_EN, color: C.green, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(cutoffs[i].desc, {
      x: 6.7, y: cy + 0.35, w: 2.6, h: 0.4,
      fontSize: 11, fontFace: FONT_JP, color: C.text,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2
    });
  }
})();

// ══════════════════════════════════════════════════════
// Slide 11: 溶血の評価とDAT
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "溶血の評価とDAT");
  addFooter(s);

  // 溶血マーカーボックス
  addCard(s, 0.5, 1.25, 4.4, 1.6, C.ltOrange, C.orange);
  s.addText("溶血マーカー", {
    x: 0.65, y: 1.3, w: 4.1, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0
  });
  s.addText("・LDH ↑\n・間接ビリルビン ↑\n・ハプトグロビン ↓\n・網状赤血球 ↑", {
    x: 0.75, y: 1.7, w: 3.9, h: 1.05,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // DATフローチャート
  addCard(s, 5.2, 1.25, 4.4, 1.6, C.ltBlue, C.blue);
  s.addText("DAT（直接クームス試験）", {
    x: 5.35, y: 1.3, w: 4.1, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("免疫性溶血 vs 非免疫性溶血\nの鑑別に必須", {
    x: 5.45, y: 1.7, w: 3.9, h: 1.05,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // DAT陽性
  addCard(s, 0.5, 3.1, 4.4, 1.7, C.white, C.red);
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.1, w: 0.12, h: 1.7,
    fill: { color: C.red }
  });
  s.addText("DAT 陽性（+）", {
    x: 0.75, y: 3.15, w: 4.0, h: 0.35,
    fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true, margin: 0
  });
  s.addText("→ AIHA\n  ・温式（IgG型）\n  ・冷式（C3型 / 冷式凝集素症）\n  ・混合型", {
    x: 0.85, y: 3.5, w: 3.8, h: 1.2,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });

  // DAT陰性
  addCard(s, 5.2, 3.1, 4.4, 1.7, C.white, C.blue);
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 3.1, w: 0.12, h: 1.7,
    fill: { color: C.blue }
  });
  s.addText("DAT 陰性（-）", {
    x: 5.45, y: 3.15, w: 4.0, h: 0.35,
    fontSize: 15, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText("→ 機械的溶血, TMA\n→ 遺伝性（HS, G6PD）\n→ PNH\n※ 破砕赤血球>1% → TMA疑い", {
    x: 5.55, y: 3.5, w: 3.8, h: 1.2,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3
  });
})();

// ══════════════════════════════════════════════════════
// Slide 12: Red Flags
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };

  // タイトル（赤アクセント）
  s.addText("🚨 Red Flags ─ 見逃してはいけない所見", {
    x: 0.5, y: 0.25, w: 9, h: 0.7,
    fontSize: 28, fontFace: FONT_JP, color: C.red, bold: true, margin: 0
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 1.5, h: 0.05, fill: { color: C.red } });

  addFooter(s);

  var flags = [
    { flag: "TTP / MAHA", action: "治療を遅らせない → 血漿交換" },
    { flag: "重度貧血＋循環不安定", action: "Hb<7で輸血検討、症状優先" },
    { flag: "汎血球減少＋芽球", action: "AA / MDS / 白血病 → 緊急紹介" },
    { flag: "中高年の新規IDA", action: "消化管悪性腫瘍を除外（上下部内視鏡）" },
    { flag: "神経症状＋巨赤芽球性", action: "B12欠乏 → 不可逆あり → 早期補充" },
    { flag: "骨痛＋高Ca＋蛋白異常", action: "多発性骨髄腫を疑う" }
  ];

  for (var i = 0; i < flags.length; i++) {
    var fy = 1.2 + i * 0.62;

    // カード背景（薄い赤）
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: fy, w: 9, h: 0.52,
      fill: { color: C.ltRed },
      line: { color: C.red, width: 1 },
      rectRadius: 0.08
    });

    // 番号（赤丸＋白文字）
    s.addShape(pres.shapes.OVAL, {
      x: 0.65, y: fy + 0.06, w: 0.4, h: 0.4,
      fill: { color: C.red }
    });
    s.addText(String(i + 1), {
      x: 0.65, y: fy + 0.06, w: 0.4, h: 0.4,
      fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0
    });

    // フラグ名（濃い文字）
    s.addText(flags[i].flag, {
      x: 1.2, y: fy, w: 3.3, h: 0.52,
      fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true,
      valign: "middle", margin: 0
    });

    // アクション（通常文字色）
    s.addText("→ " + flags[i].action, {
      x: 4.6, y: fy, w: 4.7, h: 0.52,
      fontSize: 13, fontFace: FONT_JP, color: C.text,
      valign: "middle", margin: 0
    });
  }
})();

// ══════════════════════════════════════════════════════
// Slide 13: Take-home Message
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Take-home Message");
  addFooter(s);

  var messages = [
    { num: "1", text: "まず緊急性を判断（TTP, 出血性ショック, B12＋神経症状）" },
    { num: "2", text: "MCV＋網状赤血球で「産生低下 vs 喪失/破壊」を判断" },
    { num: "3", text: "フェリチンは炎症下で偽高値になる → TSAT・sTfRを併用" },
    { num: "4", text: "DATで溶血の免疫性（AIHA）vs 非免疫性（TMA等）を分岐" },
    { num: "5", text: "中高年の新規IDAは消化管悪性腫瘍の除外が必須" }
  ];

  for (var i = 0; i < messages.length; i++) {
    var my = 1.3 + i * 0.72;
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
pres.writeFile({ fileName: "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/貧血の鑑別/貧血の鑑別.pptx" })
  .then(function() { console.log("Done"); })
  .catch(function(e) { console.error(e); });
