var pptxgen = require("pptxgenjs");
var pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "医知創造ラボ";
pres.title = "意識消失発作（TLOC）の鑑別診断";

var C = {
  navy: "1B3A5C", blue: "2C5AA0", ltBlue: "E8F0FA", bg: "F0F4F8",
  white: "FFFFFF", text: "333333", red: "DC3545", yellow: "F5C518",
  green: "28A745", orange: "E8850C", ltGray: "F8F9FA", gray: "555555", subText: "666666",
  dkRed: "8B1A1A", ltRed: "FDE8E8", ltGreen: "E6F7EC", ltYellow: "FFF8E1",
  ltOrange: "FFF3E0", purple: "6F42C1", ltPurple: "F0E6FF",
  brown: "5D4037", ltBrown: "EFEBE9"
};
var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var mkShadow = function() { return { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 }; };

// ── フッター ──
function addFooter(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.15, w: 10, h: 0.475, fill: { color: C.navy } });
  slide.addText("意識消失発作（TLOC）の鑑別診断", { x: 0.5, y: 5.15, w: 7, h: 0.475, fontSize: 18, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
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

// テーブルヘルパー
function thCell(text) {
  return { text: text, options: { fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, fill: { color: C.navy }, align: "center", valign: "middle" } };
}
function tdCell(text, bg) {
  return { text: text, options: { fontSize: 12, fontFace: FONT_JP, color: C.text, fill: { color: bg || C.white }, valign: "middle" } };
}
function tdCellCenter(text, bg) {
  return { text: text, options: { fontSize: 12, fontFace: FONT_JP, color: C.text, fill: { color: bg || C.white }, align: "center", valign: "middle" } };
}

// ══════════════════════════════════════════════════════
// Slide 1: タイトルスライド
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.blue } });

  s.addText("意識消失発作（TLOC）の鑑別診断", {
    x: 0.5, y: 1.3, w: 9, h: 1.3,
    fontSize: 48, fontFace: FONT_JP, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 3, y: 2.85, w: 4, h: 0.05, fill: { color: C.yellow } });

  s.addText("― 失神・てんかん・PNESを7ステップで系統的に鑑別 ―", {
    x: 0.5, y: 3.1, w: 9, h: 0.8,
    fontSize: 24, fontFace: FONT_JP, color: C.ltBlue,
    align: "center", valign: "middle", margin: 0
  });

  s.addText("ESC 2018 / NICE ガイドライン準拠", {
    x: 0.5, y: 4.0, w: 9, h: 0.5,
    fontSize: 16, fontFace: FONT_EN, color: C.ltBlue,
    align: "center", valign: "middle", margin: 0
  });

  s.addText("医知創造ラボ", {
    x: 0.5, y: 4.6, w: 9, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.ltBlue,
    align: "center", valign: "middle", margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.57, w: 10, h: 0.06, fill: { color: C.blue } });
})();

// ══════════════════════════════════════════════════════
// Slide 2: TLOCの分類 ─ 4つの大カテゴリー
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "TLOCの分類 ─ 4つの大カテゴリー");
  addFooter(s);

  var cats = [
    { emoji: "😵", title: "失神（Syncope）", desc: "脳血流の一過性低下\n反射性（60%）/ 起立性 / 心原性", color: C.blue, bg: C.ltBlue },
    { emoji: "⚡", title: "てんかん発作", desc: "GTCS / 焦点意識減損発作\n/ 欠神発作", color: C.purple, bg: C.ltPurple },
    { emoji: "🎭", title: "PNES", desc: "心因性非てんかん性発作\n（機能性神経症状症）", color: C.orange, bg: C.ltOrange },
    { emoji: "🔍", title: "まれな原因", desc: "椎骨脳底動脈TIA / 鎖骨下盗血\n/ 低血糖 / SAH", color: C.brown, bg: C.ltBrown }
  ];

  for (var i = 0; i < cats.length; i++) {
    var cx = 0.4 + i * 2.4;
    addCard(s, cx, 1.35, 2.15, 3.4, cats[i].bg, cats[i].color);

    s.addText(cats[i].emoji, {
      x: cx, y: 1.5, w: 2.15, h: 0.55,
      fontSize: 30, align: "center", valign: "middle", margin: 0
    });
    s.addText(cats[i].title, {
      x: cx + 0.1, y: 2.1, w: 1.95, h: 0.5,
      fontSize: 15, fontFace: FONT_JP, color: cats[i].color, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(cats[i].desc, {
      x: cx + 0.1, y: 2.7, w: 1.95, h: 1.8,
      fontSize: 12, fontFace: FONT_JP, color: C.text,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.4
    });
  }
})();

// ══════════════════════════════════════════════════════
// Slide 3: 初期評価の3本柱
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "初期評価の3本柱（ESC 2018）");
  addFooter(s);

  var pillars = [
    { num: "①", title: "詳細な病歴聴取", desc: "患者本人＋目撃者\nこれだけで50-60%は\n診断可能", color: C.blue },
    { num: "②", title: "身体診察", desc: "臥位→立位血圧測定\n（3分後）を必ず含む\n起立性低血圧の検出", color: C.green },
    { num: "③", title: "12誘導心電図", desc: "全例に必須\n心原性失神の\nスクリーニング", color: C.red }
  ];

  for (var i = 0; i < pillars.length; i++) {
    var px = 0.6 + i * 3.1;
    addCard(s, px, 1.35, 2.8, 3.0, C.white, pillars[i].color);

    s.addShape(pres.shapes.OVAL, {
      x: px + 1.0, y: 1.55, w: 0.65, h: 0.65,
      fill: { color: pillars[i].color }
    });
    s.addText(pillars[i].num, {
      x: px + 1.0, y: 1.55, w: 0.65, h: 0.65,
      fontSize: 22, fontFace: FONT_JP, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(pillars[i].title, {
      x: px + 0.15, y: 2.35, w: 2.5, h: 0.5,
      fontSize: 17, fontFace: FONT_JP, color: pillars[i].color, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(pillars[i].desc, {
      x: px + 0.15, y: 2.95, w: 2.5, h: 1.2,
      fontSize: 13, fontFace: FONT_JP, color: C.text,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.4
    });
  }

  addCard(s, 0.6, 4.55, 8.8, 0.45, C.ltBlue, C.blue);
  s.addText("💡 初期評価だけで過半数が診断可能 — 病歴聴取が最も重要", {
    x: 0.8, y: 4.55, w: 8.4, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.navy, bold: true,
    align: "center", valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 4: Step 1 誘因・状況
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 1  誘因・状況 ─ 最初の分岐点");
  addFooter(s);

  var rows = [
    [thCell("誘因"), thCell("示唆される診断"), thCell("重要度")],
    [tdCell("長時間立位 / 痛み / 採血 / 感情ストレス"), tdCell("血管迷走神経性失神"), tdCellCenter("一般的", C.ltGreen)],
    [tdCell("排尿 / 咳嗽 / 嚥下時"), tdCell("状況失神"), tdCellCenter("一般的", C.ltGreen)],
    [tdCell("起立直後"), tdCell("起立性低血圧"), tdCellCenter("一般的", C.ltGreen)],
    [tdCell("運動中・労作中", C.ltRed), tdCell("心原性失神", C.ltRed), tdCellCenter("🚨 赤旗", C.ltRed)],
    [tdCell("臥位・安静時", C.ltRed), tdCell("心原性（不整脈）", C.ltRed), tdCellCenter("🚨 赤旗", C.ltRed)],
    [tdCell("睡眠中", C.ltPurple), tdCell("てんかん", C.ltPurple), tdCellCenter("⚡ 重要", C.ltPurple)],
    [tdCell("上肢運動中（片側）"), tdCell("鎖骨下動脈盗血症候群"), tdCellCenter("まれ", C.ltGray)]
  ];

  s.addTable(rows, {
    x: 0.4, y: 1.25, w: 9.2,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [3.4, 3.4, 1.6],
    rowH: [0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42],
    margin: [3, 6, 3, 6]
  });
})();

// ══════════════════════════════════════════════════════
// Slide 5: Step 2 前駆症状
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 2  前駆症状 ─ 診断の宝庫");
  addFooter(s);

  var rows = [
    [thCell("前駆症状"), thCell("示唆される診断")],
    [tdCell("発汗・悪心・蒼白・視野暗転"), tdCell("血管迷走神経性失神（自律神経症状）")],
    [tdCell("立ちくらみ・ふらつき", C.ltGreen), tdCell("起立性低血圧", C.ltGreen)],
    [tdCell("デジャビュ / 上腹部不快感の上行 / 恐怖感", C.ltPurple), tdCell("焦点発作（側頭葉てんかん）", C.ltPurple)],
    [tdCell("動悸"), tdCell("不整脈性失神")],
    [tdCell("突然の激しい頭痛", C.ltRed), tdCell("くも膜下出血", C.ltRed)],
    [tdCell("前駆症状なし（突然の意識消失）", C.ltRed), tdCell("心原性失神を最も疑う", C.ltRed)]
  ];

  s.addTable(rows, {
    x: 0.4, y: 1.25, w: 9.2,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [4.6, 4.6],
    rowH: [0.45, 0.45, 0.45, 0.5, 0.45, 0.45, 0.5],
    margin: [3, 6, 3, 6]
  });

  addCard(s, 0.4, 4.55, 9.2, 0.45, C.ltRed, C.red);
  s.addText("🚨 前駆症状なしの突然の意識消失 → 心原性失神を最も疑う", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true,
    align: "center", valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 6: Step 3 発作中の所見 ─ 目撃者情報
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 3  発作中の所見 ─ 目撃者情報が鍵");
  addFooter(s);

  var rows = [
    [thCell("所見"), thCell("失神"), thCell("てんかん"), thCell("PNES")],
    [tdCell("顔色"), tdCellCenter("蒼白", C.ltBlue), tdCellCenter("チアノーゼ", C.ltPurple), tdCellCenter("正常")],
    [tdCell("持続時間"), tdCellCenter("<30秒", C.ltBlue), tdCellCenter("1-2分", C.ltPurple), tdCellCenter(">5分", C.ltOrange)],
    [tdCell("開眼/閉眼"), tdCellCenter("開眼", C.ltBlue), tdCellCenter("開眼", C.ltPurple), tdCellCenter("閉眼★", C.ltOrange)],
    [tdCell("運動パターン"), tdCellCenter("多焦点性\nミオクローヌス\n(<10回)"), tdCellCenter("強直間代\n律動的\n(>20回)"), tdCellCenter("非同期性\n側方動揺")],
    [tdCell("舌咬傷"), tdCellCenter("尖部（稀）"), tdCellCenter("側面★\n(LR+ 49)", C.ltPurple), tdCellCenter("尖部/なし")],
    [tdCell("発作後の回復"), tdCellCenter("速やか\n(<1分)", C.ltBlue), tdCellCenter("遷延混迷\n(>5分)", C.ltPurple), tdCellCenter("変動的", C.ltOrange)]
  ];

  s.addTable(rows, {
    x: 0.3, y: 1.2, w: 9.4,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [1.8, 2.4, 2.6, 2.6],
    rowH: [0.4, 0.4, 0.4, 0.4, 0.7, 0.5, 0.55],
    margin: [2, 4, 2, 4]
  });
})();

// ══════════════════════════════════════════════════════
// Slide 7: 痙攣性失神の落とし穴
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "⚠️ 痙攣性失神（Convulsive Syncope）");
  addFooter(s);

  addCard(s, 0.6, 1.3, 8.8, 1.2, C.ltYellow, C.orange);
  s.addText("失神でも約12%に四肢のミオクローヌスが出現\n→ てんかんと誤診しないこと！", {
    x: 0.8, y: 1.35, w: 8.4, h: 1.1,
    fontSize: 18, fontFace: FONT_JP, color: C.text, bold: true,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.4
  });

  // 10/20ルール
  s.addText("10/20 ルール", {
    x: 0.6, y: 2.75, w: 8.8, h: 0.5,
    fontSize: 22, fontFace: FONT_EN, color: C.navy, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // 左カード: <10回
  addCard(s, 0.8, 3.4, 3.8, 1.4, C.ltBlue, C.blue);
  s.addText("ミオクローヌス < 10回", {
    x: 0.9, y: 3.5, w: 3.6, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.blue, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("→ 失神を示唆\n多焦点性・不規則・短時間", {
    x: 0.9, y: 3.95, w: 3.6, h: 0.75,
    fontSize: 14, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3
  });

  // 右カード: >20回
  addCard(s, 5.4, 3.4, 3.8, 1.4, C.ltPurple, C.purple);
  s.addText("ミオクローヌス > 20回", {
    x: 5.5, y: 3.5, w: 3.6, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.purple, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s.addText("→ てんかんを示唆\n律動的・全身同期性", {
    x: 5.5, y: 3.95, w: 3.6, h: 0.75,
    fontSize: 14, fontFace: FONT_JP, color: C.text,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3
  });

  // VS
  s.addShape(pres.shapes.OVAL, {
    x: 4.55, y: 3.8, w: 0.9, h: 0.55,
    fill: { color: C.navy }
  });
  s.addText("VS", {
    x: 4.55, y: 3.8, w: 0.9, h: 0.55,
    fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 8: 舌咬傷の診断的価値
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "舌咬傷の診断的価値");
  addFooter(s);

  var rows = [
    [thCell("舌咬傷の部位"), thCell("感度"), thCell("特異度"), thCell("LR+"), thCell("示唆")],
    [tdCell("舌咬傷（全般）"), tdCellCenter("33%"), tdCellCenter("96%"), tdCellCenter("8.2"), tdCell("てんかん")],
    [tdCell("舌側面の咬傷", C.ltPurple), tdCellCenter("11.3%", C.ltPurple), tdCellCenter("99.8%", C.ltPurple), tdCellCenter("49.3", C.ltPurple), tdCell("てんかんにほぼ特異的", C.ltPurple)],
    [tdCell("舌尖部の咬傷"), tdCellCenter("—"), tdCellCenter("—"), tdCellCenter("—"), tdCell("失神 or PNES")]
  ];

  s.addTable(rows, {
    x: 0.4, y: 1.3, w: 9.2,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [2.5, 1.3, 1.3, 1.1, 3.0],
    rowH: [0.45, 0.45, 0.55, 0.45],
    margin: [3, 6, 3, 6]
  });

  addCard(s, 0.6, 3.4, 8.8, 1.5, C.white, C.purple);
  s.addText("🔑 臨床的位置づけ", {
    x: 0.8, y: 3.5, w: 3, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.purple, bold: true, margin: 0
  });
  s.addText("• 舌側面の咬傷があれば → てんかんの診断はほぼ確実（LR+ 49）\n• 感度は11%と低い → なくても除外はできない\n• 舌尖部の咬傷 → 失神やPNESでも起こりうる → 鑑別には使えない", {
    x: 0.8, y: 3.95, w: 8.4, h: 0.9,
    fontSize: 14, fontFace: FONT_JP, color: C.text,
    margin: 0, lineSpacingMultiple: 1.4
  });
})();

// ══════════════════════════════════════════════════════
// Slide 9: Step 4 発作後の所見
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 4  発作後の所見 ─ 最も信頼できる鑑別点");
  addFooter(s);

  var items = [
    { label: "失神", desc: "速やかな回復（<1分）\n横になれば血流回復", color: C.blue, bg: C.ltBlue },
    { label: "てんかん", desc: "遷延する意識混濁（>5分）\n筋肉痛・頭痛・CK上昇", color: C.purple, bg: C.ltPurple },
    { label: "PNES", desc: "変動的な回復パターン\n発作中の記憶が部分的にあり", color: C.orange, bg: C.ltOrange }
  ];

  for (var i = 0; i < items.length; i++) {
    var ix = 0.5 + i * 3.15;
    addCard(s, ix, 1.3, 2.85, 2.2, items[i].bg, items[i].color);

    s.addText(items[i].label, {
      x: ix + 0.15, y: 1.4, w: 2.55, h: 0.5,
      fontSize: 20, fontFace: FONT_JP, color: items[i].color, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(items[i].desc, {
      x: ix + 0.15, y: 2.0, w: 2.55, h: 1.3,
      fontSize: 14, fontFace: FONT_JP, color: C.text,
      align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.4
    });
  }

  // Todd麻痺
  addCard(s, 0.5, 3.75, 4.5, 0.85, C.white, C.purple);
  s.addText("⚡ Todd麻痺（一過性片麻痺）", {
    x: 0.65, y: 3.8, w: 4.2, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.purple, bold: true, margin: 0
  });
  s.addText("→ 焦点起始てんかんの診断がほぼ確実", {
    x: 0.65, y: 4.15, w: 4.2, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0
  });

  // 部分記憶
  addCard(s, 5.15, 3.75, 4.35, 0.85, C.white, C.orange);
  s.addText("🎭 発作中の出来事を部分的に記憶", {
    x: 5.3, y: 3.8, w: 4.05, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0
  });
  s.addText("→ PNESを示唆（てんかん・失神では稀）", {
    x: 5.3, y: 4.15, w: 4.05, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 10: PNESの鑑別 ─ 5つのサイン
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "PNES（心因性非てんかん性発作）─ 5つのサイン");
  addFooter(s);

  var signs = [
    { num: "❶", title: "閉眼", desc: "特異度92% — てんかん・失神は通常開眼", highlight: true },
    { num: "❷", title: "5分以上の持続", desc: "真のてんかん発作は1-2分。5分超はPNESの可能性24倍", highlight: false },
    { num: "❸", title: "非同期性四肢運動", desc: "不規則・骨盤突出運動も特徴的", highlight: false },
    { num: "❹", title: "変動する経過", desc: "感度76% — 最も信頼できる徴候", highlight: true },
    { num: "❺", title: "発作中の部分的記憶", desc: "てんかん・失神ではまず見られない", highlight: false }
  ];

  for (var i = 0; i < signs.length; i++) {
    var sy = 1.3 + i * 0.72;
    var bgCol = signs[i].highlight ? C.ltOrange : C.white;
    addCard(s, 0.6, sy, 8.8, 0.62, bgCol, C.orange);

    s.addText(signs[i].num, {
      x: 0.7, y: sy, w: 0.5, h: 0.62,
      fontSize: 20, align: "center", valign: "middle", margin: 0
    });
    s.addText(signs[i].title, {
      x: 1.3, y: sy, w: 2.5, h: 0.62,
      fontSize: 16, fontFace: FONT_JP, color: C.orange, bold: true,
      valign: "middle", margin: 0
    });
    s.addText(signs[i].desc, {
      x: 3.9, y: sy, w: 5.3, h: 0.62,
      fontSize: 13, fontFace: FONT_JP, color: C.text,
      valign: "middle", margin: 0
    });
  }

  addCard(s, 0.6, 4.95, 8.8, 0.14, C.orange, C.orange);
  s.addText("確定診断：ビデオ脳波モニタリング（vEEG）— 発作中の脳波が正常であることを確認", {
    x: 0.6, y: 4.6, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.gray, align: "center", valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 11: Step 5 心原性レッドフラグ
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 5  心原性失神のレッドフラグ");
  addFooter(s);

  // 上部警告バー
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.25, w: 8.8, h: 0.55,
    fill: { color: C.red }, rectRadius: 0.08
  });
  s.addText("🚨 心原性失神は全失神の10%だが、2年死亡率17-21%", {
    x: 0.8, y: 1.25, w: 8.4, h: 0.55,
    fontSize: 17, fontFace: FONT_JP, color: C.white, bold: true,
    valign: "middle", margin: 0
  });

  var flags = [
    { emoji: "⚡", text: "前駆症状なしの突然の意識消失" },
    { emoji: "🏃", text: "運動中・労作中の失神" },
    { emoji: "🛏️", text: "臥位での失神" },
    { emoji: "📊", text: "心電図異常（QT延長、Brugada型、二束ブロック）" },
    { emoji: "🫀", text: "構造的心疾患の既往 / 心不全の徴候" },
    { emoji: "👨‍👩‍👧‍👦", text: "40歳未満の突然死の家族歴" },
    { emoji: "💓", text: "失神前の動悸" }
  ];

  for (var i = 0; i < flags.length; i++) {
    var fy = 2.0 + i * 0.42;
    var fbg = (i < 3) ? C.ltRed : C.white;
    addCard(s, 0.6, fy, 8.8, 0.37, fbg, C.red);
    s.addText(flags[i].emoji + "  " + flags[i].text, {
      x: 0.8, y: fy, w: 8.4, h: 0.37,
      fontSize: 14, fontFace: FONT_JP, color: C.text, bold: (i < 3),
      valign: "middle", margin: 0
    });
  }

  s.addText("1つでも該当 → 循環器内科への相談が必要", {
    x: 0.6, y: 4.98, w: 8.8, h: 0.14,
    fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 12: EGSYS Score
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "EGSYS Score ─ 心原性失神のスクリーニング");
  addFooter(s);

  var rows = [
    [thCell("所見"), thCell("点数")],
    [tdCell("失神前の動悸"), tdCellCenter("+4", C.ltRed)],
    [tdCell("心電図異常 and/or 心疾患"), tdCellCenter("+3", C.ltRed)],
    [tdCell("労作中の失神"), tdCellCenter("+3", C.ltRed)],
    [tdCell("臥位での失神"), tdCellCenter("+2", C.ltOrange)],
    [tdCell("自律神経性前駆症状（発汗・悪心・温感）", C.ltGreen), tdCellCenter("-1", C.ltGreen)],
    [tdCell("誘因あり（温環境、長時間立位、恐怖、痛み）", C.ltGreen), tdCellCenter("-1", C.ltGreen)]
  ];

  s.addTable(rows, {
    x: 0.6, y: 1.3, w: 8.8,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [7.0, 1.8],
    rowH: [0.42, 0.42, 0.42, 0.42, 0.42, 0.45, 0.45],
    margin: [3, 6, 3, 6]
  });

  addCard(s, 0.6, 4.45, 8.8, 0.55, C.ltRed, C.red);
  s.addText("≧3点：心原性失神の感度92%・特異度69% — 2年死亡率17-21% → 入院精査を強く考慮", {
    x: 0.8, y: 4.45, w: 8.4, h: 0.55,
    fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true,
    align: "center", valign: "middle", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 13: Step 6-7 患者背景・家族歴
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "Step 6-7  患者背景・家族歴");
  addFooter(s);

  // 左半分: 背景
  addCard(s, 0.4, 1.3, 4.4, 3.5, C.white, C.blue);
  s.addText("📋 患者背景", {
    x: 0.55, y: 1.35, w: 4.1, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });
  s.addText(
    "• 若年者 → 血管迷走神経性が最多\n" +
    "• 高齢者 → 心原性・起立性・頸動脈洞\n" +
    "• てんかん既往 → てんかん発作\n" +
    "• 心疾患既往 → 心原性失神\n" +
    "• PD/DM → 起立性低血圧（自律神経障害）\n" +
    "• 精神疾患 → PNES\n" +
    "• 降圧薬/利尿薬 → 起立性低血圧\n" +
    "• インスリン/SU薬 → 低血糖",
    {
      x: 0.55, y: 1.85, w: 4.1, h: 2.8,
      fontSize: 13, fontFace: FONT_JP, color: C.text,
      margin: 0, lineSpacingMultiple: 1.35
    }
  );

  // 右半分: 家族歴
  addCard(s, 5.2, 1.3, 4.4, 3.5, C.white, C.red);
  s.addText("👨‍👩‍👧‍👦 家族歴", {
    x: 5.35, y: 1.35, w: 4.1, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.red, bold: true, margin: 0
  });
  s.addText(
    "🚨 40歳未満の突然死\n→ QT延長症候群\n→ Brugada症候群\n→ 肥大型心筋症\n\n" +
    "⚡ てんかんの家族歴\n→ てんかんの素因\n\n" +
    "💓 チャネル病の家族歴\n→ 心電図精査必須",
    {
      x: 5.35, y: 1.85, w: 4.1, h: 2.8,
      fontSize: 13, fontFace: FONT_JP, color: C.text,
      margin: 0, lineSpacingMultiple: 1.3
    }
  );
})();

// ══════════════════════════════════════════════════════
// Slide 14: 検査アプローチ
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "検査アプローチ ─ 疑い別の追加検査");
  addFooter(s);

  var rows = [
    [thCell("疑い"), thCell("検査"), thCell("備考")],
    [tdCell("反射性失神", C.ltBlue), tdCell("チルト試験、頸動脈洞マッサージ", C.ltBlue), tdCell("再発性/不確実時。CSM>40歳", C.ltBlue)],
    [tdCell("起立性低血圧"), tdCell("臥位→立位BP（3分後）"), tdCell("SBP↓20 or DBP↓10")],
    [tdCell("不整脈性", C.ltRed), tdCell("ホルター、ILR、EPS", C.ltRed), tdCell("ILRは原因不明再発性に有用", C.ltRed)],
    [tdCell("構造的心疾患"), tdCell("心エコー、負荷試験、BNP"), tdCell("労作中失神→運動負荷試験")],
    [tdCell("てんかん", C.ltPurple), tdCell("EEG、頭部MRI", C.ltPurple), tdCell("EEG1回の感度≈50%", C.ltPurple)],
    [tdCell("PNES"), tdCell("ビデオ脳波モニタリング"), tdCell("確定診断のゴールドスタンダード")],
    [tdCell("SAH", C.ltRed), tdCell("頭部CT（6h以内）、LP", C.ltRed), tdCell("6h以内のCT感度>98%", C.ltRed)]
  ];

  s.addTable(rows, {
    x: 0.3, y: 1.25, w: 9.4,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [2.0, 3.7, 3.7],
    rowH: [0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42],
    margin: [3, 5, 3, 5]
  });

  s.addText("※ ルーチンのEEG・頭部画像は失神の初期評価としては非推奨（神経所見がない限り）", {
    x: 0.6, y: 4.85, w: 8.8, h: 0.25,
    fontSize: 11, fontFace: FONT_JP, color: C.gray, align: "center", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 15: 鑑別のまとめ表
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "鑑別のまとめ");
  addFooter(s);

  var rows = [
    [thCell(""), thCell("失神"), thCell("てんかん"), thCell("PNES"), thCell("心原性")],
    [tdCell("顔色"), tdCellCenter("蒼白", C.ltBlue), tdCellCenter("チアノーゼ", C.ltPurple), tdCellCenter("正常", C.ltOrange), tdCellCenter("—")],
    [tdCell("持続"), tdCellCenter("<30秒", C.ltBlue), tdCellCenter("1-2分", C.ltPurple), tdCellCenter(">5分", C.ltOrange), tdCellCenter("—")],
    [tdCell("回復"), tdCellCenter("速やか", C.ltBlue), tdCellCenter("遷延混迷", C.ltPurple), tdCellCenter("変動的", C.ltOrange), tdCellCenter("速やか")],
    [tdCell("開閉眼"), tdCellCenter("開眼"), tdCellCenter("開眼"), tdCellCenter("閉眼★", C.ltOrange), tdCellCenter("—")],
    [tdCell("舌咬傷"), tdCellCenter("尖部"), tdCellCenter("側面★", C.ltPurple), tdCellCenter("稀"), tdCellCenter("—")],
    [tdCell("誘因"), tdCellCenter("明確"), tdCellCenter("なし/睡眠中"), tdCellCenter("ストレス"), tdCellCenter("労作/臥位")],
    [tdCell("前兆"), tdCellCenter("自律神経症状"), tdCellCenter("デジャビュ等"), tdCellCenter("—"), tdCellCenter("動悸/なし")]
  ];

  s.addTable(rows, {
    x: 0.2, y: 1.2, w: 9.6,
    border: { type: "solid", color: C.blue, pt: 0.8 },
    colW: [1.2, 2.1, 2.1, 2.1, 2.1],
    rowH: [0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42],
    margin: [2, 4, 2, 4]
  });

  s.addText("★ = 高い特異度を持つ所見", {
    x: 0.6, y: 4.75, w: 8.8, h: 0.3,
    fontSize: 12, fontFace: FONT_JP, color: C.gray, align: "center", margin: 0
  });
})();

// ══════════════════════════════════════════════════════
// Slide 16: まとめ ─ Take Home Message
// ══════════════════════════════════════════════════════
(function() {
  var s = pres.addSlide();
  s.background = { color: C.bg };
  addSectionTitle(s, "まとめ ─ Take Home Message");
  addFooter(s);

  var messages = [
    { num: "1", text: "初期評価は「病歴＋身体診察＋心電図」— 50-60%はこれで診断可能", color: C.blue },
    { num: "2", text: "「蒼白 vs チアノーゼ」を目撃者に聞く — 失神 vs てんかんの最初の手がかり", color: C.blue },
    { num: "3", text: "舌側面の咬傷（LR+ 49）→ てんかんにほぼ特異的", color: C.purple },
    { num: "4", text: "閉眼（特異度92%）→ PNESの重要サイン", color: C.orange },
    { num: "5", text: "前駆症状なしの突然の意識消失 → 心原性を最も疑う", color: C.red },
    { num: "6", text: "運動中失神 ＋ 突然死家族歴 → チャネル病/HCMの除外必須", color: C.red },
    { num: "7", text: "EGSYS ≧3点 → 2年死亡率17-21%、入院精査を強く考慮", color: C.red }
  ];

  for (var i = 0; i < messages.length; i++) {
    var my = 1.2 + i * 0.52;
    addBadge(s, 0.6, my + 0.06, messages[i].num, messages[i].color);
    s.addText(messages[i].text, {
      x: 1.2, y: my, w: 8.2, h: 0.48,
      fontSize: 14, fontFace: FONT_JP, color: C.text,
      valign: "middle", margin: 0
    });
  }

  addCard(s, 0.6, 4.9, 8.8, 0.14, C.blue, C.blue);
})();

// ── 出力 ──
var outPath = __dirname + "/tloc_ddx_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("✅ スライド生成完了: " + outPath);
}).catch(function(err) {
  console.error("❌ エラー:", err);
});
