var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "令和8年度診療報酬改定【糖尿病・内分泌編】";

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

  s.addText("令和8年6月施行", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("令和8年度 診療報酬改定", {
    x: 0.8, y: 1.1, w: 8.4, h: 0.8,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addText("【糖尿病・内分泌編】", {
    x: 0.8, y: 1.9, w: 8.4, h: 0.7,
    fontSize: 28, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", margin: 0,
  });

  s.addText("生活習慣病管理料・眼科歯科連携加算・透析・薬価", {
    x: 0.8, y: 2.7, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.45, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("厚生労働省公式資料・糖尿病標準診療マニュアル2026年版に基づく", {
    x: 0.8, y: 3.7, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.light, italic: true, align: "center", margin: 0,
  });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.3, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.8, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 今回の改定ポイント一覧
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "糖尿病・内分泌領域の主要改定ポイント");

  var items = [
    { icon: "1", title: "生活習慣病管理料（I）（II）の見直し", desc: "包括範囲縮小、血液検査要件化、署名不要" },
    { icon: "2", title: "眼科医療機関連携強化加算【新設】", desc: "60点/年1回 ― 糖尿病網膜症の早期発見" },
    { icon: "3", title: "歯科医療機関連携強化加算【新設】", desc: "60点/年1回 ― 歯周病との双方向的関連" },
    { icon: "4", title: "充実管理加算の3段階再編", desc: "30点/20点/10点 ― 実績に基づく評価" },
    { icon: "5", title: "在宅自己注射の適用拡大", desc: "糖尿病以外の併存疾患にも算定可能" },
    { icon: "6", title: "透析：人工腎臓▲20点＋充実加算+20点", desc: "腎代替療法診療体制充実加算の新設" },
  ];

  for (var i = 0; i < items.length; i++) {
    var yBase = 1.1 + i * 0.7;
    // 番号バッジ
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: yBase, w: 0.5, h: 0.5,
      fill: { color: C.primary }, rectRadius: 0.08,
    });
    s.addText(items[i].icon, {
      x: 0.5, y: yBase, w: 0.5, h: 0.5,
      fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    // タイトル
    s.addText(items[i].title, {
      x: 1.2, y: yBase - 0.02, w: 5.5, h: 0.3,
      fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
    });
    // 説明
    s.addText(items[i].desc, {
      x: 1.2, y: yBase + 0.28, w: 8, h: 0.25,
      fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
    });
  }

  addPageNum(s, "2");
})();

// ============================================================
// SLIDE 3: 生活習慣病管理料の見直し
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "生活習慣病管理料（I）（II）の見直し");

  // カード1: 包括範囲の縮小
  addCard(s, 0.4, 1.1, 4.4, 1.8);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.25, w: 2.8, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("包括範囲の縮小", {
    x: 0.6, y: 1.25, w: 2.8, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("併算定可能な医学管理料", {
    x: 0.6, y: 1.8, w: 4, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });
  s.addText([
    { text: "16項目", options: { fontSize: 28, color: C.sub, strike: true } },
    { text: "  ", options: { fontSize: 20 } },
    { text: "→", options: { fontSize: 22, color: C.red } },
    { text: "  ", options: { fontSize: 20 } },
    { text: "37項目", options: { fontSize: 28, color: C.red, bold: true } },
  ], {
    x: 0.6, y: 2.15, w: 4, h: 0.55,
    fontFace: FONT_EN, align: "center", valign: "middle", margin: 0,
  });

  // カード2: 血液検査の要件化
  addCard(s, 5.2, 1.1, 4.4, 1.8);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.25, w: 2.8, h: 0.35, fill: { color: C.orange }, rectRadius: 0.06,
  });
  s.addText("血液検査の要件化", {
    x: 5.4, y: 1.25, w: 2.8, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("管理料（I）の算定要件として追加", {
    x: 5.4, y: 1.8, w: 4, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });
  s.addText("6か月に1回以上", {
    x: 5.4, y: 2.15, w: 4, h: 0.55,
    fontSize: 24, fontFace: FONT_JP, color: C.orange, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // カード3: 療養計画書の簡素化
  addCard(s, 0.4, 3.15, 9.2, 1.2);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.3, w: 2.8, h: 0.35, fill: { color: C.green }, rectRadius: 0.06,
  });
  s.addText("療養計画書の簡素化", {
    x: 0.6, y: 3.3, w: 2.8, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("患者署名が不要に → 事務負担軽減", {
    x: 3.6, y: 3.3, w: 5.5, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });
  s.addText("糖尿病＋他疾患を併存する患者の管理料算定の自由度が大幅に向上。併算定可能項目を確認し、漏れなく算定することが重要。", {
    x: 0.6, y: 3.8, w: 8.8, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  addPageNum(s, "3");
})();

// ============================================================
// SLIDE 4: 眼科医療機関連携強化加算
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "眼科医療機関連携強化加算【新設】");

  // メインカード
  addCard(s, 0.4, 1.1, 5.8, 2.0);

  // 60点バッジ
  s.addShape(pres.shapes.OVAL, {
    x: 0.7, y: 1.3, w: 1.3, h: 1.3,
    fill: { color: C.red },
  });
  s.addText("60点", {
    x: 0.7, y: 1.45, w: 1.3, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("年1回", {
    x: 0.7, y: 2.0, w: 1.3, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.white, align: "center", valign: "middle", margin: 0,
  });

  // 項目
  var details = [
    "対象：糖尿病を主病とする患者",
    "要件：眼科受診の必要性を認め連携を実施",
    "目的：糖尿病網膜症の早期発見・治療",
    "管理料（I）・（II）の双方で算定可",
  ];
  for (var i = 0; i < details.length; i++) {
    s.addText(details[i], {
      x: 2.2, y: 1.3 + i * 0.42, w: 3.8, h: 0.35,
      fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  }

  // 右側：臨床的意義カード
  addCard(s, 6.5, 1.1, 3.2, 2.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.7, y: 1.25, w: 2.0, h: 0.3, fill: { color: C.orange }, rectRadius: 0.06,
  });
  s.addText("臨床的意義", {
    x: 6.7, y: 1.25, w: 2.0, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("糖尿病網膜症は無症状のまま進行。増殖網膜症では失明リスクあり。内科通院中でも眼科未受診の患者は少なくない。", {
    x: 6.7, y: 1.7, w: 2.8, h: 1.2,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 下部：フロー図
  addCard(s, 0.4, 3.4, 9.2, 1.3);
  var flowItems = [
    { x: 0.8, label: "内科主治医", color: C.primary },
    { x: 3.0, label: "眼科受診の\n必要性を判断", color: C.accent },
    { x: 5.2, label: "患者同意を取得\n連携を実施", color: C.green },
    { x: 7.4, label: "眼科を受診\n60点算定", color: C.red },
  ];
  for (var i = 0; i < flowItems.length; i++) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: flowItems[i].x, y: 3.6, w: 1.8, h: 0.9,
      fill: { color: flowItems[i].color }, rectRadius: 0.1,
    });
    s.addText(flowItems[i].label, {
      x: flowItems[i].x, y: 3.6, w: 1.8, h: 0.9,
      fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 4,
    });
    if (i < flowItems.length - 1) {
      s.addText("→", {
        x: flowItems[i].x + 1.8, y: 3.6, w: 0.4, h: 0.9,
        fontSize: 22, fontFace: FONT_EN, color: C.primary, align: "center", valign: "middle", margin: 0,
      });
    }
  }

  addPageNum(s, "4");
})();

// ============================================================
// SLIDE 5: 歯科医療機関連携強化加算
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "歯科医療機関連携強化加算【新設】");

  // メインカード
  addCard(s, 0.4, 1.1, 5.8, 2.0);

  // 60点バッジ
  s.addShape(pres.shapes.OVAL, {
    x: 0.7, y: 1.3, w: 1.3, h: 1.3,
    fill: { color: C.red },
  });
  s.addText("60点", {
    x: 0.7, y: 1.45, w: 1.3, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("年1回", {
    x: 0.7, y: 2.0, w: 1.3, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.white, align: "center", valign: "middle", margin: 0,
  });

  var details = [
    "対象：糖尿病を主病とする患者",
    "要件：歯科受診の必要性を認め連携を実施",
    "目的：歯周病の予防・診断・治療",
    "管理料（I）・（II）の双方で算定可",
  ];
  for (var i = 0; i < details.length; i++) {
    s.addText(details[i], {
      x: 2.2, y: 1.3 + i * 0.42, w: 3.8, h: 0.35,
      fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  }

  // 右側：エビデンスカード
  addCard(s, 6.5, 1.1, 3.2, 2.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.7, y: 1.25, w: 2.4, h: 0.3, fill: { color: C.green }, rectRadius: 0.06,
  });
  s.addText("エビデンスの背景", {
    x: 6.7, y: 1.25, w: 2.4, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("歯周病は「第6の合併症」。歯周病治療によりHbA1cが改善するエビデンスあり。日本糖尿病学会も歯科連携を推奨。", {
    x: 6.7, y: 1.7, w: 2.8, h: 1.2,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 下部：双方向性の説明
  addCard(s, 0.4, 3.4, 9.2, 1.3);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.5, y: 3.6, w: 2.5, h: 0.9,
    fill: { color: C.primary }, rectRadius: 0.1,
  });
  s.addText("糖尿病", {
    x: 1.5, y: 3.6, w: 2.5, h: 0.9,
    fontSize: 20, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  s.addText("悪化 →\n← 改善", {
    x: 4.2, y: 3.6, w: 1.6, h: 0.9,
    fontSize: 16, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.0, y: 3.6, w: 2.5, h: 0.9,
    fill: { color: C.green }, rectRadius: 0.1,
  });
  s.addText("歯周病", {
    x: 6.0, y: 3.6, w: 2.5, h: 0.9,
    fontSize: 20, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "5");
})();

// ============================================================
// SLIDE 6: 充実管理加算の3段階再編
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "充実管理加算（旧：外来データ提出加算）の再編");

  // 3段階のカード
  var tiers = [
    { label: "加算1", pts: "30点", req: "体制整備＋十分な実績（上位20%）", color: C.red },
    { label: "加算2", pts: "20点", req: "体制整備＋相当の実績（中位30%）", color: C.orange },
    { label: "加算3", pts: "10点", req: "体制整備のみ（下位50%）", color: C.accent },
  ];

  for (var i = 0; i < tiers.length; i++) {
    var yBase = 1.2 + i * 1.15;
    addCard(s, 0.4, yBase, 9.2, 0.95);

    // ラベルバッジ
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y: yBase + 0.15, w: 1.6, h: 0.6,
      fill: { color: tiers[i].color }, rectRadius: 0.08,
    });
    s.addText(tiers[i].label, {
      x: 0.6, y: yBase + 0.1, w: 1.6, h: 0.4,
      fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(tiers[i].pts, {
      x: 0.6, y: yBase + 0.45, w: 1.6, h: 0.35,
      fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });

    // 要件テキスト
    s.addText(tiers[i].req, {
      x: 2.5, y: yBase + 0.15, w: 6.8, h: 0.6,
      fontSize: 16, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  }

  // 評価基準の説明
  addCard(s, 0.4, 4.7, 9.2, 0.7);
  s.addText("評価基準：検査実施患者割合 / 眼科・歯科連携加算の算定割合 / 受診間隔の適切性", {
    x: 0.6, y: 4.75, w: 8.8, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  s.addText("経過措置：R8年3月31日時点で継続データ提出 → R9年3月31日まで加算1相当で評価", {
    x: 0.6, y: 5.05, w: 8.8, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });

  addPageNum(s, "6");
})();

// ============================================================
// SLIDE 7: 在宅自己注射の適用拡大
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "在宅自己注射指導管理料の適用拡大");

  // 改定前
  addCard(s, 0.4, 1.2, 4.4, 2.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.35, w: 1.6, h: 0.35, fill: { color: C.sub }, rectRadius: 0.06,
  });
  s.addText("改定前", {
    x: 0.6, y: 1.35, w: 1.6, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("糖尿病に適応のある薬剤\nに限定して算定", {
    x: 0.6, y: 1.9, w: 4.0, h: 0.6,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  s.addText("例）インスリン、GLP-1受容体作動薬", {
    x: 0.6, y: 2.55, w: 4.0, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  // 矢印
  s.addText("→", {
    x: 4.5, y: 1.8, w: 1.0, h: 0.8,
    fontSize: 36, fontFace: FONT_EN, color: C.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 改定後
  addCard(s, 5.2, 1.2, 4.4, 2.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.35, w: 1.6, h: 0.35, fill: { color: C.red }, rectRadius: 0.06,
  });
  s.addText("改定後", {
    x: 5.4, y: 1.35, w: 1.6, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("糖尿病以外の併存疾患\nへの自己注射も算定可能", {
    x: 5.4, y: 1.9, w: 4.0, h: 0.6,
    fontSize: 15, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });
  s.addText("例）骨粗鬆症（テリパラチド）\n　　関節リウマチ（MTX皮下注）", {
    x: 5.4, y: 2.55, w: 4.0, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  // 下部ポイント
  addCard(s, 0.4, 3.5, 9.2, 1.6);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.65, w: 2.4, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("実務上のポイント", {
    x: 0.6, y: 3.65, w: 2.4, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var points = [
    "糖尿病患者は高齢化に伴い複数の併存疾患を抱えることが多い",
    "インスリン＋骨粗鬆症薬などの一体管理が可能に",
    "該当患者の洗い出しを行い、算定漏れを防ぐ",
  ];
  for (var i = 0; i < points.length; i++) {
    s.addText("● " + points[i], {
      x: 0.8, y: 4.15 + i * 0.3, w: 8.6, h: 0.28,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  }

  addPageNum(s, "7");
})();

// ============================================================
// SLIDE 8: 透析関連の改定
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "透析関連の改定 ― 人工腎臓と腎代替療法");

  // 人工腎臓の変更カード
  addCard(s, 0.4, 1.15, 4.4, 1.6);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.3, w: 3.0, h: 0.35, fill: { color: C.red }, rectRadius: 0.06,
  });
  s.addText("人工腎臓（J038）基本点数", {
    x: 0.6, y: 1.3, w: 3.0, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("全区分で", {
    x: 0.8, y: 1.85, w: 1.5, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  s.addText("▲20点", {
    x: 2.2, y: 1.75, w: 2.2, h: 0.55,
    fontSize: 32, fontFace: FONT_EN, color: C.red, bold: true, margin: 0,
  });

  // 充実加算カード
  addCard(s, 5.2, 1.15, 4.4, 1.6);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.3, w: 3.6, h: 0.35, fill: { color: C.green }, rectRadius: 0.06,
  });
  s.addText("腎代替療法診療体制充実加算【新設】", {
    x: 5.4, y: 1.3, w: 3.6, h: 0.35,
    fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("+20点", {
    x: 5.6, y: 1.75, w: 3.8, h: 0.55,
    fontSize: 32, fontFace: FONT_EN, color: C.green, bold: true, align: "center", margin: 0,
  });
  s.addText("取得で従来点数を維持可能", {
    x: 5.6, y: 2.3, w: 3.8, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  // 施設基準
  addCard(s, 0.4, 3.0, 9.2, 2.3);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.15, w: 2.4, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("施設基準（主な要件）", {
    x: 0.6, y: 3.15, w: 2.4, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var reqs = [
    { icon: "1", text: "災害対策：ハザードマップに基づくリスク評価・マニュアル策定・年1回以上の訓練" },
    { icon: "2", text: "患者説明：腎代替療法の選択肢（HD・PD・腎移植）について繰り返し説明" },
    { icon: "3", text: "実績要件（いずれか）：腹膜灌流指導管理24回以上/年 or 腎移植2名以上/年" },
  ];
  for (var i = 0; i < reqs.length; i++) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.7, y: 3.7 + i * 0.5, w: 0.35, h: 0.35,
      fill: { color: C.accent }, rectRadius: 0.06,
    });
    s.addText(reqs[i].icon, {
      x: 0.7, y: 3.7 + i * 0.5, w: 0.35, h: 0.35,
      fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(reqs[i].text, {
      x: 1.2, y: 3.7 + i * 0.5, w: 8.2, h: 0.35,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  }

  addPageNum(s, "8");
})();

// ============================================================
// SLIDE 9: 薬価改定＋ガイドライン更新
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "薬価改定とガイドライン更新");

  // 薬価カード
  addCard(s, 0.4, 1.1, 4.4, 1.9);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.2, w: 3.5, h: 0.32, fill: { color: C.red }, rectRadius: 0.06,
  });
  s.addText("薬価改定（R8年4月施行）", {
    x: 0.6, y: 1.2, w: 3.5, h: 0.32,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("全体：▲0.86%", {
    x: 0.8, y: 1.65, w: 3.8, h: 0.3,
    fontSize: 15, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });
  s.addText("リベルサス錠（セマグルチド）", {
    x: 0.8, y: 2.0, w: 3.8, h: 0.25,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  s.addText("改定時加算 +5%", {
    x: 0.8, y: 2.25, w: 3.8, h: 0.4,
    fontSize: 22, fontFace: FONT_EN, color: C.red, bold: true, margin: 0,
  });
  s.addText("「真の臨床的有用性の検証に係る加算」", {
    x: 0.8, y: 2.65, w: 3.8, h: 0.25,
    fontSize: 11, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  // ガイドラインカード
  addCard(s, 5.2, 1.1, 4.4, 1.9);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.2, w: 3.5, h: 0.32, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("標準診療マニュアル2026年版", {
    x: 5.4, y: 1.2, w: 3.5, h: 0.32,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("GLP-1 RA（経口）推奨強化", {
    x: 5.6, y: 1.65, w: 3.8, h: 0.28,
    fontSize: 13, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });
  s.addText("心血管既往・肥満 →\n積極的に投与開始してよい", {
    x: 5.6, y: 1.95, w: 3.8, h: 0.45,
    fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true, margin: 0,
  });
  s.addText("DPP-4阻害薬との併用禁忌", {
    x: 5.6, y: 2.45, w: 3.8, h: 0.25,
    fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });
  s.addText("GLP-1 RA・GIP/GLP-1 RAとの併用は避ける", {
    x: 5.6, y: 2.7, w: 3.8, h: 0.25,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 下部注記
  addCard(s, 0.4, 3.25, 9.2, 1.95);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.4, w: 1.4, h: 0.3, fill: { color: C.orange }, rectRadius: 0.06,
  });
  s.addText("注意", {
    x: 0.6, y: 3.4, w: 1.4, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText([
    { text: "DPP-4阻害薬 → GLP-1 RA / GIP/GLP-1 RA 切替時は\n", options: { fontSize: 13, bold: true, color: C.text } },
    { text: "DPP-4阻害薬の中止を確実に行うこと\n", options: { fontSize: 13, bold: true, color: C.red } },
    { text: "作用機序の重複 → 低血糖リスク増大、上乗せ効果なし\n", options: { fontSize: 12, color: C.sub } },
    { text: "チルゼパチド（マンジャロ）等のGIP/GLP-1 RAも同様", options: { fontSize: 12, color: C.sub } },
  ], {
    x: 0.6, y: 3.8, w: 8.8, h: 1.2,
    fontFace: FONT_JP, margin: 0,
  });

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
    x: 0.5, y: 0.2, w: 9, h: 0.55,
    fontSize: 24, fontFace: FONT_EN, color: C.white, bold: true, margin: 0,
  });

  var messages = [
    { num: "1", text: "眼科・歯科連携加算（各60点/年1回）― 合併症予防と算定の両立", color: C.accent },
    { num: "2", text: "生活習慣病管理料：併算定可能項目 16→37に拡大", color: C.green },
    { num: "3", text: "充実管理加算 3段階評価 ― 加算1（30点）は実績が必要", color: C.orange },
    { num: "4", text: "在宅自己注射：併存疾患（骨粗鬆症・RA等）にも算定拡大", color: C.accent },
    { num: "5", text: "透析：基本▲20点 → 腎代替療法充実加算+20点で相殺可", color: C.yellow },
    { num: "6", text: "GLP-1 RA：心血管既往・肥満で積極投与。DPP-4併用禁忌", color: C.red },
  ];

  for (var i = 0; i < messages.length; i++) {
    var yBase = 0.95 + i * 0.65;

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: yBase, w: 0.45, h: 0.45,
      fill: { color: messages[i].color }, rectRadius: 0.08,
    });
    s.addText(messages[i].num, {
      x: 0.5, y: yBase, w: 0.45, h: 0.45,
      fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(messages[i].text, {
      x: 1.15, y: yBase, w: 8.4, h: 0.45,
      fontSize: 14, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0,
    });
  }

  s.addShape(pres.shapes.LINE, { x: 1, y: 4.95, w: 8, h: 0, line: { color: C.accent, width: 1 } });

  s.addText("ご視聴ありがとうございました", {
    x: 0.5, y: 5.05, w: 9, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
})();

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/令和8年度診療報酬改定_糖尿病内分泌.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX saved: " + outPath);
}).catch(function(err) {
  console.error("Error: " + err);
});
