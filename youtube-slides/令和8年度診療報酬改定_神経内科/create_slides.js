var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "令和8年度診療報酬改定【神経内科編】";

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

  s.addText("【神経内科編】", {
    x: 0.8, y: 1.9, w: 8.4, h: 0.7,
    fontSize: 28, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", margin: 0,
  });

  s.addText("認知症・リハビリ・遠隔診療・難病の主要変更点", {
    x: 0.8, y: 2.7, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.45, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("厚生労働省公式資料・疑義解釈に基づく", {
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
  addHeader(s, "神経内科領域の主要改定ポイント");

  var items = [
    { icon: "1", title: "認知症地域包括診療の統合・再編", desc: "認知症患者区分で高い点数設定、要介護者は1疾患でOK" },
    { icon: "2", title: "早期リハ加算の大幅見直し", desc: "入院3日以内 25点→60点、算定期間30日→14日に短縮" },
    { icon: "3", title: "休日リハビリ加算【新設】", desc: "1単位25点 ― 脳卒中365日リハ体制を後押し" },
    { icon: "4", title: "遠隔連携診療料 ― 指定難病に拡大", desc: "全指定難病900点/月 ― 神経難病の遠隔医療" },
    { icon: "5", title: "回復期リハ病棟の基準見直し", desc: "実績指数40→42、FIM認知除外14点以下に厳格化" },
    { icon: "6", title: "離床なきリハの減算・AI活用ほか", desc: "ベッド上リハ10%減算、生成AI 1.2倍換算" },
  ];

  for (var i = 0; i < items.length; i++) {
    var yBase = 1.1 + i * 0.7;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: yBase, w: 0.5, h: 0.5,
      fill: { color: C.primary }, rectRadius: 0.08,
    });
    s.addText(items[i].icon, {
      x: 0.5, y: yBase, w: 0.5, h: 0.5,
      fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(items[i].title, {
      x: 1.2, y: yBase - 0.02, w: 5.5, h: 0.3,
      fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
    });
    s.addText(items[i].desc, {
      x: 1.2, y: yBase + 0.28, w: 8, h: 0.25,
      fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
    });
  }

  addPageNum(s, "2");
})();

// ============================================================
// SLIDE 3: 認知症地域包括診療の統合
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "認知症地域包括診療の統合・再編");

  // 説明テキスト
  s.addText("「認知症地域包括診療加算/料」が通常の「地域包括診療加算/料」に統合", {
    x: 0.5, y: 1.1, w: 9, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // カード1: 加算
  addCard(s, 0.4, 1.6, 4.4, 2.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.75, w: 2.8, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("地域包括診療加算（再診時）", {
    x: 0.6, y: 1.75, w: 2.8, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var addRows = [
    ["", "認知症患者等", "その他"],
    ["加算1", "38点", "28点"],
    ["加算2", "31点", "21点"],
  ];
  s.addTable(addRows, {
    x: 0.6, y: 2.25, w: 4.0,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: [4,6,4,6],
    border: { pt: 0.5, color: "DEE2E6" },
    colW: [1.2, 1.4, 1.4],
    rowH: [0.35, 0.35, 0.35],
    autoPage: false,
  });

  // カード2: 診療料
  addCard(s, 5.2, 1.6, 4.4, 2.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.75, w: 2.8, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("地域包括診療料（月1回）", {
    x: 5.4, y: 1.75, w: 2.8, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var feeRows = [
    ["", "認知症患者等", "その他"],
    ["診療料1", "1,682点", "1,661点"],
    ["診療料2", "1,614点", "1,601点"],
  ];
  s.addTable(feeRows, {
    x: 5.4, y: 2.25, w: 4.0,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: [4,6,4,6],
    border: { pt: 0.5, color: "DEE2E6" },
    colW: [1.2, 1.4, 1.4],
    rowH: [0.35, 0.35, 0.35],
    autoPage: false,
  });

  // 重要ポイント
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.85, w: 9.2, h: 1.6, fill: { color: C.light }, rectRadius: 0.08,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.85, w: 0.08, h: 1.6, fill: { color: C.primary },
  });
  s.addText("実務上の重要ポイント", {
    x: 0.7, y: 3.9, w: 8.5, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText(
    "・要介護者は1疾患でも算定可能に（従来2疾患以上）\n" +
    "・認知症の診断後支援 ― 地域包括支援センターとの連携を明記\n" +
    "・外来データ提出加算（月1回10点）が新設",
    {
      x: 0.7, y: 4.25, w: 8.5, h: 1.1,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  addPageNum(s, "3");
})();

// ============================================================
// SLIDE 4: 早期リハ加算の見直し
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "早期リハビリテーション加算の大幅見直し");

  // テーブル: 改定前後
  var rows = [
    [{ text: "算定時期", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "改定前", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "改定後", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "変化", options: { bold: true, color: C.white, fill: { color: C.primary } } }],
    ["入院3日以内", "25点", { text: "60点", options: { bold: true, color: C.red } }, { text: "+35点（2.4倍）", options: { bold: true, color: C.red } }],
    ["4日目～14日目", "25点", "25点", "据置"],
    ["算定可能期間", "入院から30日目", { text: "入院から14日目", options: { bold: true, color: C.red } }, { text: "16日短縮", options: { color: C.orange } }],
  ];

  s.addTable(rows, {
    x: 0.5, y: 1.2, w: 9.0,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: [6,8,6,8],
    border: { pt: 0.5, color: "DEE2E6" },
    colW: [2.5, 1.8, 2.2, 2.5],
    rowH: [0.45, 0.45, 0.45, 0.45],
    autoPage: false,
  });

  // 注意点ボックス
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.3, w: 9.2, h: 2.0, fill: { color: "FFF3CD" }, rectRadius: 0.08,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.3, w: 0.08, h: 2.0, fill: { color: C.yellow },
  });
  s.addText("注意点", {
    x: 0.7, y: 3.35, w: 8.5, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });
  s.addText(
    "・転院患者は転院前の入院日が起算日\n" +
    "・外来患者（退院後）は退院前の入院日が起算日\n" +
    "・「超急性期72時間に集中的にリハを入れる」\n" +
    "  インセンティブの明確化",
    {
      x: 0.7, y: 3.7, w: 8.5, h: 1.4,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  addPageNum(s, "4");
})();

// ============================================================
// SLIDE 5: 休日リハ加算 + 遠隔連携診療料
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "休日リハ加算【新設】＋ 遠隔連携診療料の拡大");

  // 左カード: 休日リハ
  addCard(s, 0.4, 1.1, 4.4, 2.2);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.25, w: 3.2, h: 0.35, fill: { color: C.green }, rectRadius: 0.06,
  });
  s.addText("休日リハビリテーション加算", {
    x: 0.6, y: 1.25, w: 3.2, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・休日にリハビリ実施で 1単位25点 加算\n" +
    "・発症/手術から7日目 or\n" +
    "  治療開始から30日目まで\n" +
    "・脳卒中365日リハ体制を評価",
    {
      x: 0.6, y: 1.8, w: 4, h: 1.3,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // 右カード: 遠隔連携診療料
  addCard(s, 5.2, 1.1, 4.4, 2.2);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.25, w: 3.2, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("遠隔連携診療料の対象拡大", {
    x: 5.4, y: 1.25, w: 3.2, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・全指定難病が対象に追加\n" +
    "・外来/訪問/入院 各900点/月\n" +
    "・医療受給者証の有無を問わず\n" +
    "・軽症者も対象",
    {
      x: 5.4, y: 1.8, w: 4, h: 1.3,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // 下部: 対象神経難病リスト
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.6, w: 9.2, h: 1.8, fill: { color: C.light }, rectRadius: 0.08,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.6, w: 0.08, h: 1.8, fill: { color: C.primary },
  });
  s.addText("遠隔連携診療の対象となる主な神経難病", {
    x: 0.7, y: 3.65, w: 8.5, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText(
    "ALS（筋萎縮性側索硬化症）  /  パーキンソン病  /  多発性硬化症\n" +
    "脊髄小脳変性症  /  重症筋無力症  /  多系統萎縮症\n" +
    "進行性核上性麻痺  /  ハンチントン病  /  CIDP  ほか全指定難病",
    {
      x: 0.7, y: 4.0, w: 8.5, h: 1.2,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
    }
  );

  addPageNum(s, "5");
})();

// ============================================================
// SLIDE 6: D to P with N
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "D to P with N ― 看護師同席のオンライン診療");

  // 概念図
  addCard(s, 0.4, 1.1, 9.2, 1.5);
  // 左: 専門医
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 1.3, w: 2.2, h: 1.0, fill: { color: C.primary }, rectRadius: 0.1,
  });
  s.addText("専門医\n（遠隔）", {
    x: 0.8, y: 1.3, w: 2.2, h: 1.0,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 矢印
  s.addText("ICT接続", {
    x: 3.3, y: 1.5, w: 1.5, h: 0.3,
    fontSize: 11, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3.2, y: 1.9, w: 1.7, h: 0, line: { color: C.accent, width: 2, dashType: "dash" } });

  // 中: 看護師
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.1, y: 1.3, w: 2.0, h: 1.0, fill: { color: C.green }, rectRadius: 0.1,
  });
  s.addText("看護師\n（患者側）", {
    x: 5.1, y: 1.3, w: 2.0, h: 1.0,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 矢印
  s.addShape(pres.shapes.LINE, { x: 7.3, y: 1.9, w: 0.6, h: 0, line: { color: C.sub, width: 2 } });

  // 右: 患者
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 8.0, y: 1.3, w: 1.4, h: 1.0, fill: { color: C.accent }, rectRadius: 0.1,
  });
  s.addText("患者", {
    x: 8.0, y: 1.3, w: 1.4, h: 1.0,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 算定可能な項目カード
  addCard(s, 0.4, 2.9, 4.4, 2.4);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.05, w: 3.5, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("看護師が実施可能な項目", {
    x: 0.6, y: 3.05, w: 3.5, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・検査（採血等） → 検査実施料を算定\n" +
    "・注射 → 注射実施料を算定\n" +
    "・処置 → 処置実施料を算定",
    {
      x: 0.6, y: 3.6, w: 4, h: 1.2,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
    }
  );

  // 活用シナリオ
  addCard(s, 5.2, 2.9, 4.4, 2.4);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 3.05, w: 3.5, h: 0.35, fill: { color: C.green }, rectRadius: 0.06,
  });
  s.addText("神経内科での活用シナリオ", {
    x: 5.4, y: 3.05, w: 3.5, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・ALS在宅患者の定期フォロー\n" +
    "・パーキンソン病の服薬調整\n" +
    "・遠隔地の神経難病患者に\n" +
    "  訪問看護と専門医を橋渡し",
    {
      x: 5.4, y: 3.6, w: 4, h: 1.4,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  addPageNum(s, "6");
})();

// ============================================================
// SLIDE 7: 回復期リハ病棟 入院料・実績指数
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "回復期リハ病棟の基準見直し ①入院料・実績指数");

  // 左テーブル: 入院料
  var feeRows = [
    [{ text: "区分", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "改定前", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "改定後", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "増点", options: { bold: true, color: C.white, fill: { color: C.primary } } }],
    ["入院料1", "2,229", { text: "2,346", options: { bold: true } }, "+117"],
    ["入院料2", "2,166", { text: "2,274", options: { bold: true } }, "+108"],
    ["入院料3", "1,917", { text: "2,062", options: { bold: true } }, "+145"],
    ["入院料4", "1,859", { text: "2,000", options: { bold: true } }, "+141"],
    ["入院料5", "1,696", { text: "1,794", options: { bold: true } }, "+98"],
  ];
  s.addTable(feeRows, {
    x: 0.3, y: 1.1, w: 5.0,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: [4,6,4,6],
    border: { pt: 0.5, color: "DEE2E6" },
    colW: [1.2, 1.2, 1.2, 1.0],
    rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
    autoPage: false,
  });

  // 右テーブル: 実績指数
  var indexRows = [
    [{ text: "区分", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "改定前", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "改定後", options: { bold: true, color: C.white, fill: { color: C.primary } } }],
    ["入院料1", "40以上", { text: "42以上", options: { bold: true, color: C.red } }],
    ["入院料2", "未設定", { text: "32以上", options: { bold: true, color: C.red } }],
    ["入院料3", "35以上", { text: "37以上", options: { bold: true, color: C.red } }],
    ["入院料4", "未設定", { text: "32以上", options: { bold: true, color: C.red } }],
  ];
  s.addTable(indexRows, {
    x: 5.5, y: 1.1, w: 4.2,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: [4,6,4,6],
    border: { pt: 0.5, color: "DEE2E6" },
    colW: [1.2, 1.4, 1.4],
    rowH: [0.35, 0.35, 0.35, 0.35, 0.35],
    autoPage: false,
  });

  // 強化体制加算
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.6, w: 9.2, h: 1.7, fill: { color: C.light }, rectRadius: 0.08,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.6, w: 0.08, h: 1.7, fill: { color: C.primary },
  });
  s.addText("回復期リハビリテーション強化体制加算【新設】", {
    x: 0.7, y: 3.65, w: 8.5, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText(
    "・1日につき 80点（入院料1対象）\n" +
    "・実績指数 48以上 が要件\n" +
    "・退院前訪問指導の実績＋排尿自立支援加算の届出が必要",
    {
      x: 0.7, y: 4.1, w: 8.5, h: 1.0,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  addPageNum(s, "7");
})();

// ============================================================
// SLIDE 8: 回復期リハ病棟 重症患者定義
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "回復期リハ病棟の基準見直し ②重症患者・除外基準");

  // 緩和された基準カード
  addCard(s, 0.4, 1.1, 4.4, 2.1);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.25, w: 2.4, h: 0.35, fill: { color: C.green }, rectRadius: 0.06,
  });
  s.addText("緩和された基準", {
    x: 0.6, y: 1.25, w: 2.4, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・重症患者割合\n" +
    "  入院料1・2: 4割 → 3.5割\n" +
    "  入院料3・4: 3割 → 2.5割\n" +
    "・「退院時3割改善」要件を 削除",
    {
      x: 0.6, y: 1.75, w: 4, h: 1.2,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // 厳格化された基準カード
  addCard(s, 5.2, 1.1, 4.4, 2.1);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.25, w: 2.4, h: 0.35, fill: { color: C.red }, rectRadius: 0.06,
  });
  s.addText("厳格化された基準", {
    x: 5.4, y: 1.25, w: 2.4, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・80歳以上の除外規定 → 廃止\n" +
    "・FIM認知除外: 24点以下→14点以下\n" +
    "・除外割合上限: 3割→2割\n" +
    "・実績指数下限: 27→30",
    {
      x: 5.4, y: 1.75, w: 4, h: 1.2,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // 重症患者の新定義
  addCard(s, 0.4, 3.5, 9.2, 1.9);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.65, w: 2.6, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("重症患者の新定義", {
    x: 0.6, y: 3.65, w: 2.6, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "以下のいずれか：\n" +
    "① 日常生活機能評価10点以上 または FIM総得点21点以上55点以下\n" +
    "② 高次脳機能障害と診断された患者\n" +
    "③ 脊髄損傷と診断された患者",
    {
      x: 0.6, y: 4.15, w: 8.8, h: 1.0,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  addPageNum(s, "8");
})();

// ============================================================
// SLIDE 9: 離床なきリハ減算 + 特定疾患
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "離床なきリハの減算 ＋ 特定疾患療養管理料");

  // 左カード: 離床なきリハ
  addCard(s, 0.4, 1.1, 4.4, 2.3);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.25, w: 3.2, h: 0.35, fill: { color: C.red }, rectRadius: 0.06,
  });
  s.addText("離床なきリハビリの減算【新設】", {
    x: 0.6, y: 1.25, w: 3.2, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・所定点数の 90%で算定\n" +
    " （10%減算）\n" +
    "・1日2単位まで に制限\n" +
    "・医学的理由がある場合は除外可能\n" +
    "  → カルテ記載が重要",
    {
      x: 0.6, y: 1.8, w: 4, h: 1.4,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // 右カード: 特定疾患
  addCard(s, 5.2, 1.1, 4.4, 2.3);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.25, w: 3.2, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("特定疾患療養管理料の変更", {
    x: 5.4, y: 1.25, w: 3.2, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・ギラン・バレー症候群は\n" +
    "  引き続き対象\n" +
    "・長期処方/リフィル処方箋の\n" +
    "  掲示義務を追加\n" +
    "・胃十二指腸潰瘍にNSAIDs除外",
    {
      x: 5.4, y: 1.8, w: 4, h: 1.4,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // 影響を受ける患者
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.7, w: 9.2, h: 1.6, fill: { color: "FFF3CD" }, rectRadius: 0.08,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.7, w: 0.08, h: 1.6, fill: { color: C.yellow },
  });
  s.addText("離床なきリハ減算 ― 影響を受ける神経疾患患者", {
    x: 0.7, y: 3.75, w: 8.5, h: 0.3,
    fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });
  s.addText(
    "・重症脳卒中（意識障害・高度片麻痺）、ALS・進行性筋疾患\n" +
    "・ただし「医学的理由がある場合」は除外\n" +
    "  → 診療録に理由を記載し、離床困難な医学的根拠を明確にする",
    {
      x: 0.7, y: 4.15, w: 8.5, h: 1.0,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  addPageNum(s, "9");
})();

// ============================================================
// SLIDE 10: ベースアップ + AI
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ベースアップ評価料の拡充 ＋ AI・ICT活用");

  // ベースアップ評価料テーブル
  var bRows = [
    [{ text: "場面", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "R6参考", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "R8（新規）", options: { bold: true, color: C.white, fill: { color: C.primary } } },
     { text: "R8（継続）", options: { bold: true, color: C.white, fill: { color: C.primary } } }],
    ["初診時", "6点", { text: "17点", options: { bold: true } }, { text: "23点", options: { bold: true, color: C.red } }],
    ["再診時", "2点", "4点", { text: "6点", options: { bold: true } }],
    ["訪問診療", "28点", "79点", { text: "107点", options: { bold: true, color: C.red } }],
  ];
  s.addTable(bRows, {
    x: 0.3, y: 1.1, w: 5.4,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: [4,6,4,6],
    border: { pt: 0.5, color: "DEE2E6" },
    colW: [1.3, 1.1, 1.5, 1.5],
    rowH: [0.35, 0.35, 0.35, 0.35],
    autoPage: false,
  });

  // AI活用カード
  addCard(s, 5.9, 1.1, 3.8, 2.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.1, y: 1.25, w: 2.6, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("生成AI活用の評価", {
    x: 6.1, y: 1.25, w: 2.6, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "医師事務作業補助者\n" +
    "・AI文書作成 → 1.2倍換算\n" +
    "・AI＋音声入力 → 1.3倍換算",
    {
      x: 6.1, y: 1.75, w: 3.4, h: 1.0,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // R9年度の予告
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.4, w: 9.2, h: 0.6, fill: { color: C.light }, rectRadius: 0.08,
  });
  s.addText("令和9年度にはベースアップ評価料が さらに200%（2倍）に引き上げ予定", {
    x: 0.7, y: 3.45, w: 8.5, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  // 電子的診療情報連携
  addCard(s, 0.4, 4.2, 9.2, 1.2);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 4.35, w: 4.0, h: 0.35, fill: { color: C.primary }, rectRadius: 0.06,
  });
  s.addText("電子的診療情報連携体制整備加算【新設】", {
    x: 0.6, y: 4.35, w: 4.0, h: 0.35,
    fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("初診: 加算1 15点 / 加算2 9点 / 加算3 4点   再診: 2点", {
    x: 0.6, y: 4.85, w: 8.8, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  addPageNum(s, "10");
})();

// ============================================================
// SLIDE 11: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "まとめ ― 神経内科へのインパクト");

  // プラスの影響
  addCard(s, 0.4, 1.1, 4.4, 3.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.25, w: 2.0, h: 0.35, fill: { color: C.green }, rectRadius: 0.06,
  });
  s.addText("プラスの影響", {
    x: 0.6, y: 1.25, w: 2.0, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・認知症外来の包括評価が充実\n" +
    "・脳卒中超急性期リハの評価強化\n" +
    "  （3日以内 60点）\n" +
    "・指定難病の遠隔診療が全面可能に\n" +
    "・回復期リハ入院料の増点\n" +
    "・休日リハ加算で365日体制\n" +
    "・ベースアップ評価料の増収",
    {
      x: 0.6, y: 1.75, w: 4, h: 2.2,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // 注意が必要な変更
  addCard(s, 5.2, 1.1, 4.4, 3.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.4, y: 1.25, w: 2.4, h: 0.35, fill: { color: C.red }, rectRadius: 0.06,
  });
  s.addText("注意が必要な変更", {
    x: 5.4, y: 1.25, w: 2.4, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText(
    "・早期リハ加算の期間短縮\n" +
    "  （30日→14日）\n" +
    "・回復期リハ実績指数の引き上げ\n" +
    "・FIM認知除外基準の厳格化\n" +
    "・離床困難患者のリハ減算\n" +
    "  （医学的理由の記載要）\n" +
    "・リフィル処方箋の掲示義務",
    {
      x: 5.4, y: 1.75, w: 4, h: 2.2,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // 3軸メッセージ
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 4.35, w: 9.2, h: 1.0, fill: { color: C.dark }, rectRadius: 0.08,
  });
  s.addText("3つの軸: 「認知症の包括管理」「急性期リハの集中化」「神経難病の遠隔医療」", {
    x: 0.7, y: 4.4, w: 8.5, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", margin: 0,
  });
  s.addText("令和8年6月1日施行 ― 各施設での体制整備と算定準備を", {
    x: 0.7, y: 4.85, w: 8.5, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  addPageNum(s, "11");
})();

// ============================================================
// SLIDE 12: エンディング
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("ご視聴ありがとうございました", {
    x: 0.8, y: 1.0, w: 8.4, h: 0.8,
    fontSize: 30, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 2.1, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("チャンネル登録・高評価をお願いします", {
    x: 0.8, y: 2.5, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText(
    "参考文献\n" +
    "1. 厚生労働省「令和8年度診療報酬改定について」\n" +
    "2. 厚生労働省「令和8年度診療報酬改定説明資料」\n" +
    "3. 厚生労働省 疑義解釈資料（その1・その2）\n" +
    "4. GemMed 2026年度診療報酬改定答申シリーズ",
    {
      x: 1.5, y: 3.2, w: 7, h: 2.0,
      fontSize: 12, fontFace: FONT_JP, color: C.light, margin: 0, lineSpacingMultiple: 1.4,
    }
  );
})();

// ============================================================
// 出力
// ============================================================
var outPath = __dirname + "/令和8年度診療報酬改定_神経内科.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
