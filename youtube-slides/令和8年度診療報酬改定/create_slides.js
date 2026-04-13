var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "令和8年度診療報酬改定 ― +3.09%改定の全体像と医科の主要変更点";

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
    x: 0.8, y: 1.2, w: 8.4, h: 1.0,
    fontSize: 38, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addText("+3.09%改定の全体像と医科の主要変更点", {
    x: 0.8, y: 2.3, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.15, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText([
    { text: "厚生労働省保険局医療課  令和8年3月10日版資料に基づく", options: { color: C.light, fontSize: 13, italic: true } },
  ], { x: 0.8, y: 3.4, w: 8.4, h: 0.4, fontFace: FONT_JP, align: "center", margin: 0 });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 改定率の全体像
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "改定率の全体像 ― +3.09%の内訳");

  // メインの数字カード
  addCard(s, 0.4, 1.1, 3.0, 1.6);
  s.addText("+3.09%", {
    x: 0.5, y: 1.15, w: 2.8, h: 0.8,
    fontSize: 42, fontFace: FONT_EN, color: C.primary, bold: true, align: "center", margin: 0,
  });
  s.addText("診療報酬本体\n（2年度平均）", {
    x: 0.5, y: 1.95, w: 2.8, h: 0.65,
    fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.1,
  });

  // 年度別
  addCard(s, 0.4, 2.9, 1.4, 0.9);
  s.addText("R8年度", { x: 0.45, y: 2.92, w: 1.3, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0 });
  s.addText("+2.41%", { x: 0.45, y: 3.25, w: 1.3, h: 0.4, fontSize: 20, fontFace: FONT_EN, color: C.primary, bold: true, align: "center", margin: 0 });

  addCard(s, 2.0, 2.9, 1.4, 0.9);
  s.addText("R9年度", { x: 2.05, y: 2.92, w: 1.3, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0 });
  s.addText("+3.77%", { x: 2.05, y: 3.25, w: 1.3, h: 0.4, fontSize: 20, fontFace: FONT_EN, color: C.primary, bold: true, align: "center", margin: 0 });

  // 内訳リスト
  addCard(s, 3.7, 1.1, 6.0, 2.7);
  var items = [
    { label: "賃上げ分", val: "+1.70%", color: C.green },
    { label: "物価対応分", val: "+0.76%", color: C.primary },
    { label: "食費・光熱水費", val: "+0.09%", color: C.accent },
    { label: "R6改定以降 経営環境悪化対応", val: "+0.44%", color: C.orange },
    { label: "後発医薬品等 効率化", val: "▲0.15%", color: C.red },
    { label: "その他", val: "+0.25%", color: C.sub },
  ];
  for (var i = 0; i < items.length; i++) {
    var yy = 1.2 + i * 0.42;
    s.addShape(pres.shapes.RECTANGLE, { x: 3.9, y: yy, w: 0.15, h: 0.3, fill: { color: items[i].color } });
    s.addText(items[i].label, { x: 4.2, y: yy - 0.02, w: 3.5, h: 0.34, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
    s.addText(items[i].val, { x: 7.8, y: yy - 0.02, w: 1.7, h: 0.34, fontSize: 16, fontFace: FONT_EN, color: items[i].color, bold: true, align: "right", margin: 0 });
  }

  // 薬価等
  addCard(s, 0.4, 4.0, 9.3, 0.8);
  s.addText([
    { text: "薬価等 ", options: { fontSize: 14, bold: true, color: C.text } },
    { text: "薬価 ▲0.86%（R8年4月施行）+ 材料 ▲0.01% = 合計 ", options: { fontSize: 13, color: C.sub } },
    { text: "▲0.87%", options: { fontSize: 16, bold: true, color: C.red } },
    { text: "　　各科改定率 ", options: { fontSize: 13, color: C.sub } },
    { text: "医科+0.28%  歯科+0.31%  調剤+0.08%", options: { fontSize: 13, color: C.primary, bold: true } },
  ], { x: 0.7, y: 4.1, w: 8.8, h: 0.55, fontFace: FONT_JP, margin: 0 });

  addPageNum(s, "2");
})();

// ============================================================
// SLIDE 3: 4つの基本方針
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "改定の4つの基本方針");

  var cards = [
    { num: "1", title: "物価・賃金・人手不足\nへの対応", items: "医療従事者の処遇改善\nICT・AI・IoT活用\nタスク・シェアリング推進", color: C.green },
    { num: "2", title: "2040年を見据えた\n機能分化・連携", items: "「治す」と「治し、支える」\nかかりつけ医機能の評価\n外来の機能分化と連携", color: C.primary },
    { num: "3", title: "安心・安全で質の高い\n医療の推進", items: "アウトカム評価の推進\n医療DX・ICT連携\n口腔管理のデジタル化", color: C.orange },
    { num: "4", title: "効率化・適正化", items: "後発品の使用促進\n費用対効果評価の活用\n長期処方・リフィル推進", color: C.red },
  ];

  for (var i = 0; i < cards.length; i++) {
    var cx = 0.35 + i * 2.4;
    addCard(s, cx, 1.15, 2.2, 3.9);
    // 番号バッジ
    s.addShape(pres.shapes.OVAL, { x: cx + 0.8, y: 1.25, w: 0.6, h: 0.6, fill: { color: cards[i].color } });
    s.addText(cards[i].num, { x: cx + 0.8, y: 1.28, w: 0.6, h: 0.55, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0 });
    // タイトル
    s.addText(cards[i].title, { x: cx + 0.1, y: 1.95, w: 2.0, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, align: "center", margin: 0, lineSpacingMultiple: 1.2 });
    // 区切り線
    s.addShape(pres.shapes.LINE, { x: cx + 0.3, y: 2.9, w: 1.6, h: 0, line: { color: cards[i].color, width: 2 } });
    // 内容
    s.addText(cards[i].items, { x: cx + 0.15, y: 3.0, w: 1.9, h: 1.9, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "left", margin: 4, lineSpacingMultiple: 1.4, valign: "top" });
  }

  addPageNum(s, "3");
})();

// ============================================================
// SLIDE 4: 賃上げ対応の全体像
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "賃上げ対応 ― ベースアップ評価料の大幅見直し");

  // 目標値カード
  addCard(s, 0.4, 1.15, 4.5, 1.5);
  s.addText("賃上げ目標", { x: 0.5, y: 1.2, w: 2.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0 });
  s.addText([
    { text: "R8年度 ", options: { fontSize: 13, color: C.text } },
    { text: "+3.2%", options: { fontSize: 28, bold: true, color: C.green } },
    { text: "   R9年度 ", options: { fontSize: 13, color: C.text } },
    { text: "+3.2%", options: { fontSize: 28, bold: true, color: C.green } },
  ], { x: 0.5, y: 1.55, w: 4.3, h: 0.5, fontFace: FONT_EN, margin: 0 });
  s.addText("看護補助者・事務職員は +5.7%", { x: 0.5, y: 2.1, w: 4.3, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0 });

  // 対象拡大カード
  addCard(s, 5.2, 1.15, 4.5, 1.5);
  s.addText("対象の拡大", { x: 5.3, y: 1.2, w: 2.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0 });
  s.addText([
    { text: "事務職員 ", options: { fontSize: 16, bold: true, color: C.primary } },
    { text: "を新たに対象に追加", options: { fontSize: 14, color: C.text } },
  ], { x: 5.3, y: 1.6, w: 4.3, h: 0.3, fontFace: FONT_JP, margin: 0 });
  s.addText([
    { text: "保険薬局 ", options: { fontSize: 16, bold: true, color: C.primary } },
    { text: "も対象施設に追加", options: { fontSize: 14, color: C.text } },
  ], { x: 5.3, y: 1.95, w: 4.3, h: 0.3, fontFace: FONT_JP, margin: 0 });
  s.addText("40歳以上の医師・歯科医師は除く", { x: 5.3, y: 2.3, w: 4.3, h: 0.25, fontSize: 11, fontFace: FONT_JP, color: C.sub, margin: 0 });

  // 評価体系の変更ポイント
  addCard(s, 0.4, 2.85, 9.3, 2.2);
  s.addText("主な変更ポイント", { x: 0.6, y: 2.92, w: 3.0, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var points = [
    { icon: "1", text: "R8年度とR9年度で段階的評価（R9年度は点数が倍増する設計）" },
    { icon: "2", text: "継続的に賃上げを実施している施設と新規施設で異なる評価" },
    { icon: "3", text: "入院BU評価料は165区分→250区分に拡大（R9年6月は500区分）" },
    { icon: "4", text: "賃上げ未実施の医療機関には入院基本料の減算規定を新設" },
  ];
  for (var i = 0; i < points.length; i++) {
    var py = 3.35 + i * 0.42;
    s.addShape(pres.shapes.OVAL, { x: 0.65, y: py, w: 0.3, h: 0.3, fill: { color: C.primary } });
    s.addText(points[i].icon, { x: 0.65, y: py, w: 0.3, h: 0.3, fontSize: 12, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(points[i].text, { x: 1.1, y: py - 0.02, w: 8.4, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  }

  addPageNum(s, "4");
})();

// ============================================================
// SLIDE 5: BU評価料の具体的な点数変更
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "外来・在宅BU評価料（I） ― 具体的な点数変更");

  // テーブル
  var tableRows = [
    [
      { text: "区分", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
      { text: "現行", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
      { text: "R8 新規", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
      { text: "R8 継続", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
      { text: "R9 新規", options: { bold: true, color: "FFFFFF", fill: { color: C.green } } },
      { text: "R9 継続", options: { bold: true, color: "FFFFFF", fill: { color: C.green } } },
    ],
    [
      { text: "初診時" }, { text: "6点" },
      { text: "17点", options: { bold: true, color: C.primary } },
      { text: "23点", options: { bold: true, color: C.primary } },
      { text: "34点", options: { bold: true, color: C.green } },
      { text: "40点", options: { bold: true, color: C.green } },
    ],
    [
      { text: "再診時" }, { text: "2点" },
      { text: "4点", options: { bold: true, color: C.primary } },
      { text: "6点", options: { bold: true, color: C.primary } },
      { text: "8点", options: { bold: true, color: C.green } },
      { text: "10点", options: { bold: true, color: C.green } },
    ],
    [
      { text: "訪問診療\n（同一建物以外）" }, { text: "28点" },
      { text: "79点", options: { bold: true, color: C.primary } },
      { text: "107点", options: { bold: true, color: C.primary } },
      { text: "158点", options: { bold: true, color: C.green } },
      { text: "186点", options: { bold: true, color: C.green } },
    ],
    [
      { text: "訪問診療\n（同一訪問時）" }, { text: "7点" },
      { text: "19点", options: { bold: true, color: C.primary } },
      { text: "26点", options: { bold: true, color: C.primary } },
      { text: "38点", options: { bold: true, color: C.green } },
      { text: "45点", options: { bold: true, color: C.green } },
    ],
  ];

  s.addTable(tableRows, {
    x: 0.4, y: 1.15, w: 9.2,
    fontSize: 13, fontFace: FONT_JP, color: C.text,
    border: { type: "solid", pt: 0.5, color: "DEE2E6" },
    colW: [2.0, 1.0, 1.3, 1.3, 1.3, 1.3],
    rowH: [0.4, 0.45, 0.45, 0.65, 0.65],
    align: "center", valign: "middle",
    autoPage: false,
  });

  // 注釈
  addCard(s, 0.4, 4.0, 9.2, 1.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.0, w: 0.15, h: 1.0, fill: { color: C.orange } });
  s.addText([
    { text: "注意：", options: { bold: true, color: C.orange } },
    { text: "賃上げ未実施の医療機関には入院基本料の", options: { color: C.text } },
    { text: "減算規定を新設", options: { bold: true, color: C.red } },
    { text: "\n例）急性期一般入院料1の場合：1日あたり ", options: { color: C.text } },
    { text: "121点減算", options: { bold: true, color: C.red } },
    { text: "\n夜勤手当の増額にも使途拡大が可能に", options: { color: C.text } },
  ], { x: 0.7, y: 4.05, w: 8.7, h: 0.9, fontSize: 13, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "5");
})();

// ============================================================
// SLIDE 6: 物価対応料の新設
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "物価対応 ― 「物価対応料」の新設");

  s.addText([
    { text: "R8年度以降の物価上昇に段階的に対応するため、初・再診料や入院料とは", options: { color: C.text } },
    { text: "別に算定可能", options: { bold: true, color: C.primary } },
    { text: "な加算として新設。R9年度は", options: { color: C.text } },
    { text: "R8年度の2倍", options: { bold: true, color: C.red } },
    { text: "の点数となる。", options: { color: C.text } },
  ], { x: 0.5, y: 1.05, w: 9.0, h: 0.55, fontSize: 14, fontFace: FONT_JP, margin: 0 });

  // 外来テーブル
  addCard(s, 0.4, 1.7, 4.3, 2.0);
  s.addText("外来・在宅物価対応料", { x: 0.5, y: 1.75, w: 4.0, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var outRows = [
    [
      { text: "", options: { fill: { color: C.primary } } },
      { text: "R8年度", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
      { text: "R9年度", options: { bold: true, color: "FFFFFF", fill: { color: C.green } } },
    ],
    [ { text: "初診時" }, { text: "2点" }, { text: "4点", options: { bold: true, color: C.green } } ],
    [ { text: "再診時等" }, { text: "2点" }, { text: "4点", options: { bold: true, color: C.green } } ],
    [ { text: "訪問診療時" }, { text: "3点" }, { text: "6点", options: { bold: true, color: C.green } } ],
  ];
  s.addTable(outRows, {
    x: 0.5, y: 2.15, w: 4.1,
    fontSize: 13, fontFace: FONT_JP, color: C.text,
    border: { type: "solid", pt: 0.5, color: "DEE2E6" },
    colW: [1.5, 1.3, 1.3], rowH: [0.35, 0.35, 0.35, 0.35],
    align: "center", valign: "middle", autoPage: false,
  });

  // 入院テーブル
  addCard(s, 5.1, 1.7, 4.6, 3.3);
  s.addText("入院物価対応料（1日につき）", { x: 5.2, y: 1.75, w: 4.3, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var inRows = [
    [
      { text: "施設区分", options: { fill: { color: C.primary }, color: "FFFFFF", bold: true } },
      { text: "R8", options: { fill: { color: C.primary }, color: "FFFFFF", bold: true } },
      { text: "R9", options: { fill: { color: C.green }, color: "FFFFFF", bold: true } },
    ],
    [ { text: "急性期病院A" }, { text: "66点" }, { text: "132点", options: { bold: true, color: C.green } } ],
    [ { text: "急性期一般1" }, { text: "58点" }, { text: "116点", options: { bold: true, color: C.green } } ],
    [ { text: "特定機能病院7:1" }, { text: "84点" }, { text: "168点", options: { bold: true, color: C.green } } ],
    [ { text: "地域包括医療1" }, { text: "49点" }, { text: "98点", options: { bold: true, color: C.green } } ],
    [ { text: "回復期リハ1" }, { text: "19点" }, { text: "38点", options: { bold: true, color: C.green } } ],
    [ { text: "療養病棟1" }, { text: "18点" }, { text: "36点", options: { bold: true, color: C.green } } ],
  ];
  s.addTable(inRows, {
    x: 5.2, y: 2.15, w: 4.4,
    fontSize: 12, fontFace: FONT_JP, color: C.text,
    border: { type: "solid", pt: 0.5, color: "DEE2E6" },
    colW: [2.2, 1.1, 1.1], rowH: [0.32, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34],
    align: "center", valign: "middle", autoPage: false,
  });

  addPageNum(s, "6");
})();

// ============================================================
// SLIDE 7: 入院基本料等の引き上げ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "入院基本料等の引き上げ ― 主な点数変更");

  var tableRows = [
    [
      { text: "項目", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
      { text: "現行", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
      { text: "改定後", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
      { text: "増点", options: { bold: true, color: "FFFFFF", fill: { color: C.green } } },
    ],
    [ { text: "再診料" }, { text: "75点" }, { text: "76点", options: { bold: true } }, { text: "+1", options: { color: C.green } } ],
    [ { text: "急性期一般入院料1" }, { text: "1,688点" }, { text: "1,874点", options: { bold: true } }, { text: "+186", options: { color: C.green, bold: true } } ],
    [ { text: "特定機能病院A 7:1" }, { text: "1,822点" }, { text: "2,146点", options: { bold: true } }, { text: "+324", options: { color: C.green, bold: true } } ],
    [ { text: "地域包括医療病棟1" }, { text: "3,050点" }, { text: "3,367点", options: { bold: true } }, { text: "+317", options: { color: C.green, bold: true } } ],
    [ { text: "回復期リハ病棟1" }, { text: "2,229点" }, { text: "2,346点", options: { bold: true } }, { text: "+117", options: { color: C.green, bold: true } } ],
    [ { text: "療養病棟入院料1" }, { text: "1,964点" }, { text: "2,035点", options: { bold: true } }, { text: "+71", options: { color: C.green, bold: true } } ],
    [ { text: "精神病棟10:1" }, { text: "1,306点" }, { text: "1,471点", options: { bold: true } }, { text: "+165", options: { color: C.green, bold: true } } ],
    [ { text: "地域包括ケア1(40日以内)" }, { text: "2,838点" }, { text: "2,955点", options: { bold: true } }, { text: "+117", options: { color: C.green, bold: true } } ],
  ];

  s.addTable(tableRows, {
    x: 0.4, y: 1.15, w: 9.2,
    fontSize: 13, fontFace: FONT_JP, color: C.text,
    border: { type: "solid", pt: 0.5, color: "DEE2E6" },
    colW: [3.5, 1.8, 1.8, 1.1],
    rowH: [0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4],
    align: "center", valign: "middle", autoPage: false,
  });

  s.addText("※初診料は291点で据置。初再診料が包括されるその他の入院料等についても同様に対応。", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.3, fontSize: 11, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  addPageNum(s, "7");
})();

// ============================================================
// SLIDE 8: 急性期病院の再編
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "急性期病院の再編 ― 急性期病院入院基本料の新設");

  s.addText([
    { text: "「病棟単位」の評価から「", options: { fontSize: 14, color: C.text } },
    { text: "病院全体の急性期機能", options: { fontSize: 14, bold: true, color: C.primary } },
    { text: "」に着目した評価体系に転換", options: { fontSize: 14, color: C.text } },
  ], { x: 0.5, y: 1.05, w: 9.0, h: 0.4, fontFace: FONT_JP, margin: 0 });

  // 急性期病院A
  addCard(s, 0.4, 1.6, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.6, w: 4.4, h: 0.45, fill: { color: C.primary } });
  s.addText("急性期病院A 一般入院料", { x: 0.5, y: 1.62, w: 4.2, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });

  s.addText("1,930点", { x: 0.6, y: 2.15, w: 2.0, h: 0.5, fontSize: 28, fontFace: FONT_EN, color: C.primary, bold: true, margin: 0 });
  s.addText("看護配置 7対1", { x: 2.6, y: 2.2, w: 2.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });

  s.addText([
    { text: "要件（全て満たす）\n", options: { bold: true, color: C.text, fontSize: 12 } },
    { text: "・救急搬送 年間2,000件以上\n・全身麻酔手術 年間1,200件以上\n・DPC対象病院\n・地域包括医療病棟の届出なし", options: { color: C.sub, fontSize: 11 } },
  ], { x: 0.6, y: 2.7, w: 4.0, h: 1.3, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.3 });

  // 急性期病院B
  addCard(s, 5.2, 1.6, 4.5, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.6, w: 4.5, h: 0.45, fill: { color: C.accent } });
  s.addText("急性期病院B 一般入院料", { x: 5.3, y: 1.62, w: 4.3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });

  s.addText("1,643点", { x: 5.4, y: 2.15, w: 2.0, h: 0.5, fontSize: 28, fontFace: FONT_EN, color: C.accent, bold: true, margin: 0 });
  s.addText("看護配置 10対1", { x: 7.4, y: 2.2, w: 2.0, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0 });

  s.addText([
    { text: "要件（いずれか）\n", options: { bold: true, color: C.text, fontSize: 12 } },
    { text: "・救急搬送 1,500件以上\n・救急搬送500件 + 全麻手術500件\n・人口20万人未満の最大救急搬送病院\n・離島地域の最大救急搬送病院", options: { color: C.sub, fontSize: 11 } },
  ], { x: 5.4, y: 2.7, w: 4.1, h: 1.3, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.3 });

  // 下部：既存との関係
  addCard(s, 0.4, 4.3, 9.3, 0.8);
  s.addText([
    { text: "既存の急性期一般入院料1〜6も引き続き選択可能", options: { fontSize: 13, color: C.text } },
    { text: "  →  実績等に応じ医療機関が選択", options: { fontSize: 13, color: C.primary, bold: true } },
    { text: "\n看護・多職種協働加算を新設：", options: { fontSize: 12, color: C.sub } },
    { text: "加算1 = 277点、加算2 = 255点", options: { fontSize: 12, color: C.primary, bold: true } },
    { text: "（PT/OT/ST/管理栄養士/臨検技師が病棟で協働）", options: { fontSize: 11, color: C.sub } },
  ], { x: 0.6, y: 4.32, w: 8.9, h: 0.75, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "8");
})();

// ============================================================
// SLIDE 9: その他の重要変更
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "その他の重要変更");

  // 食費・光熱水費
  addCard(s, 0.4, 1.15, 4.5, 1.6);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 0.12, h: 1.6, fill: { color: C.orange } });
  s.addText("入院時の食費・光熱水費", { x: 0.7, y: 1.2, w: 4.0, h: 0.3, fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0 });
  s.addText([
    { text: "食費基準額：690円 → ", options: { fontSize: 13, color: C.text } },
    { text: "730円", options: { fontSize: 16, bold: true, color: C.orange } },
    { text: "（+40円/食）\n", options: { fontSize: 13, color: C.text } },
    { text: "光熱水費：398円 → ", options: { fontSize: 13, color: C.text } },
    { text: "458円", options: { fontSize: 16, bold: true, color: C.orange } },
    { text: "（+60円/日）\n", options: { fontSize: 13, color: C.text } },
    { text: "患者負担：所得に応じ+20〜40円/食", options: { fontSize: 12, color: C.sub } },
  ], { x: 0.7, y: 1.55, w: 4.0, h: 1.1, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.4 });

  // 特定機能病院
  addCard(s, 5.2, 1.15, 4.5, 1.6);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.15, w: 0.12, h: 1.6, fill: { color: C.primary } });
  s.addText("特定機能病院入院基本料", { x: 5.5, y: 1.2, w: 4.0, h: 0.3, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "A・B・Cの3区分", options: { fontSize: 16, bold: true, color: C.primary } },
    { text: "に再編\n", options: { fontSize: 13, color: C.text } },
    { text: "A：幅広い診療科設置の特定機能病院\nB：中長期目標を設定する特定機能病院\nC：A・B以外の特定機能病院", options: { fontSize: 12, color: C.sub } },
  ], { x: 5.5, y: 1.55, w: 4.0, h: 1.1, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.3 });

  // ICU
  addCard(s, 0.4, 2.95, 4.5, 2.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.95, w: 0.12, h: 2.1, fill: { color: C.red } });
  s.addText("救命救急・ICUの見直し", { x: 0.7, y: 3.0, w: 4.0, h: 0.3, fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText([
    { text: "救命救急入院料：4区分 → ", options: { fontSize: 12, color: C.text } },
    { text: "2区分", options: { fontSize: 13, bold: true, color: C.red } },
    { text: "\n特定集中治療室管理料：6区分 → ", options: { fontSize: 12, color: C.text } },
    { text: "3区分", options: { fontSize: 13, bold: true, color: C.red } },
    { text: "\nSOFA患者割合：1割 → ", options: { fontSize: 12, color: C.text } },
    { text: "2割以上", options: { fontSize: 13, bold: true, color: C.red } },
    { text: "\n病院実績要件（搬送件数・手術件数）新設", options: { fontSize: 11, color: C.sub } },
  ], { x: 0.7, y: 3.35, w: 4.0, h: 1.6, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.5 });

  // 重症度
  addCard(s, 5.2, 2.95, 4.5, 2.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.95, w: 0.12, h: 2.1, fill: { color: C.green } });
  s.addText("重症度、医療・看護必要度", { x: 5.5, y: 3.0, w: 4.0, h: 0.3, fontSize: 14, fontFace: FONT_JP, color: C.green, bold: true, margin: 0 });
  s.addText([
    { text: "A・C項目に対象コード追加\n", options: { fontSize: 12, color: C.text } },
    { text: "該当患者割合に救急搬送応需係数を加味した", options: { fontSize: 11, color: C.text } },
    { text: "指数", options: { fontSize: 12, bold: true, color: C.green } },
    { text: "に見直し\n", options: { fontSize: 11, color: C.text } },
    { text: "B項目測定：入院5日目以降 ", options: { fontSize: 12, color: C.text } },
    { text: "7日ごとに1回でも可", options: { fontSize: 12, bold: true, color: C.green } },
  ], { x: 5.5, y: 3.35, w: 4.0, h: 1.6, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.5 });

  addPageNum(s, "9");
})();

// ============================================================
// SLIDE 10: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.8, y: 0.25, w: 8.4, h: 0.6,
    fontSize: 28, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var msgs = [
    { num: "1", text: "診療報酬本体 +3.09%（2年度平均）\n― 賃上げ+1.70%と物価対応+0.76%が二本柱" },
    { num: "2", text: "BU評価料が大幅拡充 ― 事務職員を含む幅広い職員が対象\n賃上げ未実施には入院基本料の減算規定を新設" },
    { num: "3", text: "「物価対応料」を新設 ― R9年度はR8年度の2倍\n入院基本料等も大幅増点（急性期一般1：+186点）" },
    { num: "4", text: "急性期病院A/B入院基本料を新設\n― 病院全体の急性期機能に着目した新評価体系" },
    { num: "5", text: "2年度同時改定の枠組み\n― R9年度に段階的引上げ、経済動向に応じた調整も" },
  ];

  for (var i = 0; i < msgs.length; i++) {
    var yy = 1.05 + i * 0.87;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: yy + 0.08, w: 0.45, h: 0.45, fill: { color: C.accent } });
    s.addText(msgs[i].num, { x: 0.6, y: yy + 0.1, w: 0.45, h: 0.42, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(msgs[i].text, { x: 1.25, y: yy, w: 8.2, h: 0.8, fontSize: 14, fontFace: FONT_JP, color: C.light, margin: 0, lineSpacingMultiple: 1.3 });
  }

  s.addText("医知創造ラボ｜今村久司　脳神経内科専門医", {
    x: 0.8, y: 5.15, w: 8.4, h: 0.35,
    fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "10");
})();

// ============================================================
// OUTPUT
// ============================================================
var outPath = __dirname + "/令和8年度診療報酬改定.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
