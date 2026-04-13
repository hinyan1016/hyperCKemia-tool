var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "お酒を飲むと記憶がなくなる理由 ― ブラックアウトのメカニズムを脳神経内科医が解説";

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

  s.addText("脳神経内科医が答える からだの不思議 #12", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("お酒を飲むと\n記憶がなくなる理由", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― ブラックアウトのメカニズムを解説 ―", {
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
    "「昨日、どうやって帰ったか覚えていない……」",
    "「飲み会の後半の記憶がまったくない」",
    "「周りはしっかりしてたって言うのに、\n  自分だけ覚えていない」",
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
  s.addText("これが「アルコール性ブラックアウト」 ― 意識はあるのに記憶だけが作られない", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: 海馬とLTP ― 記憶はどうやって作られるか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "記憶の要 ― 海馬とLTP（長期増強）");

  // 海馬の役割カード
  addCard(s, 0.5, 1.1, 9.0, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 0.14, h: 1.3, fill: { color: C.accent } });
  s.addText("海馬（かいば）", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.45,
    fontSize: 20, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("側頭葉の内側にある構造。経験した出来事をシナプスで強化して「長期記憶」へ変換する記憶の固定（memory consolidation）を担う", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.7,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // LTPのステップ
  var steps = [
    { icon: "①", label: "神経活動", desc: "シナプスに繰り返し\n信号が入る" },
    { icon: "②", label: "LTP発動", desc: "シナプスが強化され\n接続が確立する" },
    { icon: "③", label: "長期記憶", desc: "体験が記憶として\n保存される" },
  ];
  steps.forEach(function(st, i) {
    var xPos = 0.5 + i * 3.1;
    addCard(s, xPos, 2.7, 2.9, 2.0);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.7, w: 2.9, h: 0.5, fill: { color: C.primary } });
    s.addText(st.icon + " " + st.label, { x: xPos, y: 2.7, w: 2.9, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.desc, { x: xPos, y: 3.3, w: 2.9, h: 1.2, fontSize: 15, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });
  });

  s.addText("→", { x: 3.15, y: 3.3, w: 0.5, h: 0.5, fontSize: 24, color: C.accent, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 6.25, y: 3.3, w: 0.5, h: 0.5, fontSize: 24, color: C.accent, align: "center", valign: "middle", margin: 0 });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: NMDA受容体 / GABA-A受容体 ― アルコールの作用機序
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "アルコールが記憶を妨げる2つの機序");

  // NMDA受容体
  addCard(s, 0.4, 1.1, 4.4, 3.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 0.55, fill: { color: C.red } });
  s.addText("NMDA受容体 を抑制", { x: 0.5, y: 1.1, w: 4.2, h: 0.55, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "本来の役割：", options: { bold: true, color: C.primary, breakLine: true } },
    { text: "LTP（長期増強）を開始するスイッチ", options: { color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "アルコールの影響：", options: { bold: true, color: C.red, breakLine: true } },
    { text: "このスイッチをブロック\n" +
             "→ LTPが発動せず\n" +
             "→ 記憶の固定が起きない", options: { color: C.text } },
  ], { x: 0.6, y: 1.75, w: 4.0, h: 2.5, fontSize: 15, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.4, w: 4.4, h: 0.4, fill: { color: C.light } });
  s.addText("記憶が最初から作られない", { x: 0.5, y: 4.4, w: 4.2, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: 0 });

  // GABA-A受容体
  addCard(s, 5.2, 1.1, 4.4, 3.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.55, fill: { color: C.orange } });
  s.addText("GABA-A受容体 を増強", { x: 5.3, y: 1.1, w: 4.2, h: 0.55, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "本来の役割：", options: { bold: true, color: C.primary, breakLine: true } },
    { text: "脳の興奮を抑える（抑制性）", options: { color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "アルコールの影響：", options: { bold: true, color: C.orange, breakLine: true } },
    { text: "過剰に活性化\n" +
             "→ 海馬ニューロンの発火が低下\n" +
             "→ 記憶回路全体が鈍る", options: { color: C.text } },
  ], { x: 5.4, y: 1.75, w: 4.0, h: 2.5, fontSize: 15, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 4.4, w: 4.4, h: 0.4, fill: { color: C.light } });
  s.addText("脳全体の活動が抑制される", { x: 5.3, y: 4.4, w: 4.2, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, valign: "middle", margin: 0 });

  // 中央の＋
  s.addText("+", { x: 4.6, y: 2.5, w: 0.4, h: 0.5, fontSize: 30, fontFace: FONT_EN, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0 });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: 断片型 vs 完全型 ブラックアウト 比較表
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "2種類のブラックアウト ― 断片型 vs 完全型");

  var rows = [
    { item: "別名", frag: "グレーアウト", enbloc: "完全ブラックアウト" },
    { item: "記憶の欠落", frag: "部分的・断片的な空白", enbloc: "一定時間の記憶がまるごと消失" },
    { item: "ヒントで\n思い出せるか", frag: "断片的に思い出せることがある", enbloc: "いくらヒントを出されても\n思い出せない" },
    { item: "血中アルコール\n濃度", frag: "比較的低め", enbloc: "高値（0.15%以上が多い）" },
    { item: "深刻度", frag: "より一般的", enbloc: "より深刻（問題飲酒と関連）" },
  ];

  // ヘッダ行
  var tblY = 1.1;
  var hdrH = 0.5;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: tblY, w: 1.8, h: hdrH, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.1, y: tblY, w: 3.7, h: hdrH, fill: { color: C.orange } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: tblY, w: 3.9, h: hdrH, fill: { color: C.red } });
  s.addText("項目", { x: 0.3, y: tblY, w: 1.8, h: hdrH, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("断片型（fragmentary）", { x: 2.1, y: tblY, w: 3.7, h: hdrH, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("完全型（en bloc）", { x: 5.8, y: tblY, w: 3.9, h: hdrH, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var rowH = 0.64;
  rows.forEach(function(r, i) {
    var yPos = tblY + hdrH + i * rowH;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: rowH, fill: { color: bgColor } });
    s.addText(r.item, { x: 0.3, y: yPos, w: 1.8, h: rowH, fontSize: 11, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2, lineSpacingMultiple: 1.15 });
    s.addText(r.frag, { x: 2.1, y: yPos, w: 3.7, h: rowH, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 8, 0, 8], lineSpacingMultiple: 1.15 });
    s.addText(r.enbloc, { x: 5.8, y: yPos, w: 3.9, h: rowH, fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: [0, 8, 0, 8], lineSpacingMultiple: 1.15 });
  });

  // 統計注記
  var noteY = tblY + hdrH + rows.length * rowH + 0.15;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: noteY, w: 9.4, h: 0.42, fill: { color: C.light } });
  s.addText("飲酒経験のある大学生の51%がブラックアウトを経験（Duke大学 2002）", {
    x: 0.5, y: noteY, w: 9.0, h: 0.42,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0,
  });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: リスク要因
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "記憶が飛びやすい条件 ― リスク要因");

  // 飲み方リスク
  addCard(s, 0.3, 1.1, 4.5, 3.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 4.5, h: 0.5, fill: { color: C.red } });
  s.addText("飲み方のリスク", { x: 0.4, y: 1.1, w: 4.3, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var drinkRisks = [
    "急速な飲酒（一気・短時間での大量摂取）\n→ BACが急上昇し海馬への影響が強まる",
    "空腹での飲酒\n→ 胃からの吸収が速くBACが急上昇",
    "度数の高いアルコール（蒸留酒）\n→ 同量でも血中濃度上昇が早い",
  ];
  drinkRisks.forEach(function(d, i) {
    var yPos = 1.75 + i * 1.0;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: yPos + 0.15, w: 0.4, h: 0.4, fill: { color: C.red } });
    s.addText("" + (i + 1), { x: 0.5, y: yPos + 0.15, w: 0.4, h: 0.4, fontSize: 13, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(d, { x: 1.1, y: yPos, w: 3.5, h: 0.8, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // 個人差リスク
  addCard(s, 5.2, 1.1, 4.5, 3.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.5, h: 0.5, fill: { color: C.orange } });
  s.addText("個人差のリスク", { x: 5.3, y: 1.1, w: 4.3, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var personRisks = [
    "女性：水分量が少なくBACが上がりやすい",
    "若年者：前頭前野が未発達で\n飲み過ぎやすい",
    "遺伝的素因：GABA受容体の個人差",
    "睡眠不足・疲労：脳の予備能が低下",
  ];
  personRisks.forEach(function(p, i) {
    var yPos = 1.75 + i * 0.77;
    s.addShape(pres.shapes.OVAL, { x: 5.4, y: yPos + 0.1, w: 0.35, h: 0.35, fill: { color: C.orange } });
    s.addText("" + (i + 1), { x: 5.4, y: yPos + 0.1, w: 0.35, h: 0.35, fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p, { x: 5.9, y: yPos, w: 3.6, h: 0.65, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // 警告ボックス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.5, fill: { color: C.yellow } });
  s.addText("「酒が強い＝記憶が飛ばない」は誤解！ 耐性と記憶形成への影響は別物", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.dark, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: 長期的影響 ― ウェルニッケ・コルサコフ症候群
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "繰り返すとどうなるか ― 長期的な脳への影響");

  // ウェルニッケ・コルサコフ
  addCard(s, 0.3, 1.1, 9.4, 2.35);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.4, h: 0.5, fill: { color: C.red } });
  s.addText("ウェルニッケ脳症 → コルサコフ症候群", { x: 0.5, y: 1.1, w: 9.0, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "原因：", options: { bold: true, color: C.primary, breakLine: false } },
    { text: "慢性アルコール依存によるビタミンB1（チアミン）欠乏", options: { color: C.text, breakLine: true } },
    { text: "ウェルニッケ脳症：", options: { bold: true, color: C.red, breakLine: false } },
    { text: "眼球運動障害・歩行失調・意識障害の三徴候（急性期）", options: { color: C.text, breakLine: true } },
    { text: "コルサコフ症候群：", options: { bold: true, color: C.red, breakLine: false } },
    { text: "前向性＆逆向性健忘・作話。不可逆的な記憶障害が残ることが多い", options: { color: C.text } },
  ], { x: 0.5, y: 1.7, w: 9.0, h: 1.6, fontSize: 14, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.35 });

  // アルコール性認知症
  addCard(s, 0.3, 3.6, 9.4, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.6, w: 9.4, h: 0.5, fill: { color: C.orange } });
  s.addText("アルコール性認知症", { x: 0.5, y: 3.6, w: 9.0, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("長期多量飲酒 → 海馬・前頭前野の神経細胞喪失・白質変性 → 記憶だけでなく判断力・実行機能・社会行動にも影響", {
    x: 0.5, y: 4.15, w: 9.0, h: 0.65,
    fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
  });

  // 目安ボックス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.5, fill: { color: C.light } });
  s.addText("生活習慣病リスク増加の目安：男性 純アルコール40g/日以上、女性 20g/日以上（厚労省ガイドライン 2024）", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.5,
    fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: 要注意チェックリスト
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなときは要注意 ― 受診を考えるサイン");

  var checks = [
    "少量（缶ビール1〜2本程度）でもブラックアウトが起きるようになった",
    "ブラックアウトの頻度が増えてきた",
    "飲んでいないときでも物忘れが目立つようになった",
    "目がかすれる・ふらつく・手が震えるなどが飲酒後に出る",
    "断酒すると手が震えたり、汗が出たり、気分が悪くなる（離脱症状の可能性）",
    "「飲まないと眠れない」「飲まないとイライラする」と感じる",
    "周囲から「最近ひどく酔っていることが多い」と言われる",
  ];

  checks.forEach(function(c, i) {
    var col = (i < 4) ? 0 : 1;
    var row = (i < 4) ? i : i - 4;
    var xPos = 0.3 + col * 4.9;
    var yPos = 1.1 + row * 0.97;

    addCard(s, xPos, yPos, 4.6, 0.82);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.1, y: yPos + 0.18, w: 0.38, h: 0.38, fill: { color: C.red } });
    s.addText("" + (i + 1), { x: xPos + 0.1, y: yPos + 0.18, w: 0.38, h: 0.38, fontSize: 12, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c, { x: xPos + 0.58, y: yPos, w: 3.9, h: 0.82, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  // 救急受診
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.5, fill: { color: C.red } });
  s.addText("救急受診が必要：意識消失・呼びかけに反応しない・けいれん・眼が動かない＋ふらつき", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "8/9");
})();

// ============================================================
// SLIDE 9: まとめ（Take Home Message）
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
    { num: "1", text: "ブラックアウトは記憶が「消えた」のではなく「最初から作られなかった」\n― 海馬のLTPがブロックされるため" },
    { num: "2", text: "NMDA受容体遮断 ＋ GABA-A受容体増強の2重作用が記憶形成を妨げる" },
    { num: "3", text: "断片型と完全型の2種類があり、完全型のほうが深刻" },
    { num: "4", text: "急速飲酒・空腹・女性・若年者がリスク大。「酒が強い人」でも油断禁物" },
    { num: "5", text: "繰り返す多量飲酒 → ウェルニッケ・コルサコフ症候群やアルコール性認知症" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.05 + i * 0.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.72, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: yPos + 0.08, w: 0.52, h: 0.52, fill: { color: C.accent } });
    s.addText(m.num, { x: 0.8, y: yPos + 0.08, w: 0.52, h: 0.52, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.55, y: yPos, w: 7.65, h: 0.72, fontSize: 13, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.15, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.2, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "9/9");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/飲酒と記憶_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
