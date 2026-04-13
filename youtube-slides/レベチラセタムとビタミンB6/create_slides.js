var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "イーケプラのイライラ、ビタミンB6で改善する？ エビデンスとBRVとの使い分け";

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
  s.addText(num, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_JP, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("脳神経内科医が解説する薬の話", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("イーケプラのイライラ\nビタミンB6で改善する？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― エビデンスとブリバラセタムとの使い分け ―", {
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
    fontSize: 14, fontFace: FONT_JP, color: C.light, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: レベチラセタムの精神症状とは
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "レベチラセタム（イーケプラ）の精神症状");

  // 導入テキスト
  s.addText("使いやすく有効性の高い抗てんかん薬だが、精神・行動面の副作用がしばしば問題に", {
    x: 0.5, y: 1.1, w: 9.0, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 症状カード群
  var symptoms = [
    { icon: "😤", label: "易刺激性", desc: "ちょっとしたことでイライラする" },
    { icon: "😡", label: "攻撃性・興奮", desc: "突然怒り出す、暴言" },
    { icon: "😰", label: "焦燥・不安", desc: "落ち着きがない、焦り" },
    { icon: "😢", label: "抑うつ", desc: "気分の落ち込み、意欲低下" },
  ];

  symptoms.forEach(function(item, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var xPos = 0.5 + col * 4.7;
    var yPos = 1.8 + row * 1.6;
    addCard(s, xPos, yPos, 4.3, 1.35);
    s.addText(item.icon, { x: xPos + 0.2, y: yPos + 0.15, w: 0.8, h: 0.8, fontSize: 32, align: "center", margin: 0 });
    s.addText(item.label, { x: xPos + 1.0, y: yPos + 0.15, w: 3.0, h: 0.55, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
    s.addText(item.desc, { x: xPos + 1.0, y: yPos + 0.7, w: 3.0, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.sub, margin: 0 });
  });

  // 添付文書注意
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9.0, h: 0.55, fill: { color: "FFF3CD" } });
  s.addText("添付文書：精神症状 → 自殺企図への注意が明記されている", {
    x: 0.7, y: 5.05, w: 8.6, h: 0.45,
    fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });

  addPageNum(s, "2");
})();

// ============================================================
// SLIDE 3: なぜビタミンB6が注目されるのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "なぜビタミンB6が注目されるのか？");

  // PLP経路の説明
  addCard(s, 0.5, 1.2, 9.0, 2.0);
  s.addText("ビタミンB6の活性型 PLP（ピリドキサール5'-リン酸）", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  s.addText("100以上の酵素反応の補酵素として機能", {
    x: 0.8, y: 1.8, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  // 3つの神経伝達物質
  var nts = [
    { name: "セロトニン", color: C.green, desc: "気分の安定" },
    { name: "GABA", color: C.primary, desc: "抑制系" },
    { name: "ドパミン", color: C.orange, desc: "報酬・意欲" },
  ];

  nts.forEach(function(nt, i) {
    var xPos = 1.0 + i * 2.9;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.3, w: 2.5, h: 0.7, fill: { color: nt.color }, rectRadius: 0.1 });
    s.addText(nt.name, { x: xPos, y: 2.3, w: 2.5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(nt.desc, { x: xPos, y: 2.65, w: 2.5, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.white, align: "center", margin: 0 });
  });

  // 仮説ボックス
  addCard(s, 0.5, 3.5, 9.0, 2.0);
  s.addText("仮説：LEVの精神症状メカニズム", {
    x: 0.8, y: 3.6, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  var hypotheses = [
    "・AMPA受容体拮抗作用による影響",
    "・GABA系・セロトニン系への関与",
    "→ B6補充で神経伝達物質合成を支援し、症状を緩和できる可能性",
  ];

  hypotheses.forEach(function(h, i) {
    var clr = (i === 2) ? C.red : C.text;
    var bld = (i === 2) ? true : false;
    s.addText(h, {
      x: 1.0, y: 4.15 + i * 0.4, w: 8.0, h: 0.35,
      fontSize: 14, fontFace: FONT_JP, color: clr, bold: bld, margin: 0,
    });
  });

  addPageNum(s, "3");
})();

// ============================================================
// SLIDE 4: エビデンス① 小児データ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "エビデンス① 小児データ");

  // テーブルヘッダ
  var tableX = 0.3;
  var tableY = 1.15;
  var colW = [2.2, 1.6, 1.6, 4.0];
  var headers = ["研究", "デザイン", "対象", "結果"];
  var headerX = tableX;
  headers.forEach(function(h, i) {
    s.addShape(pres.shapes.RECTANGLE, { x: headerX, y: tableY, w: colW[i], h: 0.45, fill: { color: C.primary } });
    s.addText(h, { x: headerX, y: tableY, w: colW[i], h: 0.45, fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    headerX += colW[i];
  });

  // テーブル行
  var rows = [
    ["Romoli 2020\n(SR)", "SR\n(12研究)", "小児中心", "72.5%（108/149例）改善\n※バイアス大、エビデンス質低い"],
    ["Marino 2018", "前向き\nケースコントロール", "小児", "92%がLEV継続\n平均9日で改善"],
    ["Mahmoud 2021", "RCT", "小児", "B6群で改善幅が有意に大きい\n（p < 0.001）"],
    ["Lob 2022", "後ろ向き\nコホート", "小児\n240名", "LEV中止率 49% vs 88%\n（p = 0.001）"],
    ["Thananowan\n2025", "RCT\n102例", "小児\n思春期", "主要評価項目は陰性\n多変量解析でB6群が改善大"],
  ];

  rows.forEach(function(row, rIdx) {
    var rowY = tableY + 0.45 + rIdx * 0.75;
    var bgColor = (rIdx % 2 === 0) ? "F8F9FA" : C.white;
    var cellX = tableX;
    row.forEach(function(cell, cIdx) {
      s.addShape(pres.shapes.RECTANGLE, { x: cellX, y: rowY, w: colW[cIdx], h: 0.75, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      s.addText(cell, {
        x: cellX + 0.08, y: rowY, w: colW[cIdx] - 0.16, h: 0.75,
        fontSize: 10, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
        lineSpacingMultiple: 1.1,
      });
      cellX += colW[cIdx];
    });
  });

  // ボトムメッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.55, fill: { color: C.light } });
  s.addText("小児では「改善シグナル」あり ← ただし最大のRCT（2025年）は主要評価項目が陰性", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.45,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  addPageNum(s, "4");
})();

// ============================================================
// SLIDE 5: エビデンス② 成人データ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "エビデンス② 成人データ");

  // 成人研究カード
  var studies = [
    {
      title: "Alsaadi 2015（UAE ケースシリーズ）",
      result: "51例中34例（66.6%）改善",
      note: "後ろ向き、プラセボなし",
      color: C.green,
    },
    {
      title: "Dreischmeier 2021（米国 退役軍人）",
      result: "20例中9例（45%）改善、55%は無効",
      note: "後ろ向き、B6 100mg/日",
      color: C.orange,
    },
    {
      title: "Cheraghmakani 2022（成人RCT）",
      result: "B6群 vs プラセボ群：有意差なし",
      note: "唯一の成人RCT → 陰性結果",
      color: C.red,
    },
  ];

  studies.forEach(function(st, i) {
    var yPos = 1.2 + i * 1.3;
    addCard(s, 0.5, yPos, 9.0, 1.1);

    // 色付きサイドバー
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 0.15, h: 1.1, fill: { color: st.color } });

    s.addText(st.title, {
      x: 0.9, y: yPos + 0.08, w: 8.3, h: 0.35,
      fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
    });
    s.addText(st.result, {
      x: 0.9, y: yPos + 0.4, w: 8.3, h: 0.35,
      fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
    });
    s.addText(st.note, {
      x: 0.9, y: yPos + 0.72, w: 8.3, h: 0.3,
      fontSize: 12, fontFace: FONT_JP, color: C.sub, margin: 0,
    });
  });

  // 結論ボックス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9.0, h: 0.55, fill: { color: "FDECEA" } });
  s.addText("成人のエビデンスは小児より弱い ── 唯一のRCTは陰性結果", {
    x: 0.7, y: 5.05, w: 8.6, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });

  addPageNum(s, "5");
})();

// ============================================================
// SLIDE 6: B6 vs BRV 比較
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ビタミンB6 vs ブリバラセタム（BRV）");

  // 比較テーブル
  var tX = 0.5;
  var tY = 1.2;
  var cW = [2.5, 3.5, 3.5];
  var tHeaders = ["", "ビタミンB6", "ブリバラセタム（BRV）"];

  var hX = tX;
  tHeaders.forEach(function(h, i) {
    var bgClr = (i === 0) ? C.dark : (i === 1) ? C.green : C.primary;
    s.addShape(pres.shapes.RECTANGLE, { x: hX, y: tY, w: cW[i], h: 0.5, fill: { color: bgClr } });
    s.addText(h, { x: hX, y: tY, w: cW[i], h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    hX += cW[i];
  });

  var compRows = [
    ["改善率", "45〜72.5%\n（研究による差大）", "33.3〜83.0%\n（加重平均66.6%）"],
    ["エビデンス質", "低い\n（後ろ向き中心、RCT陰性）", "中程度\n（前向き研究あり）"],
    ["薬価\n（1錠あたり）", "ピドキサール30mg\n5.9〜6.3円", "ブリィビアクト50mg\n609.3円"],
    ["効果発現", "約1〜3週間", "切り替え後 早期"],
    ["位置づけ", "LEVを救済する\n安価な一手", "より確実性の高い\n本命"],
  ];

  compRows.forEach(function(row, rIdx) {
    var rowY = tY + 0.5 + rIdx * 0.75;
    var bgColor = (rIdx % 2 === 0) ? "F8F9FA" : C.white;
    var cX = tX;
    row.forEach(function(cell, ci) {
      s.addShape(pres.shapes.RECTANGLE, { x: cX, y: rowY, w: cW[ci], h: 0.75, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      var fClr = (ci === 0) ? C.dark : C.text;
      var fBold = (ci === 0) ? true : false;
      s.addText(cell, {
        x: cX + 0.1, y: rowY, w: cW[ci] - 0.2, h: 0.75,
        fontSize: 11, fontFace: FONT_JP, color: fClr, bold: fBold, valign: "middle", margin: 0,
        lineSpacingMultiple: 1.1,
      });
      cX += cW[ci];
    });
  });

  addPageNum(s, "6");
})();

// ============================================================
// SLIDE 7: 実践フローチャート
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "実践フローチャート");

  // フローチャート
  // Step 1: LEV精神症状出現
  s.addShape(pres.shapes.RECTANGLE, { x: 3.0, y: 1.2, w: 4.0, h: 0.65, fill: { color: C.primary }, rectRadius: 0.1 });
  s.addText("LEVで精神症状が出現", { x: 3.0, y: 1.2, w: 4.0, h: 0.65, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // 矢印
  s.addShape(pres.shapes.LINE, { x: 5.0, y: 1.85, w: 0, h: 0.45, line: { color: C.sub, width: 2 } });

  // 分岐: 重症度判定
  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 2.3, w: 5.0, h: 0.65, fill: { color: C.yellow }, rectRadius: 0.1 });
  s.addText("重症度を評価する", { x: 2.5, y: 2.3, w: 5.0, h: 0.65, fontSize: 16, fontFace: FONT_JP, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0 });

  // 左分岐（軽症〜中等症）
  s.addShape(pres.shapes.LINE, { x: 3.5, y: 2.95, w: 0, h: 0.45, line: { color: C.green, width: 2 } });
  s.addText("軽症〜中等症", { x: 1.5, y: 2.95, w: 2.0, h: 0.4, fontSize: 12, fontFace: FONT_JP, color: C.green, bold: true, align: "center", margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 3.4, w: 3.5, h: 0.6, fill: { color: C.green }, rectRadius: 0.1 });
  s.addText("B6 短期試行（2-4週）", { x: 1.0, y: 3.4, w: 3.5, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  s.addShape(pres.shapes.LINE, { x: 2.75, y: 4.0, w: 0, h: 0.4, line: { color: C.green, width: 2 } });

  // 分岐：改善あり/なし
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.4, w: 2.2, h: 0.55, fill: { color: "D4EDDA" }, rectRadius: 0.1 });
  s.addText("改善あり\n→ B6継続", { x: 0.3, y: 4.4, w: 2.2, h: 0.55, fontSize: 12, fontFace: FONT_JP, color: C.green, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 2.8, y: 4.4, w: 2.2, h: 0.55, fill: { color: "FDECEA" }, rectRadius: 0.1 });
  s.addText("改善なし\n→ BRVへ切替", { x: 2.8, y: 4.4, w: 2.2, h: 0.55, fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.0 });

  // 右分岐（重症）
  s.addShape(pres.shapes.LINE, { x: 6.5, y: 2.95, w: 0, h: 0.45, line: { color: C.red, width: 2 } });
  s.addText("重症", { x: 6.0, y: 2.95, w: 2.0, h: 0.4, fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 3.4, w: 3.5, h: 0.6, fill: { color: C.red }, rectRadius: 0.1 });
  s.addText("LEV減量/中止 or BRVへ", { x: 5.5, y: 3.4, w: 3.5, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // 重症の定義
  addCard(s, 5.5, 4.2, 3.7, 1.2);
  s.addText("重症の基準：", { x: 5.7, y: 4.25, w: 3.3, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  var severeItems = ["・自殺念慮/企図", "・精神病症状", "・他害リスク", "・急速な悪化"];
  severeItems.forEach(function(item, i) {
    s.addText(item, { x: 5.7, y: 4.55 + i * 0.2, w: 3.3, h: 0.2, fontSize: 11, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  addPageNum(s, "7");
})();

// ============================================================
// SLIDE 8: 日本で使うときの注意点
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "日本で使うときの注意点");

  // 注意点1: 製剤の違い
  addCard(s, 0.5, 1.2, 9.0, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 0.15, h: 1.5, fill: { color: C.orange } });
  s.addText("1. 製剤の違い：Pyridoxine vs ピドキサール", {
    x: 0.9, y: 1.25, w: 8.3, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("海外研究 → pyridoxine（ピリドキシン）\n日本の処方 → ピドキサール（活性型PLP製剤）\n同じ「ビタミンB6」でも厳密には異なる製剤", {
    x: 0.9, y: 1.7, w: 8.3, h: 0.9,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // 注意点2: 用量
  addCard(s, 0.5, 2.9, 9.0, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.9, w: 0.15, h: 1.2, fill: { color: C.orange } });
  s.addText("2. 用量の問題", {
    x: 0.9, y: 2.95, w: 8.3, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("海外エビデンス：小児 50-350 mg/日、成人 50-100 mg/日（pyridoxine）\n日本のピドキサール：通常成人 10-60 mg/日 → 単純な換算は不可", {
    x: 0.9, y: 3.35, w: 8.3, h: 0.7,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // 注意点3: 安全性
  addCard(s, 0.5, 4.3, 9.0, 1.25);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.3, w: 0.15, h: 1.25, fill: { color: C.red } });
  s.addText("3. 過量投与のリスク", {
    x: 0.9, y: 4.35, w: 8.3, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("Pyridoxine 1-6 g/日の長期投与 → 感覚ニューロパチー\n上限量：米国FNB 100 mg/日 vs EFSA 12 mg/日\n→ 漫然と長期継続する薬ではない", {
    x: 0.9, y: 4.75, w: 8.3, h: 0.7,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  addPageNum(s, "8");
})();

// ============================================================
// SLIDE 9: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.8, y: 0.3, w: 8.4, h: 0.6,
    fontSize: 28, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "LEV精神症状に対するB6には複数の報告あり\nしかしエビデンスの質は低く、特に成人では弱い" },
    { num: "2", text: "B6は「標準策」ではなく\n条件のよい症例で先に試せる低コスト介入" },
    { num: "3", text: "軽症〜中等症 → B6短期試行（2-4週）\n重症・B6無効 → BRVへ早めに切替" },
    { num: "4", text: "日本ではpyridoxine vs ピドキサールの製剤差に注意\n用量設計の根拠はまだ弱い" },
  ];

  messages.forEach(function(msg, i) {
    var yPos = 1.1 + i * 1.1;
    // 番号バッジ
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: yPos + 0.05, w: 0.55, h: 0.55, fill: { color: C.accent } });
    s.addText(msg.num, { x: 0.8, y: yPos + 0.05, w: 0.55, h: 0.55, fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // テキスト
    s.addText(msg.text, {
      x: 1.6, y: yPos, w: 7.6, h: 0.85,
      fontSize: 16, fontFace: FONT_JP, color: C.light, margin: 0,
      lineSpacingMultiple: 1.2,
    });
  });

  // フッター
  s.addText("医知創造ラボ  今村久司  脳神経内科専門医", {
    x: 0.8, y: 5.1, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
})();

// ============================================================
// Output
// ============================================================
var outPath = __dirname + "/レベチラセタムとビタミンB6_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
