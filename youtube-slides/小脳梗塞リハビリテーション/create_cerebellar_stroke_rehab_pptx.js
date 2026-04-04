var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "小脳梗塞後のリハビリテーション ― エビデンスに基づく回復戦略";

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
var TOTAL = 17;

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
  s.addText(num + "/" + TOTAL, { x: 9.0, y: 5.2, w: 0.8, h: 0.4, fontSize: 11, fontFace: FONT_B, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: Title
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("小脳梗塞後のリハビリテーション", {
    x: 0.8, y: 0.8, w: 8.4, h: 1.2,
    fontSize: 38, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("Cerebellar Stroke Rehabilitation", {
    x: 0.8, y: 1.9, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_EN, color: C.accent, align: "center", italic: true, margin: 0,
  });
  s.addText("― エビデンスに基づく回復戦略と最新治療 ―", {
    x: 0.8, y: 2.6, w: 8.4, h: 0.7,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 3.5, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.8, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("最新ガイドライン・RCT・メタアナリシスに基づく", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.accent, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 本日のアウトライン
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "本日のアウトライン");
  addPageNum(s, 2);

  var items = [
    { num: "1", text: "小脳梗塞の臨床的特徴と後遺症" },
    { num: "2", text: "予後と回復の時間経過（KOSCOスタディ）" },
    { num: "3", text: "失調の評価スケール（SARA）" },
    { num: "4", text: "従来型リハビリテーション" },
    { num: "5", text: "非侵襲的脳刺激療法（NIBS）" },
    { num: "6", text: "新規治療（DBS・VR）" },
    { num: "7", text: "まとめ：5つの臨床ポイント" },
  ];

  items.forEach(function(item, i) {
    var yPos = 1.2 + i * 0.55;
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: yPos + 0.05, w: 0.4, h: 0.4, fill: { color: C.primary } });
    s.addText(item.num, { x: 0.8, y: yPos + 0.05, w: 0.4, h: 0.4, fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(item.text, { x: 1.4, y: yPos, w: 8, h: 0.5, fontSize: 20, fontFace: FONT_JP, color: C.text, margin: 0 });
  });
})();

// ============================================================
// SLIDE 3: 小脳梗塞の疫学と臨床的特徴
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "小脳梗塞の疫学と臨床的特徴");
  addPageNum(s, 3);

  // 疫学カード
  addCard(s, 0.5, 1.2, 4.2, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.2, h: 0.4, fill: { color: C.primary } });
  s.addText("疫学", { x: 0.7, y: 1.2, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("全脳卒中の 2〜3%\n比較的まれだが特有の後遺症パターン", { x: 0.7, y: 1.75, w: 3.8, h: 0.8, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0 });

  // 血管領域カード
  addCard(s, 5.3, 1.2, 4.2, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.2, w: 4.2, h: 0.4, fill: { color: C.primary } });
  s.addText("責任血管", { x: 5.5, y: 1.2, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("PICA / AICA / SCA\nSCA: 急性期最重度だが3ヶ月で差消失", { x: 5.5, y: 1.75, w: 3.8, h: 0.8, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0 });

  // 後遺症カード
  addCard(s, 0.5, 3.0, 9, 2.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 9, h: 0.4, fill: { color: C.red } });
  s.addText("主な後遺症", { x: 0.7, y: 3.0, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });

  var symptoms = [
    ["運動失調", "体幹・四肢の協調運動障害"],
    ["歩行障害", "ワイドベース歩行、速度低下"],
    ["姿勢障害", "バランス障害、体幹動揺"],
    ["めまい", "前庭系障害、眼振"],
    ["構音障害", "断綴性発語（scanning speech）"],
    ["嚥下障害", "虫部・両側病変で出現"],
  ];

  symptoms.forEach(function(sym, i) {
    var col = i < 3 ? 0 : 1;
    var row = i % 3;
    var xBase = col === 0 ? 0.7 : 5.0;
    var yBase = 3.55 + row * 0.5;
    s.addText(sym[0], { x: xBase, y: yBase, w: 1.5, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
    s.addText(sym[1], { x: xBase + 1.5, y: yBase, w: 3, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  });
})();

// ============================================================
// SLIDE 4: KOSCOスタディ — 回復の時間経過
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "KOSCOスタディ：機能ドメイン別の回復時間経過");
  addPageNum(s, 4);

  // 研究概要
  addCard(s, 0.5, 1.1, 9, 0.6);
  s.addText("Lee HS et al. Front Neurol 2025 | 孤立性小脳梗塞 183例 | 36ヶ月追跡", {
    x: 0.7, y: 1.15, w: 8.6, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // タイムライン
  var timeline = [
    { period: "3ヶ月", domain: "認知（K-MMSE 27.6）\n嚥下機能", color: C.green },
    { period: "6ヶ月", domain: "運動（FMA 98.8）\n言語機能", color: C.accent },
    { period: "12ヶ月", domain: "歩行（FAC 4.7）\n機能的自立度（FIM 122.2）", color: C.primary },
    { period: "36ヶ月", domain: "安定した予後を維持", color: C.dark },
  ];

  timeline.forEach(function(item, i) {
    var yPos = 1.95 + i * 0.85;
    // 期間バッジ
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 1.5, h: 0.6, fill: { color: item.color }, rectRadius: 0.1 });
    s.addText(item.period, { x: 0.8, y: yPos, w: 1.5, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // 矢印線
    if (i < 3) {
      s.addShape(pres.shapes.LINE, { x: 1.55, y: yPos + 0.6, w: 0, h: 0.25, line: { color: C.sub, width: 2 } });
    }
    // ドメイン説明
    addCard(s, 2.6, yPos - 0.05, 6.8, 0.7);
    s.addText(item.domain, { x: 2.8, y: yPos, w: 6.4, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  // 予後不良因子
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.02, fill: { color: C.red } });
  s.addText("予後不良因子: SCA領域 / 高齢 / 女性 / 初期重症度高 / 併存疾患", {
    x: 0.5, y: 5.05, w: 9, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 5: リハビリ後の機能回復（Kelly 2001）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "リハビリ後の機能回復（Kelly et al. 2001）");
  addPageNum(s, 5);

  s.addText("小脳脳卒中 58例（梗塞49、出血9）| 平均年齢 69.2歳", {
    x: 0.5, y: 1.1, w: 9, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // FIMスコアカード
  var fimData = [
    { label: "入院時FIM", value: "65.5", color: C.red },
    { label: "退院時FIM", value: "89.8", color: C.orange },
    { label: "F/U FIM", value: "123.5", color: C.green },
  ];

  fimData.forEach(function(item, i) {
    var xPos = 0.8 + i * 3.1;
    addCard(s, xPos, 1.7, 2.8, 1.6);
    s.addText(item.value, { x: xPos, y: 1.8, w: 2.8, h: 1.0, fontSize: 42, fontFace: FONT_EN, color: item.color, bold: true, align: "center", margin: 0 });
    s.addText(item.label, { x: xPos, y: 2.75, w: 2.8, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0 });
  });

  // 矢印
  s.addText("\u2192", { x: 3.4, y: 2.0, w: 0.6, h: 0.8, fontSize: 30, fontFace: FONT_EN, color: C.accent, align: "center", margin: 0 });
  s.addText("\u2192", { x: 6.5, y: 2.0, w: 0.6, h: 0.8, fontSize: 30, fontFace: FONT_EN, color: C.accent, align: "center", margin: 0 });

  // 予後予測因子
  addCard(s, 0.5, 3.6, 9, 1.6);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.6, w: 9, h: 0.4, fill: { color: C.primary } });
  s.addText("独立した予後予測因子", { x: 0.7, y: 3.6, w: 5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });

  s.addText("1. 入院時FIM（機能的自立度）  ← 最大の予測因子\n2. Charlson併存疾患スコア\n3. 初期症候群（めまい/嘔吐/失調/頭痛 → 良好転帰）\n4. SCA梗塞 → 予後不良", {
    x: 0.8, y: 4.1, w: 8.4, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: C.text, lineSpacingMultiple: 1.3, margin: 0,
  });
})();

// ============================================================
// SLIDE 6: SARAスケール
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "失調の評価：SARA（Scale for Assessment and Rating of Ataxia）");
  addPageNum(s, 6);

  // SARA概要
  addCard(s, 0.5, 1.1, 4.5, 0.8);
  s.addText("40点満点（0=正常 / 40=最重度失調）\n18点が姿勢・歩行関連項目", {
    x: 0.7, y: 1.15, w: 4.1, h: 0.7, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // カットオフ値
  addCard(s, 5.3, 1.1, 4.2, 0.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.1, w: 4.2, h: 0.35, fill: { color: C.red } });
  s.addText("臨床カットオフ値", { x: 5.5, y: 1.1, w: 3, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("\u2265 5.5: ADL依存 | \u2265 14.25: 重度依存", {
    x: 5.5, y: 1.55, w: 3.8, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });

  // SARA項目テーブル
  var rows = [
    ["項目", "配点", "評価内容"],
    ["歩行", "0-8", "歩行パターン、安定性"],
    ["立位", "0-6", "閉眼立位、タンデム立位"],
    ["座位", "0-4", "座位バランス"],
    ["言語障害", "0-6", "構音障害の程度"],
    ["指追い試験", "0-4", "上肢測定障害"],
    ["鼻指試験", "0-4", "企図振戦"],
    ["手の回内外", "0-4", "変換運動障害"],
    ["踵すね試験", "0-4", "下肢協調運動"],
  ];

  rows.forEach(function(row, i) {
    var yPos = 2.1 + i * 0.35;
    var bgColor = i === 0 ? C.primary : (i % 2 === 0 ? "F0F4FF" : C.white);
    var txtColor = i === 0 ? C.white : C.text;
    var isBold = i === 0;

    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9, h: 0.35, fill: { color: bgColor } });
    s.addText(row[0], { x: 0.7, y: yPos, w: 2.5, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: txtColor, bold: isBold, margin: 0 });
    s.addText(row[1], { x: 3.2, y: yPos, w: 1.5, h: 0.35, fontSize: 13, fontFace: FONT_EN, color: txtColor, bold: isBold, align: "center", margin: 0 });
    s.addText(row[2], { x: 4.7, y: yPos, w: 4.6, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: txtColor, bold: isBold, margin: 0 });
  });

  // BBS相関
  s.addText("BBS との相関: r = -0.946  →  SARAはNIHSSで捉えきれない後方循環障害を評価", {
    x: 0.5, y: 5.0, w: 9, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 7: 従来型リハビリの柱
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "従来型リハビリテーションの柱");
  addPageNum(s, 7);

  var methods = [
    { title: "Frenkel体操", desc: "視覚FBを用いた協調運動訓練", color: C.primary },
    { title: "PNF", desc: "固有受容性神経筋促通法\n体幹・骨盤の固有受容覚強化", color: C.accent },
    { title: "バランス訓練", desc: "静的→動的バランスの段階的負荷", color: C.green },
    { title: "歩行訓練", desc: "ロフストランドクラッチ使用\n重錘負荷訓練", color: C.orange },
    { title: "体重免荷TM", desc: "安全な環境での歩行練習\nトレッドミル訓練", color: C.red },
  ];

  methods.forEach(function(m, i) {
    var col = i < 3 ? 0 : 1;
    var row = i < 3 ? i : i - 3;
    var xPos = col === 0 ? 0.5 : 5.3;
    var yPos = 1.2 + row * 1.35;

    addCard(s, xPos, yPos, 4.2, 1.15);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 4.2, h: 0.35, fill: { color: m.color } });
    s.addText(m.title, { x: xPos + 0.15, y: yPos, w: 3.9, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
    s.addText(m.desc, { x: xPos + 0.15, y: yPos + 0.4, w: 3.9, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  // エビデンス注記
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.4, fill: { color: C.light } });
  s.addText("Marquer et al. 2014: バランス・協調訓練の集中的プログラムが有効（中等度エビデンス）", {
    x: 0.7, y: 5.0, w: 8.6, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.primary, margin: 0,
  });
})();

// ============================================================
// SLIDE 8: 姿勢障害 vs 歩行失調
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "回復パターン：姿勢障害 vs 歩行失調");
  addPageNum(s, 8);

  s.addText("Bultmann et al. Gait Posture 2014 | 急性小脳梗塞 23例 | 3ヶ月追跡", {
    x: 0.5, y: 1.1, w: 9, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // 姿勢障害カード
  addCard(s, 0.5, 1.7, 4.2, 2.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 4.2, h: 0.4, fill: { color: C.green } });
  s.addText("姿勢障害", { x: 0.7, y: 1.7, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("3ヶ月で\n完全回復", { x: 0.7, y: 2.3, w: 3.8, h: 1.0, fontSize: 28, fontFace: FONT_JP, color: C.green, bold: true, align: "center", margin: 0 });
  s.addText("\u2714 回復しやすい", { x: 0.7, y: 3.4, w: 3.8, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.green, bold: true, align: "center", margin: 0 });

  // 歩行失調カード
  addCard(s, 5.3, 1.7, 4.2, 2.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.7, w: 4.2, h: 0.4, fill: { color: C.orange } });
  s.addText("歩行失調", { x: 5.5, y: 1.7, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("軽度の失調が\n残存", { x: 5.5, y: 2.3, w: 3.8, h: 1.0, fontSize: 28, fontFace: FONT_JP, color: C.orange, bold: true, align: "center", margin: 0 });
  s.addText("\u26A0 特に歩行速度の低下", { x: 5.5, y: 3.4, w: 3.8, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true, align: "center", margin: 0 });

  // 追加情報
  addCard(s, 0.5, 4.2, 9, 1.0);
  s.addText("小脳小葉 IV-VI の病変 → 立位・歩行・下肢失調と関連\n梗塞体積が大きいほど失調が重度", {
    x: 0.7, y: 4.3, w: 8.6, h: 0.8, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });
})();

// ============================================================
// SLIDE 9: 訓練強度のエビデンス
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "訓練強度のエビデンス");
  addPageNum(s, 9);

  s.addText("Lubetzky-Vilnai & Kartin. J Neurol Phys Ther 2010 | 22研究のSR", {
    x: 0.5, y: 1.1, w: 9, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // 急性期
  addCard(s, 0.5, 1.7, 4.3, 1.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 4.3, h: 0.4, fill: { color: C.green } });
  s.addText("急性期", { x: 0.7, y: 1.7, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("週 2-3 回", { x: 0.7, y: 2.3, w: 3.9, h: 0.6, fontSize: 30, fontFace: FONT_JP, color: C.green, bold: true, align: "center", margin: 0 });
  s.addText("個別バランス訓練で十分な効果\n（中等度エビデンス）", { x: 0.7, y: 2.95, w: 3.9, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", margin: 0 });

  // 亜急性期/慢性期
  addCard(s, 5.2, 1.7, 4.3, 1.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.7, w: 4.3, h: 0.4, fill: { color: C.primary } });
  s.addText("亜急性期/慢性期", { x: 5.4, y: 1.7, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("個別 or グループ", { x: 5.4, y: 2.3, w: 3.9, h: 0.6, fontSize: 26, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0 });
  s.addText("個別訓練またはグループ療法\nが有効", { x: 5.4, y: 2.95, w: 3.9, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", margin: 0 });

  // 注意事項
  addCard(s, 0.5, 3.8, 9, 1.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.8, w: 9, h: 0.4, fill: { color: C.red } });
  s.addText("\u26A0 注意：過剰な訓練", { x: 0.7, y: 3.8, w: 5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("1日 90分以上 x 週5日 の訓練は逆効果の可能性\n\u2192 適切な強度設定（overwork weaknessの回避）が重要", {
    x: 0.7, y: 4.3, w: 8.6, h: 0.8, fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });
})();

// ============================================================
// SLIDE 10: NIBS概論
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "非侵襲的脳刺激療法（NIBS）：なぜ小脳？");
  addPageNum(s, 10);

  addCard(s, 0.5, 1.2, 9, 1.5);
  s.addText("小脳-視床-皮質路（cerebello-thalamo-cortical pathway）", {
    x: 0.7, y: 1.3, w: 8.6, h: 0.45, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("\u2022 運動制御・運動学習の中枢\n\u2022 皮質との豊富な結合 → 間接的な皮質可塑性修飾が可能\n\u2022 刺激標的として安全かつアクセスしやすい", {
    x: 0.7, y: 1.8, w: 8.6, h: 0.8, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // NIBS手法の比較
  var methods = [
    { name: "iTBS", full: "間欠的シータバースト刺激", desc: "TMS（経頭蓋磁気刺激）の一種\n短時間（約3分）で小脳興奮性を増加", color: C.primary },
    { name: "tDCS", full: "経頭蓋直流電気刺激", desc: "微弱直流電流（1-2mA）で\n神経活動を修飾（20分程度）", color: C.accent },
  ];

  methods.forEach(function(m, i) {
    var xPos = 0.5 + i * 4.7;
    addCard(s, xPos, 3.0, 4.3, 2.2);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.0, w: 4.3, h: 0.45, fill: { color: m.color } });
    s.addText(m.name, { x: xPos + 0.15, y: 3.0, w: 2, h: 0.45, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, margin: 0 });
    s.addText(m.full, { x: xPos + 0.15, y: 3.55, w: 4, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
    s.addText(m.desc, { x: xPos + 0.15, y: 3.95, w: 4, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  });
})();

// ============================================================
// SLIDE 11: Koch RCT — 小脳iTBSの歩行・バランス回復
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "小脳iTBSの歩行・バランス回復（Koch et al. JAMA Neurol 2019）");
  addPageNum(s, 11);

  s.addText("二重盲検RCT | 慢性MCA脳梗塞 34例 | 小脳iTBS+PT 3週間 vs sham+PT", {
    x: 0.5, y: 1.1, w: 9, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // BBS結果
  addCard(s, 0.5, 1.7, 6, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 6, h: 0.4, fill: { color: C.green } });
  s.addText("BBS（Berg Balance Scale）の変化  p < 0.001", { x: 0.7, y: 1.7, w: 5.6, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });

  var bbsData = [
    { label: "ベースライン", value: "34.5", color: C.red },
    { label: "3週後", value: "43.4", color: C.orange },
    { label: "6週後", value: "47.5", color: C.green },
  ];

  bbsData.forEach(function(d, i) {
    var xPos = 0.8 + i * 1.9;
    s.addText(d.value, { x: xPos, y: 2.3, w: 1.6, h: 0.7, fontSize: 32, fontFace: FONT_EN, color: d.color, bold: true, align: "center", margin: 0 });
    s.addText(d.label, { x: xPos, y: 2.95, w: 1.6, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0 });
    if (i < 2) {
      s.addText("\u2192", { x: xPos + 1.5, y: 2.3, w: 0.5, h: 0.7, fontSize: 24, fontFace: FONT_EN, color: C.accent, align: "center", margin: 0 });
    }
  });

  // Step width
  addCard(s, 6.8, 1.7, 2.7, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 1.7, w: 2.7, h: 0.4, fill: { color: C.accent } });
  s.addText("歩幅（cm）", { x: 7.0, y: 1.7, w: 2.3, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("16.8 \u2192 14.3", { x: 7.0, y: 2.3, w: 2.3, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", margin: 0 });
  s.addText("p < 0.05", { x: 7.0, y: 2.95, w: 2.3, h: 0.35, fontSize: 12, fontFace: FONT_EN, color: C.sub, align: "center", margin: 0 });

  // メカニズム
  addCard(s, 0.5, 4.0, 9, 1.2);
  s.addText("メカニズム", { x: 0.7, y: 4.1, w: 2, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("\u2022 小脳-皮質可塑性の促進（cerebello-cortical plasticity）\n\u2022 後頭頂皮質（PPC）の神経活動増加を確認\n\u2022 歩行自立度の向上 + 転倒リスク低減に寄与", {
    x: 0.7, y: 4.5, w: 8.6, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
  });
})();

// ============================================================
// SLIDE 12: M1 vs 小脳iTBS（Liao 2024）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "M1 vs 小脳iTBS（Liao et al. Stroke 2024）");
  addPageNum(s, 12);

  s.addText("二重盲検RCT | 亜急性脳卒中 36例 | M1群 vs 小脳群 vs sham群（各12例）", {
    x: 0.5, y: 1.1, w: 9, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // 結果テーブル
  var rows = [
    ["", "M1-iTBS", "小脳-iTBS", "Sham"],
    ["BBS改善", "+10.7*", "+14.2*", "基準"],
    ["FMA-LE改善", "NS", "+7.8*", "基準"],
  ];

  rows.forEach(function(row, i) {
    var yPos = 1.7 + i * 0.55;
    var bgColor = i === 0 ? C.primary : (i % 2 === 0 ? "F0F4FF" : C.white);
    var txtColor = i === 0 ? C.white : C.text;

    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9, h: 0.55, fill: { color: bgColor } });
    s.addText(row[0], { x: 0.7, y: yPos + 0.05, w: 2.5, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: txtColor, bold: true, margin: 0 });
    s.addText(row[1], { x: 3.2, y: yPos + 0.05, w: 2.2, h: 0.45, fontSize: 16, fontFace: FONT_EN, color: i === 0 ? txtColor : C.accent, bold: true, align: "center", margin: 0 });
    s.addText(row[2], { x: 5.4, y: yPos + 0.05, w: 2.2, h: 0.45, fontSize: 16, fontFace: FONT_EN, color: i === 0 ? txtColor : C.green, bold: true, align: "center", margin: 0 });
    s.addText(row[3], { x: 7.6, y: yPos + 0.05, w: 1.7, h: 0.45, fontSize: 16, fontFace: FONT_EN, color: txtColor, align: "center", margin: 0 });
  });

  s.addText("* p < 0.05 vs sham", { x: 0.5, y: 3.45, w: 3, h: 0.3, fontSize: 12, fontFace: FONT_EN, color: C.sub, margin: 0 });

  // キーメッセージ
  addCard(s, 0.5, 3.9, 9, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.9, w: 9, h: 0.4, fill: { color: C.green } });
  s.addText("Key Finding", { x: 0.7, y: 3.9, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, margin: 0 });
  s.addText("\u2022 バランス改善: M1-iTBSと小脳-iTBSの両方で有効\n\u2022 下肢運動回復: 小脳-iTBSのみが有意な効果\n  → 小脳iTBSは歩行バランスに加え運動回復にも有効\n  → M1より優位な可能性", {
    x: 0.7, y: 4.4, w: 8.6, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });
})();

// ============================================================
// SLIDE 13: メタアナリシス（Jiang 2025）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "小脳NIBSのメタアナリシス（Jiang et al. 2025）");
  addPageNum(s, 13);

  s.addText("14 RCT / 382例 | Eur J Phys Rehabil Med", {
    x: 0.5, y: 1.1, w: 9, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // 結果テーブル
  var headers = ["評価指標", "効果量", "95% CI", "p値", "確実性"];
  var data = [
    ["バランス（BBS）", "MD = 4.17", "2.28 - 6.05", "< 0.01", "低"],
    ["歩行速度", "SMD = -0.36", "-0.68 to -0.03", "0.03", "中等度"],
    ["移動性（TUG）", "MD = -3.34", "-5.14 to -1.54", "< 0.01", "低"],
  ];

  headers.forEach(function(h, i) {
    var widths = [2.2, 1.8, 2.2, 1.2, 1.2];
    var xPos = 0.5;
    for (var j = 0; j < i; j++) { xPos += widths[j]; }
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.7, w: widths[i], h: 0.5, fill: { color: C.primary } });
    s.addText(h, { x: xPos + 0.1, y: 1.7, w: widths[i] - 0.2, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  data.forEach(function(row, i) {
    var yPos = 2.2 + i * 0.55;
    var bgColor = i % 2 === 0 ? C.white : "F0F4FF";
    var widths = [2.2, 1.8, 2.2, 1.2, 1.2];

    row.forEach(function(cell, j) {
      var xPos = 0.5;
      for (var k = 0; k < j; k++) { xPos += widths[k]; }
      s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: widths[j], h: 0.55, fill: { color: bgColor } });
      var cellColor = j === 1 ? C.red : C.text;
      s.addText(cell, { x: xPos + 0.1, y: yPos, w: widths[j] - 0.2, h: 0.55, fontSize: 14, fontFace: j >= 1 && j <= 3 ? FONT_EN : FONT_JP, color: cellColor, bold: j === 1, align: "center", valign: "middle", margin: 0 });
    });
  });

  // 結論
  addCard(s, 0.5, 4.1, 9, 1.2);
  s.addText("\u2022 小脳NIBS（tDCS + TMS）はバランス・歩行・機能的移動性を改善\n\u2022 エビデンスの確実性は低〜中等度\n\u2022 今後の課題：最適な刺激プロトコルの確立\n\u2022 行動介入（PT）との併用が効果を最大化", {
    x: 0.7, y: 4.2, w: 8.6, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });
})();

// ============================================================
// SLIDE 14: 上肢tDCS（Gong 2023）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "小脳tDCSの上肢運動機能回復（Gong et al. 2023）");
  addPageNum(s, 14);

  s.addText("二重盲検RCT | 脳卒中 77例 | 陽極 2mA/20分 x 4週間 vs sham", {
    x: 0.5, y: 1.1, w: 9, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // FMA-UE結果
  var results = [
    { period: "4週後", tdcs: "+10.7", sham: "+5.8", diff: "4.9", p: "0.013" },
    { period: "60日後", tdcs: "+18.9", sham: "+12.7", diff: "6.2", p: "0.043" },
  ];

  results.forEach(function(r, i) {
    var yPos = 1.7 + i * 1.3;
    addCard(s, 0.5, yPos, 9, 1.1);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 1.8, h: 0.4, fill: { color: C.primary } });
    s.addText(r.period, { x: 0.5, y: yPos, w: 1.8, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });

    s.addText("tDCS", { x: 2.5, y: yPos + 0.05, w: 1, h: 0.3, fontSize: 12, fontFace: FONT_EN, color: C.sub, margin: 0 });
    s.addText(r.tdcs, { x: 2.5, y: yPos + 0.35, w: 1.5, h: 0.5, fontSize: 24, fontFace: FONT_EN, color: C.green, bold: true, margin: 0 });

    s.addText("vs", { x: 4.0, y: yPos + 0.35, w: 0.5, h: 0.5, fontSize: 14, fontFace: FONT_EN, color: C.sub, align: "center", margin: 0 });

    s.addText("Sham", { x: 4.5, y: yPos + 0.05, w: 1, h: 0.3, fontSize: 12, fontFace: FONT_EN, color: C.sub, margin: 0 });
    s.addText(r.sham, { x: 4.5, y: yPos + 0.35, w: 1.5, h: 0.5, fontSize: 24, fontFace: FONT_EN, color: C.orange, bold: true, margin: 0 });

    s.addText("\u0394 = " + r.diff + "  p = " + r.p, { x: 6.5, y: yPos + 0.3, w: 2.8, h: 0.5, fontSize: 16, fontFace: FONT_EN, color: C.red, bold: true, margin: 0 });
  });

  // MCID率
  addCard(s, 0.5, 4.5, 9, 0.9);
  s.addText("臨床的有意改善（MCID）達成率", { x: 0.7, y: 4.55, w: 4, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("tDCS群: 70.3%  vs  Sham群: 34.3%   (p = 0.002)", {
    x: 0.7, y: 4.95, w: 8.6, h: 0.35, fontSize: 18, fontFace: FONT_EN, color: C.red, bold: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 15: 小脳DBS — Nature Medicine 2023
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "小脳深部脳刺激 DBS（Baker et al. Nat Med 2023）");
  addPageNum(s, 15);

  s.addText("Phase I / 非盲検 / 慢性期（1-3年後）中等度〜重度上肢障害 12例", {
    x: 0.5, y: 1.1, w: 9, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // 方法
  addCard(s, 0.5, 1.7, 4.3, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 4.3, h: 0.4, fill: { color: C.primary } });
  s.addText("方法", { x: 0.7, y: 1.7, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("小脳歯状核（dentate nucleus）に\n電極留置 + リハビリテーション併用", { x: 0.7, y: 2.2, w: 3.9, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0 });

  // 結果
  addCard(s, 5.2, 1.7, 4.3, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.7, w: 4.3, h: 0.4, fill: { color: C.green } });
  s.addText("結果", { x: 5.4, y: 1.7, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("FMA-UE 中央値 7点改善（全体）\n遠位運動保存例: 15点改善（全例MCID超）", { x: 5.4, y: 2.2, w: 3.9, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0 });

  // Key findings
  addCard(s, 0.5, 3.2, 9, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.2, w: 9, h: 0.4, fill: { color: C.accent } });
  s.addText("画期的な知見", { x: 0.7, y: 3.2, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("\u2022 慢性期（1-3年後）の「回復プラトー」を打破\n\u2022 機能改善は同側皮質代謝増加と有意に相関\n  → late-stage neuroplasticity の修飾が可能\n\u2022 重篤な周術期・刺激関連有害事象なし\n\u2022 Phase II/III試験の結果が待たれる", {
    x: 0.7, y: 3.7, w: 8.6, h: 1.4, fontSize: 15, fontFace: FONT_JP, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });
})();

// ============================================================
// SLIDE 16: VRリハビリ — 日本からの報告
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "VRリハビリテーション（Takimoto et al. BMJ Case Rep 2021）");
  addPageNum(s, 16);

  s.addText("岸和田リハビリテーション病院 | 右小脳・脳幹梗塞後失調 | 40代男性", {
    x: 0.5, y: 1.1, w: 9, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  // 経過
  addCard(s, 0.5, 1.7, 4.3, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 4.3, h: 0.4, fill: { color: C.orange } });
  s.addText("従来型PT 3週間", { x: 0.7, y: 1.7, w: 3, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("FIM: 101 \u2192 124\nADLは改善\nしかし高次バランスはプラトー\n（フォークリフト復職に不十分）", { x: 0.7, y: 2.2, w: 3.9, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });

  addCard(s, 5.2, 1.7, 4.3, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.7, w: 4.3, h: 0.4, fill: { color: C.green } });
  s.addText("VRバランス訓練 2週間", { x: 5.4, y: 1.7, w: 3, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("平日40分のVR介入で：\nSARA: 5 \u2192 1\nFBS: 48 \u2192 56\nMini-BESTest: 20 \u2192 28", { x: 5.4, y: 2.2, w: 3.9, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });

  // 転帰
  addCard(s, 0.5, 3.5, 9, 1.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9, h: 0.4, fill: { color: C.primary } });
  s.addText("転帰", { x: 0.7, y: 3.5, w: 2, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
  s.addText("体幹動揺消失 \u2192 自信回復 \u2192 フォークリフト運転手として復職\n従来療法でプラトーに達した症例にVRが追加効果を発揮", {
    x: 0.7, y: 4.0, w: 8.6, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.8, w: 9, h: 0.5, fill: { color: C.light } });
  s.addText("\u2192 VRリハビリは、従来療法で改善が頭打ちとなった小脳失調患者への有力な選択肢", {
    x: 0.7, y: 4.85, w: 8.6, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 17: まとめ — 5つの臨床ポイント
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.2, w: 9, h: 0.6,
    fontSize: 28, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var points = [
    { num: "1", text: "回復は12ヶ月まで続く ― 機能ドメインで回復時期が異なる", color: C.green },
    { num: "2", text: "姿勢障害は回復しやすい、歩行失調は残りやすい", color: C.orange },
    { num: "3", text: "SARAで評価 ― NIHSSだけでは後方循環障害を過小評価", color: C.accent },
    { num: "4", text: "小脳NIBS（iTBS/tDCS）は有望な補助療法", color: C.primary },
    { num: "5", text: "適切な訓練強度（急性期: 週2-3回）― 過剰訓練を避ける", color: C.red },
  ];

  points.forEach(function(p, i) {
    var yPos = 1.0 + i * 0.85;
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.05, w: 0.5, h: 0.5, fill: { color: p.color } });
    s.addText(p.num, { x: 0.7, y: yPos + 0.05, w: 0.5, h: 0.5, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.text, { x: 1.4, y: yPos, w: 8, h: 0.6, fontSize: 18, fontFace: FONT_JP, color: C.light, margin: 0 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.1, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司", {
    x: 0.5, y: 5.15, w: 9, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });
  addPageNum(s, 17);
})();

// ============================================================
// Output
// ============================================================
var outPath = __dirname + "/cerebellar_stroke_rehab_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
