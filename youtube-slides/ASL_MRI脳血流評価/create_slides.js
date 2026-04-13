var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "ASL MRI ― 造影剤なしの脳血流評価を徹底解説｜PLD設定・1.5T vs 3T・Multi-PLD";

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

  s.addText("脳神経内科医が解説する MRI講座", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("ASL MRI\n造影剤なしの脳血流評価", {
    x: 0.8, y: 1.1, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("PLD設定 / 1.5T vs 3T / Multi-PLD / 脳梗塞 / てんかん", {
    x: 0.8, y: 2.9, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
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
// SLIDE 2: ASLとは ― 3つの利点
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ASLとは ― 造影剤なしの脳灌流画像");

  // 左側: 説明テキスト
  addCard(s, 0.4, 1.1, 4.4, 2.0);
  s.addText("ASL = Arterial Spin Labeling", {
    x: 0.6, y: 1.2, w: 4.0, h: 0.5,
    fontSize: 18, fontFace: FONT_EN, color: C.primary, bold: true, margin: 0,
  });
  s.addText("MRIの高周波パルスで動脈血中の\nプロトンを磁気的に標識し、\n脳組織への血流を画像化する方法", {
    x: 0.6, y: 1.7, w: 4.0, h: 1.2,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });

  // 右側: 3つの利点カード
  var benefits = [
    { icon: "💉", label: "造影剤不要", desc: "腎機能障害でも安全" },
    { icon: "☢️", label: "被曝ゼロ", desc: "繰り返し撮像が可能" },
    { icon: "⏱️", label: "数分で撮像", desc: "ルーチンに追加しやすい" },
  ];
  benefits.forEach(function(b, i) {
    var yPos = 1.1 + i * 0.7;
    addCard(s, 5.2, yPos, 4.4, 0.6);
    s.addText(b.icon, { x: 5.4, y: yPos + 0.05, w: 0.5, h: 0.5, fontSize: 22, margin: 0, valign: "middle" });
    s.addText(b.label, { x: 5.9, y: yPos + 0.05, w: 1.8, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0, valign: "middle" });
    s.addText(b.desc, { x: 7.7, y: yPos + 0.05, w: 1.8, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0, valign: "middle" });
  });

  // 下部: SPECT/CTp比較
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.4, w: 9.2, h: 0.5, fill: { color: C.light } });
  s.addText("SPECT（放射性トレーサー必要）や CT perfusion（造影剤+被曝）に対する大きな利点", {
    x: 0.4, y: 3.4, w: 9.2, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // pCASLが推奨
  addCard(s, 0.4, 4.1, 9.2, 0.9);
  s.addText("現在の臨床標準: pCASL（pseudo-Continuous ASL）", {
    x: 0.6, y: 4.15, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("2015年ISMRM合意で推奨 ― SNRと定量精度のバランスが最良", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  addPageNum(s, "2/12");
})();

// ============================================================
// SLIDE 3: ASLの原理 ― 3ステップ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ASLの原理 ― 3つのステップ");

  var steps = [
    { num: "1", title: "ラベリング（標識）", desc: "頸部で動脈血のプロトンに\n高周波パルスを照射し\nスピンを反転させる", color: C.primary },
    { num: "2", title: "待機（PLD）", desc: "標識された血液が\n脳実質に到達するまで\n一定時間待つ", color: C.accent },
    { num: "3", title: "撮像・差分計算", desc: "Label画像とControl画像の\n差分を計算して\n灌流画像を得る", color: C.green },
  ];

  steps.forEach(function(st, i) {
    var xPos = 0.4 + i * 3.2;
    // 番号サークル
    s.addShape(pres.shapes.OVAL, { x: xPos + 1.1, y: 1.1, w: 0.7, h: 0.7, fill: { color: st.color } });
    s.addText(st.num, { x: xPos + 1.1, y: 1.1, w: 0.7, h: 0.7, fontSize: 24, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // カード
    addCard(s, xPos, 2.0, 2.9, 2.6);
    s.addText(st.title, { x: xPos + 0.2, y: 2.1, w: 2.5, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: st.color, bold: true, align: "center", margin: 0 });
    s.addText(st.desc, { x: xPos + 0.2, y: 2.7, w: 2.5, h: 1.7, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.4 });
  });

  // 矢印（ステップ間）
  s.addText("→", { x: 3.2, y: 1.15, w: 0.5, h: 0.6, fontSize: 28, fontFace: FONT_EN, color: C.sub, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 6.4, y: 1.15, w: 0.5, h: 0.6, fontSize: 28, fontFace: FONT_EN, color: C.sub, align: "center", valign: "middle", margin: 0 });

  // 注釈
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.8, w: 9.2, h: 0.5, fill: { color: C.light } });
  s.addText("差分信号は脳組織信号のわずか 0.5〜1.5% → 十分な加算が必要", {
    x: 0.4, y: 4.8, w: 9.2, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "3/12");
})();

// ============================================================
// SLIDE 4: PLD（Post Labeling Delay）の設定
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "PLD（Post Labeling Delay）― ASL最重要パラメータ");

  // PLD定義
  addCard(s, 0.4, 1.1, 4.4, 1.6);
  s.addText("PLDとは", {
    x: 0.6, y: 1.15, w: 4.0, h: 0.35,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("ラベリング終了から撮像開始までの待ち時間。\n標識血液が脳実質に到達するために\n必要な時間を確保する。", {
    x: 0.6, y: 1.5, w: 4.0, h: 1.1,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });

  // PLD短すぎ
  addCard(s, 5.2, 1.1, 4.4, 0.75);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 0.15, h: 0.75, fill: { color: C.red } });
  s.addText("短すぎ: Bright Vessel Artifact", {
    x: 5.5, y: 1.13, w: 4.0, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, margin: 0, valign: "middle",
  });
  s.addText("標識血液が血管内に滞留、灌流過小評価", {
    x: 5.5, y: 1.48, w: 4.0, h: 0.32,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle",
  });

  // PLD長すぎ
  addCard(s, 5.2, 2.0, 4.4, 0.75);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.0, w: 0.15, h: 0.75, fill: { color: C.orange } });
  s.addText("長すぎ: SNR低下", {
    x: 5.5, y: 2.03, w: 4.0, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0, valign: "middle",
  });
  s.addText("T1緩和で標識信号が減衰、画質劣化", {
    x: 5.5, y: 2.38, w: 4.0, h: 0.32,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle",
  });

  // 推奨PLD表
  s.addText("2015年 ISMRM合意 推奨PLD（pCASL）", {
    x: 0.4, y: 2.95, w: 9.2, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  var tblRows = [
    [{ text: "対象", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
     { text: "推奨PLD", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
     { text: "LD", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } }],
    ["小児", "1500 ms", "1800 ms"],
    ["健常成人（<70歳）", "1800 ms", "1800 ms"],
    ["健常成人（≧70歳）", "2000 ms", "1800 ms"],
    [{ text: "臨床患者（成人）", options: { bold: true, color: C.primary } },
     { text: "2000 ms", options: { bold: true, color: C.red } },
     "1800 ms"],
  ];

  s.addTable(tblRows, {
    x: 0.4, y: 3.35, w: 9.2,
    fontSize: 13, fontFace: FONT_JP, color: C.text,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    rowH: [0.35, 0.32, 0.32, 0.32, 0.38],
    colW: [3.5, 3.0, 2.7],
    align: "center",
    valign: "middle",
  });

  addPageNum(s, "4/12");
})();

// ============================================================
// SLIDE 5: 1.5T vs 3T
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "1.5T vs 3T ― 磁場強度の影響");

  // メインメッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 9.2, h: 0.5, fill: { color: C.light } });
  s.addText("ISMRM合意: 「3Tが利用可能な場合は3Tを推奨」", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 比較表
  var tblRows = [
    [{ text: "パラメータ", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
     { text: "1.5T", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
     { text: "3T", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } }],
    ["血液T1値", "約1350 ms", { text: "約1650 ms", options: { bold: true, color: C.green } }],
    ["SNR", "基準", { text: "約1.5〜2倍", options: { bold: true, color: C.green } }],
    ["ラベリング効率", "やや低い", { text: "高い・安定", options: { bold: true, color: C.green } }],
    ["SAR（比吸収率）", { text: "低い", options: { color: C.green } }, { text: "高い（要注意）", options: { color: C.orange } }],
    ["磁場不均一性", { text: "小さい", options: { color: C.green } }, { text: "大きい（頭蓋底注意）", options: { color: C.orange } }],
    ["撮像時間", "長め", { text: "短縮可能", options: { bold: true, color: C.green } }],
  ];

  s.addTable(tblRows, {
    x: 0.4, y: 1.8, w: 9.2,
    fontSize: 13, fontFace: FONT_JP, color: C.text,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
    colW: [3.0, 3.1, 3.1],
    align: "center",
    valign: "middle",
  });

  // 1.5Tでもできる
  addCard(s, 0.4, 4.35, 9.2, 0.9);
  s.addText("1.5Tでも使える！", {
    x: 0.6, y: 4.4, w: 2.5, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.green, bold: true, margin: 0,
  });
  s.addText("加算回数を増やす（撮像5〜6分）/ Background suppression最適化 / 3D readout使用", {
    x: 3.2, y: 4.4, w: 6.2, h: 0.7,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  addPageNum(s, "5/12");
})();

// ============================================================
// SLIDE 6: Single-PLD vs Multi-PLD
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "Single-PLD vs Multi-PLD ― 最新トレンド");

  // Single-PLD
  addCard(s, 0.4, 1.1, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 4.4, h: 0.45, fill: { color: C.orange } });
  s.addText("Single-PLD", {
    x: 0.4, y: 1.1, w: 4.4, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("1つの固定PLDで撮像", {
    x: 0.6, y: 1.6, w: 4.0, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  var singleItems = [
    { mark: "○", text: "シンプル・撮像時間短い", col: C.green },
    { mark: "○", text: "読影が容易", col: C.green },
    { mark: "×", text: "ATTの影響を受けやすい", col: C.red },
    { mark: "×", text: "側副路の遅い血流を見逃す", col: C.red },
  ];
  singleItems.forEach(function(item, i) {
    s.addText(item.mark + " " + item.text, {
      x: 0.6, y: 1.95 + i * 0.35, w: 4.0, h: 0.32,
      fontSize: 13, fontFace: FONT_JP, color: item.col, margin: 0, valign: "middle",
    });
  });

  // Multi-PLD
  addCard(s, 5.2, 1.1, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.4, h: 0.45, fill: { color: C.green } });
  s.addText("Multi-PLD", {
    x: 5.2, y: 1.1, w: 4.4, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("複数のPLDで撮像", {
    x: 5.4, y: 1.6, w: 4.0, h: 0.3,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  var multiItems = [
    { mark: "○", text: "CBF + ATT 同時定量", col: C.green },
    { mark: "○", text: "ATA軽減、定量精度向上", col: C.green },
    { mark: "×", text: "撮像時間やや長い（5〜7分）", col: C.red },
    { mark: "×", text: "解析ソフトが必要", col: C.red },
  ];
  multiItems.forEach(function(item, i) {
    s.addText(item.mark + " " + item.text, {
      x: 5.4, y: 1.95 + i * 0.35, w: 4.0, h: 0.32,
      fontSize: 13, fontFace: FONT_JP, color: item.col, margin: 0, valign: "middle",
    });
  });

  // Multi-PLD推奨プロトコル
  s.addText("Multi-PLD 推奨プロトコル例（Woods et al. 2024）", {
    x: 0.4, y: 3.8, w: 9.2, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  var tblRows = [
    [{ text: "", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
     { text: "若年健常者", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } },
     { text: "高齢者・臨床患者", options: { bold: true, color: "FFFFFF", fill: { color: C.primary } } }],
    ["ATTカバー", "0.5〜1.8 s", "0.5〜2.5 s"],
    ["PLD数", "4〜8ポイント", "4〜8ポイント"],
    ["代表的PLD", "0.5/1.0/1.5/2.0 s", "0.5/1.0/1.5/2.0/2.5/3.0 s"],
  ];

  s.addTable(tblRows, {
    x: 0.4, y: 4.2, w: 9.2,
    fontSize: 12, fontFace: FONT_JP, color: C.text,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" },
    rowH: [0.3, 0.28, 0.28, 0.28],
    colW: [1.8, 3.7, 3.7],
    align: "center",
    valign: "middle",
  });

  addPageNum(s, "6/12");
})();

// ============================================================
// SLIDE 7: 脳梗塞でのASL① ― DWIより早い？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳梗塞でのASL ― DWIより早い？");

  // タイムライン
  addCard(s, 0.4, 1.1, 9.2, 1.5);
  s.addText("発症からの時間軸", {
    x: 0.6, y: 1.15, w: 3.0, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  // タイムラインバー
  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 1.8, w: 8.0, h: 0.15, fill: { color: "DEE2E6" } });
  // ASL検出
  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 1.8, w: 2.0, h: 0.15, fill: { color: C.red } });
  s.addText("ASL: 灌流低下", { x: 0.8, y: 2.0, w: 2.5, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText("発症直後〜", { x: 0.8, y: 2.25, w: 2.5, h: 0.25, fontSize: 11, fontFace: FONT_JP, color: C.sub, margin: 0 });

  // DWI検出
  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 1.8, w: 2.0, h: 0.15, fill: { color: C.orange } });
  s.addText("DWI: 高信号", { x: 3.3, y: 2.0, w: 2.5, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0 });
  s.addText("〜30分以降", { x: 3.3, y: 2.25, w: 2.5, h: 0.25, fontSize: 11, fontFace: FONT_JP, color: C.sub, margin: 0 });

  // ASAP-ASL
  addCard(s, 0.4, 2.8, 4.4, 2.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.8, w: 4.4, h: 0.4, fill: { color: C.red } });
  s.addText("ASAP-ASL（超急性期プロトコル）", {
    x: 0.4, y: 2.8, w: 4.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  var asapItems = [
    "撮像時間: 1分17秒",
    "PLD: 1232 ms（短縮）",
    "撮像範囲: 基底核レベル限定",
    "目的: 3分で灌流評価",
    "   → 血栓回収の適応判断",
  ];
  asapItems.forEach(function(item, i) {
    s.addText("• " + item, {
      x: 0.6, y: 3.3 + i * 0.33, w: 4.0, h: 0.3,
      fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle",
    });
  });

  // OTAL法
  addCard(s, 5.2, 2.8, 4.4, 2.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.8, w: 4.4, h: 0.4, fill: { color: C.primary } });
  s.addText("OTAL法（2相ASL）", {
    x: 5.2, y: 2.8, w: 4.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("第1相: PLD 1200 ms\n  → 主幹動脈閉塞の検出\n\n第2相: PLD 2200 ms\n  → 側副路経由の遅い血流検出\n\n約90%の患者で十分な灌流評価", {
    x: 5.4, y: 3.3, w: 4.0, h: 1.6,
    fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2,
  });

  addPageNum(s, "7/12");
})();

// ============================================================
// SLIDE 8: 脳梗塞でのASL② ― 読影ポイント
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳梗塞ASL ― 4つの読影パターン");

  var patterns = [
    { title: "灌流低下（DWI陽性域一致）", desc: "梗塞コアの灌流障害", color: C.red, label: "1" },
    { title: "灌流低下（DWI陰性域）", desc: "ペナンブラの可能性 → 治療介入の根拠", color: C.orange, label: "2" },
    { title: "Bright Vessel（高信号血管）", desc: "側副路の存在を示唆。灌流正常と誤読しない！", color: C.yellow, label: "3" },
    { title: "再灌流後の高灌流", desc: "tPA/血栓回収後の再開通。出血性変化リスク評価", color: C.green, label: "4" },
  ];

  patterns.forEach(function(p, i) {
    var yPos = 1.1 + i * 0.98;
    addCard(s, 0.4, yPos, 9.2, 0.85);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 0.15, h: 0.85, fill: { color: p.color } });
    // 番号サークル
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.18, w: 0.5, h: 0.5, fill: { color: p.color } });
    s.addText(p.label, { x: 0.7, y: yPos + 0.18, w: 0.5, h: 0.5, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.title, { x: 1.4, y: yPos + 0.1, w: 3.6, h: 0.65, fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, margin: 0, valign: "middle" });
    s.addText(p.desc, { x: 5.1, y: yPos + 0.1, w: 4.3, h: 0.65, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0, valign: "middle" });
  });

  addPageNum(s, "8/12");
})();

// ============================================================
// SLIDE 9: てんかんでのASL
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "てんかんでのASL ― SPECTに代わる非侵襲評価");

  // 3カード
  var states = [
    { title: "発作時（ictal）", finding: "焦点の高灌流", meaning: "焦点の局在同定に有用\nSPECTの代替として期待", color: C.red },
    { title: "発作直後（postictal）", finding: "高灌流→低灌流へ移行", meaning: "タイミング依存\n撮像時の状態把握が重要", color: C.orange },
    { title: "発作間欠期（interictal）", finding: "焦点の軽度低灌流", meaning: "構造MRI正常例で\n補助情報として価値", color: C.primary },
  ];

  states.forEach(function(st, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.1, 3.0, 2.6);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 3.0, h: 0.45, fill: { color: st.color } });
    s.addText(st.title, { x: xPos, y: 1.1, w: 3.0, h: 0.45, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText("ASL所見:", { x: xPos + 0.2, y: 1.7, w: 2.6, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.sub, margin: 0, valign: "middle" });
    s.addText(st.finding, { x: xPos + 0.2, y: 2.0, w: 2.6, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: st.color, bold: true, margin: 0, valign: "middle" });
    s.addText(st.meaning, { x: xPos + 0.2, y: 2.5, w: 2.6, h: 1.0, fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4 });
  });

  // SE評価
  addCard(s, 0.3, 3.95, 9.4, 1.2);
  s.addText("てんかん重積状態（SE）の評価", {
    x: 0.5, y: 4.0, w: 5.0, h: 0.35,
    fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });
  var seItems = [
    "SEと自然消退する発作の鑑別にASLは最も感度が高いとの報告",
    "非けいれん性SE（NCSE）でも高灌流パターンで診断補助",
    "高灌流消失 = SE離脱の客観的指標",
  ];
  seItems.forEach(function(item, i) {
    s.addText("• " + item, {
      x: 0.5, y: 4.38 + i * 0.25, w: 9.0, h: 0.23,
      fontSize: 11, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle",
    });
  });

  addPageNum(s, "9/12");
})();

// ============================================================
// SLIDE 10: 脳梗塞 vs てんかんの鑑別
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳梗塞 vs てんかん ― ASLで鑑別のヒント");

  s.addText("急性片麻痺・失語 → 脳梗塞？ Todd麻痺？", {
    x: 0.4, y: 1.15, w: 9.2, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.text, bold: true, align: "center", margin: 0,
  });

  // 左: 脳梗塞
  addCard(s, 0.4, 1.8, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.8, w: 4.4, h: 0.45, fill: { color: C.red } });
  s.addText("脳梗塞を示唆", {
    x: 0.4, y: 1.8, w: 4.4, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("ASL: 灌流低下 ⬇️\nDWI: 高信号（拡散制限）\n\n→ 虚血性脳血管障害\n   血栓溶解・回収の適応を検討", {
    x: 0.6, y: 2.4, w: 4.0, h: 1.6,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });

  // VS
  s.addShape(pres.shapes.OVAL, { x: 4.55, y: 2.7, w: 0.9, h: 0.9, fill: { color: C.dark } });
  s.addText("VS", { x: 4.55, y: 2.7, w: 0.9, h: 0.9, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // 右: てんかん
  addCard(s, 5.2, 1.8, 4.4, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.8, w: 4.4, h: 0.45, fill: { color: C.green } });
  s.addText("てんかん（Todd麻痺）を示唆", {
    x: 5.2, y: 1.8, w: 4.4, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addText("ASL: 高灌流 ⬆️\nDWI: 正常〜軽度変化\n\n→ てんかん発作後の一過性症状\n   抗てんかん薬の調整を検討", {
    x: 5.4, y: 2.4, w: 4.0, h: 1.6,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4,
  });

  // 注釈
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.5, w: 9.2, h: 0.7, fill: { color: C.light } });
  s.addText("DWI + ASL の併用で鑑別精度が向上\nただし ASL 単独での確定診断は避け、臨床情報と総合的に判断する", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.6,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
  });

  addPageNum(s, "10/12");
})();

// ============================================================
// SLIDE 11: ピットフォール
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ASL読影の5つのピットフォール");

  var pitfalls = [
    { num: "1", title: "Bright Vessel Artifact", desc: "血管内高信号 ≠ 灌流正常\n側副路の遅い血流を反映", color: C.red },
    { num: "2", title: "Border Zone Artifact", desc: "分水嶺領域はATTが長く\n灌流過小評価されやすい", color: C.orange },
    { num: "3", title: "ラベリング面の設定不良", desc: "Willis動脈輪に近すぎると\nラベリング効率が低下", color: C.orange },
    { num: "4", title: "年齢によるATT変化", desc: "高齢者はATT延長\n→ PLD延長が必要", color: C.yellow },
    { num: "5", title: "動きアーチファクト", desc: "Label/Control間の頭部移動で\n差分画像に大きなノイズ", color: C.yellow },
  ];

  pitfalls.forEach(function(p, i) {
    var yPos = 1.1 + i * 0.82;
    addCard(s, 0.4, yPos, 9.2, 0.72);
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: yPos + 0.11, w: 0.5, h: 0.5, fill: { color: p.color } });
    s.addText(p.num, { x: 0.6, y: yPos + 0.11, w: 0.5, h: 0.5, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.title, { x: 1.3, y: yPos + 0.05, w: 3.0, h: 0.62, fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, margin: 0, valign: "middle" });
    s.addText(p.desc, { x: 4.4, y: yPos + 0.05, w: 5.0, h: 0.62, fontSize: 12, fontFace: FONT_JP, color: C.sub, margin: 0, valign: "middle", lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "11/12");
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "まとめ ― 明日からのASL活用チェックリスト");

  var items = [
    { icon: "✅", text: "撮像法: pCASL を第一選択（ISMRM合意推奨）" },
    { icon: "✅", text: "PLD設定: 成人臨床 → PLD 2000 ms / LD 1800 ms を基本" },
    { icon: "✅", text: "磁場強度: 3Tが望ましいが 1.5Tでも実用可能（撮像時間延長）" },
    { icon: "✅", text: "Multi-PLD: 脳血管疾患が多い施設では ATT情報の恩恵大" },
    { icon: "✅", text: "脳梗塞: DWI + ASL で ペナンブラ評価・側副路評価" },
    { icon: "✅", text: "てんかん: 撮像時の臨床状態を記録し灌流所見と対照" },
    { icon: "⚠️", text: "ピットフォール: ATA誤読・高齢者PLD延長を忘れない" },
  ];

  items.forEach(function(item, i) {
    var yPos = 1.1 + i * 0.58;
    addCard(s, 0.4, yPos, 9.2, 0.5);
    s.addText(item.icon, { x: 0.6, y: yPos + 0.02, w: 0.5, h: 0.46, fontSize: 18, margin: 0, valign: "middle" });
    s.addText(item.text, { x: 1.2, y: yPos + 0.02, w: 8.2, h: 0.46, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle" });
  });

  // 最後のメッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.0, w: 9.2, h: 0.5, fill: { color: C.dark } });
  s.addText("造影剤なし・被曝なし・繰り返し可能 ― ルーチンMRIに数分追加するだけ", {
    x: 0.4, y: 5.0, w: 9.2, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "12/12");
})();

// --- 出力 ---
var outPath = __dirname + "/ASL_MRI脳血流評価_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
