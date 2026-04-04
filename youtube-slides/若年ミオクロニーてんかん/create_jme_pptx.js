var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "若年ミオクロニーてんかん（JME）― 見逃されやすい思春期てんかんの診断と治療";

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
// SLIDE 1: Title
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("若年ミオクロニーてんかん", {
    x: 0.8, y: 0.8, w: 8.4, h: 1.2,
    fontSize: 40, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("Juvenile Myoclonic Epilepsy (JME)", {
    x: 0.8, y: 1.9, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_EN, color: C.accent, align: "center", italic: true, margin: 0,
  });
  s.addText("― 見逃されやすい思春期てんかんの診断と治療 ―", {
    x: 0.8, y: 2.6, w: 8.4, h: 0.7,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 3.5, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.8, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("2025年最新メタアナリシス・コホート研究に基づく", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 疫学
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "疫学 ― 思ったより多い「よくあるてんかん」");

  // Key stats cards
  var stats = [
    { num: "5〜10%", desc: "全てんかんに\n占める割合" },
    { num: "0.4/1000", desc: "成人での\n有病率" },
    { num: "12〜18歳", desc: "発症年齢の\nピーク" },
    { num: "~50%", desc: "家族歴\nあり" },
  ];
  stats.forEach(function(item, i) {
    var xPos = 0.3 + i * 2.4;
    addCard(s, xPos, 1.2, 2.2, 1.8);
    s.addText(item.num, { x: xPos, y: 1.3, w: 2.2, h: 0.8, fontSize: 24, fontFace: FONT_EN, color: C.primary, bold: true, align: "center", margin: 0 });
    s.addText(item.desc, { x: xPos, y: 2.1, w: 2.2, h: 0.8, fontSize: 14, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
  });

  // Genetics
  addCard(s, 0.5, 3.3, 9, 2.1);
  s.addText("遺伝学的背景", { x: 0.8, y: 3.4, w: 3, h: 0.5, fontSize: 18, fontFace: FONT_H, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "遺伝形式: ", options: { bold: true, color: C.text } },
    { text: "多因子遺伝（polygenic）が主体", options: { color: C.text, breakLine: true } },
    { text: "GWAS感受性座位: ", options: { bold: true, color: C.text } },
    { text: "4p12, 8q23.1, 16p11.2（JME特異的）", options: { color: C.text, fontFace: FONT_EN, breakLine: true } },
    { text: "病態基盤: ", options: { bold: true, color: C.text } },
    { text: "前頭-視床-皮質ネットワークの遺伝的異常", options: { color: C.text } },
  ], { x: 0.8, y: 3.9, w: 8.4, h: 1.4, fontSize: 15, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "2/15");
})();

// ============================================================
// SLIDE 3: 3つの発作型
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "臨床的特徴 ― 3つの発作型");

  var types = [
    { name: "ミオクロニー発作", freq: "100%", desc: "覚醒直後の両側上肢の\nピクつき\n意識は保持される", color: C.primary, icon: "⚡" },
    { name: "GTCS\n（強直間代発作）", freq: "80〜95%", desc: "ミオクロニー発作の\n数年後に出現\n診断契機となることが最多", color: C.red, icon: "🔴" },
    { name: "欠神発作", freq: "10〜40%", desc: "短時間の意識消失\n薬剤抵抗性の\nリスク因子", color: C.orange, icon: "⚠" },
  ];

  types.forEach(function(t, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.2, 3.0, 4.1);
    // Colored header bar
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 3.0, h: 0.7, fill: { color: t.color } });
    s.addText(t.name, { x: xPos, y: 1.2, w: 3.0, h: 0.7, fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // Frequency circle
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.7, y: 2.1, w: 1.6, h: 1.2, fill: { color: C.light } });
    s.addText(t.freq, { x: xPos + 0.7, y: 2.1, w: 1.6, h: 1.2, fontSize: 22, fontFace: FONT_EN, color: t.color, bold: true, align: "center", valign: "middle", margin: 0 });
    // Description
    s.addText(t.desc, { x: xPos + 0.2, y: 3.5, w: 2.6, h: 1.6, fontSize: 14, fontFace: FONT_B, color: C.text, align: "center", valign: "top", margin: 0 });
  });

  addPageNum(s, "3/15");
})();

// ============================================================
// SLIDE 4: 誘発因子
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "誘発因子");

  var factors = [
    { label: "睡眠不足", desc: "最重要。概日リズムとの関連が明確", color: C.red, icon: "💤" },
    { label: "アルコール", desc: "過度の摂取で発作閾値が低下", color: C.orange, icon: "🍺" },
    { label: "ストレス・疲労", desc: "精神的・身体的ストレス", color: C.yellow, icon: "😰" },
    { label: "光刺激", desc: "光過敏性 30〜40%。脳波で光突発反応", color: C.accent, icon: "💡" },
    { label: "月経周期", desc: "catamenial seizures → 薬剤抵抗性リスク", color: C.green, icon: "📅" },
  ];

  factors.forEach(function(f, i) {
    var yPos = 1.15 + i * 0.88;
    addCard(s, 0.5, yPos, 9, 0.78);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 0.25, h: 0.78, fill: { color: f.color } });
    s.addText(f.label, { x: 1.0, y: yPos, w: 2.5, h: 0.78, fontSize: 18, fontFace: FONT_H, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(f.desc, { x: 3.6, y: yPos, w: 5.6, h: 0.78, fontSize: 15, fontFace: FONT_B, color: C.sub, valign: "middle", margin: 0 });
  });

  addPageNum(s, "4/15");
})();

// ============================================================
// SLIDE 5: なぜ見逃されるのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "なぜ見逃されるのか？ ― 診断遅延の3大原因");

  // Warning card
  addCard(s, 0.5, 1.2, 9, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 0.25, h: 1.5, fill: { color: C.red } });
  s.addText([
    { text: "発症から診断まで平均 8〜10年", options: { fontSize: 24, bold: true, color: C.red } },
  ], { x: 1.0, y: 1.2, w: 8.2, h: 1.5, fontFace: FONT_H, align: "center", valign: "middle", margin: 0 });

  var reasons = [
    { num: "1", txt: "患者がミオクロニー発作を\n「病気」と認識していない", detail: "「朝は体がだるくてピクピクする」" },
    { num: "2", txt: "「てんかん＝全身けいれん」\nのイメージ", detail: "ピクつきでは受診しない" },
    { num: "3", txt: "医師がミオクロニーの\n既往を聴取し忘れる", detail: "GTCSだけを診断してしまう" },
  ];

  reasons.forEach(function(r, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 3.0, 3.0, 2.4);
    s.addShape(pres.shapes.OVAL, { x: xPos + 1.1, y: 3.1, w: 0.8, h: 0.7, fill: { color: C.red } });
    s.addText(r.num, { x: xPos + 1.1, y: 3.1, w: 0.8, h: 0.7, fontSize: 22, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(r.txt, { x: xPos + 0.1, y: 3.9, w: 2.8, h: 0.9, fontSize: 14, fontFace: FONT_B, color: C.text, align: "center", valign: "top", margin: 0 });
    s.addText(r.detail, { x: xPos + 0.1, y: 4.7, w: 2.8, h: 0.5, fontSize: 12, fontFace: FONT_B, color: C.sub, italic: true, align: "center", valign: "middle", margin: 0 });
  });

  addPageNum(s, "5/15");
})();

// ============================================================
// SLIDE 6: 問診のコツ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "JMEを拾い上げる問診テクニック");

  addCard(s, 0.5, 1.2, 9, 4.2);
  s.addText("GTCSで受診した思春期〜若年成人に聞く！", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_H, color: C.primary, bold: true, align: "center", margin: 0,
  });

  var questions = [
    "「朝起きてすぐ、コップや携帯を落としたことは？」",
    "「朝食のとき、箸やスプーンが飛んだことは？」",
    "「体がビクッとなって、周囲の人に指摘されたことは？」",
  ];

  questions.forEach(function(q, i) {
    var yPos = 2.1 + i * 0.9;
    s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: yPos, w: 8, h: 0.75, fill: { color: C.light } });
    s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: yPos, w: 0.2, h: 0.75, fill: { color: C.green } });
    s.addText(q, { x: 1.5, y: yPos, w: 7.3, h: 0.75, fontSize: 17, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
  });

  // Bottom note
  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 4.85, w: 8, h: 0.45, fill: { color: C.green } });
  s.addText("→「今思えば…」と答える患者は非常に多い！", {
    x: 1.0, y: 4.85, w: 8, h: 0.45,
    fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "6/15");
})();

// ============================================================
// SLIDE 7: 脳波所見
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳波所見 ― 4〜6Hz多棘徐波");

  // Left card: main findings
  addCard(s, 0.5, 1.2, 5.2, 4.2);
  s.addText("JME脳波の特徴", { x: 0.8, y: 1.3, w: 4.6, h: 0.5, fontSize: 18, fontFace: FONT_H, color: C.primary, bold: true, margin: 0 });

  var eegFindings = [
    { label: "基礎律動", val: "正常" },
    { label: "発作間欠期", val: "4〜6Hz 多棘徐波複合\n(polyspike-and-slow wave)" },
    { label: "分布", val: "前頭〜中心部優位" },
    { label: "光突発反応", val: "30〜40%に陽性" },
    { label: "感度向上", val: "睡眠剥奪脳波を推奨" },
  ];

  eegFindings.forEach(function(f, i) {
    var yPos = 1.9 + i * 0.7;
    s.addText(f.label, { x: 0.8, y: yPos, w: 2.0, h: 0.6, fontSize: 14, fontFace: FONT_H, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(f.val, { x: 2.8, y: yPos, w: 2.7, h: 0.6, fontSize: 14, fontFace: FONT_B, color: C.sub, valign: "middle", margin: 0 });
  });

  // Right card: key point
  addCard(s, 5.9, 1.2, 3.6, 4.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: 1.2, w: 3.6, h: 0.6, fill: { color: C.orange } });
  s.addText("鑑別ポイント", { x: 5.9, y: 1.2, w: 3.6, h: 0.6, fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });

  s.addText([
    { text: "定型欠神:", options: { bold: true, breakLine: true } },
    { text: "3Hz 棘徐波", options: { color: C.sub, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "JME:", options: { bold: true, breakLine: true } },
    { text: "4〜6Hz 多棘徐波", options: { color: C.red, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "「多棘」= 1つの徐波に\n先行する複数の棘波\n→ ミオクロニー発作の\n電気生理学的対応物", options: { fontSize: 13, color: C.text } },
  ], { x: 6.1, y: 2.0, w: 3.2, h: 3.2, fontSize: 15, fontFace: FONT_B, color: C.text, valign: "top", margin: [4, 4, 4, 4] });

  addPageNum(s, "7/15");
})();

// ============================================================
// SLIDE 8: 鑑別診断
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "鑑別診断 ― 特に注意すべき疾患");

  // Table header
  var tableY = 1.15;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: tableY, w: 9, h: 0.55, fill: { color: C.primary } });
  s.addText("鑑別疾患", { x: 0.5, y: tableY, w: 2.5, h: 0.55, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("共通点", { x: 3.0, y: tableY, w: 2.5, h: 0.55, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("鑑別のポイント", { x: 5.5, y: tableY, w: 4, h: 0.55, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var ddx = [
    { name: "若年欠神てんかん\n(JAE)", common: "GGE、思春期発症", point: "欠神が主体。ミオクロニーは稀\n3Hz棘徐波が典型" },
    { name: "覚醒時GTCS\n単独てんかん", common: "覚醒時のGTCS", point: "ミオクロニー・欠神なし" },
    { name: "進行性ミオクローヌス\nてんかん (PME)", common: "ミオクロニー、GTCS", point: "進行性の認知機能低下・小脳失調\nJMEは非進行性" },
    { name: "家族性成人型\nミオクロニーてんかん", common: "ミオクロニー、GTCS", point: "AD遺伝。持続性皮質ミオクロニー\nLong-read sequencingで確定" },
    { name: "心因性非てんかん性\n発作 (PNES)", common: "発作性症状", point: "偽薬剤抵抗性の原因\n精神的併存症の評価重要" },
  ];

  ddx.forEach(function(d, i) {
    var yPos = tableY + 0.55 + i * 0.82;
    var bg = (i % 2 === 0) ? C.light : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9, h: 0.82, fill: { color: bg } });
    s.addText(d.name, { x: 0.6, y: yPos, w: 2.3, h: 0.82, fontSize: 12, fontFace: FONT_H, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(d.common, { x: 3.0, y: yPos, w: 2.5, h: 0.82, fontSize: 12, fontFace: FONT_B, color: C.sub, valign: "middle", margin: 0 });
    s.addText(d.point, { x: 5.5, y: yPos, w: 3.8, h: 0.82, fontSize: 12, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
  });

  addPageNum(s, "8/15");
})();

// ============================================================
// SLIDE 9: 治療 第一選択薬
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "治療 ― 第一選択薬（2025年エビデンス）");

  // VPA card
  addCard(s, 0.5, 1.2, 4.3, 3.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.6, fill: { color: C.primary } });
  s.addText("バルプロ酸（VPA）", { x: 0.5, y: 1.2, w: 4.3, h: 0.6, fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "最も有効", options: { bold: true, color: C.primary, breakLine: true } },
    { text: "単剤で最大90%発作抑制", options: { breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "⚠ 催奇形性・神経発達リスク", options: { bold: true, color: C.red, breakLine: true } },
    { text: "→ 妊娠可能女性には原則禁忌", options: { color: C.red, breakLine: true } },
    { text: "体重増加・OSA悪化にも注意", options: { color: C.sub } },
  ], { x: 0.8, y: 1.9, w: 3.7, h: 2.2, fontSize: 14, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  // LEV card
  addCard(s, 5.2, 1.2, 4.3, 3.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.6, fill: { color: C.green } });
  s.addText("レベチラセタム（LEV）", { x: 5.2, y: 1.2, w: 4.3, h: 0.6, fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "VPAと同等の有効性", options: { bold: true, color: C.green, breakLine: true } },
    { text: "メタ解析: OR 0.77 (有意差なし)", options: { breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "✅ 妊娠中も比較的安全", options: { bold: true, color: C.green, breakLine: true } },
    { text: "5年転帰: LEV優位の可能性", options: { color: C.primary, breakLine: true } },
    { text: "⚠ 精神症状 37.6%に注意", options: { color: C.orange } },
  ], { x: 5.5, y: 1.9, w: 3.7, h: 2.2, fontSize: 14, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  // Meta-analysis highlight
  addCard(s, 0.5, 4.4, 9, 1.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.4, w: 0.2, h: 1.0, fill: { color: C.accent } });
  s.addText([
    { text: "2025年メタアナリシス (18研究, 2,189例): ", options: { bold: true, color: C.primary, fontFace: FONT_JP } },
    { text: "LEV vs VPA → ミオクロニー発作消失率に有意差なし。LEVはLTG (OR 2.22), TPM (OR 1.93)に有意に優位", options: { color: C.text } },
  ], { x: 1.0, y: 4.4, w: 8.2, h: 1.0, fontSize: 13, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, "9/15");
})();

// ============================================================
// SLIDE 10: 妊娠可能女性の治療アルゴリズム
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "妊娠可能女性の治療アルゴリズム");

  addCard(s, 0.5, 1.2, 9, 4.2);

  // Flow
  var flowSteps = [
    { text: "妊娠可能女性の JME", bg: C.dark, y: 1.35 },
    { text: "第一選択: レベチラセタム（LEV）", bg: C.green, y: 2.15 },
    { text: "LEV無効 or 精神症状で忍容不可", bg: C.orange, y: 2.95 },
    { text: "ラモトリギン（漸増）or ゾニサミド", bg: C.accent, y: 3.75 },
    { text: "ブリバラセタム or ペランパネル追加", bg: C.primary, y: 4.55 },
  ];

  flowSteps.forEach(function(step) {
    s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: step.y, w: 5.0, h: 0.55, fill: { color: step.bg } });
    s.addText(step.text, { x: 1.5, y: step.y, w: 5.0, h: 0.55, fontSize: 15, fontFace: FONT_B, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // Arrows
  [1.9, 2.7, 3.5, 4.3].forEach(function(y) {
    s.addText("▼", { x: 3.7, y: y, w: 0.6, h: 0.3, fontSize: 16, fontFace: FONT_B, color: C.primary, align: "center", margin: 0 });
  });

  // Side note
  s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 1.5, w: 2.3, h: 3.6, fill: { color: "FFF3CD" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 1.5, w: 2.3, h: 0.5, fill: { color: C.orange } });
  s.addText("OCP相互作用", { x: 7.0, y: 1.5, w: 2.3, h: 0.5, fontSize: 13, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText([
    { text: "TPM ≥200mg/day", options: { bold: true, breakLine: true } },
    { text: "→ エチニルエストラジオール↓", options: { breakLine: true, fontSize: 11 } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "PER ≥12mg/day", options: { bold: true, breakLine: true } },
    { text: "→ レボノルゲストレル代謝↑", options: { breakLine: true, fontSize: 11 } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "→ 避妊法の確認必須", options: { bold: true, color: C.red } },
  ], { x: 7.1, y: 2.1, w: 2.1, h: 2.8, fontSize: 12, fontFace: FONT_B, color: C.text, valign: "top", margin: [4,4,4,4] });

  addPageNum(s, "10/15");
})();

// ============================================================
// SLIDE 11: 避けるべき薬剤
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "JMEで避けるべき薬剤");

  // Big red warning
  addCard(s, 0.5, 1.2, 9, 2.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 0.6, fill: { color: C.red } });
  s.addText("⚠ ミオクロニー発作を増悪させるNaチャネルブロッカー", {
    x: 0.5, y: 1.2, w: 9, h: 0.6, fontSize: 17, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var drugs = ["カルバマゼピン（CBZ）", "フェニトイン（PHT）", "ガバペンチン（GBP）"];
  drugs.forEach(function(d, i) {
    var xPos = 0.8 + i * 2.9;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.0, w: 2.6, h: 0.7, fill: { color: "F8D7DA" } });
    s.addText("✕  " + d, { x: xPos, y: 2.0, w: 2.6, h: 0.7, fontSize: 15, fontFace: FONT_H, color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.0, w: 0.15, h: 0.7, fill: { color: C.red } });
  });

  // Pitfall
  addCard(s, 0.5, 3.6, 9, 1.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.6, w: 0.25, h: 1.8, fill: { color: C.orange } });
  s.addText("よくあるピットフォール", { x: 1.0, y: 3.65, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_H, color: C.orange, bold: true, margin: 0 });
  s.addText([
    { text: "GTCSだけを見て焦点性てんかんと誤診", options: { breakLine: true } },
    { text: "　→ カルバマゼピンを処方", options: { breakLine: true } },
    { text: "　→ ミオクロニー発作が増悪", options: { breakLine: true, color: C.red, bold: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "朝のミオクロニー発作を問診で拾い上げることが誤治療を防ぐ第一歩", options: { bold: true, color: C.primary } },
  ], { x: 1.0, y: 4.1, w: 8, h: 1.2, fontSize: 14, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  addPageNum(s, "11/15");
})();

// ============================================================
// SLIDE 12: 薬剤抵抗性JME
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "薬剤抵抗性JME ― 新規治療の選択肢");

  // Left: Risk factors
  addCard(s, 0.5, 1.2, 4.3, 4.2);
  s.addText("リスク因子", { x: 0.8, y: 1.3, w: 3.7, h: 0.5, fontSize: 18, fontFace: FONT_H, color: C.red, bold: true, margin: 0 });
  s.addText("15〜33% が薬剤抵抗性", { x: 0.8, y: 1.8, w: 3.7, h: 0.4, fontSize: 14, fontFace: FONT_B, color: C.sub, margin: 0 });

  var risks = ["欠神発作の合併", "若年発症", "月経関連発作", "熱性けいれん既往", "VPA抵抗性", "脳波上のGPFA"];
  risks.forEach(function(r, i) {
    var yPos = 2.3 + i * 0.5;
    s.addText("• " + r, { x: 0.8, y: yPos, w: 3.7, h: 0.45, fontSize: 14, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
  });

  // Right: New treatments
  addCard(s, 5.2, 1.2, 4.3, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fill: { color: C.green } });
  s.addText("セノバメート", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "3M: 75%が≥50%発作減少", options: { breakLine: true } },
    { text: "最終: 38%が発作消失", options: { breakLine: true } },
    { text: "ER受診も有意に減少", options: { breakLine: true } },
    { text: "(Cleveland Clinic, 16例)", options: { color: C.sub, fontSize: 11 } },
  ], { x: 5.4, y: 1.8, w: 3.9, h: 1.3, fontSize: 13, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  addCard(s, 5.2, 3.4, 4.3, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.4, w: 4.3, h: 0.5, fill: { color: C.accent } });
  s.addText("ブリバラセタム", { x: 5.2, y: 3.4, w: 4.3, h: 0.5, fontSize: 16, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "JME: responder rate 60%", options: { breakLine: true } },
    { text: "LEVから15:1比率で切替可", options: { breakLine: true } },
    { text: "精神行動副作用がLEVより少ない", options: { breakLine: true } },
    { text: "(多施設後方視研究, 61例)", options: { color: C.sub, fontSize: 11 } },
  ], { x: 5.4, y: 4.0, w: 3.9, h: 1.3, fontSize: 13, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  addPageNum(s, "12/15");
})();

// ============================================================
// SLIDE 13: 長期予後
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "長期予後 ― 断薬リスクと経過パターン");

  // Key stat
  addCard(s, 0.5, 1.2, 9, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 0.25, h: 1.3, fill: { color: C.red } });
  s.addText([
    { text: "断薬後の再発率: ", options: { fontSize: 22, color: C.text, fontFace: FONT_JP } },
    { text: "75〜100%", options: { fontSize: 28, bold: true, color: C.red, fontFace: FONT_EN } },
    { text: " (5年以内)", options: { fontSize: 22, color: C.text, fontFace: FONT_JP } },
  ], { x: 1.0, y: 1.2, w: 8.2, h: 0.7, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });
  s.addText("→ 多くの患者で生涯にわたる服薬継続が推奨", {
    x: 1.0, y: 1.9, w: 8.2, h: 0.5, fontSize: 16, fontFace: FONT_B, color: C.primary, align: "center", margin: 0,
  });

  // Prognosis patterns
  addCard(s, 0.5, 2.7, 4.3, 2.7);
  s.addText("経過パターン (GGE, n=199)", { x: 0.8, y: 2.8, w: 3.7, h: 0.5, fontSize: 16, fontFace: FONT_H, color: C.primary, bold: true, margin: 0 });
  var patterns = [
    { label: "寛解-再発型", pct: "40.2%", color: C.accent },
    { label: "3年寛解達成", pct: "81.9%", color: C.green },
    { label: "10年寛解達成", pct: "46.2%", color: C.orange },
    { label: "持続的抵抗性", pct: "14.6%", color: C.red },
  ];
  patterns.forEach(function(p, i) {
    var yPos = 3.4 + i * 0.5;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.2, h: 0.4, fill: { color: p.color } });
    s.addText(p.label, { x: 1.2, y: yPos, w: 2.0, h: 0.4, fontSize: 13, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
    s.addText(p.pct, { x: 3.3, y: yPos, w: 1.2, h: 0.4, fontSize: 15, fontFace: FONT_EN, color: p.color, bold: true, valign: "middle", margin: 0 });
  });

  // New view
  addCard(s, 5.2, 2.7, 4.3, 2.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.7, w: 4.3, h: 0.5, fill: { color: C.accent } });
  s.addText("従来の教条 vs 最新の知見", { x: 5.2, y: 2.7, w: 4.3, h: 0.5, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "従来: ", options: { bold: true } },
    { text: "「JME = 全例生涯服薬」", options: { breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "最新 (Nat Rev Neurol 2021):", options: { bold: true, breakLine: true } },
    { text: "「加齢とともに一部の患者で", options: { breakLine: true } },
    { text: "薬剤中止が可能」", options: { breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "→ 個々のリスク評価に基づく", options: { bold: true, color: C.primary, breakLine: true } },
    { text: "   判断が求められる", options: { bold: true, color: C.primary } },
  ], { x: 5.4, y: 3.3, w: 3.9, h: 2.0, fontSize: 13, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.15 });

  addPageNum(s, "13/15");
})();

// ============================================================
// SLIDE 14: 併存症・QOL
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "併存症と QOL ― 発作だけの問題ではない");

  // Psychiatric
  addCard(s, 0.5, 1.2, 3.0, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 3.0, h: 0.5, fill: { color: C.primary } });
  s.addText("精神医学的併存症", { x: 0.5, y: 1.2, w: 3.0, h: 0.5, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "不安障害: 10%", options: { breakLine: true } },
    { text: "うつ病: 6%", options: { breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "GTCS頻度と精神的", options: { breakLine: true } },
    { text: "併存症に有意な関連", options: { bold: true, color: C.primary } },
  ], { x: 0.7, y: 1.8, w: 2.6, h: 1.8, fontSize: 14, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  // Frontal
  addCard(s, 3.7, 1.2, 2.8, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 3.7, y: 1.2, w: 2.8, h: 0.5, fill: { color: C.orange } });
  s.addText("前頭葉機能障害", { x: 3.7, y: 1.2, w: 2.8, h: 0.5, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "遂行機能障害", options: { bold: true, breakLine: true } },
    { text: "発作コントロールの", options: { breakLine: true } },
    { text: "良否にかかわらず出現", options: { breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "前頭-視床-皮質", options: { breakLine: true } },
    { text: "ネットワーク異常", options: { color: C.orange, bold: true } },
  ], { x: 3.9, y: 1.8, w: 2.4, h: 1.8, fontSize: 13, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.15 });

  // Sleep
  addCard(s, 6.7, 1.2, 2.8, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 6.7, y: 1.2, w: 2.8, h: 0.5, fill: { color: C.accent } });
  s.addText("睡眠障害", { x: 6.7, y: 1.2, w: 2.8, h: 0.5, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "入眠潜時延長", options: { breakLine: true } },
    { text: "睡眠効率低下", options: { breakLine: true } },
    { text: "REM睡眠減少", options: { breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "VPA使用者:", options: { bold: true, breakLine: true } },
    { text: "いびき・OSA増加", options: { color: C.red, bold: true } },
  ], { x: 6.9, y: 1.8, w: 2.4, h: 1.8, fontSize: 13, fontFace: FONT_B, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.15 });

  // QOL bottom
  addCard(s, 0.5, 3.9, 9, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.9, w: 0.25, h: 1.5, fill: { color: C.green } });
  s.addText([
    { text: "QOLへの影響: ", options: { bold: true, color: C.primary } },
    { text: "発作コントロール下でも「発作への心配」「認知機能」ドメインが特に低下", options: { breakLine: true } },
    { text: "→ 発作管理 ＋ 精神面・睡眠面を含めた包括的ケアが長期QOL改善の鍵", options: { bold: true, color: C.green } },
  ], { x: 1.0, y: 3.95, w: 8.2, h: 1.3, fontSize: 15, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "14/15");
})();

// ============================================================
// SLIDE 15: Take-home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take-home Message", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 28, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "GTCSの若年者 → 必ず朝のミオクロニーを問診" },
    { num: "2", text: "脳波 = 4〜6Hz多棘徐波。睡眠剥奪で感度向上" },
    { num: "3", text: "VPA ≒ LEV の有効性。妊娠可能女性 → LEV第一選択" },
    { num: "4", text: "CBZ・PHT・GBP はミオクロニー増悪 → 禁忌" },
    { num: "5", text: "薬剤抵抗性(15-33%) → セノバメート・BRV" },
    { num: "6", text: "多くは生涯服薬。一部は加齢で中止可能" },
    { num: "7", text: "精神併存症 + 睡眠障害 → 包括的ケアが鍵" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.1 + i * 0.62;
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.05, w: 0.45, h: 0.45, fill: { color: C.accent } });
    s.addText(m.num, { x: 0.7, y: yPos + 0.05, w: 0.45, h: 0.45, fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.4, y: yPos, w: 8, h: 0.55, fontSize: 16, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
  });

  s.addText("ご視聴ありがとうございました", {
    x: 0.5, y: 5.1, w: 9, h: 0.4,
    fontSize: 16, fontFace: FONT_B, color: C.sub, align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 16: 参考文献
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "参考文献");

  var refs = [
    "1. Gesche J, et al. Eur J Neurol. 2020;27(4):676-684",
    "2. Vorderwülbecke BJ, et al. Nat Rev Neurol. 2021;18(2):71-83",
    "3. Cerulli Irelli E, et al. Epilepsia. 2020;61(11):2452-2460",
    "4. Mastrangelo M, et al. Am J Med Genet B. 2025;198(7):76-87",
    "5. Nica A. Rev Neurol (Paris). 2024;180(4):271-289",
    "6. Yetkin O, et al. Seizure. 2024;120:61-71",
    "7. Lu Y, et al. Epilepsia. 2025;66(11):4107-4121",
    "8. Pereira ARO, et al. Epilepsy Behav. 2025;171:110512",
    "9. Tabrizi N, et al. Seizure. 2025;127:66-70",
    "10. Stephen LJ, Brodie MJ. CNS Drugs. 2020;34(2):147-161",
    "11. Strzelczyk A, et al. Epilepsia. 2018;59(8):1549-1556",
    "12. Woodrich N, et al. Epilepsia Open. 2025;10(6):1871-1882",
    "13. Choudhury C, et al. Cureus. 2024;16(9):e69228",
    "14. Ustun D, et al. Epilepsy Res. 2026;222:107753",
  ];

  refs.forEach(function(r, i) {
    var yPos = 1.1 + i * 0.3;
    s.addText(r, { x: 0.5, y: yPos, w: 9, h: 0.3, fontSize: 10, fontFace: FONT_EN, color: C.text, valign: "middle", margin: 0 });
  });

  addPageNum(s, "16/16");
})();

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/JME_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
