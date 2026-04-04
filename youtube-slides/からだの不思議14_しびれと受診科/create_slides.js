var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "しびれが出たら何科？ ― 脳神経内科・整形外科・内科の使い分け";

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

  s.addText("脳神経内科医が答える からだの不思議 #14", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("しびれが出たら\n何科を受診すればいい？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 脳神経内科・整形外科・内科の使い分け ―", {
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
    "「手がしびれてきた。どこの病院に行けばいいんだろう…？」",
    "「足のしびれが続くけど、整形外科でいいのかな？」",
    "「整形外科で異常なしと言われたのに、しびれが治らない」",
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
  s.addText("原因によって受診すべき科がまったく異なります", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: しびれの神経解剖 ― 中枢性 vs 末梢性
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "しびれの神経解剖 ― 中枢性 vs 末梢性");

  // 2列カード
  var types = [
    {
      label: "中枢性しびれ",
      where: "脳・脊髄が原因",
      examples: "脳梗塞・脳出血\n多発性硬化症・脊髄炎",
      pattern: "体の片側全体\n（顔＋腕＋足）\n体幹レベル以下",
      color: C.red,
      icon: "🧠",
    },
    {
      label: "末梢性しびれ",
      where: "末梢神経・脊椎が原因",
      examples: "頸椎症・手根管症候群\n糖尿病性神経障害\nビタミンB12欠乏",
      pattern: "特定の指だけ\n足の裏だけ\n両足がじわじわ",
      color: C.green,
      icon: "🦴",
    },
  ];

  types.forEach(function(t, i) {
    var xPos = 0.3 + i * 4.9;
    addCard(s, xPos, 1.1, 4.5, 4.3);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 4.5, h: 0.55, fill: { color: t.color } });
    s.addText(t.label, { x: xPos, y: 1.1, w: 4.5, h: 0.55, fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.icon, { x: xPos, y: 1.7, w: 4.5, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText("場所：" + t.where, { x: xPos + 0.2, y: 2.3, w: 4.1, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: t.color, bold: true, margin: 0 });
    s.addText("代表例", { x: xPos + 0.2, y: 2.75, w: 4.1, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.sub, bold: true, margin: 0 });
    s.addText(t.examples, { x: xPos + 0.2, y: 3.05, w: 4.1, h: 0.7, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2 });
    s.addText("パターン", { x: xPos + 0.2, y: 3.75, w: 4.1, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.sub, bold: true, margin: 0 });
    s.addText(t.pattern, { x: xPos + 0.2, y: 4.05, w: 4.1, h: 0.7, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.1, w: 9.4, h: 0.45, fill: { color: C.light } });
  s.addText("「場所のパターン」が受診科の選択に直結する最大の手がかり", {
    x: 0.5, y: 5.1, w: 9.0, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: 科別対象疾患の比較表
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "どの科に行けばいい？ ― 科別の対象疾患");

  var rows = [
    {
      dept: "脳神経内科",
      diseases: "脳梗塞・多発性硬化症\nギランバレー症候群\n糖尿病性神経障害・脊髄炎",
      pattern: "片側の顔+腕+足\n体幹レベル以下\n両足先のじわじわ感",
      deptColor: C.primary,
    },
    {
      dept: "整形外科",
      diseases: "頸椎症・椎間板ヘルニア\n手根管症候群\n脊柱管狭窄症",
      pattern: "首の動きで変化する腕のしびれ\n親指〜中指3本\n腰痛＋片足のしびれ",
      deptColor: C.green,
    },
    {
      dept: "内科\n（糖尿病・代謝）",
      diseases: "糖尿病性末梢神経障害\nビタミンB12欠乏\n甲状腺機能低下症",
      pattern: "両足先から対称性に上がる\n足先の灼熱感\n慢性・緩徐進行",
      deptColor: C.orange,
    },
  ];

  // ヘッダ行
  s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 1.05, w: 2.5, h: 0.55, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.7, y: 1.05, w: 3.5, h: 0.55, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 1.05, w: 3.6, h: 0.55, fill: { color: C.dark } });
  s.addText("診療科", { x: 0.2, y: 1.05, w: 2.5, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("主な対象疾患", { x: 2.7, y: 1.05, w: 3.5, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("しびれのパターン例", { x: 6.2, y: 1.05, w: 3.6, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.6 + i * 1.25;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: yPos, w: 9.6, h: 1.25, fill: { color: bgColor } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: yPos, w: 0.1, h: 1.25, fill: { color: r.deptColor } });
    s.addText(r.dept, { x: 0.35, y: yPos, w: 2.2, h: 1.25, fontSize: 14, fontFace: FONT_JP, color: r.deptColor, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.diseases, { x: 2.7, y: yPos, w: 3.4, h: 1.25, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.2 });
    s.addText(r.pattern, { x: 6.2, y: yPos, w: 3.6, h: 1.25, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.2 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 5.1, w: 9.6, h: 0.45, fill: { color: C.yellow } });
  s.addText("「整形外科で異常なし」でも安心しないで ― 内科的・神経内科的な原因が隠れることがある", {
    x: 0.4, y: 5.1, w: 9.2, h: 0.45,
    fontSize: 13, fontFace: FONT_JP, color: C.dark, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: フローチャート風判断ガイド
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "受診科の選び方フローチャート");

  var steps = [
    {
      q: "Q1　突然起きた（数分〜数時間）しびれ？\n　　 顔・腕・足が片側 / ろれつ / 視野異常を伴う",
      a: "→  すぐ救急へ。脳卒中の可能性",
      qColor: C.red,
      aColor: C.red,
      icon: "！",
      iconBg: C.red,
    },
    {
      q: "Q2　首や腰を動かすと悪化する？\n　　 腕や足の特定の指だけしびれる？ 首・腰の痛みを伴う",
      a: "→  整形外科　頸椎・腰椎・手根管症候群を疑う",
      qColor: C.text,
      aColor: C.green,
      icon: "骨",
      iconBg: C.green,
    },
    {
      q: "Q3　糖尿病・大量飲酒・偏食がある？\n　　 両足先からじわじわ上がってくる？",
      a: "→  内科（糖尿病内科）　代謝・栄養性の末梢神経障害",
      qColor: C.text,
      aColor: C.orange,
      icon: "内",
      iconBg: C.orange,
    },
    {
      q: "Q4　上記に当てはまらない？\n　　 体の片側だけ・特定レベル以下・再発を繰り返す？",
      a: "→  脳神経内科　脳・脊髄・複合的な末梢神経疾患を精査",
      qColor: C.text,
      aColor: C.primary,
      icon: "脳",
      iconBg: C.primary,
    },
  ];

  steps.forEach(function(st, i) {
    var yPos = 1.05 + i * 1.12;
    addCard(s, 0.3, yPos, 9.4, 1.0);
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: yPos + 0.2, w: 0.6, h: 0.6, fill: { color: st.iconBg } });
    s.addText(st.icon, { x: 0.5, y: yPos + 0.2, w: 0.6, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.q, { x: 1.25, y: yPos + 0.03, w: 5.9, h: 0.55, fontSize: 12, fontFace: FONT_JP, color: st.qColor, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    s.addText(st.a, { x: 1.25, y: yPos + 0.58, w: 7.9, h: 0.38, fontSize: 13, fontFace: FONT_JP, color: st.aColor, bold: true, valign: "middle", margin: 0 });
  });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: 脳卒中FAST ― 緊急サイン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.red } });

  s.addText("救急へ ― 脳卒中のFASTサイン", {
    x: 0.5, y: 0.15, w: 9.0, h: 0.65,
    fontSize: 26, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  var fast = [
    { letter: "F", word: "Face", ja: "顔のゆがみ", desc: "笑ったとき片側の口角が下がる", color: C.red },
    { letter: "A", word: "Arm", ja: "腕の脱力", desc: "両腕を前に出すと片側が下がる", color: C.orange },
    { letter: "S", word: "Speech", ja: "言語障害", desc: "ろれつが回らない・言葉が出ない", color: C.yellow },
    { letter: "T", word: "Time", ja: "発症時刻の確認", desc: "治療の時間窓（4.5時間以内）がある", color: C.accent },
  ];

  fast.forEach(function(f, i) {
    var xPos = 0.2 + i * 2.42;
    addCard(s, xPos, 0.95, 2.25, 3.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 0.95, w: 2.25, h: 0.8, fill: { color: f.color } });
    s.addText(f.letter, { x: xPos, y: 0.95, w: 1.1, h: 0.8, fontSize: 36, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.word, { x: xPos + 1.1, y: 0.95, w: 1.15, h: 0.8, fontSize: 16, fontFace: FONT_EN, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText(f.ja, { x: xPos, y: 1.8, w: 2.25, h: 0.55, fontSize: 16, fontFace: FONT_JP, color: f.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.desc, { x: xPos + 0.1, y: 2.4, w: 2.05, h: 1.0, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 4, lineSpacingMultiple: 1.3 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.6, w: 9.4, h: 0.95, fill: { color: "C0392B" } });
  s.addText([
    { text: "しびれだけの脳梗塞（感覚性脳卒中）も存在します\n", options: { fontSize: 16, bold: true, color: C.white, breakLine: false } },
    { text: "突然起きたしびれは、FASTがなくても救急を受診してください", options: { fontSize: 14, color: C.light } },
  ], { x: 0.5, y: 4.62, w: 9.0, h: 0.9, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: 末梢性しびれの代表例
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "末梢性しびれの代表例");

  var cases = [
    {
      name: "手根管症候群",
      dept: "整形外科",
      trigger: "親指〜薬指（手のひら側）\n夜間・早朝に悪化",
      note: "妊娠・甲状腺低下・透析があれば内科も",
      color: C.green,
    },
    {
      name: "糖尿病性神経障害",
      dept: "内科・神経内科",
      trigger: "両足先から対称性に\n灼熱感・しびれが上がる",
      note: "血糖コントロール改善が治療の柱",
      color: C.orange,
    },
    {
      name: "頸椎椎間板ヘルニア",
      dept: "整形外科",
      trigger: "首の動きで悪化\n腕の特定の指だけしびれる",
      note: "難治例・脊髄症は脳神経外科も",
      color: C.primary,
    },
    {
      name: "多発性硬化症",
      dept: "脳神経内科",
      trigger: "数週間で部位が変わる\n体幹レベル以下がすべてしびれる",
      note: "若い女性に多い・再発を繰り返す",
      color: C.red,
    },
  ];

  cases.forEach(function(c, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var xPos = 0.3 + col * 4.85;
    var yPos = 1.05 + row * 2.1;
    addCard(s, xPos, yPos, 4.55, 1.95);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 4.55, h: 0.45, fill: { color: c.color } });
    s.addText(c.name, { x: xPos + 0.1, y: yPos, w: 2.8, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
    s.addText(c.dept, { x: xPos + 2.9, y: yPos, w: 1.55, h: 0.45, fontSize: 12, fontFace: FONT_JP, color: C.light, align: "right", valign: "middle", margin: [0, 4, 0, 0] });
    s.addText("パターン", { x: xPos + 0.2, y: yPos + 0.5, w: 4.1, h: 0.25, fontSize: 12, fontFace: FONT_JP, color: C.sub, bold: true, margin: 0 });
    s.addText(c.trigger, { x: xPos + 0.2, y: yPos + 0.72, w: 4.1, h: 0.65, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2 });
    s.addText("ポイント：" + c.note, { x: xPos + 0.2, y: yPos + 1.55, w: 4.15, h: 0.35, fontSize: 11, fontFace: FONT_JP, color: c.color, bold: true, margin: 0 });
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: こんなときは脳神経内科へ急いで
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなときは脳神経内科への受診を急いで");

  var checks = [
    "片側の顔・腕・足がすべてしびれる（同時に起きる）",
    "数週間で体の別の場所に広がった（多発性硬化症など）",
    "体の胸・腰の高さより下の全体がしびれる（脊髄病変）",
    "足がしびれると同時に力が入りにくい・歩きにくい",
    "排尿・排便のコントロールが急に悪くなった（緊急サイン）",
    "糖尿病がないのに両足がじわじわとしびれて悪化する",
  ];

  checks.forEach(function(c, i) {
    var yPos = 1.1 + i * 0.72;
    addCard(s, 0.3, yPos, 9.4, 0.62);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.12, h: 0.62, fill: { color: C.red } });
    s.addShape(pres.shapes.OVAL, { x: 0.55, y: yPos + 0.14, w: 0.34, h: 0.34, fill: { color: C.red } });
    s.addText("" + (i + 1), { x: 0.55, y: yPos + 0.14, w: 0.34, h: 0.34, fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c, { x: 1.05, y: yPos, w: 8.5, h: 0.62, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.45, w: 9.4, h: 0.1, fill: { color: C.red } });

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
    x: 0.5, y: 0.2, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "しびれの原因は「中枢性（脳・脊髄）」と「末梢性（末梢神経・骨格系）」に大別される" },
    { num: "2", text: "「場所のパターン」が受診科の最大の手がかり\n片側の顔+腕+足 → 脳神経内科　首の動きで変化 → 整形外科　両足じわじわ → 内科" },
    { num: "3", text: "突然起きた片側のしびれはFASTを確認し、迷わず救急へ\n「しびれだけの脳梗塞」も存在する" },
    { num: "4", text: "「整形外科で異常なし」でも原因が見つかっていなければ\n脳神経内科への相談を恐れないで" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.05 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.2, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.2, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.3, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "9/9");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/しびれと受診科_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
