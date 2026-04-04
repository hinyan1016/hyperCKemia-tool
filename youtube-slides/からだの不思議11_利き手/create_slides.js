var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "利き手はなぜ決まるのか？ ― 脳の左右非対称と遺伝のしくみ";

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

  s.addText("脳神経内科医が答える からだの不思議 #11", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("利き手はなぜ決まるのか？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.0,
    fontSize: 40, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addText("― 脳の左右非対称と遺伝のしくみ ―", {
    x: 0.8, y: 2.3, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.15, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.4, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.0, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 左利きの割合 ― 世界的な偏り
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "左利きの割合 ― 世界的な偏り");

  // 大きな数字
  addCard(s, 0.5, 1.1, 9.0, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 9.0, h: 0.5, fill: { color: C.primary } });
  s.addText("世界人口の利き手分布", { x: 0.8, y: 1.1, w: 8.4, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "右利き", options: { color: C.green, bold: true, fontSize: 28 } },
    { text: "  約88〜90%", options: { color: C.text, bold: true, fontSize: 22 } },
    { text: "　　", options: { fontSize: 22 } },
    { text: "左利き", options: { color: C.orange, bold: true, fontSize: 28 } },
    { text: "  約10〜12%", options: { color: C.text, bold: true, fontSize: 22 } },
  ], { x: 0.8, y: 1.65, w: 8.4, h: 0.9, fontFace: FONT_JP, valign: "middle", align: "center", margin: 0 });

  // 3つのポイント
  var pts = [
    { icon: "📜", title: "文化・時代を超えて一定", desc: "古代の壁画・石器の使用痕にも右利き優位\n→ 人類共通の生物学的傾向" },
    { icon: "🧬", title: "遺伝だけでは説明できない", desc: "一卵性双生児でも25%は利き手が不一致\n→ 遺伝寄与率は約25%にとどまる" },
    { icon: "🧠", title: "脳の非対称性が根底にある", desc: "生まれつきの神経回路の偏りが\n「利き手」として現れる" },
  ];
  pts.forEach(function(p, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 2.9, 3.0, 2.4);
    s.addText(p.icon, { x: xPos, y: 3.0, w: 3.0, h: 0.6, fontSize: 26, align: "center", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.2, y: 3.65, w: 2.6, h: 0.45, fill: { color: C.primary } });
    s.addText(p.title, { x: xPos + 0.2, y: 3.65, w: 2.6, h: 0.45, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.desc, { x: xPos + 0.1, y: 4.15, w: 2.8, h: 1.0, fontSize: 12, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "2/10");
})();

// ============================================================
// SLIDE 3: 脳の対側支配 ― 右手を動かすのは左脳
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳の対側支配 ― 右手を動かすのは左脳");

  // メイン概念カード
  addCard(s, 0.5, 1.1, 9.0, 1.4);
  s.addText("大脳半球優位性（laterality）", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_EN, color: C.primary, bold: true, italic: true, margin: 0,
  });
  s.addText("脳は外から見ると左右対称に見えるが、機能は大きく非対称。この非対称性を「大脳半球優位性」と呼ぶ。", {
    x: 0.8, y: 1.68, w: 8.4, h: 0.7,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 左右の脳の役割カード
  var sides = [
    { label: "左脳", role: "右手・右足を制御\n言語機能（ブローカ野・ウェルニッケ野）\n細かい手指の運動", color: C.primary, note: "右利きの96%で言語優位" },
    { label: "右脳", role: "左手・左足を制御\n空間認識・顔認識\n全体的なパターン処理", color: C.orange, note: "左利きでも70%は左脳言語優位" },
  ];
  sides.forEach(function(side, i) {
    var xPos = 0.5 + i * 4.8;
    addCard(s, xPos, 2.8, 4.4, 2.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.8, w: 4.4, h: 0.55, fill: { color: side.color } });
    s.addText(side.label, { x: xPos, y: 2.8, w: 4.4, h: 0.55, fontSize: 20, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(side.role, { x: xPos + 0.2, y: 3.4, w: 4.0, h: 1.1, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.95, w: 4.4, h: 0.35, fill: { color: C.light } });
    s.addText(side.note, { x: xPos, y: 4.95, w: 4.4, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // 対側支配の矢印
  s.addText("交差", { x: 4.6, y: 3.4, w: 0.8, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("⇌", { x: 4.55, y: 3.9, w: 0.9, h: 0.5, fontSize: 22, color: C.accent, align: "center", margin: 0 });

  addPageNum(s, "3/10");
})();

// ============================================================
// SLIDE 4: 遺伝要因 ― GWAS・多遺伝子形質
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "遺伝と利き手 ― どこまで「生まれつき」か");

  // 表ヘッダ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 2.8, h: 0.5, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.1, y: 1.1, w: 6.6, h: 0.5, fill: { color: C.primary } });
  s.addText("要因", { x: 0.3, y: 1.1, w: 2.8, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("利き手への影響", { x: 3.1, y: 1.1, w: 6.6, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var rows = [
    { factor: "遺伝", detail: "寄与率は約25%。単一の「利き手遺伝子」はなく、PCSK6・LRRTM1・MHC領域など多数の遺伝子座が微小な影響を積み重ねる多遺伝子形質（GWAS研究）", color: C.red },
    { factor: "胎児期環境", detail: "子宮内の位置・ホルモン環境（テストステロン等）・脊髄の左右差形成が関与。妊娠10週ごろから右手を口に近づける動作を好む", color: C.orange },
    { factor: "出生後の学習・文化", detail: "道具・文字・習慣による強化。強い矯正圧があれば外見上の利き手は変わりうる", color: C.green },
    { factor: "確率的揺らぎ", detail: "一卵性双生児の25%に利き手の不一致あり。発生の「偶然性」も影響する", color: C.accent },
  ];
  rows.forEach(function(r, i) {
    var yPos = 1.65 + i * 0.9;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.88, fill: { color: bgColor } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.12, h: 0.88, fill: { color: r.color } });
    s.addText(r.factor, { x: 0.42, y: yPos, w: 2.56, h: 0.88, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.detail, { x: 3.1, y: yPos, w: 6.4, h: 0.88, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.2 });
  });

  // 結論ボックス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.25, w: 9.4, h: 0.35, fill: { color: C.light } });
  s.addText("両親とも右利きでも、左利きの子が生まれる確率は約9〜10%", {
    x: 0.5, y: 5.25, w: 9.0, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "4/10");
})();

// ============================================================
// SLIDE 5: 胎児期からの形成
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "胎児期にすでに決まっている？");

  // タイムライン風カード
  var steps = [
    { time: "妊娠10週", title: "右手優位の動作", desc: "超音波研究で胎児が右手を口に近づける動作を好むことが確認。生後の利き手と高い一致率", color: C.green },
    { time: "妊娠〜出生", title: "脊髄の左右差形成", desc: "大脳の指令より先に脊髄レベルで非対称性が出現。末梢神経系からのボトムアップ的な形成", color: C.accent },
    { time: "出生後2〜5歳", title: "利き手の確立", desc: "道具・文字との接触で強化され、利き手が安定する。5〜6歳を過ぎても安定しない場合は個人差の範囲が多い", color: C.primary },
    { time: "成人以降", title: "利き手は固定", desc: "成人後に利き手が変わる場合は脳卒中・神経疾患の可能性。「急な変化」は受診のサイン", color: C.red },
  ];

  steps.forEach(function(step, i) {
    var yPos = 1.1 + i * 1.05;
    addCard(s, 0.5, yPos, 9.0, 0.88);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 1.7, h: 0.88, fill: { color: step.color } });
    s.addText(step.time, { x: 0.5, y: yPos, w: 1.7, h: 0.88, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(step.title, { x: 2.4, y: yPos + 0.05, w: 2.8, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(step.desc, { x: 2.4, y: yPos + 0.45, w: 6.9, h: 0.4, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // ポイント
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.25, w: 9.0, h: 0.35, fill: { color: C.light } });
  s.addText("利き手は「習慣」ではなく、胎児期の神経発達が土台 ― 生まれつきの特性", {
    x: 0.7, y: 5.25, w: 8.6, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "5/10");
})();

// ============================================================
// SLIDE 6: 言語優位半球との関係
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "言語優位半球との関係");

  // 定義
  addCard(s, 0.5, 1.1, 9.0, 1.0);
  s.addText("ブローカ野（発話）・ウェルニッケ野（言語理解）は多くの人で", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  s.addText("左半球に存在する（左半球言語優位）", {
    x: 0.8, y: 1.55, w: 8.4, h: 0.45,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  // 比較カード
  var groups = [
    {
      label: "右利き",
      stats: [
        { val: "約96%", desc: "左半球言語優位", color: C.primary },
        { val: "約4%", desc: "右半球 or 両側", color: C.sub },
      ],
      color: C.green,
    },
    {
      label: "左利き",
      stats: [
        { val: "約70%", desc: "左半球言語優位", color: C.primary },
        { val: "約15%", desc: "右半球言語優位", color: C.orange },
        { val: "約15%", desc: "両半球", color: C.accent },
      ],
      color: C.orange,
    },
  ];
  groups.forEach(function(g, gi) {
    var xBase = 0.4 + gi * 4.7;
    addCard(s, xBase, 2.3, 4.4, 2.9);
    s.addShape(pres.shapes.RECTANGLE, { x: xBase, y: 2.3, w: 4.4, h: 0.5, fill: { color: g.color } });
    s.addText(g.label, { x: xBase, y: 2.3, w: 4.4, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    g.stats.forEach(function(st, si) {
      var yPos = 2.95 + si * 0.72;
      s.addText(st.val, { x: xBase + 0.2, y: yPos, w: 1.4, h: 0.6, fontSize: 22, fontFace: FONT_JP, color: st.color, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(st.desc, { x: xBase + 1.7, y: yPos, w: 2.5, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
    });
  });

  // 警告ボックス
  addCard(s, 0.5, 5.05, 9.0, 0.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.05, w: 0.12, h: 0.5, fill: { color: C.red } });
  s.addText("臨床的注意：脳卒中・てんかん手術では利き手だけで言語優位半球を決定できない → fMRI・Wada試験で個別評価", {
    x: 0.75, y: 5.05, w: 8.6, h: 0.5,
    fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
  });

  addPageNum(s, "6/10");
})();

// ============================================================
// SLIDE 7: 左利きの誤解と事実
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "左利きにまつわる誤解と事実");

  var myths = [
    {
      myth: "「左利きは右脳人間」",
      fact: "左利きの約70%は右利きと同じく左脳で言語処理。大規模研究では利き手と創造性・IQ・芸術的才能に有意な相関なし",
      icon: "✗",
      iconColor: C.red,
    },
    {
      myth: "「左利きは矯正すべき」",
      fact: "強制的な矯正は吃音・情緒不安定・学習困難のリスクを高める可能性が指摘されており、現在では推奨されない",
      icon: "✗",
      iconColor: C.red,
    },
    {
      myth: "「左利きは少数派＝異常」",
      fact: "約10%という割合は何千年もの人類史を通じてほぼ一定。進化的に淘汰されなかったことは左利きに適応的優位性がある可能性を示す（サウスポー効果）",
      icon: "✗",
      iconColor: C.red,
    },
  ];

  myths.forEach(function(m, i) {
    var yPos = 1.1 + i * 1.4;
    addCard(s, 0.5, yPos, 9.0, 1.2);
    // 誤解ラベル
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 3.8, h: 0.45, fill: { color: "FFEBEE" } });
    s.addText(m.icon + "  誤解  " + m.myth, { x: 0.7, y: yPos, w: 3.6, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: 0 });
    // 事実ラベル
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos + 0.45, w: 9.0, h: 0.75, fill: { color: "E8F4FD" } });
    s.addText([
      { text: "事実  ", options: { color: C.primary, bold: true, fontSize: 13 } },
      { text: m.fact, options: { color: C.text, fontSize: 13 } },
    ], { x: 0.7, y: yPos + 0.45, w: 8.6, h: 0.75, fontFace: FONT_JP, valign: "middle", margin: [0, 4, 0, 4], lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "7/10");
})();

// ============================================================
// SLIDE 8: 矯正の問題 ― なぜ左利きを直さないのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "矯正の問題 ― なぜ左利きを直さないのか");

  // 歴史的背景
  addCard(s, 0.5, 1.1, 9.0, 1.0);
  s.addText("かつての矯正の歴史", { x: 0.8, y: 1.15, w: 4.0, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText("多くの国で20世紀中盤まで学校・家庭での左手矯正が行われていた。\n現在は神経発達的根拠から「矯正しない」が国際的標準", {
    x: 0.8, y: 1.55, w: 8.4, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // リスクカード（3つ）
  var risks = [
    { icon: "🗣", title: "吃音リスク", desc: "強制矯正と吃音発症の関連が報告されている。言語野への影響が示唆される" },
    { icon: "😟", title: "情緒不安定", desc: "矯正ストレスにより情緒的な問題が生じる可能性が指摘されている" },
    { icon: "📖", title: "学習困難", desc: "利き手と異なる手の使用強制が学習への集中を妨げることがある" },
  ];
  risks.forEach(function(r, i) {
    var xPos = 0.4 + i * 3.1;
    addCard(s, xPos, 2.4, 2.9, 2.1);
    s.addText(r.icon, { x: xPos, y: 2.5, w: 2.9, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.15, w: 2.9, h: 0.45, fill: { color: C.red } });
    s.addText(r.title, { x: xPos, y: 3.15, w: 2.9, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(r.desc, { x: xPos + 0.1, y: 3.65, w: 2.7, h: 0.8, fontSize: 12, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // 結論
  addCard(s, 0.5, 4.75, 9.0, 0.75);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.75, w: 0.12, h: 0.75, fill: { color: C.green } });
  s.addText([
    { text: "左利きのままでよい。", options: { bold: true, color: C.primary, fontSize: 16, breakLine: false } },
    { text: " 日常生活・学業・スポーツに支障なし。左利き用道具も充実している。", options: { color: C.text, fontSize: 14 } },
  ], { x: 0.8, y: 4.75, w: 8.6, h: 0.75, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, "8/10");
})();

// ============================================================
// SLIDE 9: こんなときは要注意 ― 利き手が変わる場合
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなときは要注意 ― 利き手が変わる場合");

  // 説明文
  s.addText("利き手は通常、幼少期に確立すると生涯変わりません。成人後に利き手が変わる・変わった感じがする場合は神経疾患のサインの可能性があります。", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.65,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // 警告サイン（赤背景）
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.9, w: 9.0, h: 0.5, fill: { color: C.red } });
  s.addText("受診を考えたほうがよい「利き手の変化」", {
    x: 0.7, y: 1.9, w: 8.6, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });

  var signs = [
    "以前は問題なく使えていた利き手が急に不器用になった",
    "利き手の手・腕に力が入らない、または震える",
    "細かい作業（箸・ボタン・字を書く）が急に難しくなった",
    "利き手側の感覚がおかしい（しびれ・違和感）",
    "利き手とあわせて言葉が出にくい・ろれつが回らないがある",
  ];
  signs.forEach(function(sg, i) {
    var yPos = 2.5 + i * 0.52;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.48, fill: { color: (i % 2 === 0) ? "FFF5F5" : C.white } });
    s.addShape(pres.shapes.OVAL, { x: 0.65, y: yPos + 0.09, w: 0.3, h: 0.3, fill: { color: C.red } });
    s.addText(sg, { x: 1.1, y: yPos, w: 8.2, h: 0.48, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  // 緊急注記
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.15, w: 9.0, h: 0.4, fill: { color: "FFEBEE" } });
  s.addText("特に急性の変化は脳卒中の可能性あり → ためらわず速やかに受診", {
    x: 0.7, y: 5.15, w: 8.6, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "9/10");
})();

// ============================================================
// SLIDE 10: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.25, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "世界人口の約88〜90%が右利き。この偏りは人類史を通じてほぼ一定" },
    { num: "2", text: "利き手は遺伝・胎児期環境・発生の確率的揺らぎの組み合わせで決まる。単一の「利き手遺伝子」はない" },
    { num: "3", text: "左利きの約70%は右利きと同じく左脳で言語処理。「左利き＝右脳人間」は誤解" },
    { num: "4", text: "左利きの無理な矯正は推奨されない。吃音・情緒不安定のリスクがある" },
    { num: "5", text: "成人後に突然利き手が不器用になったら脳卒中を疑い、速やかに受診" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.05 + i * 0.88;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.75, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.12, w: 0.5, h: 0.5, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.12, w: 0.5, h: 0.5, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.7, y: yPos, w: 7.2, h: 0.75, fontSize: 15, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  // 底部メッセージ
  s.addShape(pres.shapes.LINE, { x: 2, y: 5.3, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.38, w: 9.0, h: 0.35,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "10/10");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/利き手_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
