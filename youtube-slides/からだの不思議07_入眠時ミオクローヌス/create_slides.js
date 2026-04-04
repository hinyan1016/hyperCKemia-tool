var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "寝てる時にビクッとなるのはなぜ？ ― 「入眠時ミオクローヌス」の正体";

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

  s.addText("脳神経内科医が答える からだの不思議 #07", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("寝てる時にビクッとなるのはなぜ？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.0,
    fontSize: 34, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });
  s.addText("「入眠時ミオクローヌス」の正体", {
    x: 0.8, y: 2.25, w: 8.4, h: 0.8,
    fontSize: 30, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", margin: 0,
  });

  s.addText("― 脳神経内科医が解説するそのメカニズム ―", {
    x: 0.8, y: 3.15, w: 8.4, h: 0.55,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.85, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.65, w: 8.4, h: 0.4,
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
    "「電車でうとうとしていたら、急にビクッとなって恥ずかしかった…」",
    "「眠りに落ちる瞬間に、崖から落ちる夢を見てビクッと目が覚める」",
    "「疲れた日の夜は特に強くて、パートナーに『大丈夫？』と心配された」",
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
  s.addText("健康な成人の60〜70%が経験する、ごく一般的な現象", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: 入眠時ミオクローヌスとは
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "入眠時ミオクローヌスとは");

  // 定義カード
  addCard(s, 0.5, 1.1, 9.0, 1.5);
  s.addText("Hypnic Jerk  /  Sleep Start", {
    x: 0.8, y: 1.2, w: 8.4, h: 0.5,
    fontSize: 22, fontFace: FONT_EN, color: C.primary, bold: true, italic: true, margin: 0,
  });
  s.addText("眠りに落ちる瞬間に全身または四肢がビクッと収縮する生理的現象", {
    x: 0.8, y: 1.7, w: 8.4, h: 0.7,
    fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 語源ボックス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.85, w: 9.0, h: 0.55, fill: { color: C.light } });
  s.addText("ミオクローヌス ＝「ミオ（筋）」＋「クローヌス（痙攣性の収縮）」　→ 生理的（正常な）ミオクローヌスに分類", {
    x: 0.7, y: 2.85, w: 8.6, h: 0.55,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, valign: "middle", margin: 0,
  });

  // 頻度カード
  var freqs = [
    { label: "成人全体", val: "60〜70%", sub: "経験あり", color: C.green },
    { label: "疲労・睡眠不足時", val: "頻度↑", sub: "強度も増す", color: C.orange },
    { label: "危険な病気？", val: "ほとんどNO", sub: "生理的現象", color: C.primary },
  ];
  freqs.forEach(function(f, i) {
    var xPos = 0.5 + i * 3.1;
    addCard(s, xPos, 3.65, 2.9, 1.8);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.65, w: 2.9, h: 0.5, fill: { color: f.color } });
    s.addText(f.label, { x: xPos, y: 3.65, w: 2.9, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.val, { x: xPos, y: 4.2, w: 2.9, h: 0.65, fontSize: 22, fontFace: FONT_JP, color: f.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.sub, { x: xPos, y: 4.9, w: 2.9, h: 0.45, fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", valign: "middle", margin: 0 });
  });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: 脳の仕組み ― 網様体と覚醒-睡眠移行
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳の中で何が起きているか ― 覚醒→睡眠の切り替え時");

  // 3つのステップカード
  var steps = [
    { num: "1", title: "覚醒モード", desc: "脳幹の網様体賦活系が\n全身筋肉に「緊張維持」の\n信号を送り続ける", color: C.orange },
    { num: "2", title: "移行期（ここでビクッ！）", desc: "網様体の信号が急減弱する\n過渡期に「誤放電」が発生\n→ これが入眠時ミオクローヌス", color: C.red },
    { num: "3", title: "睡眠モード", desc: "筋緊張が低下し\n安定した睡眠へ移行", color: C.green },
  ];
  steps.forEach(function(st, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.2, 3.0, 2.8);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 3.0, h: 0.5, fill: { color: st.color } });
    s.addText(st.num, { x: xPos, y: 1.2, w: 0.5, h: 0.5, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.title, { x: xPos + 0.4, y: 1.2, w: 2.5, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
    s.addText(st.desc, { x: xPos, y: 1.75, w: 3.0, h: 1.1, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: [8, 6, 0, 6], lineSpacingMultiple: 1.25 });
  });
  // 矢印
  s.addText("→", { x: 3.05, y: 2.2, w: 0.45, h: 0.5, fontSize: 24, color: C.accent, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 6.25, y: 2.2, w: 0.45, h: 0.5, fontSize: 24, color: C.accent, align: "center", valign: "middle", margin: 0 });

  // 「崖から落ちる夢」の説明
  addCard(s, 0.5, 4.25, 9.0, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.25, w: 9.0, h: 0.45, fill: { color: C.primary } });
  s.addText("「崖から落ちる夢」の謎", { x: 0.7, y: 4.25, w: 8.6, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "筋肉のビクつきが先 → 脳がそれを感知して「落下した」というストーリーを後付けで生成", options: { color: C.text, fontSize: 14, breakLine: true } },
    { text: "夢が原因ではなく、ビクつきへの脳の解釈が夢を生む", options: { color: C.sub, fontSize: 13 } },
  ], { x: 0.7, y: 4.75, w: 8.6, h: 0.6, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: 鑑別表 ― てんかん・RLS・PLMDとの違い
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "似ているけれど別の疾患 ― 鑑別のポイント");

  var rows = [
    {
      disease: "てんかん\n（ミオクロニー発作）",
      diff: "覚醒時にも起こる\n朝方（起床直後）に多い\n繰り返す・意識消失を伴うことも",
      sign: "明け方にコーヒーをこぼす\n箸を落とすなどが続く",
    },
    {
      disease: "むずむず脚症候群\n（RLS）",
      diff: "「動かさずにいられない不快感」が主\nビクッではなく「むずむず・ほてり」",
      sign: "夕方〜夜間に悪化\n歩くと楽になる\n鉄欠乏・腎不全に伴うことも",
    },
    {
      disease: "周期性四肢運動障害\n（PLMD）",
      diff: "睡眠中に20〜40秒ごとに\n足がピクつく\n本人は気づかないことが多い",
      sign: "「眠れない」「日中に強い眠気」\nパートナーが「足が動いている」と指摘",
    },
  ];

  // テーブルヘッダ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.0, w: 2.4, h: 0.55, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.7, y: 1.0, w: 3.8, h: 0.55, fill: { color: C.primary } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 1.0, w: 3.3, h: 0.55, fill: { color: C.accent } });
  s.addText("疾患", { x: 0.3, y: 1.0, w: 2.4, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("入眠時ミオクローヌスとの違い", { x: 2.7, y: 1.0, w: 3.8, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("特徴的なサイン", { x: 6.5, y: 1.0, w: 3.3, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.55 + i * 1.25;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.5, h: 1.2, fill: { color: bgColor } });
    s.addText(r.disease, { x: 0.3, y: yPos, w: 2.4, h: 1.2, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2, lineSpacingMultiple: 1.2 });
    s.addText(r.diff, { x: 2.7, y: yPos, w: 3.8, h: 1.2, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.2 });
    s.addText(r.sign, { x: 6.5, y: yPos, w: 3.3, h: 1.2, fontSize: 12, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.2 });
  });

  // 鑑別の核心
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.05, w: 9.5, h: 0.5, fill: { color: C.light } });
  s.addText("典型的な入眠時ミオクローヌス：「眠りに落ちる瞬間だけ」「1〜2回で終わる」「翌朝には何ともない」", {
    x: 0.5, y: 5.05, w: 9.2, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: 誘因 ― 何があると起こりやすいか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなとき起こりやすい ― 誘因");

  var triggers = [
    { icon: "☕", label: "カフェインの過剰摂取", desc: "覚醒から睡眠への移行が\n不安定になりやすい" },
    { icon: "😰", label: "強い精神的ストレス", desc: "神経系の興奮が高まり\n睡眠移行が乱れる" },
    { icon: "💤", label: "睡眠不足・慢性的疲労", desc: "疲れているほど逆説的に\n頻度・強度が増す" },
    { icon: "🏃", label: "激しい運動（就寝前）", desc: "交感神経亢進状態のまま\n眠りに入るため誤放電↑" },
  ];

  triggers.forEach(function(t, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var xPos = 0.4 + col * 4.85;
    var yPos = 1.1 + row * 2.1;
    addCard(s, xPos, yPos, 4.5, 1.85);
    s.addText(t.icon, { x: xPos + 0.1, y: yPos + 0.15, w: 1.0, h: 1.2, fontSize: 36, align: "center", valign: "middle", margin: 0 });
    s.addText(t.label, { x: xPos + 1.1, y: yPos + 0.1, w: 3.2, h: 0.55, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.desc, { x: xPos + 1.1, y: yPos + 0.65, w: 3.2, h: 1.0, fontSize: 13, fontFace: FONT_JP, color: C.sub, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
  });

  // 対策メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.05, w: 9.2, h: 0.5, fill: { color: C.light } });
  s.addText("→ これらを見直すだけで頻度・強度が改善することが多い", {
    x: 0.6, y: 5.05, w: 9.0, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: 受診を考えるサイン（チェックリスト）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなときは要注意 ― 受診を考えるサイン");

  var checks = [
    { text: "ビクつきが1晩に何十回も繰り返す（入眠できないほど）", urgent: false },
    { text: "朝方（起床直後）にも手足がビクつく", urgent: false },
    { text: "日中（覚醒中）にも突然筋肉がビクつく", urgent: false },
    { text: "ビクつきのあとに意識が一時的にぼんやりする", urgent: false },
    { text: "パートナーから「眠っている間ずっと足が動いている」と言われた", urgent: false },
    { text: "足のむずむず・ほてり感で眠れない（安静で悪化、歩くと改善）", urgent: false },
    { text: "「よく眠れているはずなのに」日中に強烈な眠気がある", urgent: false },
    { text: "ビクつきに加えて意識消失・転倒・けいれんがあった", urgent: true },
  ];

  checks.forEach(function(c, i) {
    var col = (i < 4) ? 0 : 1;
    var row = (i < 4) ? i : i - 4;
    var xPos = 0.3 + col * 4.8;
    var yPos = 1.05 + row * 0.98;
    var dotColor = c.urgent ? C.red : C.orange;

    addCard(s, xPos, yPos, 4.6, 0.82);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.1, y: yPos + 0.22, w: 0.35, h: 0.35, fill: { color: dotColor } });
    s.addText(c.urgent ? "!" : String(i + 1), { x: xPos + 0.1, y: yPos + 0.22, w: 0.35, h: 0.35, fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.text, { x: xPos + 0.55, y: yPos, w: 3.9, h: 0.82, fontSize: 12, fontFace: FONT_JP, color: c.urgent ? C.red : C.text, bold: c.urgent, valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.07, w: 9.4, h: 0.45, fill: { color: C.light } });
  s.addText("当てはまるものがあれば → 脳神経内科または睡眠外来へ", {
    x: 0.5, y: 5.07, w: 9.0, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: 受診のタイミングと検査
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "受診のタイミングと検査について");

  // 受診先カード
  addCard(s, 0.5, 1.1, 4.3, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.3, h: 0.5, fill: { color: C.primary } });
  s.addText("相談先 → 脳神経内科", { x: 0.7, y: 1.1, w: 3.9, h: 0.5, fontSize: 17, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "「ビクつきが気になる」相談は脳神経内科が対応", options: { bold: true, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "睡眠外来でも対応可", options: { color: C.sub, breakLine: true } },
    { text: "→ 終夜睡眠ポリグラフ（PSG）で詳しく評価", options: { color: C.sub } },
  ], { x: 0.7, y: 1.7, w: 3.9, h: 1.8, fontSize: 14, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  // PSG検査カード
  addCard(s, 5.2, 1.1, 4.5, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.5, h: 0.5, fill: { color: C.accent } });
  s.addText("終夜睡眠ポリグラフ（PSG）", { x: 5.4, y: 1.1, w: 4.1, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  var psgItems = [
    "睡眠中の脳波・筋電図・眼球運動を同時記録",
    "てんかん・RLS・PLMDとの鑑別が可能",
    "ほとんどは「生理的現象」と確認される",
    "受診で安心感 + 万一の見逃しを防ぐ",
  ];
  psgItems.forEach(function(p, i) {
    s.addText("● " + p, {
      x: 5.4, y: 1.75 + i * 0.47, w: 4.1, h: 0.45,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  // 安心メッセージ
  addCard(s, 0.5, 3.8, 9.0, 1.55);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.8, w: 0.12, h: 1.55, fill: { color: C.green } });
  s.addText([
    { text: "「何も問題なかった」も大きな価値", options: { bold: true, color: C.primary, breakLine: true } },
    { text: "ベースラインの記録として、将来の比較データになります", options: { color: C.text, breakLine: true } },
    { text: "→ 「こんなことで相談していいの？」と思ったときこそ、受診を", options: { color: C.green, bold: true } },
  ], { x: 0.9, y: 3.9, w: 8.4, h: 1.3, fontSize: 15, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

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
    { num: "1", text: "入眠時のビクッは「入眠時ミオクローヌス」。成人の60〜70%が経験する生理的現象" },
    { num: "2", text: "脳幹の網様体が覚醒→睡眠の切り替え時に誤放電することが原因。「崖から落ちる夢」はビクつきへの脳の後付け" },
    { num: "3", text: "カフェイン・睡眠不足・ストレス・激しい運動後で増強。これらの見直しが効果的" },
    { num: "4", text: "覚醒時にも起こる・毎晩何十回も繰り返す・日中に強い眠気があるときは脳神経内科へ" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.1 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.2, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.1, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.2, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "9/9");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/入眠時ミオクローヌス_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
