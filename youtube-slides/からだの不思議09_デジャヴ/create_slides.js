var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "デジャヴの正体を神経科学で解説 ― 脳神経内科医が答えるからだの不思議 #09";

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

  s.addText("脳神経内科医が答える からだの不思議 #09", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("デジャヴの正体を\n神経科学で解説", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 38, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 「既視感」はなぜ起きるのか ―", {
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
    "「初めて来た場所なのに、なぜか\n  『ここに来たことがある』と感じた」",
    "「初対面の人なのに、どこかで\n  会ったことがある気がした」",
    "「これから何が起きるか、知っている気がした……\n  でも実際には初めてのはずなのに」",
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
  s.addText("健康な人の 60〜80% が生涯に一度は経験する", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/10");
})();

// ============================================================
// SLIDE 3: デジャヴとは ― 定義と特徴
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "デジャヴとは ― 定義と特徴");

  addCard(s, 0.5, 1.1, 9.0, 1.5);
  s.addText("déjà vu（デジャヴ）", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.5,
    fontSize: 22, fontFace: FONT_EN, color: C.primary, bold: true, italic: true, margin: 0,
  });
  s.addText("フランス語で「既に見た」。Brown（2004）の定義：\n「以前に経験したという感覚を伴うが、いつ・どこで・どのように経験したかは思い出せない現象」", {
    x: 0.8, y: 1.65, w: 8.4, h: 0.9,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2,
  });

  var features = [
    { label: "持続時間", val: "数秒〜十数秒", color: C.accent },
    { label: "メタ認知", val: "「錯覚だ」と\nわかる", color: C.green },
    { label: "有病率", val: "生涯60〜80%", color: C.primary },
  ];
  features.forEach(function(f, i) {
    var xPos = 0.5 + i * 3.1;
    addCard(s, xPos, 2.9, 2.9, 1.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.9, w: 2.9, h: 0.5, fill: { color: f.color } });
    s.addText(f.label, { x: xPos, y: 2.9, w: 2.9, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.val, { x: xPos, y: 3.45, w: 2.9, h: 0.85, fontSize: 18, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.65, w: 9.0, h: 0.6, fill: { color: C.light } });
  s.addText("幻覚との違い：デジャヴ中も「これは錯覚だ」という自己認識（メタ認知）は保たれる", {
    x: 0.8, y: 4.65, w: 8.4, h: 0.6,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "3/10");
})();

// ============================================================
// SLIDE 4: 3つの有力仮説 ― 概要
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "デジャヴの神経科学 ― 3つの有力仮説");

  var theories = [
    { num: "1", title: "二重処理理論", subtitle: "Dual-Processing Theory", desc: "符号化とファミリアリティ評価の\n「わずかなズレ」が原因", color: C.accent },
    { num: "2", title: "パターン補完理論", subtitle: "Hologram / Pattern-Completion", desc: "過去の記憶断片と現在の情報が\n一致 → 記憶全体を誤って引き出す", color: C.green },
    { num: "3", title: "記憶回路ミスマッチ説", subtitle: "MTL Circuit Mismatch", desc: "海馬・嗅内皮質の入力経路と\n記憶評価回路のタイムラグ", color: C.orange },
  ];

  theories.forEach(function(t, i) {
    var xPos = 0.2 + i * 3.25;
    addCard(s, xPos, 1.2, 3.1, 3.8);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 3.1, h: 0.55, fill: { color: t.color } });
    s.addShape(pres.shapes.OVAL, { x: xPos + 1.25, y: 1.85, w: 0.6, h: 0.6, fill: { color: t.color } });
    s.addText(t.num, { x: xPos + 1.25, y: 1.85, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.title, { x: xPos, y: 1.2, w: 3.1, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.subtitle, { x: xPos + 0.1, y: 2.55, w: 2.9, h: 0.5, fontSize: 11, fontFace: FONT_EN, color: t.color, bold: true, align: "center", italic: true, margin: 0 });
    s.addText(t.desc, { x: xPos + 0.1, y: 3.1, w: 2.9, h: 1.7, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.1, w: 9.0, h: 0.4, fill: { color: C.dark } });
  s.addText("共通点：内側側頭葉（海馬・嗅内皮質）の記憶評価システムの一時的な誤作動", {
    x: 0.8, y: 5.1, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "4/10");
})();

// ============================================================
// SLIDE 5: 二重処理理論の詳細
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "① 二重処理理論（Dual-Processing Theory）");

  addCard(s, 0.5, 1.1, 9.0, 1.4);
  s.addText("通常：符号化（エンコード）とファミリアリティ評価は「同時」に行われる", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("疲労・注意散漫・ストレス時 → 無意識の処理が先行し、\n意識的知覚が「追いついた」ときに脳が「処理済み＝既知」と誤認する", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.8,
    fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2,
  });

  // プロセス図
  var steps = [
    { label: "無意識の\n情報処理", color: C.accent },
    { label: "わずかな\nタイムラグ", color: C.yellow },
    { label: "意識的な\n知覚", color: C.primary },
    { label: "「既知」と\n誤判定", color: C.red },
  ];
  steps.forEach(function(st, i) {
    var xPos = 0.6 + i * 2.25;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xPos, y: 2.8, w: 2.0, h: 1.0, fill: { color: st.color }, rectRadius: 0.1 });
    s.addText(st.label, { x: xPos, y: 2.8, w: 2.0, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: (st.color === C.yellow) ? C.text : C.white, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    if (i < 3) {
      s.addText("→", { x: xPos + 2.05, y: 2.95, w: 0.2, h: 0.7, fontSize: 22, color: C.sub, align: "center", valign: "middle", margin: 0 });
    }
  });

  addCard(s, 0.5, 4.1, 9.0, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.1, w: 0.12, h: 1.3, fill: { color: C.accent } });
  s.addText([
    { text: "誘発しやすい状況", options: { bold: true, color: C.primary, breakLine: true } },
    { text: "睡眠不足・過労・強いストレス・注意散漫な状態", options: { color: C.text, breakLine: true } },
    { text: "→ 記憶評価の同期が崩れやすくなる", options: { color: C.sub } },
  ], { x: 0.8, y: 4.2, w: 8.4, h: 1.1, fontSize: 14, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "5/10");
})();

// ============================================================
// SLIDE 6: 海馬とパターン補完・内側側頭葉
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "② & ③ 海馬と内側側頭葉の役割");

  var regions = [
    { name: "嗅内皮質", role: "感覚情報の\n「関門」\n大脳皮質→海馬へ", icon: "🚪" },
    { name: "海馬", role: "パターン補完\n記憶の新旧判断\n全体記憶の復元", icon: "🧠" },
    { name: "海馬CA1野", role: "新旧判断の\n最終チェック\n誤判定でデジャヴ", icon: "⚡" },
  ];
  regions.forEach(function(r, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.2, 3.0, 2.4);
    s.addText(r.icon, { x: xPos, y: 1.3, w: 3.0, h: 0.65, fontSize: 30, align: "center", margin: 0 });
    s.addText(r.name, { x: xPos, y: 1.95, w: 3.0, h: 0.55, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0 });
    s.addText(r.role, { x: xPos, y: 2.5, w: 3.0, h: 1.0, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
  });

  s.addText("→", { x: 3.1, y: 1.9, w: 0.4, h: 0.55, fontSize: 26, color: C.accent, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 6.3, y: 1.9, w: 0.4, h: 0.55, fontSize: 26, color: C.accent, align: "center", valign: "middle", margin: 0 });

  addCard(s, 0.5, 3.85, 9.0, 1.55);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.85, w: 9.0, h: 0.5, fill: { color: C.primary } });
  s.addText("パターン補完とデジャヴ", { x: 0.8, y: 3.85, w: 8.4, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "現在の場面の断片 + 過去の記憶断片が類似", options: { fontSize: 15, bold: true, color: C.orange, breakLine: true } },
    { text: "→ 海馬が過去の記憶全体を補完・引き出す → 「知っている」という感覚が生まれる", options: { fontSize: 14, color: C.text } },
  ], { x: 0.8, y: 4.4, w: 8.4, h: 0.9, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  addPageNum(s, "6/10");
})();

// ============================================================
// SLIDE 7: 年齢別・状況別の頻度
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "誰に・いつ起きやすいか");

  var rows = [
    { factor: "年齢", detail: "15〜25歳でもっとも頻繁。加齢とともに減少し、高齢者では半分以下に" },
    { factor: "疲労・ストレス", detail: "睡眠不足・過労・強いストレス下で起きやすい" },
    { factor: "旅行・新環境", detail: "未知の環境・初めての場所への曝露が多いほど頻度が高い" },
    { factor: "想像力・感受性", detail: "想像力が豊かで感受性の高い人で経験頻度が高い傾向" },
    { factor: "メディア視聴", detail: "フィクションに多く触れる人でやや頻度が高い" },
  ];

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.05, w: 2.2, h: 0.5, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 1.05, w: 7.2, h: 0.5, fill: { color: C.primary } });
  s.addText("要因", { x: 0.3, y: 1.05, w: 2.2, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("内容", { x: 2.5, y: 1.05, w: 7.2, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.55 + i * 0.78;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.78, fill: { color: bgColor } });
    s.addText(r.factor, { x: 0.3, y: yPos, w: 2.2, h: 0.78, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.detail, { x: 2.5, y: yPos, w: 7.1, h: 0.78, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6] });
  });

  addPageNum(s, "7/10");
})();

// ============================================================
// SLIDE 8: 側頭葉てんかんとの関係
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんなデジャヴは要注意 ― 側頭葉てんかん");

  addCard(s, 0.5, 1.1, 9.0, 1.1);
  s.addText("側頭葉てんかんの一部では、発作の「前兆（オーラ）」としてデジャヴが起こる", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("てんかん放電が内側側頭葉（海馬・嗅内皮質）を巻き込み、記憶評価回路を異常活性化させるため", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.35, w: 9.0, h: 0.42, fill: { color: C.red } });
  s.addText("医師への相談が必要なサイン", { x: 0.8, y: 2.35, w: 8.4, h: 0.42, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var signs = [
    "頻度が急に増えた（年数回→週・毎日）",
    "デジャヴ後に意識がぼんやり、手足がピクつく",
    "腹部の不快感・上腹感覚・口をモグモグする動作",
    "デジャヴ中に時間感覚がなくなり、数分間の記憶が抜ける",
    "明らかな誘因なく一定パターンで繰り返す",
  ];
  signs.forEach(function(sg, i) {
    var yPos = 2.82 + i * 0.45;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: yPos + 0.1, w: 0.25, h: 0.25, fill: { color: C.red } });
    s.addText(sg, { x: 1.0, y: yPos, w: 8.6, h: 0.42, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  addPageNum(s, "8/10");
})();

// ============================================================
// SLIDE 9: 健康なデジャヴ vs 要注意のデジャヴ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "健康なデジャヴ vs 要注意のデジャヴ");

  var rows = [
    { feature: "頻度", normal: "月に数回以下、不規則", warning: "週に何度も、あるいは毎日繰り返す" },
    { feature: "随伴症状", normal: "なし", warning: "腹部不快感・自動症（口・手の反復動作）・意識減損" },
    { feature: "持続時間", normal: "数秒〜十数秒", warning: "数十秒〜数分（記憶が抜けることも）" },
    { feature: "メタ認知", normal: "「錯覚だ」とわかる", warning: "発作中は気づけないことがある" },
    { feature: "誘因", normal: "疲労・旅行など", warning: "明らかな誘因なく一定パターンで繰り返す" },
  ];

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.05, w: 1.8, h: 0.52, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.1, y: 1.05, w: 3.7, h: 0.52, fill: { color: C.green } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: 1.05, w: 3.9, h: 0.52, fill: { color: C.red } });
  s.addText("特徴", { x: 0.3, y: 1.05, w: 1.8, h: 0.52, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("健康なデジャヴ", { x: 2.1, y: 1.05, w: 3.7, h: 0.52, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("要注意のデジャヴ", { x: 5.8, y: 1.05, w: 3.9, h: 0.52, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.57 + i * 0.78;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.78, fill: { color: bgColor } });
    s.addText(r.feature, { x: 0.3, y: yPos, w: 1.8, h: 0.78, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.normal, { x: 2.1, y: yPos, w: 3.7, h: 0.78, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6] });
    s.addText(r.warning, { x: 5.8, y: yPos, w: 3.9, h: 0.78, fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: [0, 6, 0, 6] });
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
    x: 0.5, y: 0.3, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "デジャヴは健康な人の60〜80%が経験する正常な現象\n― 脳の記憶処理システムの一時的な誤作動" },
    { num: "2", text: "主な原因は内側側頭葉（海馬・嗅内皮質）の\n記憶評価と入力情報のミスマッチ" },
    { num: "3", text: "15〜25歳に多く、疲労・旅行・ストレスで起きやすい" },
    { num: "4", text: "頻度の急増・随伴症状（意識減損・自動症）がある場合は\n側頭葉てんかんの可能性 → 脳神経内科へ" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.1 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.2, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.1, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  |  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.2, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "10/10");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/デジャヴ_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
