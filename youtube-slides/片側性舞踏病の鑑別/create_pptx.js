var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "片側性舞踏病の鑑別 ─ 脳梗塞・高血糖・薬剤性を見逃さない";

// --- Color Palette ---
var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "E8850C",
  light: "E8F4FD",
  warmBg: "F8F9FA",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "DC3545",
  green: "28A745",
  yellow: "F5C518",
  lightRed: "F8D7DA",
  lightYellow: "FFF3CD",
  lightGreen: "D4EDDA",
  purple: "7B2D8E",
};

var FJ = "BIZ UDPGothic";
var FE = "Segoe UI";

function shd() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function hdr(s, title, bgColor) {
  bgColor = bgColor || C.primary;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: bgColor } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 26, fontFace: FJ, color: C.white, bold: true, margin: 0 });
}

function card(s, x, y, w, h, fillColor, borderColor) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: fillColor || C.white }, shadow: shd(), rectRadius: 0.1 });
  if (borderColor) {
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.12, h: h, fill: { color: borderColor }, rectRadius: 0.05 });
  }
}

function ftr(s) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.38, fill: { color: C.dark } });
  s.addText("医知創造ラボ", { x: 0.4, y: 5.27, w: 3, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("🧠", { x: 0.6, y: 0.5, w: 8.8, h: 0.8, fontSize: 48, align: "center", margin: 0 });
  s.addText("片側性舞踏病の鑑別", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 38, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("脳梗塞・高血糖・薬剤性を見逃さない", {
    x: 0.6, y: 3.0, w: 8.8, h: 0.6,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.6,
    fontSize: 14, fontFace: FJ, color: "B0BEC5", align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 片側性舞踏病とは
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "片側性舞踏病とは何か");

  card(s, 0.3, 1.2, 4.5, 3.8, C.white, C.primary);
  s.addText("定義", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var items = [
    "片側の四肢・顔面に出る不規則な不随意運動",
    "素早く、非律動的で、流れるように移動",
    "近位筋優位で激しい型 = ヘミバリスム",
    "hemichorea ↔ hemiballismus は連続スペクトラム",
  ];
  items.forEach(function(t, i) {
    s.addText("• " + t, { x: 0.6, y: 1.85 + i * 0.55, w: 4.0, h: 0.5, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.1, 1.2, 4.5, 3.8, C.light, C.accent);
  s.addText("臨床的ポイント", { x: 5.3, y: 1.3, w: 4.1, h: 0.4, fontSize: 18, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("「病名」ではなく「症候」", { x: 5.4, y: 1.9, w: 4.0, h: 0.5, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("→ 背景に原因疾患がある", { x: 5.4, y: 2.4, w: 4.0, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("片側に限局 → 対側の脳に\n構造病変・代謝異常を疑う", { x: 5.4, y: 3.0, w: 4.0, h: 0.7, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("急性〜亜急性発症では\n後天性・可逆的な原因を優先", { x: 5.4, y: 3.8, w: 4.0, h: 0.7, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });

  ftr(s);
})();

// ============================================================
// SLIDE 3: 主な原因一覧
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "片側性舞踏病の主な原因");

  var causes = [
    { icon: "🩸", name: "脳血管障害", detail: "最多 68%\n脳梗塞・脳出血", color: C.red },
    { icon: "📈", name: "高血糖関連", detail: "糖尿病性線条体症\n高齢女性に多い", color: C.accent },
    { icon: "💊", name: "薬剤性", detail: "D2遮断薬、制吐薬\nレボドパなど", color: C.purple },
    { icon: "🛡️", name: "自己免疫性", detail: "抗リン脂質抗体症候群\nSLE、Sydenham", color: C.green },
    { icon: "🧬", name: "変性・遺伝性", detail: "ハンチントン病\n高齢初発は稀", color: C.sub },
  ];

  causes.forEach(function(c, i) {
    var xPos = 0.2 + i * 1.96;
    card(s, xPos, 1.2, 1.8, 3.6, C.white, c.color);
    s.addText(c.icon, { x: xPos, y: 1.35, w: 1.8, h: 0.5, fontSize: 28, align: "center", margin: 0 });
    s.addText(c.name, { x: xPos + 0.15, y: 1.9, w: 1.5, h: 0.45, fontSize: 15, fontFace: FJ, color: c.color, bold: true, align: "center", margin: 0 });
    s.addText(c.detail, { x: xPos + 0.15, y: 2.5, w: 1.5, h: 1.2, fontSize: 12, fontFace: FJ, color: C.text, align: "center", valign: "top", margin: 0 });
  });

  s.addText("成人発症散発性舞踏病の68%が脳血管障害（Bovenzi 2021）", {
    x: 0.3, y: 5.0, w: 9.4, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, italic: true, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 4: 脳血管障害（見逃し注意）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "🩸 見逃し注意①：脳血管障害", C.red);

  card(s, 0.3, 1.2, 4.5, 2.0, C.white, C.red);
  s.addText("古典的な責任病巣", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var classic = ["対側 視床下核", "被殻・尾状核", "視床"];
  classic.forEach(function(t, i) {
    s.addText("• " + t, { x: 0.6, y: 1.8 + i * 0.4, w: 4.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.1, 1.2, 4.5, 2.0, C.lightRed, C.red);
  s.addText("新しい知見", { x: 5.3, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("基底核が保たれた皮質梗塞でも\nhemichoreaが起こる！", { x: 5.4, y: 1.85, w: 4.0, h: 0.6, fontSize: 14, fontFace: FJ, color: C.text, bold: true, margin: 0 });
  s.addText("頭頂葉・島皮質・前頭葉\n（Carbayo 2020）", { x: 5.4, y: 2.5, w: 4.0, h: 0.5, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0 });

  card(s, 0.3, 3.5, 9.3, 1.4, C.lightYellow);
  s.addText("⚠ 「基底核に異常がない＝血管性を否定」にはならない", { x: 0.5, y: 3.6, w: 9.0, h: 0.4, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("DWI再読影：基底核だけでなく頭頂葉・島皮質・centrum semiovaleまで確認", { x: 0.5, y: 4.1, w: 9.0, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("突然発症なら急性脳卒中として対応", { x: 0.5, y: 4.5, w: 9.0, h: 0.3, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });

  ftr(s);
})();

// ============================================================
// SLIDE 5: 高血糖性舞踏病
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "📈 見逃し注意②：高血糖性舞踏病", C.accent);

  card(s, 0.3, 1.2, 4.5, 2.4, C.white, C.accent);
  s.addText("糖尿病性線条体症の特徴", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var features = [
    "高齢女性に多い",
    "88.1%がhemichorea/hemiballismus",
    "平均血糖 414 mg/dL",
    "平均HbA1c 13.1%",
    "22%で糖尿病の初発所見",
  ];
  features.forEach(function(t, i) {
    s.addText("• " + t, { x: 0.6, y: 1.8 + i * 0.4, w: 4.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.1, 1.2, 4.5, 2.4, C.white, C.primary);
  s.addText("画像所見", { x: 5.3, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("CT：対側線条体の高吸収", { x: 5.4, y: 1.85, w: 4.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("MRI：T1高信号（対側線条体）", { x: 5.4, y: 2.25, w: 4.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("CT/MRI不一致 ≒ 1/6", { x: 5.4, y: 2.75, w: 4.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("MRI正常例 5.2%あり", { x: 5.4, y: 3.15, w: 4.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.red, bold: true, margin: 0 });

  card(s, 0.3, 3.9, 9.3, 1.0, C.lightYellow);
  s.addText("⚠ 糖尿病の既往がなくても22%が舞踏病で糖尿病と判明 → 必ず血糖・HbA1cを確認", {
    x: 0.5, y: 4.0, w: 9.0, h: 0.4, fontSize: 14, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });
  s.addText("画像が非典型でも高血糖性舞踏病は除外しきれない（Chua 2020, Hoffmeister 2024）", {
    x: 0.5, y: 4.45, w: 9.0, h: 0.35, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 6: 薬剤性・自己免疫性
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "💊🛡️ 見逃し注意③：薬剤性・自己免疫性");

  card(s, 0.3, 1.2, 4.5, 3.8, C.white, C.purple);
  s.addText("確認すべき薬剤", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.purple, bold: true, margin: 0 });
  var drugs = [
    "抗精神病薬（遅発性ジスキネジア）",
    "メトクロプラミド・ドンペリドン",
    "レボドパ・ドパミンアゴニスト",
    "抗てんかん薬（フェニトイン等）",
    "抗うつ薬",
    "認知症治療薬の変更",
    "最近の中止薬も含める",
  ];
  drugs.forEach(function(t, i) {
    s.addText("• " + t, { x: 0.6, y: 1.8 + i * 0.42, w: 4.0, h: 0.38, fontSize: 12.5, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.1, 1.2, 4.5, 3.8, C.white, C.green);
  s.addText("自己免疫性を疑う検査", { x: 5.3, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  var autoimmune = [
    "抗核抗体（ANA）",
    "抗カルジオリピン抗体",
    "抗β2GPI抗体",
    "ループスアンチコアグラント",
  ];
  autoimmune.forEach(function(t, i) {
    s.addText("• " + t, { x: 5.4, y: 1.85 + i * 0.45, w: 4.0, h: 0.4, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  s.addText("疑うヒント：", { x: 5.4, y: 3.7, w: 4.0, h: 0.3, fontSize: 13, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("血栓症既往、livedo、\n自己免疫疾患既往", { x: 5.4, y: 4.0, w: 4.0, h: 0.6, fontSize: 12, fontFace: FJ, color: C.text, margin: 0 });

  ftr(s);
})();

// ============================================================
// SLIDE 7: 鑑別の進め方（ステップ）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "鑑別の進め方 ─ 3ステップ");

  // Step 1
  card(s, 0.3, 1.2, 3.0, 3.8, C.white, C.primary);
  s.addText("Step 1", { x: 0.5, y: 1.3, w: 2.6, h: 0.35, fontSize: 14, fontFace: FE, color: C.primary, bold: true, margin: 0 });
  s.addText("可逆的原因を外す", { x: 0.5, y: 1.65, w: 2.6, h: 0.35, fontSize: 15, fontFace: FJ, color: C.text, bold: true, margin: 0 });
  var step1 = ["血糖・HbA1c", "電解質（Ca, Mg）", "腎・肝機能", "甲状腺機能", "薬歴の総点検", "頭部MRI（DWI）", "MRA"];
  step1.forEach(function(t, i) {
    s.addText("• " + t, { x: 0.55, y: 2.15 + i * 0.37, w: 2.5, h: 0.33, fontSize: 11.5, fontFace: FJ, color: C.text, margin: 0 });
  });

  // Arrow 1
  s.addText("→", { x: 3.3, y: 2.8, w: 0.4, h: 0.4, fontSize: 24, fontFace: FE, color: C.accent, bold: true, align: "center", margin: 0 });

  // Step 2
  card(s, 3.65, 1.2, 3.0, 3.8, C.white, C.accent);
  s.addText("Step 2", { x: 3.85, y: 1.3, w: 2.6, h: 0.35, fontSize: 14, fontFace: FE, color: C.accent, bold: true, margin: 0 });
  s.addText("血管性を詰める", { x: 3.85, y: 1.65, w: 2.6, h: 0.35, fontSize: 15, fontFace: FJ, color: C.text, bold: true, margin: 0 });
  var step2 = ["脳梗塞リスク評価", "心房細動検索", "頸動脈エコー", "心エコー", "ラクナ・白質病変確認"];
  step2.forEach(function(t, i) {
    s.addText("• " + t, { x: 3.9, y: 2.15 + i * 0.37, w: 2.5, h: 0.33, fontSize: 11.5, fontFace: FJ, color: C.text, margin: 0 });
  });

  // Arrow 2
  s.addText("→", { x: 6.65, y: 2.8, w: 0.4, h: 0.4, fontSize: 24, fontFace: FE, color: C.accent, bold: true, align: "center", margin: 0 });

  // Step 3
  card(s, 7.0, 1.2, 2.7, 3.8, C.white, C.green);
  s.addText("Step 3", { x: 7.2, y: 1.3, w: 2.3, h: 0.35, fontSize: 14, fontFace: FE, color: C.green, bold: true, margin: 0 });
  s.addText("免疫・腫瘍随伴", { x: 7.2, y: 1.65, w: 2.3, h: 0.35, fontSize: 15, fontFace: FJ, color: C.text, bold: true, margin: 0 });
  var step3 = ["ANA・aPL系", "悪性腫瘍検索", "自己抗体パネル", "髄液検査"];
  step3.forEach(function(t, i) {
    s.addText("• " + t, { x: 7.25, y: 2.15 + i * 0.37, w: 2.2, h: 0.33, fontSize: 11.5, fontFace: FJ, color: C.text, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 8: 他の不随意運動との鑑別
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "本当に舞踏病か？ ─ 他の不随意運動との鑑別");

  // テーブル風レイアウト
  var rows = [
    { name: "ジストニア", feature: "同じ姿勢・方向に捻れる", diff: "舞踏病は不規則・流動的" },
    { name: "ミオクローヌス", feature: "一瞬のピクつき", diff: "舞踏病はやや持続的で移動" },
    { name: "アカシジア", feature: "内的焦燥感で動く", diff: "舞踏病は無意識に起こる" },
    { name: "てんかん性", feature: "反復的・常同性が強い", diff: "舞踏病は非常同的" },
    { name: "Pseudoathetosis", feature: "深部感覚障害に伴う", diff: "閉眼で増悪するのが特徴" },
  ];

  // ヘッダー行
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.15, w: 9.4, h: 0.45, fill: { color: C.primary } });
  s.addText("不随意運動", { x: 0.4, y: 1.18, w: 2.5, h: 0.4, fontSize: 13, fontFace: FJ, color: C.white, bold: true, margin: 0 });
  s.addText("特徴", { x: 2.9, y: 1.18, w: 3.2, h: 0.4, fontSize: 13, fontFace: FJ, color: C.white, bold: true, margin: 0 });
  s.addText("舞踏病との違い", { x: 6.3, y: 1.18, w: 3.2, h: 0.4, fontSize: 13, fontFace: FJ, color: C.white, bold: true, margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.65 + i * 0.55;
    var bgCol = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.5, fill: { color: bgCol } });
    s.addText(r.name, { x: 0.4, y: yPos + 0.02, w: 2.5, h: 0.45, fontSize: 13, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(r.feature, { x: 2.9, y: yPos + 0.02, w: 3.2, h: 0.45, fontSize: 12, fontFace: FJ, color: C.text, margin: 0 });
    s.addText(r.diff, { x: 6.3, y: yPos + 0.02, w: 3.2, h: 0.45, fontSize: 12, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 0.3, 4.5, 9.3, 0.6, C.lightGreen);
  s.addText("💡 判断に迷うときは動画記録が有用。睡眠で消えるか、歩行で悪化するかもチェック", {
    x: 0.5, y: 4.55, w: 9.0, h: 0.45, fontSize: 14, fontFace: FJ, color: C.text, bold: true, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 9: MRI/CTの落とし穴
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "MRI・CTの落とし穴", C.red);

  card(s, 0.3, 1.2, 4.5, 1.8, C.white, C.primary);
  s.addText("血管性", { x: 0.5, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("• DWI：対側基底核（被殻・尾状核）\n• 皮質梗塞のみで発症する例あり\n• 慢性虚血変化に新鮮梗塞を埋もれさせない", {
    x: 0.6, y: 1.75, w: 4.0, h: 1.1, fontSize: 12.5, fontFace: FJ, color: C.text, margin: 0,
  });

  card(s, 5.1, 1.2, 4.5, 1.8, C.white, C.accent);
  s.addText("高血糖性", { x: 5.3, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("• CT：線条体高吸収\n• MRI：T1高信号（対側線条体）\n• 治療後に画像も改善する", {
    x: 5.4, y: 1.75, w: 4.0, h: 1.1, fontSize: 12.5, fontFace: FJ, color: C.text, margin: 0,
  });

  // 落とし穴ハイライト
  card(s, 0.3, 3.3, 9.3, 1.8, C.lightRed, C.red);
  s.addText("⚠ 画像の落とし穴", { x: 0.5, y: 3.4, w: 9.0, h: 0.4, fontSize: 18, fontFace: FJ, color: C.red, bold: true, margin: 0 });

  var pitfalls = [
    { label: "CT/MRI不一致", value: "≒ 1/6", source: "Chua 2020" },
    { label: "症状と画像の時間的ずれ", value: "≒ 10%", source: "Chua 2020" },
    { label: "MRI正常例", value: "5.2%", source: "Hoffmeister 2024" },
  ];
  pitfalls.forEach(function(p, i) {
    var xPos = 0.5 + i * 3.1;
    s.addText(p.label, { x: xPos, y: 3.9, w: 2.9, h: 0.35, fontSize: 13, fontFace: FJ, color: C.text, bold: true, margin: 0 });
    s.addText(p.value, { x: xPos, y: 4.3, w: 2.9, h: 0.4, fontSize: 20, fontFace: FE, color: C.red, bold: true, margin: 0 });
    s.addText(p.source, { x: xPos, y: 4.7, w: 2.9, h: 0.25, fontSize: 10, fontFace: FJ, color: C.sub, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 10: 治療
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "治療の考え方");

  card(s, 0.3, 1.2, 5.5, 2.2, C.white, C.primary);
  s.addText("原則：原因治療が先", { x: 0.5, y: 1.3, w: 5.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var txItems = [
    "脳血管障害 → 脳卒中として対応",
    "高血糖性 → 厳格な血糖是正（インスリン・補液）",
    "薬剤性 → 原因薬の中止・変更",
    "自己免疫性 → 免疫抑制療法",
  ];
  txItems.forEach(function(t, i) {
    s.addText("• " + t, { x: 0.6, y: 1.8 + i * 0.38, w: 5.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 6.1, 1.2, 3.5, 2.2, C.light, C.accent);
  s.addText("対症療法", { x: 6.3, y: 1.3, w: 3.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("ハロペリドール\n0.375〜0.75 mg/日", { x: 6.3, y: 1.85, w: 3.1, h: 0.55, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("クロナゼパム\n0.5〜1 mg/日", { x: 6.3, y: 2.5, w: 3.1, h: 0.55, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("(Shiraiwa 2020)", { x: 6.3, y: 3.1, w: 3.1, h: 0.25, fontSize: 10, fontFace: FJ, color: C.sub, margin: 0 });

  card(s, 0.3, 3.7, 9.3, 1.3, C.lightYellow);
  s.addText("⚠ 高齢者での注意点", { x: 0.5, y: 3.8, w: 9.0, h: 0.35, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var warnings = ["傾眠・転倒リスク", "嚥下低下・誤嚥", "薬剤性パーキンソニズム", "QT延長", "認知機能悪化"];
  warnings.forEach(function(w, i) {
    var xPos = 0.5 + i * 1.85;
    s.addText("• " + w, { x: xPos, y: 4.25, w: 1.8, h: 0.6, fontSize: 12, fontFace: FJ, color: C.text, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 11: 予後
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "予後 ─ 可逆的原因が多い");

  // 82%改善の大きな数字
  card(s, 0.3, 1.2, 4.5, 2.5, C.lightGreen, C.green);
  s.addText("82%", { x: 0.5, y: 1.4, w: 4.1, h: 1.0, fontSize: 56, fontFace: FE, color: C.green, bold: true, align: "center", margin: 0 });
  s.addText("自然軽快 or 軽微な後遺症", { x: 0.5, y: 2.5, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.text, bold: true, align: "center", margin: 0 });
  s.addText("Bovenzi 2021", { x: 0.5, y: 3.0, w: 4.1, h: 0.3, fontSize: 11, fontFace: FJ, color: C.sub, align: "center", margin: 0 });

  card(s, 5.1, 1.2, 4.5, 2.5, C.white, C.primary);
  s.addText("原因別の注意点", { x: 5.3, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var progItems = [
    "高血糖性：血糖管理で多くが改善",
    "ただし再発率 ≒ 28%（Chua 2020）",
    "血管性：二次予防が長期予後を左右",
    "薬剤性：遅発性ジスキネジアは不可逆的な場合も",
  ];
  progItems.forEach(function(t, i) {
    s.addText("• " + t, { x: 5.4, y: 1.85 + i * 0.45, w: 4.0, h: 0.4, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 0.3, 4.0, 9.3, 0.9, C.light);
  s.addText("💡 迅速な原因検索が予後を左右する ─ 「原因不明」で終わらせず可逆的原因を執拗に探す", {
    x: 0.5, y: 4.1, w: 9.0, h: 0.6, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ─ Take Home Message", C.dark);

  var messages = [
    { num: "1", text: "片側性舞踏病は「病名」ではなく「症候」\n→ 必ず原因疾患を探す", color: C.primary },
    { num: "2", text: "成人発症の68%が脳血管障害\n→ 画像検査を最優先に", color: C.red },
    { num: "3", text: "高血糖性舞踏病を必ず除外\n→ 糖尿病の既往がなくてもHbA1cを確認", color: C.accent },
    { num: "4", text: "画像の落とし穴に注意\n→ 皮質梗塞、MRI偽陰性、CT/MRI不一致", color: C.purple },
    { num: "5", text: "82%が改善する可逆的病態\n→ 迅速な原因検索が予後を決める", color: C.green },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.1 + i * 0.82;
    card(s, 0.3, yPos, 9.4, 0.72, C.white, m.color);
    s.addShape(pres.shapes.OVAL, { x: 0.55, y: yPos + 0.1, w: 0.5, h: 0.5, fill: { color: m.color } });
    s.addText(m.num, { x: 0.55, y: yPos + 0.1, w: 0.5, h: 0.5, fontSize: 18, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.3, y: yPos + 0.05, w: 8.2, h: 0.6, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 13: 参考文献
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "参考文献", C.dark);

  var refs = [
    "1. Bovenzi R, et al. Neurol Sci. 2021;42(12):5187-5196.",
    "2. Cincotta M, Walker RH. Tremor Other Hyperkinet Mov. 2022;12:3.",
    "3. Carbayo Á, et al. Cerebrovasc Dis. 2020;49(5):498-505.",
    "4. Chua CB, et al. Sci Rep. 2020;10(1):1594.",
    "5. Hoffmeister MC, et al. J Endocrinol Invest. 2024;47(1):21-34.",
    "6. Chen Y, et al. Front Neurol. 2022;13:956470.",
    "7. Shiraiwa N, et al. J Neurol Sci. 2020;417:117033.",
  ];

  refs.forEach(function(r, i) {
    s.addText(r, { x: 0.5, y: 1.1 + i * 0.48, w: 9.0, h: 0.4, fontSize: 12, fontFace: FJ, color: C.text, margin: 0 });
  });

  s.addText("ご視聴ありがとうございました", { x: 0.5, y: 4.7, w: 9.0, h: 0.4, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, align: "center", margin: 0 });

  ftr(s);
})();

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/hemichorea_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
