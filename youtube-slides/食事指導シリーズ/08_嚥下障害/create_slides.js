var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第8回 嚥下障害");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 8); }

// テーマカラー（嚥下障害＝ローズ〜ピンク系）
var DY = {
  rose: "BE185D",
  darkRose: "9D174D",
  pink: "DB2777",
  red: "DC3545",
  amber: "D97706",
  blue: "2563EB",
  teal: "0D9488",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 8,
  "嚥下障害の食事指導\n― 食形態選択と誤嚥予防の実際",
  "やわらかい食事が安全とは限らない"
);

// ============================================================
// SLIDE 2: 目的は誤嚥予防だけではない
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食事指導の目的は誤嚥予防だけではない", DY.rose);

  var goals = [
    { num: "1", title: "誤嚥や窒息のリスクを下げる", color: DY.red },
    { num: "2", title: "十分な栄養と水分を保つ", color: DY.amber },
    { num: "3", title: "できる限り経口摂取を続ける", color: C.primary },
    { num: "4", title: "食べる負担や苦痛を減らす", color: DY.blue },
  ];

  for (var i = 0; i < goals.length; i++) {
    var yPos = 1.2 + i * 0.85;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 8.0, h: 0.7, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 1.3, y: yPos + 0.1, w: 0.5, h: 0.5, fill: { color: goals[i].color } });
    s.addText(goals[i].num, { x: 1.3, y: yPos + 0.1, w: 0.5, h: 0.5, fontSize: 18, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(goals[i].title, { x: 2.1, y: yPos, w: 6.6, h: 0.7, fontSize: 20, fontFace: FJ, color: goals[i].color, bold: true, valign: "middle", margin: 0 });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.7, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("この複数の目標を同時に考える必要がある", {
    x: 1.0, y: 4.7, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 3: まず確認したいこと
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食事指導の前に確認すべき項目");

  var checks = [
    "背景疾患は何か",
    "むせるのは食べるときか食後か",
    "固形物と液体のどちらが苦手か",
    "体重減少・水分摂取量の低下はあるか",
    "肺炎を繰り返していないか",
    "口腔内の清潔・義歯の状態",
    "食事中の姿勢・一口量",
    "食べる意欲が保てているか",
  ];

  for (var i = 0; i < checks.length; i++) {
    var yPos = 1.1 + i * 0.52;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.44, rectRadius: 0.06, fill: { color: bgCol }, shadow: shd() });
    s.addText("□  " + checks[i], {
      x: 1.1, y: yPos, w: 7.8, h: 0.44,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 4: やわらかい≠安全
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「やわらかい」と「安全」は同じではない", DY.red);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: C.lightRed } });
  s.addText("食形態の判断は「噛みやすそうか」だけでは不十分", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 17, fontFace: FJ, color: DY.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var viewpoints = [
    { icon: "🔵", text: "まとまるか ― 口の中で食塊を作れるか", color: DY.rose },
    { icon: "💥", text: "ばらけないか ― 細かく散らばらないか", color: DY.pink },
    { icon: "💨", text: "速く流れ込みすぎないか ― 嚥下反射の前に流入しないか", color: DY.blue },
    { icon: "📎", text: "咽頭に張り付きすぎないか ― 残留の原因にならないか", color: DY.amber },
  ];

  for (var i = 0; i < viewpoints.length; i++) {
    var yPos = 2.1 + i * 0.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.6, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 0.1, h: 0.6, fill: { color: viewpoints[i].color } });
    s.addText(viewpoints[i].icon + "  " + viewpoints[i].text, {
      x: 1.2, y: yPos, w: 7.8, h: 0.6,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  s.addText("例: 豆腐＝崩れやすい / パン＝ぱさつく / みそ汁の具＝液体と固形が混在", {
    x: 0.8, y: 5.1, w: 8.4, h: 0.35,
    fontSize: 13, fontFace: FJ, color: C.sub, align: "center", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 5: 避けたい食品の特徴
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "避けたい食品の「特徴」で覚える");

  var traits = [
    { icon: "🏜️", trait: "ぱさつく", examples: "食パン、ゆで卵の黄身、焼き魚、いも類", color: DY.amber },
    { icon: "💥", trait: "ばらける", examples: "そぼろ、ひき肉、炒り卵、クッキー", color: DY.pink },
    { icon: "📎", trait: "張り付く", examples: "もち、団子、のり、わかめ", color: DY.red },
    { icon: "💧", trait: "液体+固形が混ざる", examples: "みそ汁の具、果物入りゼリー、シリアル入りヨーグルト", color: DY.blue },
    { icon: "🌶️", trait: "酸味・刺激", examples: "患者によっては咳込みが誘発される", color: DY.rose },
  ];

  for (var i = 0; i < traits.length; i++) {
    var yPos = 1.05 + i * 0.82;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.7, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText(traits[i].icon, { x: 0.7, y: yPos, w: 0.5, h: 0.7, fontSize: 18, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.3, y: yPos + 0.12, w: 1.8, h: 0.42, rectRadius: 0.1, fill: { color: traits[i].color } });
    s.addText(traits[i].trait, { x: 1.3, y: yPos + 0.12, w: 1.8, h: 0.42, fontSize: 14, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(traits[i].examples, { x: 3.3, y: yPos, w: 6.0, h: 0.7, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 6: とろみの使い方
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "とろみは有用だが万能ではない");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("✓ メリット", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・流れ込みが速すぎる液体を\n  調整しやすい\n・一口量を管理しやすい\n・患者によってはむせが減る", {
    x: 0.8, y: 1.85, w: 3.7, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: DY.red } });
  s.addText("⚠ デメリット", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・口腔内や咽頭に残りやすい\n・飲みにくさ → 水分摂取量↓\n・作り方で粘度がばらつく\n・嗜好性が低く嫌がる人がいる", {
    x: 5.5, y: 1.85, w: 3.7, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.4, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("「つければ安全」ではなく、適切な濃さでその人に合っているか確認", {
    x: 1.0, y: 4.4, w: 8.0, h: 0.55,
    fontSize: 16, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 7: 姿勢・一口量・食事環境
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "姿勢・一口量・食事環境の工夫", DY.rose);

  var items = [
    { icon: "🪑", title: "姿勢", detail: "頸部前屈が有効なことも。体幹の傾き、頸部過伸展は避ける", color: DY.rose },
    { icon: "🥄", title: "一口量", detail: "大きすぎると負担増。介助者が急がせないことも重要", color: DY.amber },
    { icon: "⏱️", title: "食事ペース", detail: "次の一口の前にきちんと飲み込めているか確認", color: DY.blue },
    { icon: "🧘", title: "注意集中", detail: "会話が多すぎる、TV視聴、眠気 → 誤嚥リスク↑", color: DY.teal },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 1.1 + i * 0.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.8, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 0.8, fill: { color: items[i].color } });
    s.addText(items[i].icon + " " + items[i].title, { x: 1.0, y: yPos + 0.02, w: 2.5, h: 0.35, fontSize: 18, fontFace: FJ, color: items[i].color, bold: true, margin: 0 });
    s.addText(items[i].detail, { x: 3.5, y: yPos + 0.02, w: 5.7, h: 0.7, fontSize: 15, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 8: 口腔ケアと栄養・水分管理
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "口腔ケアと栄養・水分管理も食事指導の一部");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 2.6, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: DY.rose } });
  s.addText("🦷 口腔ケア", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・口腔不潔 → 誤嚥時の細菌負荷↑\n  → 肺炎リスク↑\n・食前後の口腔ケア\n・義歯の適合確認\n・口腔乾燥への対応\n・歯科・歯科衛生士との連携", {
    x: 0.8, y: 1.85, w: 3.7, h: 1.7, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 2.6, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: DY.amber } });
  s.addText("⚖️ 栄養・水分管理", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・食形態↓ → 食欲↓ → 摂取量↓\n・常に確認:\n  □ 体重が落ちていないか\n  □ 水分摂取量は足りているか\n・とろみ嫌い → 水分不足 → 脱水\n  せん妄・便秘・腎機能悪化", {
    x: 5.5, y: 1.85, w: 3.7, h: 1.7, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.2, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("安全性だけに注目して水分不足を見逃さない", {
    x: 1.0, y: 4.2, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 9: よくある誤解
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある誤解");

  var myths = [
    { myth: "やわらかい食事なら安全", fact: "ばらける・張り付く・液体と固形が混ざるものは安全とは限らない" },
    { myth: "とろみをつければ安心", fact: "残留を増やしたり水分摂取量を落としたりすることがある" },
    { myth: "むせなければ問題ない", fact: "不顕性誤嚥ではむせがない。湿った声、食後の痰、反復する発熱も手がかり" },
    { myth: "誤嚥予防だけ考えればよい", fact: "低栄養、脱水、食べる意欲の低下も大きな問題" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.1 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.02, w: 0.5, h: 0.4, fontSize: 18, fontFace: FE, color: DY.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.02, w: 8.0, h: 0.35, fontSize: 16, fontFace: FJ, color: DY.red, bold: true, margin: 0 });
    s.addText("→ " + myths[i].fact, { x: 1.3, y: yPos + 0.42, w: 8.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.primary, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 10: 行動目標
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外来で伝える行動目標の例");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("「むせるから食べない」ではなく「どうすれば安全に食べ続けられるか」", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 15, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var goals = [
    "ばらける食品・張り付く食品を避ける",
    "一口量を少なくする",
    "飲み込んでから次の一口にする",
    "食後すぐに横にならない",
    "食前後の口腔ケアを行う",
  ];

  for (var i = 0; i < goals.length; i++) {
    var yPos = 2.0 + i * 0.65;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 7.0, h: 0.52, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 0.1, h: 0.52, fill: { color: C.accent } });
    s.addText("▸ " + goals[i], {
      x: 1.9, y: yPos, w: 6.4, h: 0.52,
      fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 11-13: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 8, [
  "やわらかいことと安全なことは同じではない",
  "食形態はまとまり・ばらつき・流れやすさ・張り付きで考える",
  "とろみは有用だが万能ではない ― 水分不足に注意",
  "姿勢・一口量・口腔ケアまで含めて調整する",
  "誤嚥予防だけでなく栄養・水分確保との両立が最重要",
]);

T.addNextEpisodeSlide(pres, 9, "消化管術後の食事指導\n― 少量頻回と症状別対応の考え方");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
