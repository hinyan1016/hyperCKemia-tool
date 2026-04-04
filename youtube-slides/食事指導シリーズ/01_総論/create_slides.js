var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第1回 総論");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 1); }

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 1,
  "食事指導は「何を禁止するか」ではなく\n「何を守るか」から始める",
  "外来で土台になる考え方を整理する"
);

// ============================================================
// SLIDE 2: よくある食事指導
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある食事指導");

  var phrases = [
    { text: "「血圧が高いので塩分を減らしましょう」", color: "DC3545" },
    { text: "「血糖が高いので甘いものを控えましょう」", color: "E67E22" },
    { text: "「コレステロールが高いので油ものはやめましょう」", color: "8E44AD" },
  ];
  for (var i = 0; i < phrases.length; i++) {
    var yPos = 1.3 + i * 1.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 8.0, h: 0.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 0.12, h: 0.8, rectRadius: 0, fill: { color: phrases[i].color } });
    s.addText(phrases[i].text, {
      x: 1.4, y: yPos, w: 7.4, h: 0.8,
      fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: 4.7, w: 7.0, h: 0.5, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("間違いではないが、これだけでは食事指導はうまくいかない", {
    x: 1.5, y: 4.7, w: 7.0, h: 0.5,
    fontSize: 16, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 3: なぜうまくいかないのか
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食事指導がうまくいかない3つの理由");

  var reasons = [
    { num: "1", title: "説明が抽象的", detail: "「塩分を控えて」→みそ汁は？漬物は？\n「甘いものを控えて」→主食は減らすの？", color: "DC3545" },
    { num: "2", title: "制限ばかりが前に出る", detail: "「ダメなものを増やす話」になり\n毎日のことなので長く続かない", color: "E67E22" },
    { num: "3", title: "低栄養の見落とし", detail: "制限が強すぎると高齢者・複数疾患で\n体重減少・筋肉量低下を招く", color: "8E44AD" },
  ];

  for (var i = 0; i < reasons.length; i++) {
    var yPos = 1.2 + i * 1.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.1, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    // Number badge
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.2, w: 0.7, h: 0.7, fill: { color: reasons[i].color } });
    s.addText(reasons[i].num, { x: 0.9, y: yPos + 0.2, w: 0.7, h: 0.7, fontSize: 22, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // Title
    s.addText(reasons[i].title, { x: 1.9, y: yPos + 0.05, w: 3.0, h: 0.4, fontSize: 20, fontFace: FJ, color: reasons[i].color, bold: true, margin: 0 });
    // Detail
    s.addText(reasons[i].detail, { x: 1.9, y: yPos + 0.45, w: 7.2, h: 0.6, fontSize: 15, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 4: 食事は検査値だけでは決められない
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食事は検査値だけでは決められない");

  // Central message
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.0, rectRadius: 0.1, fill: { color: C.light }, shadow: shd() });
  s.addText("年齢・体格・筋肉量・臓器機能・嚥下・生活背景を総合して「その人に合った食事」が見える", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.0,
    fontSize: 18, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Two case examples
  // Case A
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: "DC3545" } });
  s.addText("Case A: 制限の積み重ね", { x: 0.5, y: 2.6, w: 4.3, h: 0.55, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("高血圧＋糖尿病＋肥満\n→ 減塩＋糖質制限＋脂質制限\n→ 制限が多すぎて続かない", {
    x: 0.8, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  // Case B
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: "E67E22" } });
  s.addText("Case B: 制限が逆効果", { x: 5.2, y: 2.6, w: 4.3, h: 0.55, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("腎臓病＋食欲低下＋体重減少\n→ たんぱく制限＋減塩\n→ 低栄養が進行してしまう", {
    x: 5.5, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 5: まず確認したい7つの項目（タイトル）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };

  s.addText("食事指導の出発点", {
    x: 0.6, y: 1.0, w: 8.8, h: 0.6,
    fontSize: 18, fontFace: FJ, color: "A5D6A7", align: "center", margin: 0,
  });
  s.addText("まず確認したい\n7つの項目", {
    x: 0.6, y: 1.8, w: 8.8, h: 2.0,
    fontSize: 42, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3.0, y: 4.0, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("病気ごとの細かい話に入る前に、外来でまず見ておくべきこと", {
    x: 0.6, y: 4.3, w: 8.8, h: 0.5,
    fontSize: 16, fontFace: FJ, color: "B0BEC5", align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 6: 7つの項目（一覧表）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まず確認したい7つの項目");

  var items = [
    { icon: "⚖️", name: "体重と体重変化", reason: "意図しない増減がないか" },
    { icon: "🧂", name: "食塩の取り方", reason: "高血圧・腎臓病・心不全" },
    { icon: "🍱", name: "食事量全体", reason: "多すぎ or 少なすぎ" },
    { icon: "🥩", name: "たんぱく質の取り方", reason: "CKD vs フレイル・肝硬変" },
    { icon: "💧", name: "水分の取り方", reason: "脱水 vs 心不全・腎障害" },
    { icon: "🍽️", name: "食形態", reason: "嚥下・咀嚼・術後の適合" },
    { icon: "🏠", name: "生活背景", reason: "実行可能性の評価" },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 1.1 + i * 0.57;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: yPos, w: 9.2, h: 0.5, rectRadius: 0.05, fill: { color: bgCol } });
    s.addText(items[i].icon, { x: 0.5, y: yPos, w: 0.6, h: 0.5, fontSize: 18, align: "center", valign: "middle", margin: 0 });
    s.addText(items[i].name, { x: 1.2, y: yPos, w: 3.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(items[i].reason, { x: 4.8, y: yPos, w: 4.6, h: 0.5, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 7: 体重と体重変化
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "① 体重と体重変化");

  // Key message
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 0.7, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("単に「重いか軽いか」ではなく、ここ数か月での変化が重要", {
    x: 1.0, y: 1.2, w: 8.0, h: 0.7,
    fontSize: 18, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Left: 増加
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.3, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.3, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: "E67E22" } });
  s.addText("体重が増えている", { x: 0.5, y: 2.3, w: 4.3, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・食事量の見直し\n・間食・甘い飲み物の確認\n・食べる速度・タイミング\n・活動量との釣り合い", {
    x: 0.8, y: 2.9, w: 3.7, h: 2.0, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  // Right: 減少
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.3, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.3, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: "DC3545" } });
  s.addText("体重が減っている", { x: 5.2, y: 2.3, w: 4.3, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・まず低栄養を疑う\n・制限の前に「十分食べているか」\n・高齢者・慢性疾患では特に注意\n・筋肉量低下のサインを見る", {
    x: 5.5, y: 2.9, w: 3.7, h: 2.0, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 8: 食塩・食事量・たんぱく質
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "② 食塩  ③ 食事量  ④ たんぱく質");

  // 食塩
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: 1.2, w: 3.0, h: 3.7, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: 1.2, w: 3.0, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("🧂 食塩", { x: 0.3, y: 1.2, w: 3.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("自覚なく多いことが多い\n\n隠れ塩分源:\n・みそ汁、麺の汁\n・漬物、練り物\n・惣菜、弁当\n・カップ麺", {
    x: 0.5, y: 1.85, w: 2.6, h: 2.9, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });

  // 食事量
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.5, y: 1.2, w: 3.0, h: 3.7, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.5, y: 1.2, w: 3.0, h: 0.5, rectRadius: 0.1, fill: { color: C.accent } });
  s.addText("🍱 食事量全体", { x: 3.5, y: 1.2, w: 3.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("種類より総量が大事\n\n確認すべき点:\n・主食/主菜/副菜\n・間食・夜食\n・甘い飲み物\n・アルコール", {
    x: 3.7, y: 1.85, w: 2.6, h: 2.9, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });

  // たんぱく質
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.7, y: 1.2, w: 3.0, h: 3.7, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.7, y: 1.2, w: 3.0, h: 0.5, rectRadius: 0.1, fill: { color: "8E44AD" } });
  s.addText("🥩 たんぱく質", { x: 6.7, y: 1.2, w: 3.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("一律に制限は危険\n\nCKD → 制限が必要\n高齢者 → 不足が問題\nフレイル → 不足が問題\n肝硬変 → 不足が問題\n術後 → 不足が問題", {
    x: 6.9, y: 1.85, w: 2.6, h: 2.9, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 9: 水分・食形態・生活背景
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "⑤ 水分  ⑥ 食形態  ⑦ 生活背景");

  // 水分
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: 1.2, w: 3.0, h: 3.7, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: 1.2, w: 3.0, h: 0.5, rectRadius: 0.1, fill: { color: "2980B9" } });
  s.addText("💧 水分", { x: 0.3, y: 1.2, w: 3.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("見落としやすい項目\n\n脱水リスク:\n・高齢者・下痢\n・嚥下障害\n\n制限が必要:\n・心不全\n・進行腎障害", {
    x: 0.5, y: 1.85, w: 2.6, h: 2.9, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });

  // 食形態
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.5, y: 1.2, w: 3.0, h: 3.7, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.5, y: 1.2, w: 3.0, h: 0.5, rectRadius: 0.1, fill: { color: "16A085" } });
  s.addText("🍽️ 食形態", { x: 3.5, y: 1.2, w: 3.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("安全に食べられるか\n\n確認:\n・嚥下障害の有無\n・咀嚼力・入れ歯\n・術後の食事制限\n・少量頻回の必要性", {
    x: 3.7, y: 1.85, w: 2.6, h: 2.9, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });

  // 生活背景
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.7, y: 1.2, w: 3.0, h: 3.7, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.7, y: 1.2, w: 3.0, h: 0.5, rectRadius: 0.1, fill: { color: "7F8C8D" } });
  s.addText("🏠 生活背景", { x: 6.7, y: 1.2, w: 3.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("実行可能性が変わる\n\n考慮すべき点:\n・一人暮らし/調理力\n・仕事の帰宅時間\n・家族との食事\n・買い物のしやすさ", {
    x: 6.9, y: 1.85, w: 2.6, h: 2.9, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 10: 制限だけの指導が危ない理由
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: "DC3545" } });
  s.addText("制限だけの指導が危ない理由", {
    x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, fontFace: FJ, color: C.white, bold: true, margin: 0,
  });

  var cases = [
    { label: "糖尿病", problem: "主食を極端に減らす → 食事量全体が落ちる", icon: "📉" },
    { label: "腎臓病", problem: "たんぱく制限しすぎ → 筋肉量が落ちる", icon: "💪" },
    { label: "嚥下障害", problem: "食べやすいものだけ → 水分・栄養不足", icon: "⚠️" },
  ];

  for (var i = 0; i < cases.length; i++) {
    var yPos = 1.2 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 0.12, h: 0.8, rectRadius: 0, fill: { color: "DC3545" } });
    s.addText(cases[i].icon + " " + cases[i].label, { x: 1.2, y: yPos, w: 2.0, h: 0.8, fontSize: 18, fontFace: FJ, color: "DC3545", bold: true, valign: "middle", margin: 0 });
    s.addText(cases[i].problem, { x: 3.2, y: yPos, w: 5.8, h: 0.8, fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  }

  // Key insight
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.5, w: 8.4, h: 0.8, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("特に高齢者：「食べすぎ」より「食べなさすぎ」の方が深刻になることがある", {
    x: 1.0, y: 4.5, w: 8.0, h: 0.8,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 11: 食事指導で考えるべき2つの軸
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食事指導で常に考える2つの軸");

  // Left axis
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.5, rectRadius: 0.15, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.6, rectRadius: 0.15, fill: { color: C.primary } });
  s.addText("何を減らすべきか", { x: 0.5, y: 1.3, w: 4.3, h: 0.6, fontSize: 22, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・過剰な塩分\n・過剰なカロリー\n・病態に応じた制限\n  （たんぱく質、カリウム等）", {
    x: 0.8, y: 2.1, w: 3.7, h: 1.5, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  // Right axis
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.5, rectRadius: 0.15, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.6, rectRadius: 0.15, fill: { color: C.accent } });
  s.addText("何を不足させてはいけないか", { x: 5.2, y: 1.3, w: 4.3, h: 0.6, fontSize: 20, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・十分なエネルギー\n・必要なたんぱく質\n・水分\n・ビタミン・微量栄養素", {
    x: 5.5, y: 2.1, w: 3.7, h: 1.5, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  // Bottom message
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: 4.2, w: 7.0, h: 0.8, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("このバランスが崩れると、食事療法はうまくいかない", {
    x: 1.5, y: 4.2, w: 7.0, h: 0.8,
    fontSize: 19, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 12: 外来で使いやすい進め方
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外来で使いやすい食事指導の進め方");

  var steps = [
    { num: "1", title: "1回で変える項目は1〜2つに絞る", detail: "一度にたくさん → 続かない。具体的な行動目標1つから", color: C.primary },
    { num: "2", title: "数字より行動に落とす", detail: "「減塩」→「みそ汁は1日1回」「麺の汁は残す」", color: C.accent },
    { num: "3", title: "できていることを確認する", detail: "問題探しだけでなく、守れている点を評価 → 焦点が絞れる", color: "2980B9" },
    { num: "4", title: "次回につながる指導にする", detail: "前回からの変化を確認。うまくいかなければ目標を再設定", color: "8E44AD" },
  ];

  for (var i = 0; i < steps.length; i++) {
    var yPos = 1.15 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 0.85, y: yPos + 0.15, w: 0.55, h: 0.55, fill: { color: steps[i].color } });
    s.addText(steps[i].num, { x: 0.85, y: yPos + 0.15, w: 0.55, h: 0.55, fontSize: 20, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(steps[i].title, { x: 1.7, y: yPos + 0.02, w: 7.5, h: 0.4, fontSize: 18, fontFace: FJ, color: steps[i].color, bold: true, margin: 0 });
    s.addText(steps[i].detail, { x: 1.7, y: yPos + 0.42, w: 7.5, h: 0.38, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 13: 具体的な行動例
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "具体的な行動目標の例");

  // Bad examples
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 3.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: "DC3545" } });
  s.addText("✕ 抽象的な指示", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・「塩分を控えてください」\n\n・「甘いものを控えてください」\n\n・「バランスよく食べてください」\n\n・「カロリーに気をつけて」", {
    x: 0.8, y: 1.9, w: 3.7, h: 2.9, fontSize: 16, fontFace: FJ, color: C.sub, margin: 0,
  });

  // Good examples
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 3.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("○ 具体的な行動", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・「みそ汁は1日1回に」\n\n・「毎日の缶コーヒーをやめる」\n\n・「麺の汁は半分残す」\n\n・「夜食の菓子パンをやめる」", {
    x: 5.5, y: 1.9, w: 3.7, h: 2.9, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 14: 栄養士につなぐべき場面
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "栄養士につなぐべき場面");

  var cases = [
    "腎臓病で塩分・たんぱく質・カリウムの整理が必要",
    "肝硬変で低栄養や筋肉量低下が気になる",
    "消化管手術後で食事量が安定しない",
    "嚥下障害があり食形態・水分調整が必要",
    "複数の生活習慣病が重なり優先順位の整理が困難",
    "高齢で体重減少・低栄養が疑われる",
  ];

  for (var i = 0; i < cases.length; i++) {
    var yPos = 1.15 + i * 0.65;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.55, rectRadius: 0.08, fill: { color: (i % 2 === 0) ? C.white : C.light } });
    s.addText("▸ " + cases[i], {
      x: 1.1, y: yPos, w: 7.8, h: 0.55,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  // Bottom message
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 5.0, w: 8.0, h: 0.45, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("医師が方向性を示し、栄養士が生活に落とし込む ― この連携が重要", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.45,
    fontSize: 15, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 15: シリーズ全体の見通し
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "このシリーズで扱うこと（全10回）");

  var episodes = [
    { num: "1", name: "総論（今回）", point: "「何を守るか」から始める", color: C.primary },
    { num: "2", name: "高血圧", point: "減塩を中心に", color: "DC3545" },
    { num: "3", name: "脂質異常症", point: "脂質の「質」を考える", color: "E67E22" },
    { num: "4", name: "糖尿病", point: "食事全体の組み立て", color: "8E44AD" },
    { num: "5", name: "肥満", point: "崩れにくい食べ方", color: "2980B9" },
    { num: "6", name: "CKD", point: "制限と低栄養予防の両立", color: "16A085" },
    { num: "7", name: "肝疾患", point: "脂肪肝と肝硬変で真逆", color: "D35400" },
    { num: "8", name: "嚥下障害", point: "食形態と安全な食べ方", color: "7F8C8D" },
    { num: "9", name: "術後消化管", point: "段階的な食事拡大", color: "2C3E50" },
    { num: "10", name: "フレイル", point: "制限より「守る」優先順位", color: "1B5E20" },
  ];

  for (var i = 0; i < episodes.length; i++) {
    var row = i < 5 ? i : i - 5;
    var col = i < 5 ? 0 : 1;
    var xPos = 0.3 + col * 4.85;
    var yPos = 1.1 + row * 0.8;

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xPos, y: yPos, w: 4.65, h: 0.68, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.15, y: yPos + 0.12, w: 0.44, h: 0.44, fill: { color: episodes[i].color } });
    s.addText(episodes[i].num, { x: xPos + 0.15, y: yPos + 0.12, w: 0.44, h: 0.44, fontSize: 14, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(episodes[i].name, { x: xPos + 0.7, y: yPos + 0.02, w: 1.8, h: 0.32, fontSize: 14, fontFace: FJ, color: episodes[i].color, bold: true, margin: 0 });
    s.addText(episodes[i].point, { x: xPos + 0.7, y: yPos + 0.34, w: 3.7, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 16-18: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 1, [
  "食事指導は「禁止の列挙」ではなく「何を優先して守るか」を決めること",
  "まず7つの項目を確認：体重変化・食塩・食事量・たんぱく質・水分・食形態・生活背景",
  "制限だけでは低栄養を招く ― 「減らすもの」と「守るもの」の両方を考える",
  "1回で変える目標は1〜2つ、抽象的な指示ではなく具体的な行動に落とす",
  "複雑な症例は栄養士と連携 ― 医師が方向性、栄養士が生活に落とし込む",
]);

T.addNextEpisodeSlide(pres, 2, "高血圧の食事指導\n― 減塩だけでは不十分な理由");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
