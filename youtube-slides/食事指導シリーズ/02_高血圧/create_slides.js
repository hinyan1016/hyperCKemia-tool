var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第2回 高血圧");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 2); }

// テーマカラー（高血圧＝赤系をアクセントに使用）
var HT = {
  red: "DC3545",
  darkRed: "A71D2A",
  orange: "E67E22",
  blue: "2980B9",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 2,
  "高血圧の食事指導\n― 減塩だけでは不十分な理由",
  "食塩摂取 6g/日未満を目指す外来の実践"
);

// ============================================================
// SLIDE 2: 高血圧で食事が重要な理由
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "高血圧で食事が重要な理由");

  // Risk chain
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.0, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("高血圧 → 脳卒中・心不全・虚血性心疾患・CKD の重要な危険因子", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.0,
    fontSize: 20, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Two columns
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("食事療法の位置づけ", { x: 0.5, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・薬物療法と並ぶ基礎的介入\n・減塩は治療にも予防にも有効\n・食事だけで全て解決はしないが\n  薬物療法との併用で効果向上", {
    x: 0.8, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: HT.orange } });
  s.addText("血圧に影響する因子", { x: 5.2, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・食塩摂取\n・肥満（内臓脂肪）\n・飲酒\n・運動不足・ストレス\n・腎機能・薬剤", {
    x: 5.5, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 3: 食事指導の3つの柱
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "高血圧の食事指導で目指す3つのこと");

  var pillars = [
    { num: "1", title: "食塩摂取を減らす", detail: "最優先。外食・加工食品・汁物が多い患者で効果大", color: HT.red, priority: "最優先" },
    { num: "2", title: "体重が過剰なら少しでも減らす", detail: "肥満があると血圧は下がりにくい。食塩を減らしても体重増加が続けば改善乏しい", color: HT.orange, priority: "重要" },
    { num: "3", title: "続けられる形にする", detail: "短期間だけ厳しくしても続かない。無理なく続けられて初めて意味がある", color: C.primary, priority: "基本" },
  ];

  for (var i = 0; i < pillars.length; i++) {
    var yPos = 1.15 + i * 1.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.1, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.2, w: 0.7, h: 0.7, fill: { color: pillars[i].color } });
    s.addText(pillars[i].num, { x: 0.9, y: yPos + 0.2, w: 0.7, h: 0.7, fontSize: 24, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(pillars[i].title, { x: 1.9, y: yPos + 0.05, w: 5.5, h: 0.4, fontSize: 20, fontFace: FJ, color: pillars[i].color, bold: true, margin: 0 });
    s.addText(pillars[i].detail, { x: 1.9, y: yPos + 0.48, w: 6.5, h: 0.55, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
    // Priority badge
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.2, y: yPos + 0.3, w: 1.0, h: 0.4, rectRadius: 0.2, fill: { color: pillars[i].color } });
    s.addText(pillars[i].priority, { x: 8.2, y: yPos + 0.3, w: 1.0, h: 0.4, fontSize: 12, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 4: 隠れた塩分源
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "患者さんが気づいていない塩分源", HT.red);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "F8D7DA" } });
  s.addText("「塩辛いものはそんなに食べていない」← 実はここに塩分が隠れている", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: HT.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var sources = [
    { icon: "🍜", name: "汁物", detail: "みそ汁・ラーメン・うどんの汁", salt: "1.5〜6g/杯" },
    { icon: "🥒", name: "漬物・佃煮", detail: "少量でも高塩分。毎食で積み重なる", salt: "1〜2g/回" },
    { icon: "🥓", name: "加工食品", detail: "ハム・ソーセージ・練り物・カップ麺", salt: "1〜5g/食" },
    { icon: "🍱", name: "惣菜・弁当", detail: "外食・コンビニ弁当・丼物", salt: "3〜6g/食" },
  ];

  for (var i = 0; i < sources.length; i++) {
    var yPos = 2.0 + i * 0.82;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.7, rectRadius: 0.08, fill: { color: bgCol }, shadow: shd() });
    s.addText(sources[i].icon, { x: 0.7, y: yPos, w: 0.6, h: 0.7, fontSize: 20, align: "center", valign: "middle", margin: 0 });
    s.addText(sources[i].name, { x: 1.4, y: yPos, w: 2.0, h: 0.7, fontSize: 18, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(sources[i].detail, { x: 3.5, y: yPos, w: 3.8, h: 0.7, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.6, y: yPos + 0.15, w: 1.7, h: 0.4, rectRadius: 0.1, fill: { color: HT.red } });
    s.addText(sources[i].salt, { x: 7.6, y: yPos + 0.15, w: 1.7, h: 0.4, fontSize: 13, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 5: 食卓塩だけでは減塩にならない
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「塩をかけない」だけでは減塩にならない");

  // Misconception
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 0.12, h: 2.0, fill: { color: HT.red } });
  s.addText("患者さんの認識", { x: 0.9, y: 1.4, w: 3.6, h: 0.35, fontSize: 16, fontFace: FJ, color: HT.red, bold: true, margin: 0 });
  s.addText("「食卓塩は使っていません」\n「そんなに塩辛いものは\n 食べていないです」", {
    x: 0.9, y: 1.8, w: 3.6, h: 1.3, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  // Reality
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 0.12, h: 2.0, fill: { color: C.primary } });
  s.addText("実際の塩分源", { x: 5.6, y: 1.4, w: 3.6, h: 0.35, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("料理に振る塩 ≪ 食品・惣菜に\nもともと含まれる塩分\n\n→ 塩分の多い食品を見抜く力が重要", {
    x: 5.6, y: 1.8, w: 3.6, h: 1.3, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  // Checklist
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.7, w: 9.0, h: 1.6, rectRadius: 0.1, fill: { color: C.light }, shadow: shd() });
  s.addText("外来で聞くべきチェックポイント", { x: 0.8, y: 3.8, w: 8.4, h: 0.35, fontSize: 16, fontFace: FJ, color: C.dark, bold: true, margin: 0 });
  s.addText("□ 汁物の回数　□ 麺の汁を飲むか　□ 漬物の頻度\n□ 弁当・惣菜の利用頻度　□ カップ麺・加工肉の習慣", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.9, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 6: 減塩の具体策
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食塩を減らしやすい5つの具体策");

  var tips = [
    { icon: "🍲", title: "汁物の回数と量を調整", detail: "1日3回→1回、具を多く汁を少なめに" },
    { icon: "🍜", title: "麺の汁を残す", detail: "説明しやすく、実行しやすい。大きな減塩効果" },
    { icon: "🫙", title: "調味料は少量を意識", detail: "小皿に取る、全体にかけない。しょうゆ・ソース" },
    { icon: "🍋", title: "だし・酸味・香辛料を活用", detail: "レモン、酢、こしょう、七味、だしのうまみ" },
    { icon: "🚫", title: "加工食品を毎日の習慣にしない", detail: "たまに食べるのはOK。毎日続くと塩分がかさむ" },
  ];

  for (var i = 0; i < tips.length; i++) {
    var yPos = 1.1 + i * 0.82;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.7, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText(tips[i].icon, { x: 0.7, y: yPos, w: 0.6, h: 0.7, fontSize: 20, align: "center", valign: "middle", margin: 0 });
    s.addText(tips[i].title, { x: 1.4, y: yPos, w: 3.2, h: 0.7, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(tips[i].detail, { x: 4.7, y: yPos, w: 4.6, h: 0.7, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 7: 野菜・果物とカリウムの注意
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "野菜・果物はどう考えるか");

  // General
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("一般的には有利", { x: 0.5, y: 1.3, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・食物繊維・カリウムを含む\n・血圧管理に有利な食事パターン\n  の一部\n・野菜や果物の摂取を一般に推奨", {
    x: 0.8, y: 2.0, w: 3.7, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  // Caution
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: HT.red } });
  s.addText("⚠ CKD合併では注意", { x: 5.2, y: 1.3, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・腎機能低下＋高K血症があれば\n  カリウム摂取を無条件に推奨\n  できない\n・腎機能・血清K値を確認してから\n・「どんどん増やしましょう」は危険", {
    x: 5.5, y: 2.0, w: 3.7, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.5, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("一律ではなく、腎機能に応じて個別に判断する", {
    x: 1.0, y: 4.5, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 8: 体重・飲酒・食事パターン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "減塩だけでは不十分 ― 体重・飲酒・食事全体");

  var items = [
    { icon: "⚖️", title: "体重", detail: "内臓脂肪↑ → 血圧↑。食塩を減らしても体重増加が続けば改善乏しい", color: HT.orange },
    { icon: "🍺", title: "飲酒", detail: "毎日＋量が多い＋つまみが塩辛い → 悪循環。量と頻度の見直し", color: HT.red },
    { icon: "🍽️", title: "食事パターン全体", detail: "外食多い・野菜少ない・加工食品多い・夜遅い食事 → 日常全体を聞く", color: HT.blue },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 1.2 + i * 1.25;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.05, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 1.05, fill: { color: items[i].color } });
    s.addText(items[i].icon + " " + items[i].title, { x: 1.0, y: yPos + 0.05, w: 2.5, h: 0.4, fontSize: 20, fontFace: FJ, color: items[i].color, bold: true, margin: 0 });
    s.addText(items[i].detail, { x: 1.0, y: yPos + 0.5, w: 8.0, h: 0.45, fontSize: 15, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 9: 外食・コンビニでの選び方
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外食・コンビニでの選び方");

  // 外食
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 3.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("🍽️ 外食のコツ", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・麺類は汁を飲み干さない\n\n・丼物単品より定食の方が調整\n  しやすい\n\n・漬物や小鉢は全部食べなくてOK\n\n・追加のしょうゆ・ソースは控える", {
    x: 0.8, y: 1.85, w: 3.7, h: 2.9, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  // コンビニ
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 3.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: HT.orange } });
  s.addText("🏪 コンビニのコツ", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・カップ麺・加工肉中心を避ける\n\n・おにぎり＋サラダ＋ゆで卵\n  など組み合わせる\n\n・スープ類は毎回つけない\n\n・栄養成分表示で食塩相当量を\n  チェックする習慣をつける", {
    x: 5.5, y: 1.85, w: 3.7, h: 2.9, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 10: CKD合併時の注意
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: HT.red } });
  s.addText("腎機能低下・高カリウム血症がある場合の注意", {
    x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FJ, color: C.white, bold: true, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.2, rectRadius: 0.1, fill: { color: "F8D7DA" } });
  s.addText("高血圧という病名だけで画一的に話すと危険\n一般的な高血圧向け食事指導をそのまま当てはめると危うい場面がある", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.2,
    fontSize: 18, fontFace: FJ, color: HT.darkRed, bold: true, align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 2.7, w: 8.4, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addText("確認すべき項目", { x: 1.1, y: 2.85, w: 7.8, h: 0.35, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, margin: 0 });

  var checks = [
    "腎機能（eGFR）の程度",
    "尿蛋白の有無",
    "血清カリウム値",
    "利尿薬・RAS阻害薬の使用状況",
  ];
  for (var i = 0; i < checks.length; i++) {
    s.addText("☑ " + checks[i], {
      x: 1.3, y: 3.3 + i * 0.45, w: 7.4, h: 0.4,
      fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 11: よくある誤解
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある誤解");

  var myths = [
    { myth: "減塩しょうゆならたくさん使ってよい", fact: "量が増えれば塩分も増える。「減塩」は万能ではない" },
    { myth: "塩を使わなければ減塩できている", fact: "加工食品・汁物・惣菜・外食由来の塩分が圧倒的に多い" },
    { myth: "体質だから食事を変えても意味がない", fact: "食事介入、特に減塩・体重管理は基本。薬物療法との併用で効果向上" },
    { myth: "一気に厳しくしないと意味がない", fact: "続かないことが多い。小さな変更の積み重ねの方が有効" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.1 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.02, w: 0.5, h: 0.4, fontSize: 18, fontFace: FE, color: HT.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.02, w: 8.0, h: 0.35, fontSize: 16, fontFace: FJ, color: HT.red, bold: true, margin: 0 });
    s.addText("→ " + myths[i].fact, { x: 1.3, y: yPos + 0.42, w: 8.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.primary, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 12: 外来での行動目標の例
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外来で伝える行動目標の例");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("「減塩しましょう」で終わらせず、1〜2個の行動目標に落とす", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var goals = [
    "みそ汁は1日2回まで",
    "麺の汁は飲まない",
    "漬物は毎食ではなく1日1回まで",
    "カップ麺は週1回まで",
    "しょうゆは小皿に取って使う",
  ];

  for (var i = 0; i < goals.length; i++) {
    var yPos = 2.0 + i * 0.7;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 7.0, h: 0.55, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 0.1, h: 0.55, fill: { color: C.accent } });
    s.addText("▸ " + goals[i], {
      x: 1.9, y: yPos, w: 6.4, h: 0.55,
      fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: 5.0, w: 7.0, h: 0.4, rectRadius: 0.1, fill: { color: "F8D7DA" } });
  s.addText("✕「塩分を控えて」「薄味にして」だけでは行動変容につながりにくい", {
    x: 1.5, y: 5.0, w: 7.0, h: 0.4,
    fontSize: 13, fontFace: FJ, color: HT.red, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 13-15: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 2, [
  "高血圧の食事指導の中心は減塩。目標は食塩 6g/日未満",
  "「塩をかけない」だけでは不十分 ― 汁物・漬物・加工食品・惣菜が本丸",
  "野菜・果物は一般に有利だが、CKD合併では腎機能・K値に注意",
  "減塩だけでなく体重・飲酒・食事パターン全体を見る",
  "抽象的な指示ではなく、1〜2個の具体的行動目標に落とす",
]);

T.addNextEpisodeSlide(pres, 3, "脂質異常症の食事指導\n― LDLと中性脂肪で戦略が違う");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
