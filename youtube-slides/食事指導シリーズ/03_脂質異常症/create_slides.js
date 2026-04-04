var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第3回 脂質異常症");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 3); }

// テーマカラー（脂質異常症＝オレンジ〜琥珀系をアクセントに使用）
var DL = {
  amber: "F59E0B",
  darkAmber: "D97706",
  red: "DC3545",
  blue: "2563EB",
  teal: "0D9488",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 3,
  "脂質異常症の食事指導\n― LDLと中性脂肪で戦略が違う",
  "油を全部悪者にしない外来の実践"
);

// ============================================================
// SLIDE 2: 脂質異常症で食事が重要な理由
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脂質異常症で食事指導が重要な理由");

  // Risk chain
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.0, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("脂質異常症 → 動脈硬化 → 心筋梗塞・脳梗塞の重要な危険因子", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.0,
    fontSize: 20, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Two columns
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("薬物療法との関係", { x: 0.5, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・スタチンが必要な場面は多い\n・しかし薬を始めても食事は重要\n・食事が乱れたままだと体重増加・\n  糖尿病・脂肪肝が進みやすい", {
    x: 0.8, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: DL.amber } });
  s.addText("食事指導の意義", { x: 5.2, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・コレステロールの数値だけでなく\n  動脈硬化全体のリスクを下げる\n・生活調整としての位置づけ\n・他の合併症の予防にもつながる", {
    x: 5.5, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 3: LDLと中性脂肪で戦略が違う
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "LDLと中性脂肪で戦略が違う", DL.darkAmber);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("脂質異常症といっても中身はひとつではない ― まず何が高いかを確認する", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: DL.darkAmber, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // LDL column
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.1, w: 4.3, h: 3.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.1, w: 0.12, h: 3.0, fill: { color: DL.red } });
  s.addText("LDLコレステロール高値", { x: 0.9, y: 2.2, w: 3.6, h: 0.4, fontSize: 18, fontFace: FJ, color: DL.red, bold: true, margin: 0 });
  s.addText("見直すべきもの:", { x: 0.9, y: 2.65, w: 3.6, h: 0.3, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText("・飽和脂肪酸（肉の脂身・バター）\n・揚げ物の頻度\n・加工肉\n・菓子パン・スナック菓子\n・洋菓子", {
    x: 0.9, y: 3.0, w: 3.6, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  // TG column
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.1, w: 4.3, h: 3.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.1, w: 0.12, h: 3.0, fill: { color: DL.amber } });
  s.addText("中性脂肪高値", { x: 5.6, y: 2.2, w: 3.6, h: 0.4, fontSize: 18, fontFace: FJ, color: DL.darkAmber, bold: true, margin: 0 });
  s.addText("見直すべきもの:", { x: 5.6, y: 2.65, w: 3.6, h: 0.3, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText("・アルコール（量と頻度）\n・甘い飲み物\n・間食・菓子パン\n・夜遅い食事\n・糖質の過剰摂取", {
    x: 5.6, y: 3.0, w: 3.6, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 4: LDL高値で見直したいもの
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "LDL高値でまず見直したい食品", DL.red);

  var items = [
    { icon: "🥩", name: "飽和脂肪酸", detail: "肉の脂身・バター・生クリーム・ラード", impact: "最重要" },
    { icon: "🍗", name: "揚げ物", detail: "唐揚げ・とんかつ・天ぷら・フライの頻度", impact: "高" },
    { icon: "🥓", name: "加工肉", detail: "ハム・ソーセージ・ベーコン（塩分も多い）", impact: "高" },
    { icon: "🍩", name: "菓子類", detail: "菓子パン・スナック・洋菓子（脂質+糖質）", impact: "高" },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 1.1 + i * 0.95;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.8, rectRadius: 0.08, fill: { color: bgCol }, shadow: shd() });
    s.addText(items[i].icon, { x: 0.7, y: yPos, w: 0.6, h: 0.8, fontSize: 22, align: "center", valign: "middle", margin: 0 });
    s.addText(items[i].name, { x: 1.4, y: yPos + 0.02, w: 2.2, h: 0.38, fontSize: 18, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(items[i].detail, { x: 1.4, y: yPos + 0.4, w: 5.8, h: 0.35, fontSize: 14, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.8, y: yPos + 0.2, w: 1.5, h: 0.4, rectRadius: 0.1, fill: { color: DL.red } });
    s.addText(items[i].impact, { x: 7.8, y: yPos + 0.2, w: 1.5, h: 0.4, fontSize: 13, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 5: 中性脂肪高値で見直したいもの
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "中性脂肪高値で見直したいもの", DL.darkAmber);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("「油ものを減らせば十分」とは限らない ― 糖質・アルコールの影響に注目", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: DL.darkAmber, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var items = [
    { icon: "🍺", title: "アルコール", detail: "種類よりも量。週末まとめ飲みでも影響あり", color: DL.red },
    { icon: "🥤", title: "甘い飲み物", detail: "清涼飲料・加糖缶コーヒー・スポーツドリンク。満腹感なく摂取量が増えやすい", color: DL.amber },
    { icon: "🌙", title: "夜食・菓子パン", detail: "夜遅い食事・夕食後の間食。食べる時間帯も重要", color: DL.blue },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 2.1 + i * 1.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.95, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 0.95, fill: { color: items[i].color } });
    s.addText(items[i].icon + " " + items[i].title, { x: 1.0, y: yPos + 0.05, w: 3.0, h: 0.35, fontSize: 20, fontFace: FJ, color: items[i].color, bold: true, margin: 0 });
    s.addText(items[i].detail, { x: 1.0, y: yPos + 0.45, w: 8.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 6: むしろ増やしたいもの
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "むしろ増やしたい食品", C.primary);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("減らす話だけでなく「増やすとバランスが良くなるもの」も伝える", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 16, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var foods = [
    { icon: "🐟", title: "魚", detail: "肉に偏りがちな食事を整える。焼き魚・煮魚など\n調理法が揚げ物ばかりでは意味が薄れるので注意", color: DL.teal },
    { icon: "🥬", title: "食物繊維", detail: "野菜・海藻・きのこ・豆類・全粒穀物\n主食や脂っこいおかずへの偏りを減らしやすくなる", color: C.primary },
    { icon: "🫘", title: "大豆製品", detail: "豆腐・納豆・豆乳で主菜の一部を置き換え\n肉中心の食事を少し整えたい患者さんに提案しやすい", color: DL.amber },
  ];

  for (var i = 0; i < foods.length; i++) {
    var yPos = 2.0 + i * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.15, w: 0.7, h: 0.7, fill: { color: foods[i].color } });
    s.addText(foods[i].icon, { x: 0.9, y: yPos + 0.15, w: 0.7, h: 0.7, fontSize: 22, align: "center", valign: "middle", margin: 0 });
    s.addText(foods[i].title, { x: 1.9, y: yPos + 0.05, w: 2.0, h: 0.35, fontSize: 20, fontFace: FJ, color: foods[i].color, bold: true, margin: 0 });
    s.addText(foods[i].detail, { x: 1.9, y: yPos + 0.4, w: 7.2, h: 0.55, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 7: 卵だけに注目しすぎない
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「卵は週に何個まで？」への対応");

  // Question
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 0.12, h: 2.0, fill: { color: DL.amber } });
  s.addText("患者さんの関心", { x: 0.9, y: 1.4, w: 3.6, h: 0.35, fontSize: 16, fontFace: FJ, color: DL.darkAmber, bold: true, margin: 0 });
  s.addText("「卵は何個まで？」\n「コレステロールが多い食品は\n 避けた方がいいですよね？」", {
    x: 0.9, y: 1.8, w: 3.6, h: 1.3, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  // Reality
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 0.12, h: 2.0, fill: { color: C.primary } });
  s.addText("見落とされがちな問題", { x: 5.6, y: 1.4, w: 3.6, h: 0.35, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("・毎日菓子パンを食べている\n・揚げ物が多い\n・加工肉が多い\n・甘い飲み物がある\n・体重が増えている", {
    x: 5.6, y: 1.8, w: 3.6, h: 1.3, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 3.7, w: 8.0, h: 1.3, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("単一の食品を悪者にするより、食事全体のパターンを見る方が実際的\n\n卵だけ減らしても、他の問題が残っていれば改善は限られる", {
    x: 1.3, y: 3.8, w: 7.4, h: 1.1,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 8: 外食・コンビニでの選び方
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外食・コンビニでの現実的な選び方");

  // 良い選び方
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 2.6, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("✓ 脂質の質を意識した選び方", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・主菜は揚げ物より焼き・蒸し系\n\n・菓子パン+甘い飲み物のセットを\n  避ける\n\n・加工肉より魚や豆腐の惣菜を選ぶ", {
    x: 0.8, y: 1.85, w: 3.7, h: 1.7, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  // 控えたい組み合わせ
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 2.6, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: DL.red } });
  s.addText("✕ 控えたい組み合わせ", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・菓子パン + 甘い飲み物\n\n・カップ麺 + 揚げ物\n\n・パスタ単品 + デザート\n\n・弁当 + スナック菓子", {
    x: 5.5, y: 1.85, w: 3.7, h: 1.7, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.2, w: 8.4, h: 0.7, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("「脂っこいおかずを減らすだけでなく、菓子や甘い飲み物も一緒に見直しましょう」", {
    x: 1.0, y: 4.2, w: 8.0, h: 0.7,
    fontSize: 16, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
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
    { myth: "油は全部悪い", fact: "油を極端に恐れると食事が貧弱に → かえって菓子・間食が増えることも" },
    { myth: "サラダ油なら安心", fact: "植物油でも量・調理法・他の食品との組み合わせが重要" },
    { myth: "和食なら大丈夫", fact: "和食でも揚げ物が多い、甘い味付け、白米過剰なら改善しにくい" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.2 + i * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.0, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.05, w: 0.5, h: 0.4, fontSize: 20, fontFace: FE, color: DL.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.05, w: 8.0, h: 0.4, fontSize: 18, fontFace: FJ, color: DL.red, bold: true, margin: 0 });
    s.addText("→ " + myths[i].fact, { x: 1.3, y: yPos + 0.5, w: 8.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.primary, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 10: 外来での行動目標の例
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外来で伝える行動目標の例");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("「全部やめる」より「いちばん影響の大きいものから減らす」", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var goals = [
    "毎日の菓子パンをやめる",
    "甘い飲み物を無糖に変える",
    "揚げ物を週の半分以下にする",
    "加工肉を減らして魚や豆腐に置き換える",
  ];

  for (var i = 0; i < goals.length; i++) {
    var yPos = 2.0 + i * 0.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 7.0, h: 0.6, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 0.1, h: 0.6, fill: { color: C.accent } });
    s.addText("▸ " + goals[i], {
      x: 1.9, y: yPos, w: 6.4, h: 0.6,
      fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: 5.0, w: 7.0, h: 0.4, rectRadius: 0.1, fill: { color: C.lightRed } });
  s.addText("✕「油ものを控えてください」だけでは行動変容につながりにくい", {
    x: 1.5, y: 5.0, w: 7.0, h: 0.4,
    fontSize: 13, fontFace: FJ, color: DL.red, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 11-13: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 3, [
  "脂質異常症の食事指導は「油を減らす」だけでは不十分",
  "LDL高値 → 飽和脂肪酸・揚げ物・加工肉・菓子がターゲット",
  "中性脂肪高値 → アルコール・甘い飲み物・夜食・間食も重要",
  "卵など単一食品に注目しすぎず、食事全体のパターンを見る",
  "最も影響が大きいポイントから現実的に変えていく",
]);

T.addNextEpisodeSlide(pres, 4, "糖尿病の食事指導\n― 血糖を安定させる3つのルール");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
