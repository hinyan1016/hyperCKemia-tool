var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第7回 肝疾患");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 7); }

// テーマカラー（肝疾患＝ティール〜シアン系）
var LV = {
  teal: "0F766E",
  darkTeal: "115E59",
  cyan: "0891B2",
  red: "DC3545",
  amber: "D97706",
  orange: "EA580C",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 7,
  "肝疾患の食事指導\n― 脂肪肝と肝硬変で真逆になる理由",
  "「肝臓が悪いから一律に控える」ではない"
);

// ============================================================
// SLIDE 2: 食事指導の方向が大きく異なる
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脂肪肝と肝硬変で方向性が真逆になる", LV.teal);

  // 脂肪肝
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: LV.orange } });
  s.addText("脂肪肝 → 減らす方向", { x: 0.5, y: 1.3, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・肥満、内臓脂肪が背景\n・糖尿病、脂質異常症、高血圧\n・過剰エネルギー摂取の是正が中心\n・甘い飲み物、夜食、アルコール\n  間食の見直し\n・体重を少しずつ是正", {
    x: 0.8, y: 2.0, w: 3.7, h: 1.9, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  // 肝硬変
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: LV.teal } });
  s.addText("肝硬変 → 不足させない方向", { x: 5.2, y: 1.3, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・食欲低下、早期膨満感\n・腹水、味覚変化\n・十分に食べられていない\n・筋肉量低下が進みやすい\n・食事を減らしすぎると\n  状態が悪化する", {
    x: 5.5, y: 2.0, w: 3.7, h: 1.9, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.5, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("まず「脂肪肝」か「肝硬変」かを整理してから食事指導に入る", {
    x: 1.0, y: 4.5, w: 8.0, h: 0.55,
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
    { icon: "🔍", text: "脂肪肝が中心か、肝硬変まで進んでいるか" },
    { icon: "⚖️", text: "体重が増えているか、減っているか" },
    { icon: "💊", text: "糖尿病・脂質異常症・高血圧の合併" },
    { icon: "🍺", text: "アルコール摂取の有無" },
    { icon: "😔", text: "食欲低下や早期膨満感がないか" },
    { icon: "💧", text: "腹水や浮腫があるか" },
    { icon: "💪", text: "筋肉量低下やフレイルが疑われないか" },
    { icon: "🧠", text: "肝性脳症の既往があるか" },
  ];

  for (var i = 0; i < checks.length; i++) {
    var yPos = 1.1 + i * 0.52;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.44, rectRadius: 0.06, fill: { color: bgCol }, shadow: shd() });
    s.addText(checks[i].icon + "  " + checks[i].text, {
      x: 1.1, y: yPos, w: 7.8, h: 0.44,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 4: 脂肪肝で見直したい食習慣
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脂肪肝でまず見直したい食習慣", LV.orange);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("特別な食品を足すより、日常の習慣を見直す方が効果的", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: LV.amber, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var items = [
    { icon: "⚖️", text: "体重が増えていないか", color: LV.orange },
    { icon: "🍪", text: "間食が多くないか", color: LV.amber },
    { icon: "🥤", text: "甘い飲み物を習慣的に飲んでいないか", color: LV.red },
    { icon: "🌙", text: "夜遅い食事や夜食が多くないか", color: LV.cyan },
    { icon: "🍺", text: "アルコール摂取がないか", color: LV.teal },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 2.0 + i * 0.68;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 8.0, h: 0.55, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 0.1, h: 0.55, fill: { color: items[i].color } });
    s.addText(items[i].icon + "  " + items[i].text, {
      x: 1.4, y: yPos, w: 7.4, h: 0.55,
      fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 5: 脂肪肝で減らすもの
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脂肪肝では「何を減らすか」が中心");

  var tips = [
    { icon: "🥤", title: "甘い飲み物をやめる", detail: "清涼飲料・加糖缶コーヒー・スポーツドリンク" },
    { icon: "🍩", title: "菓子・菓子パンの頻度を減らす", detail: "脂質＋糖質が同時に多い食品" },
    { icon: "🌙", title: "夜食を減らす", detail: "夜遅い食事＋間食の積み重ね" },
    { icon: "🍗", title: "揚げ物・脂質の多い惣菜を続けない", detail: "週の頻度を意識する" },
    { icon: "🍺", title: "アルコールを見直す", detail: "量・頻度・つまみも含めて" },
  ];

  for (var i = 0; i < tips.length; i++) {
    var yPos = 1.1 + i * 0.82;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.7, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText(tips[i].icon, { x: 0.7, y: yPos, w: 0.6, h: 0.7, fontSize: 20, align: "center", valign: "middle", margin: 0 });
    s.addText(tips[i].title, { x: 1.4, y: yPos, w: 3.5, h: 0.7, fontSize: 17, fontFace: FJ, color: LV.orange, bold: true, valign: "middle", margin: 0 });
    s.addText(tips[i].detail, { x: 5.0, y: yPos, w: 4.3, h: 0.7, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 6: 肝硬変では低栄養と筋肉量低下が問題
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "肝硬変で問題になるのは低栄養と筋肉量低下", LV.red);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: C.lightRed } });
  s.addText("脂肪肝と同じ感覚で食事制限をすると危険なことがある", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: LV.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var points = [
    { icon: "⚠️", text: "腹水・浮腫で太って見えるが、実は筋肉量↓・栄養↓", color: LV.red },
    { icon: "⏰", text: "空腹時間が長いこと自体が不利に働く", color: LV.amber },
    { icon: "🚫", text: "食事回数が少ない・朝食を抜く → 低栄養を進める", color: LV.orange },
  ];

  for (var i = 0; i < points.length; i++) {
    var yPos = 2.1 + i * 0.85;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.7, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 0.7, fill: { color: points[i].color } });
    s.addText(points[i].icon + "  " + points[i].text, {
      x: 1.0, y: yPos, w: 8.0, h: 0.7,
      fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  // Check items
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.7, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("□ 食事量は足りているか　□ たんぱく質が不足していないか　□ 筋肉量↓はないか", {
    x: 1.0, y: 4.7, w: 8.0, h: 0.6,
    fontSize: 14, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 7: 制限しすぎてはいけない理由
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "肝硬変で制限しすぎてはいけない理由");

  // Cascade
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.2, rectRadius: 0.1, fill: { color: C.lightRed } });
  s.addText("自己判断で肉も脂も控える → 食事全体が細くなる\n→ 筋肉量↓ → 活動性↓ → 転倒・感染症・入院↑", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.2,
    fontSize: 18, fontFace: FJ, color: LV.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var cautions = [
    { icon: "🧂", text: "塩分制限で「味が薄くて食べられない」状態になっていないか", color: LV.red },
    { icon: "🥩", text: "肝性脳症を恐れてたんぱく質を極端に減らしていないか", color: LV.amber },
  ];

  for (var i = 0; i < cautions.length; i++) {
    var yPos = 2.8 + i * 0.85;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.7, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 0.12, h: 0.7, fill: { color: cautions[i].color } });
    s.addText(cautions[i].icon + "  " + cautions[i].text, {
      x: 1.2, y: yPos, w: 7.8, h: 0.7,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.8, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("制限の結果として食べられなくなること自体が大きな問題", {
    x: 1.0, y: 4.8, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 8: 腹水時の塩分と就寝前補食
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "腹水があるときの塩分 & 就寝前補食", LV.teal);

  // 塩分
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: LV.teal } });
  s.addText("🧂 塩分の工夫", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・みそ汁やスープの回数を減らす\n・麺類の汁を飲み切らない\n・漬物を控える\n・加工食品を重ねない\n・だし・酸味・香辛料を活かす\n\n⚠ 味が薄すぎて食事量↓は本末転倒", {
    x: 0.8, y: 1.85, w: 3.7, h: 1.9, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  // 就寝前補食
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: LV.cyan } });
  s.addText("🌙 就寝前補食（LES）", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・長い空腹時間を避ける考え方\n・少量を分けて食べる方が\n  合う患者もいる\n\n検討すべき場面:\n・食事量が少ない\n・体重減少・フレイルが心配\n・一度にたくさん食べられない", {
    x: 5.5, y: 1.85, w: 3.7, h: 1.9, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
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
    { myth: "肝臓が悪いなら絶食気味の方がよい", fact: "脂肪肝ではエネルギー是正が必要。肝硬変では食べなさすぎが問題" },
    { myth: "肝性脳症→たんぱく質をできるだけ減らす", fact: "たんぱく質不足→筋肉量↓→かえって悪化。薬物療法も含めて考える" },
    { myth: "肝臓に良い健康食品を足せばよい", fact: "まず日常の食習慣、体重、アルコール、食事量の整理が優先" },
    { myth: "脂肪肝と肝硬変は同じ食事でよい", fact: "最も重要な誤解。減らす方向と不足させない方向は真逆" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.1 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.02, w: 0.5, h: 0.4, fontSize: 18, fontFace: FE, color: LV.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.02, w: 8.0, h: 0.35, fontSize: 16, fontFace: FJ, color: LV.red, bold: true, margin: 0 });
    s.addText("→ " + myths[i].fact, { x: 1.3, y: yPos + 0.42, w: 8.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.primary, margin: 0 });
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

  // 脂肪肝
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: LV.orange } });
  s.addText("脂肪肝の患者さんへ", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("▸ 甘い飲み物をやめる\n▸ 間食を減らす\n▸ 夜食を見直す\n▸ 体重を月1回測定する", {
    x: 0.8, y: 1.9, w: 3.7, h: 1.5, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  // 肝硬変
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: LV.teal } });
  s.addText("肝硬変の患者さんへ", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("▸ 食事を抜かない\n▸ 1日3食は確保する\n▸ 薄味で食べられないなら工夫\n▸ 肉や魚を自己判断で減らしすぎない", {
    x: 5.5, y: 1.9, w: 3.7, h: 1.5, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.1, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("病態に応じて行動目標の方向性が違う ― ここが最重要ポイント", {
    x: 1.0, y: 4.1, w: 8.0, h: 0.55,
    fontSize: 16, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 11-13: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 7, [
  "脂肪肝と肝硬変で食事指導の方向性は大きく異なる",
  "脂肪肝 → 体重管理・間食・飲み物・アルコールの是正が中心",
  "肝硬変 → 低栄養・筋肉量低下を防ぐ視点が極めて重要",
  "腹水時は塩分調整が必要だが制限しすぎに注意",
  "「肝臓が悪いから一律に控える」ではなく病態ごとに優先順位を決める",
]);

T.addNextEpisodeSlide(pres, 8, "嚥下障害の食事指導\n― 食形態選択と誤嚥予防の実際");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
