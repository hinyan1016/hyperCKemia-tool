var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第6回 CKD");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 6); }

// テーマカラー（CKD＝紫〜インディゴ系）
var CK = {
  indigo: "4338CA",
  darkIndigo: "3730A3",
  purple: "7C3AED",
  red: "DC3545",
  amber: "D97706",
  teal: "0D9488",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 6,
  "CKDの食事指導\n― たんぱく質・塩分・カリウムの三本柱",
  "一律に制限するのではなく優先順位を決める"
);

// ============================================================
// SLIDE 2: 腎臓病の食事指導はなぜ難しいか
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "腎臓病の食事指導はなぜ難しいか", CK.indigo);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.0, rectRadius: 0.1, fill: { color: "EDE9FE" } });
  s.addText("「これを食べてはいけない」と単純に言えないから難しい", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.0,
    fontSize: 20, fontFace: FJ, color: CK.darkIndigo, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var factors = [
    { icon: "🧂", text: "塩分が大事な患者", color: CK.indigo },
    { icon: "🥩", text: "たんぱく質の調整が重要な患者", color: CK.purple },
    { icon: "🍌", text: "カリウム・リンを意識すべき患者", color: CK.amber },
    { icon: "👴", text: "まず十分に食べることを優先すべき患者", color: CK.red },
  ];

  for (var i = 0; i < factors.length; i++) {
    var yPos = 2.5 + i * 0.7;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 7.0, h: 0.55, rectRadius: 0.08, fill: { color: bgCol }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 0.1, h: 0.55, fill: { color: factors[i].color } });
    s.addText(factors[i].icon + "  " + factors[i].text, {
      x: 1.9, y: yPos, w: 6.4, h: 0.55,
      fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
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
    { icon: "🔬", text: "腎機能低下の程度（eGFR）" },
    { icon: "💧", text: "尿蛋白の有無" },
    { icon: "📊", text: "血圧・浮腫の状況" },
    { icon: "⚡", text: "カリウム値が高くなっていないか" },
    { icon: "🦴", text: "リンが問題になっていないか" },
    { icon: "💉", text: "糖尿病の有無" },
    { icon: "🏥", text: "透析を受けているか" },
    { icon: "⚖️", text: "体重減少や食欲低下がないか" },
  ];

  for (var i = 0; i < checks.length; i++) {
    var yPos = 1.1 + i * 0.45;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.40, rectRadius: 0.06, fill: { color: bgCol }, shadow: shd() });
    s.addText(checks[i].icon + "  " + checks[i].text, {
      x: 1.1, y: yPos, w: 7.8, h: 0.40,
      fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.85, w: 8.4, h: 0.4, rectRadius: 0.1, fill: { color: "EDE9FE" } });
  s.addText("最初にここを整理しないと食事指導がぼやける", {
    x: 1.0, y: 4.85, w: 8.0, h: 0.4,
    fontSize: 14, fontFace: FJ, color: CK.darkIndigo, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 4: まず取り組みたいのは塩分
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "最初に取り組みたいのは塩分制限");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("見直しやすく、効果が出やすいのは塩分 — 血圧↓ 浮腫↓ 腎負担↓", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Hidden salt sources
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.0, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.0, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: CK.red } });
  s.addText("隠れた塩分源", { x: 0.5, y: 2.0, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・みそ汁・麺類の汁\n・漬物・佃煮\n・練り物\n・ハム・ソーセージ\n・惣菜・弁当\n・カップ麺", {
    x: 0.8, y: 2.65, w: 3.7, h: 1.9, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  // Practical tips
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.0, w: 4.3, h: 2.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.0, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("具体的な減塩策", { x: 5.2, y: 2.0, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・みそ汁の回数を減らす\n・麺の汁を全部飲まない\n・加工食品を続けて重ねない\n・惣菜は味の濃いものを毎日\n  使わない\n・調味料を少し意識して使う", {
    x: 5.5, y: 2.65, w: 3.7, h: 1.9, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 5: たんぱく質は多すぎも少なすぎも問題
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "たんぱく質 ― 多すぎも少なすぎも問題", CK.purple);

  // Too much
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.2, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 0.12, h: 2.2, fill: { color: CK.red } });
  s.addText("⚠ 多すぎ", { x: 0.9, y: 1.4, w: 3.6, h: 0.35, fontSize: 18, fontFace: FJ, color: CK.red, bold: true, margin: 0 });
  s.addText("・腎機能低下時に腎臓への負担↑\n・老廃物（尿素窒素など）が蓄積\n・進行を早める可能性", {
    x: 0.9, y: 1.85, w: 3.6, h: 1.4, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  // Too little
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.2, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 0.12, h: 2.2, fill: { color: CK.amber } });
  s.addText("⚠ 少なすぎ", { x: 5.6, y: 1.4, w: 3.6, h: 0.35, fontSize: 18, fontFace: FJ, color: CK.amber, bold: true, margin: 0 });
  s.addText("・筋肉量低下 → フレイル\n・低栄養 → 体力・免疫力↓\n・高齢者では特に注意が必要", {
    x: 5.6, y: 1.85, w: 3.6, h: 1.4, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 3.9, w: 8.0, h: 1.2, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("たんぱく質を完全に悪者にしない\n\n腎機能・年齢・栄養状態・フレイルの有無を見ながら\n「多すぎず、少なすぎず」の視点で考える", {
    x: 1.3, y: 4.0, w: 7.4, h: 1.0,
    fontSize: 16, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 6: カリウム
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "カリウム ― 全員に一律制限ではない");

  // Need restriction
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: CK.red } });
  s.addText("制限が必要な場合", { x: 0.5, y: 1.3, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・進行した腎機能低下\n・高カリウム血症を繰り返す\n\n注意食品:\n  野菜・果物・いも類・豆類\n  果汁飲料・ドライフルーツ", {
    x: 0.8, y: 1.95, w: 3.7, h: 1.6, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  // Don't over-restrict
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("過度な制限は不要な場合", { x: 5.2, y: 1.3, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・カリウム値が安定している\n・軽度の腎機能低下\n\n過度に野菜・果物を避けると:\n  食事の質↓ 便秘 食欲低下", {
    x: 5.5, y: 1.95, w: 3.7, h: 1.6, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.2, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: "EDE9FE" } });
  s.addText("実際のカリウム値と病状を見て判断する — 一律制限は避ける", {
    x: 1.0, y: 4.2, w: 8.0, h: 0.55,
    fontSize: 16, fontFace: FJ, color: CK.darkIndigo, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 7: リンは加工食品の影響も大きい
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "リン ― 加工食品の影響が見落とされやすい", CK.amber);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("「肉や魚を減らせばよい」ではなく、加工食品をどう減らすかが重要", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: CK.amber, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var sources = [
    { icon: "🥓", name: "加工肉", detail: "ハム・ソーセージ・ベーコン" },
    { icon: "🍢", name: "練り製品", detail: "かまぼこ・はんぺん・ちくわ" },
    { icon: "🍜", name: "インスタント食品", detail: "カップ麺・インスタントラーメン" },
    { icon: "🍪", name: "スナック菓子", detail: "ポテトチップス・せんべい等" },
    { icon: "🥛", name: "一部の乳製品・飲料", detail: "プロセスチーズ・一部の清涼飲料" },
  ];

  for (var i = 0; i < sources.length; i++) {
    var yPos = 2.0 + i * 0.65;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 8.0, h: 0.52, rectRadius: 0.06, fill: { color: bgCol }, shadow: shd() });
    s.addText(sources[i].icon + "  " + sources[i].name, { x: 1.3, y: yPos, w: 2.5, h: 0.52, fontSize: 16, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(sources[i].detail, { x: 4.0, y: yPos, w: 4.8, h: 0.52, fontSize: 15, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 8: 糖尿病合併時
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "糖尿病を合併したときの考え方");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: C.lightRed } });
  s.addText("糖質だけを見ても腎臓だけを見ても不十分 ― 全体の整理が必要", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: CK.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Pitfall 1
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.1, w: 4.3, h: 1.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.1, w: 0.12, h: 1.5, fill: { color: CK.red } });
  s.addText("⚠ よくある落とし穴①", { x: 0.9, y: 2.2, w: 3.6, h: 0.3, fontSize: 14, fontFace: FJ, color: CK.red, bold: true, margin: 0 });
  s.addText("糖尿病があるから主食を極端に\n減らす → 代わりに肉・脂質が\n増えて腎臓に負担", {
    x: 0.9, y: 2.6, w: 3.6, h: 0.9, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  // Pitfall 2
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.1, w: 4.3, h: 1.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.1, w: 0.12, h: 1.5, fill: { color: CK.amber } });
  s.addText("⚠ よくある落とし穴②", { x: 5.6, y: 2.2, w: 3.6, h: 0.3, fontSize: 14, fontFace: FJ, color: CK.amber, bold: true, margin: 0 });
  s.addText("腎臓病を気にして食事量が\n減りすぎ → 血糖管理も\n栄養状態も崩れる", {
    x: 5.6, y: 2.6, w: 3.6, h: 0.9, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  // Solution
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.0, w: 8.4, h: 1.2, rectRadius: 0.1, fill: { color: C.lightGreen } });
  s.addText("整理のポイント", { x: 1.1, y: 4.1, w: 7.8, h: 0.3, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("・食塩を減らす　・食事全体の量を整える\n・主食量を極端にしない　・たんぱく質を自己流で削りすぎない", {
    x: 1.1, y: 4.45, w: 7.8, h: 0.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 9: 外食・惣菜の注意点
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外食・惣菜での現実的な工夫");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("全面否定より、現実的な工夫を伝える方が続く", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var tips = [
    "汁物を重ねない（みそ汁＋スープなど）",
    "麺類の汁を飲み切らない",
    "漬物や佃煮を毎回つけない",
    "ハムや練り物ばかりに偏らない",
    "味の濃いものを連続しない",
  ];

  for (var i = 0; i < tips.length; i++) {
    var yPos = 2.0 + i * 0.7;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 7.0, h: 0.55, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 0.1, h: 0.55, fill: { color: CK.indigo } });
    s.addText("▸ " + tips[i], {
      x: 1.9, y: yPos, w: 6.4, h: 0.55,
      fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 10: よくある誤解
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある誤解");

  var myths = [
    { myth: "腎臓病なら野菜は全部だめ", fact: "カリウム制限が必要な人はいるが全員ではない。一律回避は食事の質↓" },
    { myth: "とにかくたんぱく質を減らせばよい", fact: "減らしすぎは低栄養・筋力低下に。特に高齢者では注意" },
    { myth: "塩分だけ気をつけていればよい", fact: "カリウム・リン・食事量全体・低栄養の有無まで見る必要がある" },
    { myth: "健康食品なら安心", fact: "野菜ジュース・サプリがK・P・糖質の面で問題になることがある" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.1 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.02, w: 0.5, h: 0.4, fontSize: 18, fontFace: FE, color: CK.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.02, w: 8.0, h: 0.35, fontSize: 16, fontFace: FJ, color: CK.red, bold: true, margin: 0 });
    s.addText("→ " + myths[i].fact, { x: 1.3, y: yPos + 0.42, w: 8.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.primary, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 11: 外来での行動目標の例
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外来で伝える行動目標の例");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("「いろいろ制限して」ではなく、最も重要な1〜2個に絞る", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var goals = [
    "みそ汁は1日1回まで",
    "麺の汁は残す",
    "加工食品を毎日重ねない",
    "肉や魚を自己判断で極端に減らさない",
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
  s.addText("✕「いろいろ制限してください」では何をすべきか分からない", {
    x: 1.5, y: 5.0, w: 7.0, h: 0.4,
    fontSize: 13, fontFace: FJ, color: CK.red, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 12-14: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 6, [
  "腎臓病の食事指導は一律に制限ではなく優先順位を決めること",
  "まず塩分の見直しが実行しやすく効果も大きい",
  "たんぱく質は多すぎも少なすぎも問題 ― 完全に悪者にしない",
  "カリウム・リンは必要な患者で重点的に見る",
  "高齢者では制限しすぎによる低栄養・筋力低下に注意",
]);

T.addNextEpisodeSlide(pres, 7, "肝疾患の食事指導\n― 脂肪肝と肝硬変で真逆になる理由");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
