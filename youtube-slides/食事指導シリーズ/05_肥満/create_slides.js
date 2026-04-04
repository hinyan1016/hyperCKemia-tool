var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第5回 肥満");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 5); }

// テーマカラー（肥満＝暖色系・テラコッタ）
var OB = {
  terra: "C2410C",
  darkTerra: "9A3412",
  orange: "EA580C",
  amber: "D97706",
  teal: "0D9488",
  blue: "2563EB",
  red: "DC3545",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 5,
  "肥満の食事指導\n― 食べ方を変える行動療法的アプローチ",
  "「食べるな」ではなく崩れにくい食べ方を作る"
);

// ============================================================
// SLIDE 2: 肥満で食事指導が大切な理由
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "肥満で食事指導が大切な理由");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.0, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("肥満は見た目の問題ではなく代謝の問題（BMI ≥ 25）", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.0,
    fontSize: 20, fontFace: FJ, color: OB.darkTerra, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Two columns
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: OB.terra } });
  s.addText("肥満が影響する疾患", { x: 0.5, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・高血圧\n・糖尿病\n・脂質異常症\n・脂肪肝\n・関節症・睡眠時無呼吸", {
    x: 0.8, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: OB.amber } });
  s.addText("重要な視点", { x: 5.2, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・全員が同じ食べ方ではない\n・「太っている＝食べすぎ」で\n  片づけない\n・どの場面で余分な摂取が\n  起こっているかを具体的に見る", {
    x: 5.5, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 3: 体重を減らすより先に整えたいこと
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "体重を減らすより先に整えたいこと", OB.terra);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("「何キロ減らす」より、まず毎日の食べ方を整理する方が実際的", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: OB.darkTerra, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var checks = [
    { icon: "🕐", text: "1日3食のリズムがあるか" },
    { icon: "🌅", text: "朝食を抜いていないか" },
    { icon: "🌙", text: "夜に食事量が偏っていないか" },
    { icon: "🍪", text: "間食や飲み物で余分なカロリーが入っていないか" },
    { icon: "🍱", text: "外食や惣菜が多くないか" },
    { icon: "⏩", text: "早食いになっていないか" },
  ];

  for (var i = 0; i < checks.length; i++) {
    var yPos = 2.0 + i * 0.58;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 8.0, h: 0.48, rectRadius: 0.06, fill: { color: bgCol }, shadow: shd() });
    s.addText(checks[i].icon + "  " + checks[i].text, {
      x: 1.3, y: yPos, w: 7.4, h: 0.48,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 4: 太りやすい食べ方の特徴
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "太りやすい食べ方 ― 4つのパターン", OB.orange);

  var patterns = [
    { icon: "⏩", title: "早食い", detail: "満腹感の前に食べ終わる → 摂取量↑\n一口ごとに箸を置く、丼物を減らす", color: OB.terra },
    { icon: "🥤", title: "液体カロリー", detail: "砂糖入り飲み物は満腹感なくカロリー摂取\n最も見直しやすく効果も出やすい項目", color: OB.orange },
    { icon: "🌙", title: "夜遅い食事", detail: "強い空腹→一気食い。主食+揚げ物+酒が重なる\n夕方の軽い補食でドカ食いを防ぐ", color: OB.amber },
    { icon: "🍪", title: "間食の習慣化", detail: "空腹でなく習慣で食べている場合は要注意\nTV・車・仕事中の無意識の摂取が積み重なる", color: OB.blue },
  ];

  for (var i = 0; i < patterns.length; i++) {
    var yPos = 1.05 + i * 0.98;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.82, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 0.82, fill: { color: patterns[i].color } });
    s.addText(patterns[i].icon + " " + patterns[i].title, { x: 1.0, y: yPos + 0.02, w: 2.5, h: 0.35, fontSize: 18, fontFace: FJ, color: patterns[i].color, bold: true, margin: 0 });
    s.addText(patterns[i].detail, { x: 3.5, y: yPos + 0.02, w: 5.7, h: 0.75, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 5: 「食べないダイエット」が続かない理由
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「食べないダイエット」が続かない理由", OB.red);

  // Bad patterns
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: OB.red } });
  s.addText("✕ よくある自己流の制限", { x: 0.5, y: 1.3, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・朝食を抜く\n・主食を極端に減らす\n・夕食をサラダだけにする\n・数日だけ厳しく我慢する\n\n→ 空腹感が強く反動で食べすぎる\n→ 筋肉量低下・だるさ → 活動量↓", {
    x: 0.8, y: 1.95, w: 3.7, h: 1.7, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  // Better approach
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("✓ 目指すべきアプローチ", { x: 5.2, y: 1.3, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・短期間で大きく減らさない\n・食欲が暴走しにくい形に組み替え\n・崩れても戻しやすい食べ方を作る\n・少し不完全でも継続できる方法\n\n→ 長期的な行動変容が目標", {
    x: 5.5, y: 1.95, w: 3.7, h: 1.7, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.2, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("気合いで乗り切るのではなく、食欲が暴走しにくい形を作る", {
    x: 1.0, y: 4.2, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 6: 満腹感を保ちやすい食べ方
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "満腹感を保ちやすい食べ方", C.primary);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("ただ減らすより、満腹感が続きやすい構成にする方がうまくいく", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 16, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var tips = [
    { icon: "🥩", title: "たんぱく質を抜きすぎない", detail: "肉・魚・卵・大豆・乳製品が少ないと満足感↓ → 間食に走りやすい\n主食だけや菓子パンだけより、たんぱく質を含む食事の方が崩れにくい", color: OB.terra },
    { icon: "🥬", title: "食物繊維を増やす", detail: "野菜・きのこ・海藻・豆類で食事のかさを増やす\n少ない量でも満足しやすくする工夫として説明する", color: C.primary },
    { icon: "🍲", title: "汁物や副菜を上手に使う", detail: "食事が単品になりにくくなる\n高血圧や腎臓病がある場合は塩分量に注意", color: OB.teal },
  ];

  for (var i = 0; i < tips.length; i++) {
    var yPos = 2.0 + i * 1.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.95, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.12, w: 0.7, h: 0.7, fill: { color: tips[i].color } });
    s.addText(tips[i].icon, { x: 0.9, y: yPos + 0.12, w: 0.7, h: 0.7, fontSize: 22, align: "center", valign: "middle", margin: 0 });
    s.addText(tips[i].title, { x: 1.9, y: yPos + 0.02, w: 3.0, h: 0.35, fontSize: 18, fontFace: FJ, color: tips[i].color, bold: true, margin: 0 });
    s.addText(tips[i].detail, { x: 1.9, y: yPos + 0.38, w: 7.2, h: 0.5, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 7: まず減らしやすい食品
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まず減らしやすい食品 ― 効果が出やすいものから");

  var items = [
    { icon: "🍪", name: "菓子類", detail: "食事とは別に入り、摂取量を過小評価しやすい", action: "量を半分・回数を減らす・家に置かない", color: OB.terra },
    { icon: "🥤", name: "甘い飲み物", detail: "最優先で見直したい項目", action: "毎日→週数回、500mL→小サイズ", color: OB.orange },
    { icon: "🍺", name: "アルコール", detail: "飲酒カロリー＋つまみ・夜食が増えやすい", action: "回数と量、つまみの内容を一緒に見直す", color: OB.amber },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 1.2 + i * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addText(items[i].icon, { x: 0.8, y: yPos, w: 0.6, h: 1.0, fontSize: 24, align: "center", valign: "middle", margin: 0 });
    s.addText(items[i].name, { x: 1.5, y: yPos + 0.05, w: 2.0, h: 0.35, fontSize: 20, fontFace: FJ, color: items[i].color, bold: true, margin: 0 });
    s.addText(items[i].detail, { x: 3.5, y: yPos + 0.05, w: 5.7, h: 0.35, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.5, y: yPos + 0.5, w: 5.7, h: 0.35, rectRadius: 0.06, fill: { color: C.lightGreen } });
    s.addText("→ " + items[i].action, { x: 3.7, y: yPos + 0.5, w: 5.3, h: 0.35, fontSize: 14, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 8: 外食で崩れにくくする工夫
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外食で崩れにくくする工夫");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("崩れやすいのはメニュー選びより空腹の勢いで注文する場面", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: OB.darkTerra, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var tips = [
    "空腹が強すぎる状態で店に入らない（事前に軽く補食）",
    "大盛りやおかわりを「とりあえず」で頼まない",
    "セットで炭水化物を重ねない（ラーメン+チャーハン等）",
    "「せっかくだから」で追加注文しない",
    "深夜の外食をできるだけ避ける",
  ];

  for (var i = 0; i < tips.length; i++) {
    var yPos = 2.0 + i * 0.65;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 8.0, h: 0.52, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 0.1, h: 0.52, fill: { color: OB.terra } });
    s.addText("▸ " + tips[i], {
      x: 1.4, y: yPos, w: 7.4, h: 0.52,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
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
    { myth: "朝を抜けば痩せる", fact: "昼夜に強い空腹 → 食べすぎやすい。1日の食事配分が崩れることが問題" },
    { myth: "糖質だけ切ればよい", fact: "脂質・アルコール・間食・液体カロリーも大きく関与。全体像を見る" },
    { myth: "運動すれば食事はそのままでよい", fact: "食行動が原因なら運動だけで帳尻を合わせにくい。両方現実的に考える" },
    { myth: "少し減ったら元の食事に戻してよい", fact: "食事パターンが戻れば体重も戻る。続けられる食べ方の定着が目標" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.1 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.02, w: 0.5, h: 0.4, fontSize: 18, fontFace: FE, color: OB.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.02, w: 8.0, h: 0.35, fontSize: 16, fontFace: FJ, color: OB.red, bold: true, margin: 0 });
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

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("「食べすぎないで」ではなく具体的な行動目標に落とす", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var goals = [
    "甘い飲み物を水かお茶に変える",
    "間食は決めた時間と量だけにする",
    "夕食は21時より前に済ませる",
    "早食いを避けて一口ごとに箸を置く",
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
  s.addText("✕「食べすぎないようにしてください」だけでは行動変容につながりにくい", {
    x: 1.5, y: 5.0, w: 7.0, h: 0.4,
    fontSize: 13, fontFace: FJ, color: OB.red, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 11-13: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 5, [
  "肥満の食事指導は「食べるな」ではなく「崩れにくい食べ方を作る」",
  "早食い・液体カロリー・夜遅い食事・間食がよくある原因",
  "極端な制限は続かない ― 食欲が暴走しにくい形に組み替える",
  "満腹感を保ちやすい構成（たんぱく質＋食物繊維）にする",
  "最も減らしやすい菓子・飲み物・アルコールから着手する",
]);

T.addNextEpisodeSlide(pres, 6, "CKDの食事指導\n― たんぱく質・塩分・カリウムの三本柱");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
