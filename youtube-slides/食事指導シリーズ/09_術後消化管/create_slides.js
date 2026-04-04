var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第9回 消化管術後");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 9); }

// テーマカラー（消化管術後＝スレート〜グレーブルー系）
var PO = {
  slate: "475569",
  darkSlate: "334155",
  blue: "3B82F6",
  red: "DC3545",
  amber: "D97706",
  teal: "0D9488",
  orange: "EA580C",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 9,
  "消化管術後の食事指導\n― 少量頻回と症状別対応の考え方",
  "「消化のよいもの」だけでは不十分な外来の実践"
);

// ============================================================
// SLIDE 2: 食事指導が重要な理由
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "消化管手術後に食事指導が重要な理由", PO.slate);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.0, rectRadius: 0.1, fill: { color: "F1F5F9" } });
  s.addText("食後症状がつらい → 食事量↓ → 体重↓ → 筋肉↓ → 食欲↓ の悪循環", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.0,
    fontSize: 18, fontFace: FJ, color: PO.darkSlate, bold: true, align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: PO.red } });
  s.addText("起こりやすい問題", { x: 0.5, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・早期膨満感　・食後の腹部膨満\n・悪心・下痢　・便通不安定\n・ダンピング症候群\n・体重減少　・脱水　・低栄養\n・ビタミン・微量元素の不足", {
    x: 0.8, y: 3.3, w: 3.7, h: 1.6, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("食事指導の2つの目標", { x: 5.2, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("① 症状を減らすこと\n\n② 必要な栄養を確保すること\n\n→ この両立を考える", {
    x: 5.5, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
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
    "どの部位の手術か",
    "術後どのくらいの時期か",
    "体重はどの程度減っているか",
    "食後にどのような症状が出るか",
    "下痢・便秘・腹部膨満はあるか",
    "水分はどの程度取れているか",
    "1回に食べる量と1日の回数",
    "貧血や栄養障害を示唆する所見はないか",
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
// SLIDE 4: 少量頻回が基本
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "術後食事の基本は少量頻回", PO.blue);

  var reasons = [
    { num: "1", title: "一度に大量に流し込まないため", detail: "術後の消化管は容量・通過様式が変化。大量→症状悪化", color: PO.blue },
    { num: "2", title: "必要量を確保しやすくするため", detail: "3食で足りないなら4-6回に分けて入りやすくする", color: PO.teal },
    { num: "3", title: "血糖変動やダンピングを和らげるため", detail: "小分けにして食べることが食後症状の軽減につながる", color: PO.amber },
  ];

  for (var i = 0; i < reasons.length; i++) {
    var yPos = 1.2 + i * 1.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.95, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.12, w: 0.7, h: 0.7, fill: { color: reasons[i].color } });
    s.addText(reasons[i].num, { x: 0.9, y: yPos + 0.12, w: 0.7, h: 0.7, fontSize: 22, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(reasons[i].title, { x: 1.9, y: yPos + 0.02, w: 7.2, h: 0.38, fontSize: 18, fontFace: FJ, color: reasons[i].color, bold: true, margin: 0 });
    s.addText(reasons[i].detail, { x: 1.9, y: yPos + 0.45, w: 7.2, h: 0.4, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.8, w: 8.0, h: 0.5, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("「3食きちんと」より「1回量を欲張らず回数で補う」が実践的", {
    x: 1.0, y: 4.8, w: 8.0, h: 0.5,
    fontSize: 15, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 5: 胃の手術後
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "胃の手術後に起こりやすい問題", PO.red);

  var items = [
    { icon: "⚡", title: "早期ダンピング", detail: "食後30分前後。腹痛、下痢、動悸、発汗\n糖質の多いものが急速に小腸へ流れ込む", color: PO.red },
    { icon: "📉", title: "後期ダンピング", detail: "食後1-3時間。ふらつき、冷汗、空腹感\n急な血糖上昇→低血糖が関わる", color: PO.amber },
    { icon: "😣", title: "早期膨満感", detail: "胃の容量↓→少量でおなかがいっぱいに\n必要なエネルギーが入りにくい→体重減少", color: PO.slate },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 1.1 + i * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 1.0, fill: { color: items[i].color } });
    s.addText(items[i].icon + " " + items[i].title, { x: 1.0, y: yPos + 0.02, w: 3.0, h: 0.38, fontSize: 18, fontFace: FJ, color: items[i].color, bold: true, margin: 0 });
    s.addText(items[i].detail, { x: 1.0, y: yPos + 0.45, w: 8.0, h: 0.5, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.7, w: 8.4, h: 0.5, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("一度に多く食べない / 糖質に偏らない / たんぱく質を毎回意識", {
    x: 1.0, y: 4.7, w: 8.0, h: 0.5,
    fontSize: 15, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 6: 小腸・大腸手術後
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "小腸・大腸手術後に起こりやすい問題");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 3.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: PO.orange } });
  s.addText("小腸切除後", { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・下痢（脂肪便になることも）\n・脱水 — 便量が多い患者では\n  見かけ以上に水分・電解質を喪失\n・吸収障害 — 体重↓ 低Alb 貧血\n  ビタミン・微量元素不足\n・回腸末端切除 → VitB12吸収↓\n  胆汁酸吸収↓", {
    x: 0.8, y: 1.85, w: 3.7, h: 2.1, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 3.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, rectRadius: 0.1, fill: { color: PO.slate } });
  s.addText("大腸手術後", { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・下痢や頻便\n  脂っこい食事・刺激物・冷たい\n  飲み物で悪化することも\n・便秘や残便感\n  活動量↓ 食事量↓ 水分不足\n・腹部膨満\n  ガスがたまりやすい食品", {
    x: 5.5, y: 1.85, w: 3.7, h: 2.1, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.6, w: 8.0, h: 0.5, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("食べられる量を確保しつつ、便の性状を見ながら調整する", {
    x: 1.0, y: 4.6, w: 8.0, h: 0.5,
    fontSize: 15, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 7: 食べ方の工夫
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食べ方の工夫 ― 食品選びより効くことがある", PO.teal);

  var tips = [
    { icon: "🦷", title: "よく噛む", detail: "急いで飲み込むのを防ぎ食後症状の軽減にも" },
    { icon: "🍽️", title: "一度に食べすぎない", detail: "腹八分よりさらに少ない量から始めてよい" },
    { icon: "⏱️", title: "急いで食べない", detail: "早食いは過量摂取と急速通過につながる" },
    { icon: "🥤", title: "飲み物を一緒に流し込みすぎない", detail: "食事と飲水のタイミングをずらすと楽になる場合も" },
    { icon: "🥩", title: "たんぱく質を毎回意識する", detail: "炭水化物に偏ると栄養不足になりやすい" },
  ];

  for (var i = 0; i < tips.length; i++) {
    var yPos = 1.05 + i * 0.82;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.7, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText(tips[i].icon, { x: 0.7, y: yPos, w: 0.5, h: 0.7, fontSize: 18, align: "center", valign: "middle", margin: 0 });
    s.addText(tips[i].title, { x: 1.3, y: yPos, w: 2.8, h: 0.7, fontSize: 17, fontFace: FJ, color: PO.teal, bold: true, valign: "middle", margin: 0 });
    s.addText(tips[i].detail, { x: 4.2, y: yPos, w: 5.1, h: 0.7, fontSize: 14, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 8: 症状別の食事調整
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "症状別 ― 食事をどう調整するか", PO.darkSlate);

  var symptoms = [
    { symptom: "下痢が強い", actions: "1回量↓ / 脂っこい食事控える / 甘い飲み物↓ / 脱水確認", color: PO.amber },
    { symptom: "食後すぐ冷汗・動悸", actions: "早期ダンピング → 1回量↓ / 糖質偏り避ける / 飲食分離", color: PO.red },
    { symptom: "食後1-3h ふらつき", actions: "後期ダンピング → 砂糖↓ / たんぱく質を組み合わせる", color: PO.orange },
    { symptom: "腹部膨満", actions: "1回量↓ / 早食い避ける / 炭酸↓ / ガス食品を個別に確認", color: PO.slate },
  ];

  for (var i = 0; i < symptoms.length; i++) {
    var yPos = 1.1 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 2.8, h: 0.85, rectRadius: 0.08, fill: { color: symptoms[i].color } });
    s.addText(symptoms[i].symptom, { x: 0.6, y: yPos, w: 2.8, h: 0.85, fontSize: 14, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(symptoms[i].actions, { x: 3.6, y: yPos, w: 5.6, h: 0.85, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
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
    { myth: "消化のよいものならいくらでも食べてよい", fact: "一度に食べる量そのものが症状に関わる。何をと同時にどのくらいかが重要" },
    { myth: "やわらかければ安心", fact: "やわらかい菓子や甘い飲み物がかえってダンピング等の症状を悪化させる" },
    { myth: "下痢があるから食べない方がよい", fact: "絶食気味にすると体重↓脱水↓低栄養が進む。量と内容を調整して確保" },
    { myth: "術後だからしばらく様子見でよい", fact: "体重減少大・貧血進行・長く食べられないなら栄養障害の評価が必要" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.1 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.02, w: 0.5, h: 0.4, fontSize: 18, fontFace: FE, color: PO.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.02, w: 8.0, h: 0.35, fontSize: 16, fontFace: FJ, color: PO.red, bold: true, margin: 0 });
    s.addText("→ " + myths[i].fact, { x: 1.3, y: yPos + 0.42, w: 8.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.primary, margin: 0 });
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
  s.addText("「消化のよいものを」だけでは具体性に欠ける ― 症状に応じた目標を", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 15, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var goals = [
    "1回の食事量を少なめにして回数で補う",
    "甘い飲み物を食事中に一気に飲まない",
    "毎食たんぱく質を1品加える",
    "体重を週1回測定する",
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
  footer(s);
})();

// ============================================================
// SLIDE 11-13: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 9, [
  "消化管術後の食事指導の基本は少量頻回",
  "胃切除後はダンピング症状を意識し糖質に偏らない",
  "小腸術後は下痢・脱水・吸収障害を見落とさない",
  "大腸術後は便通に合わせて調整する",
  "症状を減らすことと低栄養・脱水を防ぐことの両立が最重要",
]);

T.addNextEpisodeSlide(pres, 10, "フレイル・多疾患併存の食事指導\n― 制限よりも「守る」優先順位");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
