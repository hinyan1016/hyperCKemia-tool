var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第4回 糖尿病");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 4); }

// テーマカラー（糖尿病＝ブルー系をアクセントに使用）
var DM = {
  blue: "2563EB",
  darkBlue: "1D4ED8",
  indigo: "4338CA",
  red: "DC3545",
  teal: "0D9488",
  orange: "E67E22",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 4,
  "糖尿病の食事指導\n― 血糖を安定させる3つのルール",
  "「甘いもの禁止」で終わらせない外来の実践"
);

// ============================================================
// SLIDE 2: 糖尿病で食事療法が重要な理由
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "糖尿病で食事療法が重要な理由");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.2, w: 8.4, h: 1.0, rectRadius: 0.1, fill: { color: "EBF5FB" } });
  s.addText("食事は血糖に直結 ― 薬物療法を行っていても食事が乱れれば安定しない", {
    x: 1.0, y: 1.2, w: 8.0, h: 1.0,
    fontSize: 20, fontFace: FJ, color: DM.darkBlue, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // Two columns
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: DM.blue } });
  s.addText("血糖だけではない", { x: 0.5, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・体重、脂質、血圧、脂肪肝\n・将来の動脈硬化リスク\n・「血糖だけを見る食事」ではなく\n  全身の代謝を整える食事", {
    x: 0.8, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 2.6, w: 4.3, h: 0.55, rectRadius: 0.1, fill: { color: DM.teal } });
  s.addText("継続がカギ", { x: 5.2, y: 2.6, w: 4.3, h: 0.55, fontSize: 18, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("・一時的なイベントではない\n・完璧だが続かない方法より\n  不完全でも継続できる方法に\n  価値がある", {
    x: 5.5, y: 3.3, w: 3.7, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 3: 血糖を安定させる3つのルール
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "血糖を安定させる3つのルール", DM.darkBlue);

  var rules = [
    { num: "1", title: "血糖を乱しにくい食事にする", detail: "食後高血糖を起こしやすい食べ方を見直し、食事量のぶれを小さくする", color: DM.blue, tag: "最重要" },
    { num: "2", title: "必要なら体重を調整する", detail: "肥満→減量で血糖改善。高齢者・やせ→減らさないことが重要な場合も", color: DM.orange, tag: "個別判断" },
    { num: "3", title: "続けられる方法にする", detail: "生活に合わない方法は続かない。一人暮らし・仕事・家族・外食を踏まえる", color: C.primary, tag: "基本" },
  ];

  for (var i = 0; i < rules.length; i++) {
    var yPos = 1.15 + i * 1.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.1, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.2, w: 0.7, h: 0.7, fill: { color: rules[i].color } });
    s.addText(rules[i].num, { x: 0.9, y: yPos + 0.2, w: 0.7, h: 0.7, fontSize: 24, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(rules[i].title, { x: 1.9, y: yPos + 0.05, w: 5.5, h: 0.4, fontSize: 20, fontFace: FJ, color: rules[i].color, bold: true, margin: 0 });
    s.addText(rules[i].detail, { x: 1.9, y: yPos + 0.48, w: 6.5, h: 0.55, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.2, y: yPos + 0.3, w: 1.0, h: 0.4, rectRadius: 0.2, fill: { color: rules[i].color } });
    s.addText(rules[i].tag, { x: 8.2, y: yPos + 0.3, w: 1.0, h: 0.4, fontSize: 12, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 4: 「甘いもの禁止」だけでは不十分
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「甘いもの禁止」だけでは不十分な理由");

  // Patient side
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 2.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 0.12, h: 2.0, fill: { color: DM.blue } });
  s.addText("患者さんの認識", { x: 0.9, y: 1.4, w: 3.6, h: 0.35, fontSize: 16, fontFace: FJ, color: DM.blue, bold: true, margin: 0 });
  s.addText("「甘いものは控えています」\n「砂糖は使っていません」\n「果物は体に良いから…」", {
    x: 0.9, y: 1.85, w: 3.6, h: 1.2, fontSize: 17, fontFace: FJ, color: C.text, margin: 0,
  });

  // Reality
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 2.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: 1.3, w: 0.12, h: 2.0, fill: { color: DM.red } });
  s.addText("見落とされがちな血糖源", { x: 5.6, y: 1.4, w: 3.6, h: 0.35, fontSize: 16, fontFace: FJ, color: DM.red, bold: true, margin: 0 });
  s.addText("・主食の量が多い（大盛り・丼）\n・菓子パンが食事代わり\n・飲み物に糖分が多い\n・食事量が日によって大きくぶれる", {
    x: 5.6, y: 1.85, w: 3.6, h: 1.2, fontSize: 16, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 3.7, w: 8.0, h: 1.3, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("血糖に強く影響するのは食事全体の炭水化物量\n\n何をどれだけ、どの時間帯に食べているかを見ることが本質", {
    x: 1.3, y: 3.8, w: 7.4, h: 1.1,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 5: まず整えたいのは主食の量
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まず整えたいのは主食の量");

  var staples = [
    { icon: "🍚", name: "ご飯", detail: "大盛り・丼もの・おかわりの習慣を見直す", action: "「ゼロ」ではなく「毎食の量をそろえる」" },
    { icon: "🍞", name: "パン", detail: "菓子パンは糖質＋脂質が同時に多い", action: "菓子パン＋甘い飲み物のセットは特に注意" },
    { icon: "🍜", name: "麺類", detail: "単品で済ませやすく、主食量が多くなりがち", action: "野菜やたんぱく質を組み合わせる" },
  ];

  for (var i = 0; i < staples.length; i++) {
    var yPos = 1.2 + i * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addText(staples[i].icon, { x: 0.8, y: yPos, w: 0.6, h: 1.0, fontSize: 24, align: "center", valign: "middle", margin: 0 });
    s.addText(staples[i].name, { x: 1.5, y: yPos + 0.05, w: 1.5, h: 0.4, fontSize: 20, fontFace: FJ, color: DM.blue, bold: true, margin: 0 });
    s.addText(staples[i].detail, { x: 3.0, y: yPos + 0.05, w: 6.2, h: 0.35, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
    s.addText("→ " + staples[i].action, { x: 3.0, y: yPos + 0.5, w: 6.2, h: 0.35, fontSize: 14, fontFace: FJ, color: C.primary, margin: 0 });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.8, w: 8.4, h: 0.5, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("主食をゼロにする必要はない ― 量のぶれを整えるだけでも効果あり", {
    x: 1.0, y: 4.8, w: 8.0, h: 0.5,
    fontSize: 15, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 6: 血糖を乱しやすい食品
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "血糖を乱しやすい食品 ― 主食以外で確認すべき3つ", DM.red);

  var items = [
    { icon: "🥤", title: "甘い飲み物", detail: "最も見逃されやすい。満腹感なく血糖だけ急上昇\n食事を見直しても改善しない場合、原因が飲み物のことも", color: DM.red },
    { icon: "🍪", title: "間食", detail: "糖質中心の間食が習慣化すると血糖は不安定に\n菓子・せんべい・チョコ・アイスが毎日入っていないか", color: DM.orange },
    { icon: "🥐", title: "菓子パン", detail: "糖質も脂質も多く、朝食・昼食代わりに選ばれやすい\n血糖管理の面では不利な食品", color: DM.indigo },
  ];

  for (var i = 0; i < items.length; i++) {
    var yPos = 1.1 + i * 1.25;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.05, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 1.05, fill: { color: items[i].color } });
    s.addText(items[i].icon + " " + items[i].title, { x: 1.0, y: yPos + 0.05, w: 3.0, h: 0.4, fontSize: 20, fontFace: FJ, color: items[i].color, bold: true, margin: 0 });
    s.addText(items[i].detail, { x: 1.0, y: yPos + 0.48, w: 8.0, h: 0.5, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 7: たんぱく質と野菜を組み合わせる
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "たんぱく質と野菜を組み合わせる", C.primary);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, rectRadius: 0.1, fill: { color: C.light } });
  s.addText("「何を減らすか」だけでなく「何を組み合わせるか」も大切", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var combos = [
    { before: "ご飯だけで済ます", after: "魚・肉・卵・大豆製品を加える", icon: "🍚→🍚🐟" },
    { before: "麺類だけで終わらせる", after: "野菜やたんぱく質を組み合わせる", icon: "🍜→🍜🥬" },
    { before: "菓子パンだけの朝食", after: "卵・ヨーグルト・サラダを加える", icon: "🥐→🥐🥚" },
  ];

  for (var i = 0; i < combos.length; i++) {
    var yPos = 2.0 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText(combos[i].icon, { x: 0.8, y: yPos, w: 1.0, h: 0.85, fontSize: 16, align: "center", valign: "middle", margin: 0 });
    s.addText("✕ " + combos[i].before, { x: 1.8, y: yPos + 0.02, w: 3.5, h: 0.38, fontSize: 15, fontFace: FJ, color: DM.red, margin: 0 });
    s.addText("→ " + combos[i].after, { x: 5.3, y: yPos + 0.02, w: 4.0, h: 0.38, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 5.0, w: 8.4, h: 0.4, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("⚠ 腎機能低下がある場合はたんぱく質の勧め方を慎重に", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.4,
    fontSize: 13, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 8: 食べる順番より大事なこと
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食べる順番より大事なこと");

  // Left: What patients ask
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 3.5, h: 1.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 3.5, h: 0.5, rectRadius: 0.1, fill: { color: DM.blue } });
  s.addText("よく聞かれる質問", { x: 0.5, y: 1.3, w: 3.5, h: 0.5, fontSize: 16, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("「野菜から食べると\n よいですか？」\n\n→ 一定の意味はあるが\n  それだけでは不十分", {
    x: 0.8, y: 1.9, w: 2.9, h: 1.0, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });

  // Right: More important things
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.3, y: 1.3, w: 5.2, h: 1.8, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.3, y: 1.3, w: 5.2, h: 0.5, rectRadius: 0.1, fill: { color: C.primary } });
  s.addText("もっと重要なチェックポイント", { x: 4.3, y: 1.3, w: 5.2, h: 0.5, fontSize: 16, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("□ 主食の量が多すぎないか\n□ 飲み物に糖分が入っていないか\n□ 間食が習慣化していないか\n□ 食事量が日によってぶれていないか", {
    x: 4.6, y: 1.9, w: 4.6, h: 1.0, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 3.5, w: 8.0, h: 0.55, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("食べる順番は補助的な工夫。根本は食事全体の量と内容を整えること", {
    x: 1.0, y: 3.5, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 9: 外食・コンビニでの選び方
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外食・コンビニで気をつけたい3つのポイント");

  var points = [
    { icon: "🍚", title: "主食量が多くなりやすい", detail: "丼もの・麺類単品・大盛りで炭水化物が増える\n→ 定食を選ぶ、大盛りにしない", color: DM.blue },
    { icon: "🍞", title: "単品食いになりやすい", detail: "菓子パンだけ、おにぎりだけ → 血糖が急上昇\n→ たんぱく質や野菜を組み合わせる", color: DM.orange },
    { icon: "🥤", title: "甘い飲み物をセットにしない", detail: "食事に気をつけても飲み物で効果が薄れる\n→ 水やお茶を基本にする", color: DM.red },
  ];

  for (var i = 0; i < points.length; i++) {
    var yPos = 1.1 + i * 1.25;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.05, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 1.05, fill: { color: points[i].color } });
    s.addText(points[i].icon + " " + points[i].title, { x: 1.0, y: yPos + 0.05, w: 4.0, h: 0.4, fontSize: 20, fontFace: FJ, color: points[i].color, bold: true, margin: 0 });
    s.addText(points[i].detail, { x: 1.0, y: yPos + 0.48, w: 8.0, h: 0.5, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 10: 高齢者・低栄養・腎機能低下の注意
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: DM.red } });
  s.addText("全員に同じ制限をかけない ― 個別の注意が必要な場面", {
    x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FJ, color: C.white, bold: true, margin: 0,
  });

  var cautions = [
    { icon: "👴", title: "高齢者", detail: "血糖を気にするあまり食事量↓ → 体重減少・筋肉量低下\n若年者や肥満患者と同じ制限は危険。食べられているか・体重が落ちていないかを確認", color: DM.red },
    { icon: "⚖️", title: "やせている患者", detail: "「減らす」より「整える」が重要\nエネルギー不足にならず血糖を安定させる方法を考える", color: DM.orange },
    { icon: "🫘", title: "腎機能低下", detail: "たんぱく質・塩分・カリウムの考え方が関わる\n自己流の食事制限は危険 — 栄養士との連携が有用", color: DM.indigo },
  ];

  for (var i = 0; i < cautions.length; i++) {
    var yPos = 1.2 + i * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.0, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 1.0, fill: { color: cautions[i].color } });
    s.addText(cautions[i].icon + " " + cautions[i].title, { x: 1.0, y: yPos + 0.02, w: 3.0, h: 0.35, fontSize: 20, fontFace: FJ, color: cautions[i].color, bold: true, margin: 0 });
    s.addText(cautions[i].detail, { x: 1.0, y: yPos + 0.4, w: 8.0, h: 0.55, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
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
    { myth: "果物は体によいからいくらでも食べてよい", fact: "量を考えずに食べれば血糖に影響する。「健康に良い」と「血糖に無関係」は別" },
    { myth: "砂糖を使っていなければ安心", fact: "ご飯・パン・麺類も血糖に関わる。「甘くない＝安全」は誤り" },
    { myth: "主食を抜けばよい", fact: "極端に減らすと続かず低栄養にも。ゼロより量を整える方が現実的" },
    { myth: "糖尿病用の食品だけ選べばよい", fact: "特別な食品より日常の食習慣（飲み物・間食・主食量）を整える方が効果的" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.1 + i * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.85, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.02, w: 0.5, h: 0.4, fontSize: 18, fontFace: FE, color: DM.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.02, w: 8.0, h: 0.35, fontSize: 16, fontFace: FJ, color: DM.red, bold: true, margin: 0 });
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
  s.addText("1〜2個の具体的な目標に落とし込む", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var goals = [
    "主食は毎食同じくらいの量にそろえる",
    "甘い飲み物を水かお茶に変える",
    "菓子パンを食事の代わりにしない",
    "おかずにたんぱく質を1品加える",
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
  s.addText("✕「甘いものをやめてください」だけでは行動変容につながりにくい", {
    x: 1.5, y: 5.0, w: 7.0, h: 0.4,
    fontSize: 13, fontFace: FJ, color: DM.red, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 13-15: まとめ・次回予告・エンド
// ============================================================
T.addSummarySlide(pres, 4, [
  "糖尿病の食事指導は「甘いもの禁止」だけでは不十分",
  "まず整えたいのは主食の量 ― 毎食のぶれを小さくする",
  "甘い飲み物・間食・菓子パンは見逃されやすい血糖源",
  "たんぱく質と野菜を組み合わせ、血糖変動を緩やかに",
  "高齢者では制限しすぎによるフレイル・低栄養に注意",
]);

T.addNextEpisodeSlide(pres, 5, "肥満の食事指導\n― 食べ方を変える行動療法的アプローチ");

T.addEndSlide(pres);

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
