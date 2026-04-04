var T = require("./../_shared/series_template.js");
var pres = T.createPres("食事指導シリーズ 第10回 フレイル・多疾患併存");
var C = T.C;
var FJ = T.FJ;
var FE = T.FE;

function shd() { return T.shd(); }
function hdr(s, title, bgColor) { T.hdr(s, pres, title, bgColor); }
function footer(s) { T.footerBar(s, pres, 10); }

// テーマカラー（最終回＝ゴールド〜ウォームブラウン系）
var FR = {
  gold: "B45309",
  darkGold: "92400E",
  warm: "D97706",
  red: "DC3545",
  teal: "0D9488",
  blue: "2563EB",
  indigo: "4338CA",
};

// ============================================================
// SLIDE 1: タイトル
// ============================================================
T.addTitleSlide(pres, 10,
  "フレイル・多疾患併存の食事指導\n― 制限よりも「守る」優先順位",
  "シリーズ最終回 ― 食事指導は「何を禁止するか」ではなく「何を守るか」"
);

// ============================================================
// SLIDE 2: 最も難しいのは病気が重なる患者
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "食事指導で最も難しいのは「病気が重なる患者」", FR.gold);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("各疾患の制限を足し算すると「何を食べてよいか分からない」状態に", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: FR.darkGold, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var combos = [
    { text: "糖尿病 + 高血圧 + 肥満", color: FR.warm },
    { text: "糖尿病 + 腎臓病", color: FR.indigo },
    { text: "腎臓病 + 高齢 + 体重減少", color: FR.teal },
    { text: "嚥下障害 + 脱水 + 低栄養", color: FR.red },
    { text: "肝硬変 + 筋肉量低下 + 食欲低下", color: FR.blue },
    { text: "がん術後 + 下痢 + 体重減少", color: FR.gold },
  ];

  for (var i = 0; i < combos.length; i++) {
    var yPos = 2.0 + i * 0.55;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 7.0, h: 0.45, rectRadius: 0.06, fill: { color: bgCol }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: yPos, w: 0.1, h: 0.45, fill: { color: combos[i].color } });
    s.addText(combos[i].text, {
      x: 1.9, y: yPos, w: 6.4, h: 0.45,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 3: 制限がぶつかる典型例
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "制限がぶつかる典型例", FR.red);

  var cases = [
    { title: "糖尿病 + 腎臓病", detail: "説明↑ → 「ご飯もだめ、果物もだめ、肉もだめ」\n→ 食べる意欲そのものが落ちる", color: FR.warm },
    { title: "高血圧 + 低栄養", detail: "減塩で味を薄くしすぎ → 食事量がさらに落ちる\n体重減少・筋力低下が前景にある場合は要注意", color: FR.blue },
    { title: "嚥下障害 + 脱水", detail: "とろみ水が飲みにくい → 水分摂取量↓\n誤嚥予防の代償で脱水・便秘・全身状態悪化", color: FR.teal },
    { title: "肝硬変 + 筋肉量低下", detail: "必要以上の制限 → 全身状態をむしろ悪化させる\n低栄養・筋肉量低下が大きな問題", color: FR.indigo },
  ];

  for (var i = 0; i < cases.length; i++) {
    var yPos = 1.05 + i * 0.98;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 0.82, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 0.12, h: 0.82, fill: { color: cases[i].color } });
    s.addText(cases[i].title, { x: 1.0, y: yPos + 0.02, w: 3.0, h: 0.3, fontSize: 16, fontFace: FJ, color: cases[i].color, bold: true, margin: 0 });
    s.addText(cases[i].detail, { x: 1.0, y: yPos + 0.35, w: 8.0, h: 0.42, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 4: まず優先すべきもの
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まず優先すべきものは何か ― 3つの順番", FR.darkGold);

  var priorities = [
    { num: "1", title: "生命に直結する問題がないか", detail: "脱水・誤嚥リスク・著しい食事量低下・急速な体重減少・低血糖\nLDLやHbA1cより先にここを見る", color: FR.red, tag: "最優先" },
    { num: "2", title: "低栄養を進めないこと", detail: "体重減少→筋力↓→転倒→感染→入院→ADL↓\n「この指導で食事量が落ちないか」を常に考える", color: FR.warm, tag: "重要" },
    { num: "3", title: "継続できること", detail: "独居・買い物困難・調理困難・家族と別メニュー不可\n実行できなければ意味がない", color: C.primary, tag: "基本" },
  ];

  for (var i = 0; i < priorities.length; i++) {
    var yPos = 1.15 + i * 1.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.1, rectRadius: 0.1, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.2, w: 0.7, h: 0.7, fill: { color: priorities[i].color } });
    s.addText(priorities[i].num, { x: 0.9, y: yPos + 0.2, w: 0.7, h: 0.7, fontSize: 24, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(priorities[i].title, { x: 1.9, y: yPos + 0.05, w: 5.5, h: 0.4, fontSize: 20, fontFace: FJ, color: priorities[i].color, bold: true, margin: 0 });
    s.addText(priorities[i].detail, { x: 1.9, y: yPos + 0.48, w: 6.5, h: 0.55, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.2, y: yPos + 0.3, w: 1.0, h: 0.4, rectRadius: 0.2, fill: { color: priorities[i].color } });
    s.addText(priorities[i].tag, { x: 8.2, y: yPos + 0.3, w: 1.0, h: 0.4, fontSize: 12, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 5: フレイル高齢者では何を守るか
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "フレイル高齢者では何を守るべきか", FR.warm);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("数値の管理以上に「食べる力を保つ」「体重・筋肉量を落とさない」", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: FR.darkGold, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var checks = [
    { icon: "⚖️", text: "体重減少がないか" },
    { icon: "😔", text: "食欲低下がないか" },
    { icon: "🍚", text: "主食だけに偏っていないか" },
    { icon: "🥩", text: "たんぱく質源が毎食あるか" },
    { icon: "💧", text: "水分は足りているか" },
    { icon: "🦷", text: "嚥下や咀嚼に問題がないか" },
    { icon: "🛒", text: "買い物や調理ができているか" },
  ];

  for (var i = 0; i < checks.length; i++) {
    var yPos = 2.0 + i * 0.5;
    var bgCol = (i % 2 === 0) ? C.white : C.light;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 8.0, h: 0.42, rectRadius: 0.06, fill: { color: bgCol }, shadow: shd() });
    s.addText(checks[i].icon + "  " + checks[i].text, {
      x: 1.3, y: yPos, w: 7.4, h: 0.42,
      fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 6: 数値と食べられることのバランス
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「数値の改善」と「食べられること」のバランス");

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: C.lightRed } });
  s.addText("数値の改善だけを目標にすると全体を見失うことがある", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 17, fontFace: FJ, color: FR.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var traps = [
    { left: "HbA1cを下げるために", right: "食事量全体が落ちる", color: FR.warm },
    { left: "塩分を減らしすぎて", right: "食欲が低下する", color: FR.blue },
    { left: "腎臓を気にしすぎて", right: "たんぱく質が不足する", color: FR.indigo },
    { left: "嚥下に配慮しすぎて", right: "水分摂取が減る", color: FR.teal },
  ];

  for (var i = 0; i < traps.length; i++) {
    var yPos = 2.1 + i * 0.7;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.55, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 0.1, h: 0.55, fill: { color: traps[i].color } });
    s.addText(traps[i].left, { x: 1.2, y: yPos, w: 3.8, h: 0.55, fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText("→", { x: 5.0, y: yPos, w: 0.5, h: 0.55, fontSize: 18, fontFace: FE, color: FR.red, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(traps[i].right, { x: 5.5, y: yPos, w: 3.5, h: 0.55, fontSize: 16, fontFace: FJ, color: FR.red, bold: true, valign: "middle", margin: 0 });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 4.9, w: 8.0, h: 0.45, rectRadius: 0.1, fill: { color: C.lightGreen } });
  s.addText("数値＋食べられているか＋体重＋日常生活を一緒に見る", {
    x: 1.0, y: 4.9, w: 8.0, h: 0.45,
    fontSize: 15, fontFace: FJ, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });
  footer(s);
})();

// ============================================================
// SLIDE 7: 家族指導
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "家族指導で伝えるべきこと", FR.gold);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: "FEF3C7" } });
  s.addText("家族が熱心なほど「これもだめ、あれもだめ」になりやすい", {
    x: 1.0, y: 1.15, w: 8.0, h: 0.6,
    fontSize: 16, fontFace: FJ, color: FR.darkGold, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var messages = [
    "完璧な制限食を目指さなくてよい",
    "まずは食事量を保つことが重要な場合がある",
    "体重減少や食欲低下は見逃してはいけない",
    "一度にすべてを変えなくてよい",
    "食べやすい形で必要量を入れる工夫が大切",
  ];

  for (var i = 0; i < messages.length; i++) {
    var yPos = 2.0 + i * 0.65;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 8.0, h: 0.52, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: yPos, w: 0.1, h: 0.52, fill: { color: FR.gold } });
    s.addText("▸ " + messages[i], {
      x: 1.4, y: yPos, w: 7.4, h: 0.52,
      fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }
  footer(s);
})();

// ============================================================
// SLIDE 8: 最終チェックポイント
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "外来で使いやすい最終チェックポイント", FR.darkGold);

  var steps = [
    { num: "1", text: "脱水や低栄養など急いで対応すべき問題はないか", color: FR.red },
    { num: "2", text: "体重は減っていないか", color: FR.warm },
    { num: "3", text: "1日の総摂取量は確保できているか", color: FR.gold },
    { num: "4", text: "たんぱく質源は毎日入っているか", color: FR.teal },
    { num: "5", text: "水分は足りているか", color: FR.blue },
    { num: "6", text: "そのうえで塩分・糖質・脂質・カリウムの個別調整", color: FR.indigo },
    { num: "7", text: "本人と家族が実行できる内容になっているか", color: C.primary },
  ];

  for (var i = 0; i < steps.length; i++) {
    var yPos = 1.05 + i * 0.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.46, rectRadius: 0.06, fill: { color: C.white }, shadow: shd() });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.05, w: 0.36, h: 0.36, fill: { color: steps[i].color } });
    s.addText(steps[i].num, { x: 1.0, y: yPos + 0.05, w: 0.36, h: 0.36, fontSize: 14, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(steps[i].text, { x: 1.6, y: yPos, w: 7.4, h: 0.46, fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  }

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.9, w: 8.4, h: 0.45, rectRadius: 0.1, fill: { color: C.lightYellow } });
  s.addText("この順番で考えると、制限の足し算に陥りにくくなる", {
    x: 1.0, y: 4.9, w: 8.0, h: 0.45,
    fontSize: 15, fontFace: FJ, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0,
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
    { myth: "病気が多いほど制限を厳しくすべき", fact: "制限の足し算は低栄養やフレイルを進める危険がある" },
    { myth: "数値が改善すれば食事指導は成功", fact: "食事量↓や体重減少が起きていたら成功とは言えない" },
    { myth: "高齢者にも若年者と同じ制限を", fact: "高齢者では制限しすぎが最も危険なことがある" },
  ];

  for (var i = 0; i < myths.length; i++) {
    var yPos = 1.2 + i * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: yPos, w: 8.8, h: 1.0, rectRadius: 0.08, fill: { color: C.white }, shadow: shd() });
    s.addText("✕", { x: 0.8, y: yPos + 0.05, w: 0.5, h: 0.4, fontSize: 20, fontFace: FE, color: FR.red, bold: true, align: "center", margin: 0 });
    s.addText(myths[i].myth, { x: 1.3, y: yPos + 0.05, w: 8.0, h: 0.4, fontSize: 18, fontFace: FJ, color: FR.red, bold: true, margin: 0 });
    s.addText("→ " + myths[i].fact, { x: 1.3, y: yPos + 0.52, w: 8.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.primary, margin: 0 });
  }
  footer(s);
})();

// ============================================================
// SLIDE 10: シリーズ全体のまとめ（特別スライド）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("シリーズ全10回を通じた原則", {
    x: 0.6, y: 0.6, w: 8.8, h: 0.5,
    fontSize: 16, fontFace: FJ, color: "A5D6A7", align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 2.5, y: 1.3, w: 5, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("食事指導は\n「何を禁止するか」ではなく\n「何を守るか」を決めること", {
    x: 0.6, y: 1.6, w: 8.8, h: 2.0,
    fontSize: 32, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 2.5, y: 3.8, w: 5, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("完璧な制限を目指すより\n崩してはいけない土台を守る方が\n臨床的に価値がある", {
    x: 0.6, y: 4.0, w: 8.8, h: 1.2,
    fontSize: 18, fontFace: FJ, color: "B0BEC5", align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 11: まとめ
// ============================================================
T.addSummarySlide(pres, 10, [
  "複数疾患では制限の足し算をしない",
  "まず生命に直結する問題を確認し低栄養・脱水を防ぐ",
  "フレイル高齢者では食べる力と体重・筋肉量を守ることが優先",
  "数値改善と食べられることのバランスを常に意識",
  "食事指導は「何を禁止するか」ではなく「何を守るか」",
]);

// ============================================================
// SLIDE 12: エンド（最終回特別版）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };

  s.addText("シリーズ全10回\nご視聴ありがとうございました", {
    x: 0.6, y: 0.8, w: 8.8, h: 1.0,
    fontSize: 30, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 2.0, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  var items = [
    "📋 再生リストから他の回もぜひご覧ください",
    "👍 高評価・チャンネル登録よろしくお願いします",
    "💬 コメント欄で質問・リクエストお待ちしています",
  ];
  for (var i = 0; i < items.length; i++) {
    s.addText(items[i], {
      x: 1.5, y: 2.4 + i * 0.7, w: 7, h: 0.5,
      fontSize: 18, fontFace: FJ, color: "E8F5E9", align: "center", margin: 0,
    });
  }

  s.addText(T.SERIES_NAME, {
    x: 0.6, y: 4.6, w: 8.8, h: 0.4,
    fontSize: 12, fontFace: FJ, color: "A5D6A7", align: "center", margin: 0,
  });
})();

// ============================================================
// SAVE
// ============================================================
T.save(pres, "slides.pptx");
