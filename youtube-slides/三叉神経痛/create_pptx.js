var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "三叉神経痛 ― 神経内科医視点の鑑別診断と治療戦略";

// --- Color Palette ---
var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "4A90D9",
  light: "E8F0FE",
  warmBg: "F5F7FA",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "DC3545",
  orange: "E8850C",
  yellow: "F5C518",
  green: "28A745",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var FONT_H = FONT_JP;
var FONT_B = FONT_JP;
var TOTAL = 18;

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
}

function addCard(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: C.white }, shadow: makeShadow() });
}

function addPageNum(s, n) {
  s.addText(n + "/" + TOTAL, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_B, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: Title
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("三叉神経痛", {
    x: 0.8, y: 0.8, w: 8.4, h: 1.2,
    fontSize: 44, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("Trigeminal Neuralgia", {
    x: 0.8, y: 1.9, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_EN, color: C.accent, align: "center", italic: true, margin: 0,
  });
  s.addText("― 神経内科医視点の鑑別診断と治療戦略 ―", {
    x: 0.8, y: 2.6, w: 8.4, h: 0.7,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 3.5, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.8, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("最新レビュー・ガイドラインに基づく", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 三叉神経痛とは
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "三叉神経痛とは");

  addCard(s, 0.5, 1.2, 9, 1.5);
  s.addText("三叉神経（第V脳神経）の支配領域に生じる", { x: 0.8, y: 1.3, w: 8.4, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0 });
  s.addText("発作性の電撃様激痛", { x: 0.8, y: 1.8, w: 8.4, h: 0.7, fontSize: 30, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0 });

  // Key stats
  var stats = [
    { num: "4-5.5", unit: "/10万人・年", desc: "発症率" },
    { num: "F:M=3:2", unit: "", desc: "男女比" },
    { num: "50歳〜", unit: "", desc: "好発年齢" },
    { num: "0.1-0.3%", unit: "", desc: "生涯有病率" },
  ];
  stats.forEach(function(item, i) {
    var xPos = 0.3 + i * 2.4;
    addCard(s, xPos, 3.0, 2.2, 1.8);
    s.addText(item.num + item.unit, { x: xPos, y: 3.1, w: 2.2, h: 0.8, fontSize: 20, fontFace: FONT_EN, color: C.primary, bold: true, align: "center", margin: 0 });
    s.addText(item.desc, { x: xPos, y: 3.9, w: 2.2, h: 0.6, fontSize: 15, fontFace: FONT_B, color: C.text, align: "center", margin: 0 });
  });

  addPageNum(s, 2);
})();

// ============================================================
// SLIDE 3: 三叉神経の解剖
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "三叉神経の解剖 ― V1 / V2 / V3");

  var branches = [
    { name: "V1 眼神経", via: "上眼窩裂", area: "前頭部\n上眼瞼\n鼻背", color: "4A90D9", freq: "<5%" },
    { name: "V2 上顎神経", via: "正円孔", area: "頬部\n上唇\n上歯列", color: "E8850C", freq: "35%" },
    { name: "V3 下顎神経", via: "卵円孔", area: "下唇・下顎\n下歯列\n舌前2/3", color: "28A745", freq: "30%" },
  ];

  branches.forEach(function(b, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.2, 3.0, 3.2);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 3.0, h: 0.5, fill: { color: b.color } });
    s.addText(b.name, { x: xPos, y: 1.2, w: 3.0, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText("通過孔: " + b.via, { x: xPos + 0.2, y: 1.8, w: 2.6, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0 });
    s.addText(b.area, { x: xPos + 0.2, y: 2.2, w: 2.6, h: 1.2, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3 });
    s.addText("TN頻度: " + b.freq, { x: xPos + 0.2, y: 3.5, w: 2.6, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  });

  // V2+V3 combined
  addCard(s, 0.5, 4.6, 9, 0.7);
  s.addText("V2+V3の複合: 約30%  ｜  V1単独は5%未満 → 腫瘍・TACsを鑑別", {
    x: 0.8, y: 4.7, w: 8.4, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  addPageNum(s, 3);
})();

// ============================================================
// SLIDE 4: 病態生理
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "病態生理 ― 神経血管圧迫説");

  // Flow: compression -> demyelination -> ectopic -> pain
  var steps = [
    { title: "血管圧迫", desc: "REZで動脈が\n三叉神経を圧迫", icon: "1" },
    { title: "局所脱髄", desc: "持続的圧迫により\n神経鞘が障害", icon: "2" },
    { title: "異所性発火", desc: "脱髄部位で\n自発的インパルス発生", icon: "3" },
    { title: "エファプティック伝導", desc: "触覚信号が\n疼痛信号に変換", icon: "4" },
  ];

  steps.forEach(function(st, i) {
    var xPos = 0.2 + i * 2.45;
    addCard(s, xPos, 1.2, 2.3, 2.0);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.85, y: 1.3, w: 0.6, h: 0.6, fill: { color: C.primary } });
    s.addText(st.icon, { x: xPos + 0.85, y: 1.3, w: 0.6, h: 0.6, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.title, { x: xPos + 0.1, y: 2.0, w: 2.1, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0 });
    s.addText(st.desc, { x: xPos + 0.1, y: 2.4, w: 2.1, h: 0.7, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // Arrows between steps
  [2.55, 5.0, 7.45].forEach(function(ax) {
    s.addText("\u2192", { x: ax, y: 1.8, w: 0.4, h: 0.5, fontSize: 28, fontFace: FONT_EN, color: C.accent, align: "center", margin: 0 });
  });

  // Responsible vessel
  addCard(s, 0.5, 3.5, 9, 1.8);
  s.addText("責任血管の頻度", { x: 0.8, y: 3.6, w: 3, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });

  var vessels = [
    { name: "SCA（上小脳動脈）", pct: "75-80%", w: 7.5 },
    { name: "AICA（前下小脳動脈）", pct: "10-15%", w: 2.5 },
    { name: "椎骨脳底動脈系", pct: "5-10%", w: 1.5 },
    { name: "静脈単独", pct: "<5%", w: 0.8 },
  ];

  vessels.forEach(function(v, i) {
    var yPos = 4.1 + i * 0.28;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: v.w * 0.5, h: 0.22, fill: { color: i === 0 ? C.primary : (i === 3 ? C.red : C.accent) } });
    s.addText(v.name + " " + v.pct, { x: 0.8 + v.w * 0.5 + 0.2, y: yPos - 0.02, w: 5, h: 0.26, fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  addPageNum(s, 4);
})();

// ============================================================
// SLIDE 5: ICHD-3分類
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ICHD-3に基づく臨床分類");

  var types = [
    { name: "典型的TN\n(Classical)", pct: "70-80%", desc: "MRI/手術で\n神経血管圧迫を確認", color: C.primary, sub: "純粋発作型 / 持続性疼痛併存型" },
    { name: "二次性TN\n(Secondary)", pct: "~15%", desc: "基礎疾患に起因\nMS・腫瘍・AVM等", color: C.orange, sub: "若年・両側性・感覚障害に注意" },
    { name: "特発性TN\n(Idiopathic)", pct: "10-15%", desc: "血管圧迫なし\n基礎疾患なし", color: C.sub, sub: "臨床的にTNの特徴を満たす" },
  ];

  types.forEach(function(t, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.2, 3.0, 3.6);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 3.0, h: 0.6, fill: { color: t.color } });
    s.addText(t.name, { x: xPos, y: 1.2, w: 3.0, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.9, y: 2.0, w: 1.2, h: 0.8, fill: { color: t.color }, opacity: 0.15 });
    s.addText(t.pct, { x: xPos + 0.9, y: 2.0, w: 1.2, h: 0.8, fontSize: 20, fontFace: FONT_EN, color: t.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.desc, { x: xPos + 0.2, y: 2.9, w: 2.6, h: 0.8, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.2 });
    s.addText(t.sub, { x: xPos + 0.2, y: 3.8, w: 2.6, h: 0.5, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0 });
  });

  addPageNum(s, 5);
})();

// ============================================================
// SLIDE 6: 典型的な症状
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "典型的な症状 ― 「稲妻が走る」");

  var items = [
    { label: "痛みの性質", val: "電撃様・刺すような激痛" },
    { label: "持続時間", val: "数秒〜2分（通常1分以内）" },
    { label: "頻度", val: "1日数回〜数百回" },
    { label: "部位", val: "片側性（右側60%）、V2/V3領域" },
    { label: "トリガー", val: "咀嚼・会話・洗顔・歯磨き・冷風" },
    { label: "不応期", val: "発作後に短い無痛期間" },
    { label: "寛解期", val: "数週〜数ヶ月（経過とともに短縮）" },
  ];

  items.forEach(function(item, i) {
    var yPos = 1.15 + i * 0.58;
    addCard(s, 0.5, yPos, 9, 0.52);
    s.addText(item.label, { x: 0.7, y: yPos + 0.05, w: 2.5, h: 0.42, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0, valign: "middle" });
    s.addText(item.val, { x: 3.3, y: yPos + 0.05, w: 6, h: 0.42, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle" });
  });

  // Pearl
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.5, fill: { color: "E8F5E9" } });
  s.addText("Pearl: 診断はほぼ病歴のみで可能 ― 電撃様発作痛 + 三叉神経領域 + トリガー", {
    x: 0.7, y: 5.05, w: 8.6, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true, margin: 0,
  });

  addPageNum(s, 6);
})();

// ============================================================
// SLIDE 7: 診断基準（ICHD-3）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "診断基準 ― ICHD-3");

  addCard(s, 0.5, 1.2, 9, 4.0);
  s.addText("典型的三叉神経痛の診断基準（ICHD-3: 13.1.1）", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  var criteria = [
    "A. 三叉神経の1枝以上の支配領域に生じる発作性顔面痛で、\n   基準B-Dを満たす",
    "B. 痛みは以下の特徴をすべて持つ:\n   1) 数秒〜2分持続  2) 激痛  3) 電撃様・刺痛・鋭い性質",
    "C. 患側の三叉神経支配領域への無害刺激により誘発される",
    "D. 他の ICHD-3 診断では説明できない",
    "E. MRIまたは手術により神経血管圧迫\n   （三叉神経根の形態変化を伴う）が実証されている",
  ];

  criteria.forEach(function(c, i) {
    var yPos = 1.9 + i * 0.65;
    s.addShape(pres.shapes.OVAL, { x: 0.9, y: yPos + 0.05, w: 0.35, h: 0.35, fill: { color: C.primary } });
    s.addText(String.fromCharCode(65 + i), { x: 0.9, y: yPos + 0.05, w: 0.35, h: 0.35, fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c, { x: 1.4, y: yPos, w: 7.8, h: 0.6, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.15 });
  });

  addPageNum(s, 7);
})();

// ============================================================
// SLIDE 8: トリガーゾーンと誘発因子
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "トリガーゾーンと誘発因子");

  // Trigger zones
  addCard(s, 0.5, 1.2, 4.3, 4.2);
  s.addText("トリガーゾーン", { x: 0.8, y: 1.3, w: 3.7, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  var zones = [
    "鼻唇溝（最多）",
    "上唇 / 下唇",
    "歯肉",
    "頬部",
    "鼻翼",
  ];
  zones.forEach(function(z, i) {
    s.addText("\u2022 " + z, { x: 1.0, y: 1.8 + i * 0.38, w: 3.5, h: 0.33, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.85, w: 3.7, h: 0.5, fill: { color: "FFF3CD" } });
  s.addText("軽い触覚で発作が誘発される\n（痛覚刺激は不要）", {
    x: 0.9, y: 3.85, w: 3.5, h: 0.5, fontSize: 12, fontFace: FONT_JP, color: C.orange, margin: 0, lineSpacingMultiple: 1.2,
  });

  // Triggers
  addCard(s, 5.2, 1.2, 4.3, 4.2);
  s.addText("誘発因子", { x: 5.5, y: 1.3, w: 3.7, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  var triggers = [
    { act: "咀嚼（そしゃく）", icon: "" },
    { act: "会話", icon: "" },
    { act: "洗顔", icon: "" },
    { act: "歯磨き", icon: "" },
    { act: "冷風に当たる", icon: "" },
    { act: "化粧・ひげ剃り", icon: "" },
  ];
  triggers.forEach(function(t, i) {
    s.addText("\u2022 " + t.act, { x: 5.7, y: 1.8 + i * 0.33, w: 3.3, h: 0.3, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0 });
  });

  // Refractory period
  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 3.85, w: 3.7, h: 0.5, fill: { color: "E8F0FE" } });
  s.addText("発作後に不応期あり\n（一定時間誘発されない）", {
    x: 5.6, y: 3.85, w: 3.5, h: 0.5, fontSize: 12, fontFace: FONT_JP, color: C.primary, margin: 0, lineSpacingMultiple: 1.2,
  });

  addPageNum(s, 8);
})();

// ============================================================
// SLIDE 9: Red Flags
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "Red Flags ― 二次性TNを疑う所見");

  var flags = [
    { sign: "両側性", suggest: "多発性硬化症（MS）" },
    { sign: "V1領域単独", suggest: "小脳橋角部腫瘍" },
    { sign: "感覚低下・運動障害", suggest: "構造的病変（腫瘍等）" },
    { sign: "若年発症（<40歳）", suggest: "MSの検索が必須" },
    { sign: "難聴・めまい合併", suggest: "聴神経鞘腫・髄膜腫" },
    { sign: "CBZ無効", suggest: "診断の再評価" },
    { sign: "進行性（持続痛へ移行）", suggest: "二次性の可能性を再検討" },
  ];

  flags.forEach(function(f, i) {
    var yPos = 1.15 + i * 0.58;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9, h: 0.52, fill: { color: i % 2 === 0 ? "FDF2F2" : C.white }, shadow: i % 2 === 0 ? undefined : makeShadow() });
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.08, w: 0.36, h: 0.36, fill: { color: C.red } });
    s.addText("!", { x: 0.7, y: yPos + 0.08, w: 0.36, h: 0.36, fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.sign, { x: 1.2, y: yPos + 0.05, w: 3.5, h: 0.42, fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true, margin: 0, valign: "middle" });
    s.addText("\u2192 " + f.suggest, { x: 4.8, y: yPos + 0.05, w: 4.5, h: 0.42, fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle" });
  });

  addPageNum(s, 9);
})();

// ============================================================
// SLIDE 10: 鑑別疾患一覧
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "顔面痛の鑑別診断");

  var rows = [
    ["疾患", "痛みの特徴", "持続時間", "鑑別ポイント"],
    ["歯原性疼痛", "拍動性・持続性", "持続〜数時間", "歯科検査で原因歯同定"],
    ["帯状疱疹後神経痛", "灼熱感・アロディニア", "持続（慢性）", "水疱既往・皮疹"],
    ["群発頭痛(TACs)", "激痛+自律神経症状", "15-180分", "流涙・鼻閉・縮瞳"],
    ["SUNCT/SUNA", "刺痛+結膜充血", "1-600秒", "自律神経症状が顕著"],
    ["PIFP", "鈍痛・漠然", "持続（慢性）", "局在不明瞭"],
    ["顎関節症(TMD)", "鈍痛・咀嚼時", "持続〜数時間", "開口制限・関節雑音"],
    ["舌咽神経痛", "電撃様・咽頭耳放散", "数秒〜2分", "嚥下がトリガー"],
  ];

  rows.forEach(function(row, i) {
    var yPos = 1.1 + i * 0.55;
    var isHeader = (i === 0);
    var bgColor = isHeader ? C.primary : (i % 2 === 0 ? "F8F9FA" : C.white);
    var txtColor = isHeader ? C.white : C.text;
    var ws = [2.3, 2.2, 1.8, 2.7];
    var xAcc = 0.5;
    row.forEach(function(cell, j) {
      s.addShape(pres.shapes.RECTANGLE, { x: xAcc, y: yPos, w: ws[j], h: 0.5, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      s.addText(cell, { x: xAcc + 0.1, y: yPos, w: ws[j] - 0.2, h: 0.5, fontSize: isHeader ? 12 : 11, fontFace: FONT_JP, color: txtColor, bold: isHeader, valign: "middle", margin: 0 });
      xAcc += ws[j];
    });
  });

  addPageNum(s, 10);
})();

// ============================================================
// SLIDE 11: SUNCT/SUNAとの鑑別
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "SUNCT/SUNA vs 三叉神経痛 ― 最重要鑑別");

  var rows = [
    ["特徴", "三叉神経痛", "SUNCT/SUNA"],
    ["持続時間", "数秒〜2分", "1-600秒"],
    ["自律神経症状", "通常なし", "必発（充血・流涙）"],
    ["不応期", "あり", "なし"],
    ["トリガー", "軽触覚", "皮膚刺激（不応期なし）"],
    ["CBZ反応", "著効", "無効"],
    ["有効薬", "CBZ / OXC", "ラモトリギン"],
    ["好発部位", "V2 / V3", "V1（眼窩部）"],
  ];

  rows.forEach(function(row, i) {
    var yPos = 1.1 + i * 0.55;
    var isHeader = (i === 0);
    var bgColor = isHeader ? C.primary : (i % 2 === 0 ? "F8F9FA" : C.white);
    var txtColor = isHeader ? C.white : C.text;
    var ws = [2.5, 3.5, 3.0];
    var xAcc = 0.5;
    row.forEach(function(cell, j) {
      var cellColor = txtColor;
      if (!isHeader && j === 2) cellColor = C.orange;
      s.addShape(pres.shapes.RECTANGLE, { x: xAcc, y: yPos, w: ws[j], h: 0.5, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      s.addText(cell, { x: xAcc + 0.1, y: yPos, w: ws[j] - 0.2, h: 0.5, fontSize: isHeader ? 13 : 12, fontFace: FONT_JP, color: cellColor, bold: isHeader || (j === 0 && !isHeader), valign: "middle", margin: 0 });
      xAcc += ws[j];
    });
  });

  // Pearl
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.5, fill: { color: "FFF3CD" } });
  s.addText("Key: 自律神経症状の有無 + CBZへの反応性 が最大の鑑別ポイント", {
    x: 0.7, y: 5.05, w: 8.6, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });

  addPageNum(s, 11);
})();

// ============================================================
// SLIDE 12: 画像検査
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "画像検査 ― MRI/MRA/3D-FIESTA");

  var exams = [
    { name: "MRI（通常撮像）", purpose: "二次性TNの除外", finding: "腫瘍、MS plaque、\n血管奇形の検出", level: "全例で推奨\n(AAN Level B)" },
    { name: "MRA", purpose: "責任血管の同定", finding: "SCA等の走行異常\n三叉神経との接触", level: "MVD検討時" },
    { name: "3D-FIESTA/CISS", purpose: "神経-血管関係の\n精密評価", finding: "REZでの圧迫・変形を\n高分解能で描出", level: "MVD術前評価\nに必須" },
    { name: "造影MRI", purpose: "腫瘍性病変の評価", finding: "聴神経鞘腫・髄膜腫の\n造影効果", level: "二次性TN\n疑い時" },
  ];

  exams.forEach(function(e, i) {
    var yPos = 1.15 + i * 1.05;
    addCard(s, 0.5, yPos, 9, 0.95);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 2.3, h: 0.95, fill: { color: C.primary } });
    s.addText(e.name, { x: 0.6, y: yPos + 0.1, w: 2.1, h: 0.75, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
    s.addText(e.purpose, { x: 2.9, y: yPos + 0.05, w: 2.2, h: 0.85, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0, valign: "middle" });
    s.addText(e.finding, { x: 5.2, y: yPos + 0.05, w: 2.5, h: 0.85, fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, valign: "middle", lineSpacingMultiple: 1.2 });
    s.addText(e.level, { x: 7.8, y: yPos + 0.05, w: 1.5, h: 0.85, fontSize: 11, fontFace: FONT_JP, color: C.sub, margin: 0, valign: "middle", align: "center", lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, 12);
})();

// ============================================================
// SLIDE 13: 薬物治療①（第一選択）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "薬物治療 \u2460 ― 第一選択薬");

  // CBZ card
  addCard(s, 0.5, 1.2, 4.3, 2.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.5, fill: { color: C.primary } });
  s.addText("カルバマゼピン (CBZ)", { x: 0.6, y: 1.2, w: 4.1, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0, valign: "middle" });
  s.addText([
    { text: "用量: ", options: { bold: true, color: C.primary } },
    { text: "200-1200 mg/日", options: { color: C.text, breakLine: true } },
    { text: "エビデンス: ", options: { bold: true, color: C.primary } },
    { text: "Level A (AAN)", options: { color: C.green, bold: true, breakLine: true } },
    { text: "奏効率: ", options: { bold: true, color: C.primary } },
    { text: "初期88%, NNT 1.7-1.8", options: { color: C.text, breakLine: true } },
    { text: "副作用: ", options: { bold: true, color: C.primary } },
    { text: "眠気, ふらつき, 低Na,\nSJS/TEN, 肝障害 (44%)", options: { color: C.red } },
  ], { x: 0.8, y: 1.8, w: 3.8, h: 2.0, fontSize: 13, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.4 });

  // OXC card
  addCard(s, 5.2, 1.2, 4.3, 2.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.5, fill: { color: C.accent } });
  s.addText("オキシカルバゼピン (OXC)", { x: 5.3, y: 1.2, w: 4.1, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, margin: 0, valign: "middle" });
  s.addText([
    { text: "用量: ", options: { bold: true, color: C.primary } },
    { text: "600-1800 mg/日", options: { color: C.text, breakLine: true } },
    { text: "エビデンス: ", options: { bold: true, color: C.primary } },
    { text: "Level B (AAN)", options: { color: C.accent, bold: true, breakLine: true } },
    { text: "奏効率: ", options: { bold: true, color: C.primary } },
    { text: "初期91%, CBZと同等効果", options: { color: C.text, breakLine: true } },
    { text: "副作用: ", options: { bold: true, color: C.primary } },
    { text: "CBZより少ない (30%),\n低Na血症に注意", options: { color: C.orange } },
  ], { x: 5.5, y: 1.8, w: 3.8, h: 2.0, fontSize: 13, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.4 });

  // HLA warning
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.2, w: 9, h: 1.0, fill: { color: "F8D7DA" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.2, w: 0.15, h: 1.0, fill: { color: C.red } });
  s.addText("HLA-B*1502検査（アジア人で必須）", { x: 0.9, y: 4.25, w: 8, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText("CBZ/OXC開始前にSJS/TENリスク評価。日本人<1%だが台湾~10%・東南アジア>15%\n陽性者ではCBZ/OXCは禁忌 → プレガバリン・ミロガバリン等を代替使用", {
    x: 0.9, y: 4.6, w: 8.4, h: 0.5, fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2,
  });

  addPageNum(s, 13);
})();

// ============================================================
// SLIDE 14: 薬物治療②（第二選択）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "薬物治療 \u2461 ― 第二選択薬・補助薬");

  var drugs = [
    ["薬剤", "用量", "備考"],
    ["ラモトリギン", "200-400 mg/日", "Level C, 皮疹に注意（漸増必須）"],
    ["バクロフェン", "40-80 mg/日", "CBZ併用で有効性のエビデンスあり"],
    ["プレガバリン", "150-600 mg/日", "CBZアレルギー時の代替"],
    ["ミロガバリン", "10-30 mg/日", "日本で使用可能なCa2+チャネルリガンド"],
    ["フェニトイン", "静注250mg", "急性期の疼痛コントロール"],
    ["ボツリヌス毒素A", "皮下注射", "新規エビデンス蓄積中"],
  ];

  drugs.forEach(function(row, i) {
    var yPos = 1.1 + i * 0.58;
    var isHeader = (i === 0);
    var bgColor = isHeader ? C.primary : (i % 2 === 0 ? "F8F9FA" : C.white);
    var txtColor = isHeader ? C.white : C.text;
    var ws = [2.5, 2.5, 4.0];
    var xAcc = 0.5;
    row.forEach(function(cell, j) {
      s.addShape(pres.shapes.RECTANGLE, { x: xAcc, y: yPos, w: ws[j], h: 0.52, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      s.addText(cell, { x: xAcc + 0.1, y: yPos, w: ws[j] - 0.2, h: 0.52, fontSize: isHeader ? 13 : 12, fontFace: FONT_JP, color: txtColor, bold: isHeader || (j === 0 && !isHeader), valign: "middle", margin: 0 });
      xAcc += ws[j];
    });
  });

  // Pearl
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.5, fill: { color: "E8F5E9" } });
  s.addText("Pearl: CBZへの劇的反応（NNT 1.7-1.8）は三叉神経痛の診断を強く支持する", {
    x: 0.7, y: 5.05, w: 8.6, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true, margin: 0,
  });

  addPageNum(s, 14);
})();

// ============================================================
// SLIDE 15: 外科治療の種類と比較
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "外科治療の種類と比較");

  var rows = [
    ["治療法", "疼痛軽減率", "再発率", "特徴"],
    ["MVD", "90-95%", "15-25%/10年", "唯一の原因治療"],
    ["高周波熱凝固術(RF)", "90-97%", "20-30%/5年", "顔面しびれ必発"],
    ["バルーン圧迫術", "80-90%", "20-30%/5年", "全身麻酔不要のオプション"],
    ["グリセロール注入", "70-90%", "30-50%/5年", "顔面しびれ比較的少ない"],
    ["ガンマナイフ", "70-90%", "30-50%/5年", "非侵襲, 効果発現に時間"],
  ];

  rows.forEach(function(row, i) {
    var yPos = 1.1 + i * 0.7;
    var isHeader = (i === 0);
    var bgColor = isHeader ? C.primary : (i === 1 ? "E8F0FE" : (i % 2 === 0 ? "F8F9FA" : C.white));
    var txtColor = isHeader ? C.white : C.text;
    var ws = [2.8, 1.5, 1.8, 2.9];
    var xAcc = 0.5;
    row.forEach(function(cell, j) {
      var cellColor = txtColor;
      if (i === 1 && j === 0) cellColor = C.primary;
      s.addShape(pres.shapes.RECTANGLE, { x: xAcc, y: yPos, w: ws[j], h: 0.64, fill: { color: bgColor }, line: { color: "DEE2E6", width: 0.5 } });
      s.addText(cell, { x: xAcc + 0.1, y: yPos, w: ws[j] - 0.2, h: 0.64, fontSize: isHeader ? 13 : 12, fontFace: FONT_JP, color: cellColor, bold: isHeader || (i === 1), valign: "middle", margin: 0 });
      xAcc += ws[j];
    });
  });

  // Note
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.8, w: 9, h: 0.5, fill: { color: "E8F0FE" } });
  s.addText("MVDは原因治療（血管を移動）、他は対症治療（神経を破壊/損傷）", {
    x: 0.7, y: 4.85, w: 8.6, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  addPageNum(s, 15);
})();

// ============================================================
// SLIDE 16: MVDの詳細
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "MVD（微小血管減圧術）― 詳細");

  // Procedure
  addCard(s, 0.5, 1.2, 4.3, 2.0);
  s.addText("手技", { x: 0.8, y: 1.3, w: 3.7, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "\u2022 後頭蓋窩開頭", options: { breakLine: true } },
    { text: "\u2022 責任血管をテフロンフェルトで移動", options: { breakLine: true } },
    { text: "\u2022 三叉神経への圧迫を解除", options: { breakLine: true } },
    { text: "\u2022 唯一の原因治療", options: { color: C.primary, bold: true } },
  ], { x: 0.8, y: 1.7, w: 3.8, h: 1.4, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.4 });

  // Outcomes
  addCard(s, 5.2, 1.2, 4.3, 2.0);
  s.addText("治療成績", { x: 5.5, y: 1.3, w: 3.7, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  var outcomes = [
    { label: "初期成功率", val: "90-95%", color: C.green },
    { label: "長期5年", val: "80%+", color: C.primary },
    { label: "内視鏡MVD", val: "同等", color: C.accent },
    { label: "再発率10年", val: "15-25%", color: C.orange },
  ];
  outcomes.forEach(function(o, i) {
    s.addText(o.label + ": ", { x: 5.5, y: 1.7 + i * 0.35, w: 1.8, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.text, bold: true, margin: 0 });
    s.addText(o.val, { x: 7.3, y: 1.7 + i * 0.35, w: 2, h: 0.3, fontSize: 15, fontFace: FONT_EN, color: o.color, bold: true, margin: 0 });
  });

  // Complications
  addCard(s, 0.5, 3.4, 4.3, 1.8);
  s.addText("合併症", { x: 0.8, y: 3.5, w: 3.7, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
  s.addText([
    { text: "\u2022 聴力障害: 1-4%", options: { breakLine: true } },
    { text: "\u2022 顔面しびれ: 1-10%", options: { breakLine: true } },
    { text: "\u2022 髄液漏: 1-2%", options: { breakLine: true } },
    { text: "\u2022 死亡: <0.5%", options: {} },
  ], { x: 0.8, y: 3.9, w: 3.8, h: 1.2, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3 });

  // Prognostic factors
  addCard(s, 5.2, 3.4, 4.3, 1.8);
  s.addText("予後因子", { x: 5.5, y: 3.5, w: 3.7, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "良好: ", options: { bold: true, color: C.green } },
    { text: "動脈圧迫、典型的TN", options: { color: C.text, breakLine: true } },
    { text: "不良: ", options: { bold: true, color: C.red } },
    { text: "静脈のみ圧迫、持続痛併存", options: { color: C.text, breakLine: true } },
    { text: "\n二次性TN（MS）ではMVDの\n成績はやや劣る", options: { color: C.sub, italic: true } },
  ], { x: 5.5, y: 3.9, w: 3.8, h: 1.2, fontSize: 13, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, 16);
})();

// ============================================================
// SLIDE 17: 治療アルゴリズム
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "治療アルゴリズム");

  // Step 1
  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 1.05, w: 5, h: 0.55, fill: { color: C.primary }, rectRadius: 0.1 });
  s.addText("Step 1: 診断・MRI・HLA-B*1502検査", { x: 2.5, y: 1.05, w: 5, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("\u25BC", { x: 4.7, y: 1.58, w: 0.6, h: 0.25, fontSize: 16, fontFace: FONT_EN, color: C.primary, align: "center", margin: 0 });

  // Step 2
  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 1.82, w: 5, h: 0.55, fill: { color: C.accent }, rectRadius: 0.1 });
  s.addText("Step 2: CBZ or OXC で薬物治療開始", { x: 2.5, y: 1.82, w: 5, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // Branches from Step 2 — lines centered on each target box
  // Left: effective (box center = 1.75)
  s.addShape(pres.shapes.LINE, { x: 1.75, y: 2.37, w: 0, h: 0.33, line: { color: C.green, width: 2 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.7, w: 2.5, h: 0.48, fill: { color: "E8F5E9" }, rectRadius: 0.1 });
  s.addText("\u2714 有効 \u2192 継続", { x: 0.5, y: 2.7, w: 2.5, h: 0.48, fontSize: 13, fontFace: FONT_JP, color: C.green, bold: true, align: "center", valign: "middle", margin: 0 });

  // Center: insufficient (box center = 5.0)
  s.addShape(pres.shapes.LINE, { x: 5.0, y: 2.37, w: 0, h: 0.33, line: { color: C.orange, width: 2 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 2.7, w: 3.0, h: 0.48, fill: { color: "FFF3CD" }, rectRadius: 0.1 });
  s.addText("不十分 \u2192 2nd line追加", { x: 3.5, y: 2.7, w: 3.0, h: 0.48, fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, align: "center", valign: "middle", margin: 0 });

  // Right: resistant (box center = 8.25)
  s.addShape(pres.shapes.LINE, { x: 8.25, y: 2.37, w: 0, h: 0.33, line: { color: C.red, width: 2 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 2.7, w: 2.5, h: 0.48, fill: { color: "F8D7DA" }, rectRadius: 0.1 });
  s.addText("抵抗性/副作用", { x: 7.0, y: 2.7, w: 2.5, h: 0.48, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });

  // Arrows down from "不十分" and "抵抗性" to Step 3
  s.addText("\u25BC", { x: 4.7, y: 3.16, w: 0.6, h: 0.24, fontSize: 14, fontFace: FONT_EN, color: C.orange, align: "center", margin: 0 });
  s.addText("\u25BC", { x: 7.95, y: 3.16, w: 0.6, h: 0.24, fontSize: 14, fontFace: FONT_EN, color: C.red, align: "center", margin: 0 });

  // Step 3
  s.addShape(pres.shapes.RECTANGLE, { x: 2.0, y: 3.4, w: 6.0, h: 0.5, fill: { color: C.dark }, rectRadius: 0.1 });
  s.addText("Step 3: 外科治療の検討", { x: 2.0, y: 3.4, w: 6.0, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // Three surgical options — lines centered within Step 3 bounds (x:2.0 ~ 8.0)
  var surgeries = [
    { name: "MVD", cond: "全身状態良好\n+ 典型的TN", color: C.primary },
    { name: "ガンマナイフ", cond: "高齢 or\n手術リスク高", color: C.accent },
    { name: "経皮的手技", cond: "MVD不適\nor 再発例", color: C.orange },
  ];

  surgeries.forEach(function(sg, i) {
    var xPos = 0.7 + i * 3.1;
    var boxCenter = xPos + 1.35;
    s.addShape(pres.shapes.LINE, { x: boxCenter, y: 3.9, w: 0, h: 0.25, line: { color: sg.color, width: 2 } });
    addCard(s, xPos, 4.15, 2.7, 0.85);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.15, w: 2.7, h: 0.28, fill: { color: sg.color } });
    s.addText(sg.name, { x: xPos, y: 4.15, w: 2.7, h: 0.28, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(sg.cond, { x: xPos + 0.1, y: 4.45, w: 2.5, h: 0.5, fontSize: 11, fontFace: FONT_JP, color: C.text, align: "center", margin: 0, lineSpacingMultiple: 1.1 });
  });

  addPageNum(s, 17);
})();

// ============================================================
// SLIDE 18: Take Home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.3, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "三叉神経痛の診断はほぼ病歴のみで可能\n電撃様発作痛 + 三叉神経領域 + トリガー", color: C.accent },
    { num: "2", text: "ICHD-3分類: 典型的 / 二次性 / 特発性\nRed Flagsを見逃さない（両側性、V1単独、若年）", color: C.yellow },
    { num: "3", text: "SUNCT/SUNAとの鑑別が最重要\n自律神経症状の有無 + CBZへの反応性", color: C.orange },
    { num: "4", text: "第一選択薬はCBZ（Level A）/ OXC（Level B）\nアジア人ではHLA-B*1502検査が必須", color: C.green },
    { num: "5", text: "MVDは唯一の原因治療（成功率90-95%）\n薬剤抵抗性では早期に外科コンサルト", color: C.red },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.2 + i * 0.85;
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: yPos + 0.1, w: 0.5, h: 0.5, fill: { color: m.color } });
    s.addText(m.num, { x: 0.8, y: yPos + 0.1, w: 0.5, h: 0.5, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.5, y: yPos, w: 8, h: 0.75, fontSize: 14, fontFace: FONT_JP, color: C.white, margin: 0, lineSpacingMultiple: 1.25 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  addPageNum(s, 18);
})();

// ============================================================
// OUTPUT
// ============================================================
pres.writeFile({ fileName: "三叉神経痛_slides.pptx" }).then(function() {
  console.log("Created: 三叉神経痛_slides.pptx (" + TOTAL + " slides)");
});
