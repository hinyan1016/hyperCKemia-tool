var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "脳が「痛い」と感じる仕組み ― 脳自体は痛みを感じない？【からだの不思議 #10】";

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

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
}

function addCard(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: C.white }, shadow: makeShadow() });
}

function addPageNum(s, num) {
  s.addText(num, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("脳神経内科医が答える からだの不思議 #10", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("脳が「痛い」と感じる仕組み", {
    x: 0.8, y: 1.2, w: 8.4, h: 0.9,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addText("― 脳自体は痛みを感じないのに\n　なぜ頭が痛くなるのか ―", {
    x: 0.8, y: 2.15, w: 8.4, h: 1.1,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0, lineSpacingMultiple: 1.25,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.45, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 3.7, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.25, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 脳自体は痛みを感じない！？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳自体は痛みを感じない！？");

  // インパクトメッセージ
  addCard(s, 0.5, 1.1, 9.0, 1.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 0.15, h: 1.4, fill: { color: C.red } });
  s.addText("手術中に患者さんが目を覚ましたまま、\n脳に触れられても「痛くない」のは本当？", {
    x: 0.9, y: 1.2, w: 8.4, h: 1.1,
    fontSize: 18, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
  });

  // 事実カード
  addCard(s, 0.5, 2.75, 9.0, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.75, w: 9.0, h: 0.5, fill: { color: C.red } });
  s.addText("事実", { x: 0.8, y: 2.75, w: 8.4, h: 0.5, fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "脳の実質（灰白質・白質）には、痛みを感じるセンサーが", options: { color: C.text, breakLine: false } },
    { text: "ほとんど存在しない", options: { bold: true, color: C.red, breakLine: true } },
    { text: "→ 外科医がメスで脳組織に直接触れても、患者さんは痛みを訴えない", options: { color: C.primary, bold: true } },
  ], { x: 0.8, y: 3.3, w: 8.4, h: 0.85, fontSize: 15, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  // 疑問提示
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.45, w: 9.0, h: 0.65, fill: { color: C.light } });
  s.addText("では……「頭が痛い」の正体は何か？", {
    x: 0.7, y: 4.45, w: 8.6, h: 0.65,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: 侵害受容器とは
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "痛みのセンサー「侵害受容器」とは");

  // 定義カード
  addCard(s, 0.5, 1.1, 9.0, 1.3);
  s.addText("Nociceptor（侵害受容器）", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.45,
    fontSize: 20, fontFace: FONT_EN, color: C.primary, bold: true, italic: true, margin: 0,
  });
  s.addText("組織の損傷や強い刺激（熱・機械的刺激・炎症性物質）に反応する特殊な神経終末", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.7,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 分布カード (ある / ない)
  var cols = [
    { title: "侵害受容器がある組織", color: C.green, items: ["皮膚・皮下組織", "筋肉・関節・腱", "血管壁", "内臓", "硬膜（脳を包む膜）", "骨膜"] },
    { title: "侵害受容器がない（乏しい）組織", color: C.red, items: ["脳の実質（灰白質）", "脳の実質（白質）", "脊髄の実質"] },
  ];

  cols.forEach(function(col, i) {
    var xPos = 0.3 + i * 5.0;
    addCard(s, xPos, 2.6, 4.6, 2.8);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.6, w: 4.6, h: 0.5, fill: { color: col.color } });
    s.addText(col.title, { x: xPos, y: 2.6, w: 4.6, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    col.items.forEach(function(item, j) {
      s.addText("✓  " + item, {
        x: xPos + 0.2, y: 3.2 + j * 0.41, w: 4.1, h: 0.38,
        fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
      });
    });
  });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: 頭痛の発生源 — 硬膜・血管・筋肉・脳神経
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "頭痛の本当の発生源");

  var rows = [
    { source: "硬膜\n(dura mater)", mechanism: "髄膜炎・くも膜下出血・頭蓋内圧亢進による牽引・炎症\n三叉神経（V1）が支配", headache: "髄膜炎頭痛・頭蓋内圧亢進頭痛" },
    { source: "頭蓋内動脈", mechanism: "三叉神経血管反射でCGRPなどの神経ペプチドが\n血管周囲に放出され無菌性炎症が生じる", headache: "片頭痛" },
    { source: "頭部筋肉\n・頭蓋外筋腱", mechanism: "側頭筋・後頸部筋の過緊張→筋膜の侵害受容器が活性化", headache: "緊張型頭痛" },
    { source: "三叉神経\n・後頭神経", mechanism: "神経そのものへの圧迫・脱髄", headache: "三叉神経痛・後頭神経痛" },
  ];

  // テーブルヘッダ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 1.1, w: 2.3, h: 0.5, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 1.1, w: 5.0, h: 0.5, fill: { color: C.primary } });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.5, y: 1.1, w: 2.3, h: 0.5, fill: { color: C.accent } });
  s.addText("発生源", { x: 0.2, y: 1.1, w: 2.3, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("メカニズム", { x: 2.5, y: 1.1, w: 5.0, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("関連頭痛", { x: 7.5, y: 1.1, w: 2.3, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.65 + i * 0.96;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: yPos, w: 9.6, h: 0.96, fill: { color: bgColor } });
    s.addText(r.source, { x: 0.2, y: yPos, w: 2.3, h: 0.96, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2, lineSpacingMultiple: 1.1 });
    s.addText(r.mechanism, { x: 2.5, y: yPos, w: 5.0, h: 0.96, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.15 });
    s.addText(r.headache, { x: 7.5, y: yPos, w: 2.3, h: 0.96, fontSize: 12, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 2, lineSpacingMultiple: 1.1 });
  });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: 覚醒下脳手術
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "覚醒下脳手術（Awake Craniotomy）");

  // 定義カード
  addCard(s, 0.5, 1.1, 9.0, 1.3);
  s.addText("Awake Craniotomy（覚醒下脳手術）", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_EN, color: C.primary, bold: true, italic: true, margin: 0,
  });
  s.addText("患者さんが意識のある状態で脳に電気刺激を加え、言語野・運動野の位置を確認する手術", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.7,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 手術の流れカード
  var steps = [
    { num: "1", label: "頭皮切開\n・頭蓋骨削開", note: "局所麻酔で行う（痛い）" },
    { num: "2", label: "硬膜切開", note: "局所麻酔で行う（痛い）" },
    { num: "3", label: "脳実質への\n電気刺激", note: "脳実質は痛みを感じない\n→患者と会話しながら手術" },
  ];

  steps.forEach(function(step, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 2.65, 3.0, 2.2);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.9, y: 2.75, w: 1.2, h: 1.2, fill: { color: (i < 2) ? C.orange : C.green } });
    s.addText(step.num, { x: xPos + 0.9, y: 2.75, w: 1.2, h: 1.2, fontSize: 32, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(step.label, { x: xPos, y: 4.0, w: 3.0, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
    s.addText(step.note, { x: xPos, y: 4.5, w: 3.0, h: 0.6, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.1 });
  });

  // 結論
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.0, w: 9.0, h: 0.55, fill: { color: C.light } });
  s.addText("「脳実質に痛みセンサーがない」ことが、覚醒下手術を可能にする根拠", {
    x: 0.7, y: 5.0, w: 8.6, h: 0.55,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: ペインマトリクス（脳の痛み処理ネットワーク）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳が「痛みを処理する」しくみ ― ペインマトリクス");

  // 定義
  addCard(s, 0.5, 1.1, 9.0, 0.85);
  s.addText([
    { text: "ペインマトリクス（pain matrix）", options: { bold: true, color: C.primary } },
    { text: " ― 痛みを処理する脳内ネットワーク。身体のあらゆる痛みは最終的にここで認識される", options: { color: C.text } },
  ], { x: 0.8, y: 1.15, w: 8.4, h: 0.75, fontSize: 15, fontFace: FONT_JP, valign: "middle", margin: 0 });

  // 4領域カード
  var areas = [
    { name: "一次・二次\n体性感覚野\nS1 / S2", role: "痛みの「場所」「強さ」を識別する", color: C.primary },
    { name: "前帯状皮質\nACC", role: "痛みの不快感・情動的側面を処理する", color: C.orange },
    { name: "島皮質\n(insula)", role: "内臓感覚・自律神経反応と統合する", color: C.accent },
    { name: "視床\n(thalamus)", role: "末梢からの痛み信号を大脳皮質へ中継する「中継局」", color: C.green },
  ];

  areas.forEach(function(area, i) {
    var xPos = 0.3 + (i % 2) * 4.8;
    var yPos = 2.1 + Math.floor(i / 2) * 1.65;
    addCard(s, xPos, yPos, 4.4, 1.45);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 4.4, h: 0.5, fill: { color: area.color } });
    s.addText(area.name, { x: xPos, y: yPos, w: 4.4, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.0 });
    s.addText(area.role, { x: xPos + 0.1, y: yPos + 0.55, w: 4.2, h: 0.85, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: 中枢性感作（中枢性過敏化）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "中枢性感作 ― 慢性頭痛のしくみ");

  // 定義カード
  addCard(s, 0.5, 1.1, 9.0, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 0.15, h: 1.2, fill: { color: C.red } });
  s.addText("中枢性感作（Central Sensitization）", {
    x: 0.9, y: 1.15, w: 8.3, h: 0.45,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("ペインマトリクスの過活動・感作により、実際の組織損傷がなくても痛みが継続する状態", {
    x: 0.9, y: 1.6, w: 8.3, h: 0.65,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 慢性化の流れ
  var flow = ["片頭痛を繰り返す", "三叉神経系が過敏化", "頭痛がない日も神経が過活動", "慢性片頭痛へ移行\n（月15日以上の頭痛）"];
  flow.forEach(function(f, i) {
    var xPos = 0.3 + i * 2.3;
    addCard(s, xPos, 2.5, 2.1, 1.1);
    s.addText(f, {
      x: xPos, y: 2.5, w: 2.1, h: 1.1,
      fontSize: 14, fontFace: FONT_JP, color: (i === 3) ? C.red : C.text, bold: (i === 3), align: "center", valign: "middle", margin: 4, lineSpacingMultiple: 1.15,
    });
    if (i < 3) {
      s.addText("→", {
        x: xPos + 2.1, y: 2.8, w: 0.2, h: 0.5,
        fontSize: 22, color: C.accent, align: "center", valign: "middle", margin: 0,
      });
    }
  });

  // 関連疾患
  var diseases = [
    { name: "慢性片頭痛", desc: "月15日以上の頭痛" },
    { name: "線維筋痛症", desc: "全身の慢性疼痛" },
    { name: "薬物乱用頭痛", desc: "鎮痛薬の過剰使用が悪化を招く" },
  ];

  s.addText("中枢性感作が関与する主な病態", {
    x: 0.5, y: 3.85, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, bold: true, margin: 0,
  });

  diseases.forEach(function(d, i) {
    var xPos = 0.5 + i * 3.1;
    addCard(s, xPos, 4.3, 2.9, 1.1);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.3, w: 2.9, h: 0.45, fill: { color: C.orange } });
    s.addText(d.name, { x: xPos, y: 4.3, w: 2.9, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(d.desc, { x: xPos, y: 4.8, w: 2.9, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0 });
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: 危険な頭痛のレッドフラグ（SNOOP）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんな頭痛はすぐに救急受診を ― レッドフラグ");

  var flags = [
    { flag: "突然始まる激しい頭痛\n「人生最悪の頭痛」「雷鳴頭痛」", disease: "くも膜下出血を疑う" },
    { flag: "発熱・項部硬直・意識障害を伴う", disease: "髄膜炎・脳炎を疑う" },
    { flag: "神経症状を伴う\n（片麻痺・言語障害・複視・歩行困難）", disease: "脳卒中・脳腫瘍を疑う" },
    { flag: "50歳以降の初めての頭痛", disease: "側頭動脈炎・脳腫瘍を疑う" },
    { flag: "体位で大きく変化する頭痛\n（起立で悪化、または改善）", disease: "低髄液圧症候群・頭蓋内圧亢進" },
    { flag: "免疫不全・がん治療中の頭痛", disease: "日和見感染・転移を疑う" },
  ];

  flags.forEach(function(f, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var xPos = 0.2 + col * 4.9;
    var yPos = 1.15 + row * 1.38;
    addCard(s, xPos, yPos, 4.6, 1.2);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 0.5, h: 1.2, fill: { color: C.red } });
    s.addText("!", { x: xPos, y: yPos, w: 0.5, h: 1.2, fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.flag, {
      x: xPos + 0.6, y: yPos + 0.05, w: 3.85, h: 0.7,
      fontSize: 13, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0, lineSpacingMultiple: 1.1,
    });
    s.addText("→ " + f.disease, {
      x: xPos + 0.6, y: yPos + 0.75, w: 3.85, h: 0.38,
      fontSize: 12, fontFace: FONT_JP, color: C.red, valign: "middle", margin: 0,
    });
  });

  // 注記
  s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 5.1, w: 9.6, h: 0.43, fill: { color: C.light } });
  s.addText("頭痛の約90%は命に関わらない一次性頭痛。しかし残り10%は二次性頭痛（SNOOP基準より）", {
    x: 0.4, y: 5.1, w: 9.2, h: 0.43,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, valign: "middle", margin: 0,
  });

  addPageNum(s, "8/9");
})();

// ============================================================
// SLIDE 9: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.25, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "脳の実質（灰白質・白質）には侵害受容器がほとんどなく、脳自身は痛みを感じない" },
    { num: "2", text: "頭痛の発生源は「硬膜・頭蓋内血管・頭部筋肉・脳神経」など脳の周辺組織" },
    { num: "3", text: "片頭痛は三叉神経血管反射とCGRPが重要な役割を担う" },
    { num: "4", text: "脳は痛みの「発生源」ではなく、痛みを処理・認識する「司令塔」（ペインマトリクス）" },
    { num: "5", text: "突然の激烈な頭痛・神経症状を伴う頭痛はすぐに救急受診を" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.05 + i * 0.88;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.75, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.1, w: 0.55, h: 0.55, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.1, w: 0.55, h: 0.55, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.75, y: yPos, w: 7.25, h: 0.75, fontSize: 14, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.4, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.45, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "9/9");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/脳と痛み_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
