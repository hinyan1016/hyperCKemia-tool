const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "医知創造ラボ";
pres.title = "レカネマブとAfの管理";

// ── Design constants ──
const C = {
  navy:    "1B3A5C",
  blue:    "2C5AA0",
  ltBlue:  "E8F0FA",
  bg:      "F0F4F8",
  white:   "FFFFFF",
  text:    "333333",
  red:     "DC3545",
  yellow:  "F5C518",
  green:   "28A745",
  orange:  "E8850C",
  ltGray:  "F8F9FA",
  gray:    "999999",
};

const FONT_JP = "游ゴシック";
const FONT_EN = "Segoe UI";

// Helper: fresh shadow
const mkShadow = () => ({
  type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12
});

// Helper: add footer bar to content slides
function addFooter(slide) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.15, w: 10, h: 0.475,
    fill: { color: C.navy }
  });
  slide.addText("レカネマブ × 心房細動", {
    x: 0.5, y: 5.15, w: 6, h: 0.475,
    fontSize: 18, fontFace: FONT_JP, color: C.white,
    valign: "middle", margin: 0
  });
}

// Helper: add section title text
function addSectionTitle(slide, title, opts = {}) {
  slide.addText(title, {
    x: 0.6, y: 0.35, w: 8.8, h: 0.65,
    fontSize: opts.fontSize || 28, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });
  // Accent line under title
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.05, w: 1.2, h: 0.06,
    fill: { color: C.blue }
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 1: Title
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  // Decorative accent bar at top
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.12,
    fill: { color: C.blue }
  });

  // Large decorative circle (subtle)
  s.addShape(pres.shapes.OVAL, {
    x: 7.0, y: -1.0, w: 5, h: 5,
    fill: { color: C.blue, transparency: 80 }
  });

  s.addText("レカネマブ投与を検討していた\n高齢者で心房細動（Af）が\n見つかったらどうするか", {
    x: 0.8, y: 1.0, w: 8.4, h: 2.8,
    fontSize: 32, fontFace: FONT_JP,
    color: C.white, bold: true,
    lineSpacingMultiple: 1.3, margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 3.9, w: 2.0, h: 0.06,
    fill: { color: C.yellow }
  });

  s.addText("― 抗凝固療法との両立を考える ―", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP,
    color: "AABBDD", margin: 0
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 2: 結論
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "まず結論");

  // Main conclusion card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 8.8, h: 2.2,
    fill: { color: C.white }, shadow: mkShadow()
  });
  // Left accent bar on card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 0.1, h: 2.2,
    fill: { color: C.red }
  });

  s.addText([
    { text: "長期抗凝固が必要なAfなら\n", options: { fontSize: 26, bold: true, color: C.navy, fontFace: FONT_JP, breakLine: true } },
    { text: "レカネマブは見送り・中止が妥当", options: { fontSize: 26, bold: true, color: C.red, fontFace: FONT_JP } }
  ], {
    x: 1.0, y: 1.4, w: 8.0, h: 2.0,
    valign: "middle", margin: 0
  });

  // Sub-explanation
  s.addText([
    { text: "Afが見つかったら", options: { bold: true, fontFace: FONT_JP, breakLine: true } },
    { text: "① レカネマブ開始は保留\n", options: { fontFace: FONT_JP, breakLine: true } },
    { text: "② 先に抗凝固療法の要否を決める", options: { fontFace: FONT_JP } }
  ], {
    x: 0.8, y: 3.8, w: 8.4, h: 1.1,
    fontSize: 20, color: C.text, margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 3: 日本 vs 国際比較
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "日本 vs 国際的な位置づけの違い");

  // Japan card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.35, w: 4.1, h: 3.3,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.35, w: 4.1, h: 0.6,
    fill: { color: C.blue }
  });
  s.addText("日本の添付文書", {
    x: 0.6, y: 1.35, w: 4.1, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });
  s.addText([
    { text: "抗凝固薬は\n", options: { fontSize: 20, fontFace: FONT_JP, breakLine: true } },
    { text: "「併用注意」", options: { fontSize: 26, fontFace: FONT_JP, bold: true, color: C.orange } },
    { text: "\n\n（＝併用自体は禁止されていない）", options: { fontSize: 18, fontFace: FONT_JP, color: C.gray, breakLine: true } }
  ], {
    x: 0.8, y: 2.1, w: 3.7, h: 2.3,
    color: C.text, align: "center", valign: "middle", margin: 0
  });

  // International card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.35, w: 4.1, h: 3.3,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.35, w: 4.1, h: 0.6,
    fill: { color: C.red }
  });
  s.addText("米国AUR / フランス / EMA", {
    x: 5.3, y: 1.35, w: 4.1, h: 0.6,
    fontSize: 18, fontFace: FONT_EN, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });
  s.addText([
    { text: "抗凝固中は\n", options: { fontSize: 20, fontFace: FONT_JP, breakLine: true } },
    { text: "開始しない\n", options: { fontSize: 26, fontFace: FONT_JP, bold: true, color: C.red, breakLine: true } },
    { text: "長期抗凝固が必要なら\n", options: { fontSize: 20, fontFace: FONT_JP, breakLine: true } },
    { text: "中止を推奨", options: { fontSize: 26, fontFace: FONT_JP, bold: true, color: C.red } }
  ], {
    x: 5.5, y: 2.1, w: 3.7, h: 2.3,
    color: C.text, align: "center", valign: "middle", margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 4: MRI禁忌条件
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "出血リスク ― 投与前MRI禁忌条件");

  const tableRows = [
    [
      { text: "項目", options: { fill: { color: C.navy }, color: C.white, bold: true, fontFace: FONT_JP, fontSize: 18 } },
      { text: "添付文書の扱い", options: { fill: { color: C.navy }, color: C.white, bold: true, fontFace: FONT_JP, fontSize: 18 } }
    ],
    [
      { text: "脳微小出血 5個以上", options: { fontFace: FONT_JP, fontSize: 18 } },
      { text: "禁忌", options: { fontFace: FONT_JP, fontSize: 18, bold: true, color: C.red } }
    ],
    [
      { text: "脳表ヘモジデリン沈着", options: { fontFace: FONT_JP, fontSize: 18 } },
      { text: "禁忌", options: { fontFace: FONT_JP, fontSize: 18, bold: true, color: C.red } }
    ],
    [
      { text: "1cmを超える脳出血", options: { fontFace: FONT_JP, fontSize: 18 } },
      { text: "禁忌", options: { fontFace: FONT_JP, fontSize: 18, bold: true, color: C.red } }
    ],
    [
      { text: "ApoE ε4 ホモ接合体", options: { fontFace: FONT_JP, fontSize: 18 } },
      { text: "ARIA頻度・重症度が高い\n（投与開始後14週間は特に注意）", options: { fontFace: FONT_JP, fontSize: 18, color: C.orange } }
    ]
  ];

  s.addTable(tableRows, {
    x: 0.6, y: 1.35, w: 8.8,
    colW: [4.4, 4.4],
    border: { pt: 0.5, color: "DEE2E6" },
    rowH: [0.55, 0.55, 0.55, 0.55, 0.75]
  });

  s.addText("※ 米国では禁忌ではなく警告・注意", {
    x: 0.6, y: 4.6, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: C.gray, margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 5: Clarity AD 出血データ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "抗凝固薬併用時の出血リスク");

  s.addText("Clarity AD試験データ", {
    x: 0.6, y: 1.35, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FONT_EN, color: C.blue, bold: true, margin: 0
  });

  // Stat card 1
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.0, w: 4.1, h: 2.3,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.0, w: 0.1, h: 2.3,
    fill: { color: C.blue }
  });
  s.addText("レカネマブ群全体", {
    x: 1.0, y: 2.1, w: 3.5, h: 0.45,
    fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0
  });
  s.addText("1cm超の脳内出血", {
    x: 1.0, y: 2.5, w: 3.5, h: 0.35,
    fontSize: 18, fontFace: FONT_JP, color: C.gray, margin: 0
  });
  s.addText("0.7%", {
    x: 1.0, y: 2.9, w: 3.5, h: 0.8,
    fontSize: 52, fontFace: FONT_EN, color: C.navy, bold: true, margin: 0
  });
  s.addText("（6/898例）", {
    x: 1.0, y: 3.7, w: 3.5, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: C.gray, margin: 0
  });

  // Stat card 2
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 2.0, w: 4.1, h: 2.3,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 2.0, w: 0.1, h: 2.3,
    fill: { color: C.red }
  });
  s.addText("抗凝固薬 併用例", {
    x: 5.7, y: 2.1, w: 3.5, h: 0.45,
    fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0
  });
  s.addText("1cm超の脳内出血", {
    x: 5.7, y: 2.5, w: 3.5, h: 0.35,
    fontSize: 18, fontFace: FONT_JP, color: C.gray, margin: 0
  });
  s.addText("2.5%", {
    x: 5.7, y: 2.9, w: 3.5, h: 0.8,
    fontSize: 52, fontFace: FONT_EN, color: C.red, bold: true, margin: 0
  });
  s.addText("（2/79例） プラセボ 0例", {
    x: 5.7, y: 3.7, w: 3.5, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: C.gray, margin: 0
  });

  s.addText("抗凝固が加わると危険信号が強くなる", {
    x: 0.6, y: 4.5, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.red, bold: true, margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 6: Af評価項目
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "Afが見つかったら何を見るべきか");

  s.addText("本当に長期管理が必要なAfかを循環器と詰める", {
    x: 0.6, y: 1.35, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0
  });

  const items = [
    "過去の心電図記録",
    "自覚症状の有無",
    "Holter心電図",
    "誘因（発熱・脱水・術後など一過性か）",
    "弁膜症の有無"
  ];

  s.addText(
    items.map((t, i) => ({
      text: t,
      options: { bullet: true, breakLine: i < items.length - 1, fontFace: FONT_JP, fontSize: 20 }
    })),
    {
      x: 0.8, y: 2.0, w: 8.2, h: 2.4,
      color: C.text, paraSpaceAfter: 8, margin: 0
    }
  );

  // Warning box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.2, w: 8.8, h: 0.7,
    fill: { color: "FFF3CD" }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.2, w: 0.08, h: 0.7,
    fill: { color: C.yellow }
  });
  s.addText("高齢だからという理由だけで抗凝固を避けるべきではない（JCS 2024）", {
    x: 0.9, y: 4.2, w: 8.3, h: 0.7,
    fontSize: 18, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 7: Afと認知症リスク
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "Afは認知症リスクそのもの");

  // Big stat
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.4, w: 8.8, h: 2.0,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addText([
    { text: "メタアナリシス：Af患者の認知障害リスク\n", options: { fontSize: 20, fontFace: FONT_JP, color: C.text, breakLine: true } },
    { text: "約40%上昇", options: { fontSize: 44, fontFace: FONT_EN, color: C.red, bold: true, breakLine: true } },
    { text: "（臨床的脳卒中の有無にかかわらず）", options: { fontSize: 18, fontFace: FONT_JP, color: C.gray } }
  ], {
    x: 0.8, y: 1.5, w: 8.4, h: 1.8,
    align: "center", valign: "middle", margin: 0
  });

  // Warning message
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 3.7, w: 8.8, h: 1.2,
    fill: { color: "FFF3CD" }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 3.7, w: 0.08, h: 1.2,
    fill: { color: C.yellow }
  });
  s.addText([
    { text: "「レカネマブを続けたいから抗凝固は後回し」\n", options: { fontSize: 20, fontFace: FONT_JP, bold: true, breakLine: true } },
    { text: "→ この順番は勧めにくい。Afの脳保護を先に考える", options: { fontSize: 20, fontFace: FONT_JP } }
  ], {
    x: 0.9, y: 3.8, w: 8.3, h: 1.0,
    color: C.text, valign: "middle", margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 8: 3つのシナリオ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "実際の分岐 ― 3つのシナリオ");

  const scenarios = [
    { label: "A", color: C.green, title: "長期抗凝固が不要", action: "レカネマブ継続・開始を再検討する余地あり", y: 1.4 },
    { label: "B", color: C.orange, title: "一時的な抗凝固が必要", action: "レカネマブは一時中断が妥当", y: 2.65 },
    { label: "C", color: C.red, title: "長期抗凝固が必要", action: "レカネマブは見送り・中止寄り", y: 3.9 },
  ];

  scenarios.forEach(sc => {
    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: sc.y, w: 8.8, h: 1.05,
      fill: { color: C.white }, shadow: mkShadow()
    });
    // Color accent
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: sc.y, w: 0.1, h: 1.05,
      fill: { color: sc.color }
    });
    // Label circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.9, y: sc.y + 0.15, w: 0.75, h: 0.75,
      fill: { color: sc.color }
    });
    s.addText(sc.label, {
      x: 0.9, y: sc.y + 0.15, w: 0.75, h: 0.75,
      fontSize: 28, fontFace: FONT_EN, color: C.white,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    // Title
    s.addText(sc.title, {
      x: 1.9, y: sc.y + 0.05, w: 4.0, h: 0.5,
      fontSize: 20, fontFace: FONT_JP, color: C.navy, bold: true, margin: 0
    });
    // Action
    s.addText(sc.action, {
      x: 1.9, y: sc.y + 0.5, w: 7.3, h: 0.45,
      fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0
    });
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 9: 抗血小板薬
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "抗血小板薬はどうか");

  // Card 1
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.4, w: 4.1, h: 2.5,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.4, w: 4.1, h: 0.55,
    fill: { color: C.green }
  });
  s.addText("単剤抗血小板薬", {
    x: 0.6, y: 1.4, w: 4.1, h: 0.55,
    fontSize: 20, fontFace: FONT_JP, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });
  s.addText("一律除外ではない\n\n冠動脈疾患などで\n単剤使用中の場面は\n抗凝固とは同じ重みで扱わない", {
    x: 0.8, y: 2.1, w: 3.7, h: 1.6,
    fontSize: 18, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0
  });

  // Card 2
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.4, w: 4.1, h: 2.5,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.4, w: 4.1, h: 0.55,
    fill: { color: C.orange }
  });
  s.addText("DAPT（二剤併用）", {
    x: 5.3, y: 1.4, w: 4.1, h: 0.55,
    fontSize: 20, fontFace: FONT_EN, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });
  s.addText("個別に慎重判断\n\n米国AURでは\n慎重な個別検討を\n推奨", {
    x: 5.5, y: 2.1, w: 3.7, h: 1.6,
    fontSize: 18, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0
  });

  s.addText("抗凝固と抗血小板は同じ重みでは扱わない", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.6,
    fontSize: 22, fontFace: FONT_JP, color: C.navy, bold: true, margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 10: 高齢者での注意点
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "高齢者での注意点");

  s.addText("Clarity AD試験の対象は50〜90歳。年齢内でも以下の因子で判断が変わる", {
    x: 0.6, y: 1.35, w: 8.8, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0
  });

  const factors = [
    { text: "フレイル", col: C.red },
    { text: "腎機能低下", col: C.red },
    { text: "低体重", col: C.orange },
    { text: "脳微小出血の数", col: C.red },
    { text: "CAA疑い", col: C.red },
    { text: "血圧管理状況", col: C.orange },
  ];

  factors.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.6 + col * 4.6;
    const y = 2.0 + row * 0.9;

    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 4.2, h: 0.7,
      fill: { color: C.white }, shadow: mkShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x, y: y, w: 0.08, h: 0.7,
      fill: { color: f.col }
    });
    s.addText(f.text, {
      x: x + 0.25, y: y, w: 3.8, h: 0.7,
      fontSize: 20, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
    });
  });

  s.addText("高齢者でAfが加わると安全性側の重みが一気に増す", {
    x: 0.6, y: 4.5, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.red, bold: true, margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 11: アブレーション・左心耳閉鎖
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "アブレーション・左心耳閉鎖で\n道は開けるか");

  // Red conclusion box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.4, w: 8.8, h: 0.6,
    fill: { color: "F8D7DA" }
  });
  s.addText("結論：期待しすぎない方がよい", {
    x: 0.8, y: 1.4, w: 8.4, h: 0.6,
    fontSize: 22, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: 0
  });

  // Ablation card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.2, w: 4.1, h: 2.3,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.2, w: 4.1, h: 0.5,
    fill: { color: C.navy }
  });
  s.addText("Afアブレーション", {
    x: 0.6, y: 2.2, w: 4.1, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });
  s.addText("ESC 2024：アブレーション後も\n抗凝固はリズム結果ではなく\n血栓塞栓リスクに応じて継続\n\nレカネマブのための\n近道としては不適切", {
    x: 0.8, y: 2.8, w: 3.7, h: 1.6,
    fontSize: 18, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0
  });

  // LAA closure card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 2.2, w: 4.1, h: 2.3,
    fill: { color: C.white }, shadow: mkShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 2.2, w: 4.1, h: 0.5,
    fill: { color: C.navy }
  });
  s.addText("左心耳閉鎖", {
    x: 5.3, y: 2.2, w: 4.1, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.white,
    bold: true, align: "center", valign: "middle", margin: 0
  });
  s.addText("HAS規定：routineには勧めず\n長期抗凝固禁忌がある患者に\n限定\n\n多職種カンファでの\n個別判断が必須", {
    x: 5.5, y: 2.8, w: 3.7, h: 1.6,
    fontSize: 18, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 12: 実地フロー
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  addSectionTitle(s, "実地でのフロー");

  // Key message
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.35, w: 8.8, h: 0.6,
    fill: { color: C.ltBlue }
  });
  s.addText("Afそのものが即レッドカードではない。長期抗凝固が必要になるAfがレッドフラッグ。", {
    x: 0.8, y: 1.35, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle", margin: 0
  });

  const steps = [
    { num: "1", text: "循環器で抗凝固の要否を確定\n（CHADS-VAScスコア、弁膜症、Holterなど）" },
    { num: "2", text: "脳MRIを再点検\n（微小出血の数、ヘモジデリン沈着、脳出血の有無）" },
    { num: "3", text: "リスク因子を重ねて評価\n（ApoE ε4、高血圧、脳卒中/TIA/痙攣歴）" },
    { num: "4", text: "患者・家族と共有意思決定\n（レカネマブを続けるか止めるか）" },
  ];

  steps.forEach((st, i) => {
    const y = 2.15 + i * 0.75;
    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.8, y: y + 0.05, w: 0.5, h: 0.5,
      fill: { color: C.blue }
    });
    s.addText(st.num, {
      x: 0.8, y: y + 0.05, w: 0.5, h: 0.5,
      fontSize: 20, fontFace: FONT_EN, color: C.white,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    // Text
    s.addText(st.text, {
      x: 1.5, y: y, w: 8.0, h: 0.65,
      fontSize: 18, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0
    });
  });

  addFooter(s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 13: Take-home
// ━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.12,
    fill: { color: C.blue }
  });

  s.addText("Take-home Message", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.7,
    fontSize: 32, fontFace: FONT_EN, color: C.white, bold: true, margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.3, w: 2.0, h: 0.06,
    fill: { color: C.yellow }
  });

  const msgs = [
    "Afが見つかったら、まず脳梗塞予防を決める",
    "長期抗凝固が必要なら、レカネマブは中止・見送りが基本",
    "日本の添付文書は「併用注意」だが、国際提言はより厳しい",
    "現場では安全側に倒して考える",
  ];

  s.addText(
    msgs.map((m, i) => ({
      text: m,
      options: { bullet: true, breakLine: i < msgs.length - 1, fontFace: FONT_JP, fontSize: 22 }
    })),
    {
      x: 0.8, y: 1.6, w: 8.4, h: 3.0,
      color: C.white, paraSpaceAfter: 14, margin: 0
    }
  );

  s.addText("レカネマブ × 心房細動", {
    x: 0.8, y: 4.8, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: "7799BB", margin: 0
  });
}

// ── Write file ──
const outPath = "C:/Users/jsber/OneDrive/Documents/Claude_task/lecanemab_af_slides.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("Created: " + outPath);
}).catch(err => {
  console.error("Error:", err);
});
