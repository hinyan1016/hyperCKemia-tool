const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10" x 5.625"
pres.author = "医知創造ラボ";
pres.title = "重症筋無力症に対するCAR-T療法の新時代 ― Descartes-08";

// ══════════════════════════════════════════════════════════════
// デザインチェックリスト準拠設定
// - 最小フォント: 18pt（表・脚注含む全て）
// - 本文推奨: 24pt以上
// - タイトル: 32-44pt
// - 要点: 24-28pt
// - 補足: 18-20pt（必要時のみ）
// - 1スライド = 1メッセージ
// - 文字色: 白 #FFFFFF or ほぼ黒 #111111
// - 強調色: 1色のみ
// - フォント: Meiryo + Segoe UI
// - 行間: 1.3
// - 余白優先
// ══════════════════════════════════════════════════════════════

// ── Colors (high contrast, minimal palette) ────────────────
const C = {
  navy:    "1B3A5C",  // dark background
  blue:    "2C5AA0",  // ONE accent color
  white:   "FFFFFF",
  black:   "111111",  // text on light bg
  lightBg: "F5F7FA",  // subtle light bg
  red:     "DC3545",  // warning only
  green:   "28A745",  // positive only
  muted:   "666666",  // slide number only
  gold:    "FFD700",  // accent on dark bg
};

// ── Fonts ──────────────────────────────────────────────────
const FJ = "Meiryo";
const FE = "Segoe UI";
const TOTAL = 17;

// ── Images ──────────────────────────────────────────────────
const IMG_DIR = "C:/Users/jsber/OneDrive/Documents/Claude_task/";
const IMG = {
  mg_nmj:   IMG_DIR + "img_mg_nmj.png",       // MG神経筋接合部
  cart:     IMG_DIR + "img_cart_therapy.png",   // CAR-T手順図
  bcell:    IMG_DIR + "img_bcell.png",          // B細胞活性化
};

// ── Helpers ────────────────────────────────────────────────
function header(slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.85, fill: { color: C.navy },
  });
  slide.addText(title, {
    x: 0.6, y: 0.08, w: 8.8, h: 0.7,
    fontSize: 28, fontFace: FJ, color: C.white, bold: true, margin: 0,
  });
}

function footer(slide, n) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.45, w: 10, h: 0.18, fill: { color: C.blue },
  });
  if (n) {
    slide.addText(`${n} / ${TOTAL}`, {
      x: 8.5, y: 5.15, w: 1.2, h: 0.3,
      fontSize: 18, fontFace: FE, color: C.muted, align: "right",
    });
  }
}

// ============================================================
// SLIDE 1: TITLE (dark)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.gold },
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.2, w: 0.08, h: 2.2, fill: { color: C.gold },
  });

  s.addText("重症筋無力症に対する\nCAR-T療法の新時代", {
    x: 1.0, y: 1.2, w: 8.4, h: 2.2,
    fontSize: 40, fontFace: FJ, color: C.white, bold: true,
    lineSpacingMultiple: 1.3,
  });

  s.addText("Descartes-08（AURORA試験）解説", {
    x: 1.0, y: 3.5, w: 8.4, h: 0.6,
    fontSize: 24, fontFace: FJ, color: C.gold,
  });

  s.addShape(pres.shapes.LINE, {
    x: 1.0, y: 4.3, w: 2.5, h: 0, line: { color: C.blue, width: 2 },
  });

  s.addText("医知創造ラボ  |  2026年3月", {
    x: 1.0, y: 4.5, w: 8, h: 0.5,
    fontSize: 20, fontFace: FJ, color: C.muted,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.45, w: 10, h: 0.18, fill: { color: C.gold },
  });
}

// ============================================================
// SLIDE 2: MGとは（左テキスト＋右画像）
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "重症筋無力症（MG）とは");

  // Left: text
  s.addText("自己免疫が原因で\n神経筋接合部の\n信号伝達が障害される疾患", {
    x: 0.6, y: 1.1, w: 5.0, h: 1.6,
    fontSize: 26, fontFace: FJ, color: C.black, bold: true,
    lineSpacingMultiple: 1.3,
  });

  s.addText([
    { text: "免疫系がAChRを異物と誤認", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 自己抗体を産生", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 信号が遮断される", options: { fontSize: 22, fontFace: FJ, color: C.blue, bold: true } },
  ], { x: 0.6, y: 2.8, w: 5.0, h: 1.8, lineSpacingMultiple: 1.4 });

  // Right: image with caption
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.8, y: 1.0, w: 3.8, h: 3.6,
    fill: { color: C.lightBg },
  });
  s.addImage({
    path: IMG.mg_nmj,
    x: 6.0, y: 1.1, w: 3.4, h: 3.0,
    sizing: { type: "contain", w: 3.4, h: 3.0 },
  });
  s.addText("神経筋接合部の異常", {
    x: 5.8, y: 4.1, w: 3.8, h: 0.5,
    fontSize: 18, fontFace: FJ, color: C.muted, align: "center",
  });

  footer(s, 2);
}

// ============================================================
// SLIDE 3: MGの主な症状
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "MGの主な症状");

  const symptoms = [
    "眼瞼下垂（まぶたが下がる）",
    "複視（ものが二重に見える）",
    "四肢筋力低下・嚥下障害",
  ];

  symptoms.forEach((sym, i) => {
    s.addText(sym, {
      x: 0.6, y: 1.2 + i * 1.0, w: 8.8, h: 0.8,
      fontSize: 26, fontFace: FJ, color: C.black,
      bullet: true, valign: "middle",
      paraSpaceAfter: 6,
    });
  });

  // Warning
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.2, w: 8.8, h: 0.9,
    fill: { color: "FFF0F0" },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.2, w: 0.08, h: 0.9,
    fill: { color: C.red },
  });
  s.addText("重症例では呼吸筋麻痺（クリーゼ）に至る", {
    x: 0.9, y: 4.2, w: 8.3, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.red, bold: true, valign: "middle",
  });

  footer(s, 3);
}

// ============================================================
// SLIDE 4: 従来の治療法
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "従来の治療法");

  const rows = [
    [
      { text: "治療法", options: { bold: true, color: "FFFFFF", fill: { color: C.navy }, fontSize: 20, fontFace: FJ } },
      { text: "特徴", options: { bold: true, color: "FFFFFF", fill: { color: C.navy }, fontSize: 20, fontFace: FJ } },
    ],
    [
      { text: "抗コリンエステラーゼ薬", options: { bold: true, fontSize: 20, fontFace: FJ } },
      { text: "対症療法（症状改善のみ）", options: { fontSize: 20, fontFace: FJ } },
    ],
    [
      { text: "ステロイド / 免疫抑制薬", options: { bold: true, fontSize: 20, fontFace: FJ } },
      { text: "長期副作用・効果発現に時間", options: { fontSize: 20, fontFace: FJ } },
    ],
    [
      { text: "IVIg / 血漿交換", options: { bold: true, fontSize: 20, fontFace: FJ } },
      { text: "効果は一時的", options: { fontSize: 20, fontFace: FJ } },
    ],
    [
      { text: "FcRn / 補体阻害薬", options: { bold: true, fontSize: 20, fontFace: FJ } },
      { text: "新規薬剤（高コスト）", options: { fontSize: 20, fontFace: FJ } },
    ],
  ];

  s.addTable(rows, {
    x: 0.6, y: 1.1, w: 8.8,
    colW: [4.4, 4.4],
    border: { pt: 1, color: "DEE2E6" },
    rowH: [0.6, 0.65, 0.65, 0.65, 0.65],
    autoPage: false,
  });

  footer(s, 4);
}

// ============================================================
// SLIDE 5: 従来治療の限界
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "従来治療の根本的な問題");

  s.addText("原因となるB細胞を除去できない", {
    x: 0.6, y: 1.5, w: 8.8, h: 1.2,
    fontSize: 36, fontFace: FJ, color: C.blue, bold: true,
    valign: "middle",
  });

  s.addText([
    { text: "既存治療は「免疫を弱める」「抗体を一時的に減らす」", options: { fontSize: 24, fontFace: FJ, color: C.black, breakLine: true } },
    { text: " ", options: { fontSize: 18, breakLine: true } },
    { text: "→ 治療中止で再燃リスク", options: { fontSize: 24, fontFace: FJ, color: C.red, bold: true, breakLine: true } },
    { text: "→ 長期的な免疫抑制療法が必要", options: { fontSize: 24, fontFace: FJ, color: C.red, bold: true } },
  ], { x: 0.6, y: 2.9, w: 8.8, h: 2.2, lineSpacingMultiple: 1.3 });

  footer(s, 5);
}

// ============================================================
// SLIDE 6: CAR-T療法とは（左テキスト＋右画像）
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "CAR-T細胞療法とは");

  // Left: text
  s.addText("患者自身のT細胞を\n遺伝子改変し標的細胞を\n攻撃させる治療法", {
    x: 0.6, y: 1.1, w: 5.0, h: 1.6,
    fontSize: 26, fontFace: FJ, color: C.black, bold: true,
    lineSpacingMultiple: 1.3,
  });

  s.addText([
    { text: "CARを装着したT細胞が", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "標的を効率的に攻撃", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: " ", options: { fontSize: 18, breakLine: true } },
    { text: "血液がんで承認済み", options: { fontSize: 22, fontFace: FJ, color: C.blue, bold: true, breakLine: true } },
    { text: "→ 自己免疫疾患へ応用が進展", options: { fontSize: 22, fontFace: FJ, color: C.blue, bold: true } },
  ], { x: 0.6, y: 2.8, w: 5.0, h: 2.0, lineSpacingMultiple: 1.3 });

  // Right: image with caption
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.8, y: 1.0, w: 3.8, h: 3.6,
    fill: { color: C.lightBg },
  });
  s.addImage({
    path: IMG.cart,
    x: 6.0, y: 1.1, w: 3.4, h: 3.0,
    sizing: { type: "contain", w: 3.4, h: 3.0 },
  });
  s.addText("CAR-T細胞療法の概要", {
    x: 5.8, y: 4.1, w: 3.8, h: 0.5,
    fontSize: 18, fontFace: FJ, color: C.muted, align: "center",
  });

  footer(s, 6);
}

// ============================================================
// SLIDE 7: CAR-Tの手順
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "CAR-T療法の5ステップ");

  const steps = [
    { num: "1", label: "T細胞採取" },
    { num: "2", label: "遺伝子改変" },
    { num: "3", label: "体外増殖" },
    { num: "4", label: "患者に再投与" },
    { num: "5", label: "標的細胞を破壊" },
  ];
  const stepW = 1.6, stepH = 2.8, startX = 0.5, gap = 0.25, startY = 1.2;

  steps.forEach((st, i) => {
    const x = startX + i * (stepW + gap);

    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: x + stepW / 2 - 0.3, y: startY, w: 0.6, h: 0.6,
      fill: { color: C.blue },
    });
    s.addText(st.num, {
      x: x + stepW / 2 - 0.3, y: startY, w: 0.6, h: 0.6,
      fontSize: 24, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: startY + 0.8, w: stepW, h: 1.6,
      fill: { color: C.lightBg },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: startY + 0.8, w: stepW, h: 0.06,
      fill: { color: C.blue },
    });
    s.addText(st.label, {
      x: x + 0.1, y: startY + 0.9, w: stepW - 0.2, h: 1.4,
      fontSize: 20, fontFace: FJ, color: C.black, bold: true,
      align: "center", valign: "middle",
    });

    // Arrow
    if (i < steps.length - 1) {
      s.addText("→", {
        x: x + stepW, y: startY + 0.8, w: gap, h: 1.6,
        fontSize: 24, fontFace: FE, color: C.blue, bold: true,
        align: "center", valign: "middle",
      });
    }
  });

  footer(s, 7);
}

// ============================================================
// SLIDE 8: なぜMGにCAR-Tを使うのか（左テキスト＋右画像）
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "なぜMGにCAR-Tを使うのか");

  // Left: text
  s.addText("「免疫を抑える」から\n「原因細胞を除去する」へ", {
    x: 0.6, y: 1.1, w: 5.0, h: 1.0,
    fontSize: 26, fontFace: FJ, color: C.blue, bold: true,
    lineSpacingMultiple: 1.3,
  });

  s.addText([
    { text: "B細胞表面のBCMAを標的に", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 自己反応性B細胞を除去", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 自己抗体の産生が停止", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 症状改善", options: { fontSize: 24, fontFace: FJ, color: C.blue, bold: true } },
  ], { x: 0.6, y: 2.3, w: 5.0, h: 2.0, lineSpacingMultiple: 1.4 });

  // Bottom note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.5, w: 5.0, h: 0.6,
    fill: { color: C.lightBg },
  });
  s.addText("BCMA = B cell maturation antigen", {
    x: 0.8, y: 4.5, w: 4.6, h: 0.6,
    fontSize: 18, fontFace: FE, color: C.blue, valign: "middle", italic: true,
  });

  // Right: image with caption
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.8, y: 1.0, w: 3.8, h: 3.6,
    fill: { color: C.lightBg },
  });
  s.addImage({
    path: IMG.bcell,
    x: 6.0, y: 1.1, w: 3.4, h: 3.0,
    sizing: { type: "contain", w: 3.4, h: 3.0 },
  });
  s.addText("B細胞の活性化と分化", {
    x: 5.8, y: 4.1, w: 3.8, h: 0.5,
    fontSize: 18, fontFace: FJ, color: C.muted, align: "center",
  });

  footer(s, 8);
}

// ============================================================
// SLIDE 9: Descartes-08とは
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "Descartes-08 ― RNA-CAR-T");

  s.addText("mRNAで一時的にCAR機能を付与する\n「可逆的」CAR-T", {
    x: 0.6, y: 1.2, w: 8.8, h: 1.4,
    fontSize: 28, fontFace: FJ, color: C.black, bold: true,
    lineSpacingMultiple: 1.3,
  });

  s.addText([
    { text: "従来CAR-T：", options: { fontSize: 24, fontFace: FJ, color: C.muted } },
    { text: "DNAをゲノムに永久導入", options: { fontSize: 24, fontFace: FJ, color: C.muted, breakLine: true } },
    { text: " ", options: { fontSize: 18, breakLine: true } },
    { text: "Descartes-08：", options: { fontSize: 24, fontFace: FJ, color: C.blue, bold: true } },
    { text: "mRNAで一時的に機能付与", options: { fontSize: 24, fontFace: FJ, color: C.blue, bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 18, breakLine: true } },
    { text: "→ 自然分解で効果が制御可能", options: { fontSize: 24, fontFace: FJ, color: C.black } },
  ], { x: 0.6, y: 2.8, w: 8.8, h: 2.2, lineSpacingMultiple: 1.3 });

  footer(s, 9);
}

// ============================================================
// SLIDE 10: RNA-CAR-Tの4つの利点
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "RNA-CAR-Tの4つの利点");

  const advantages = [
    { icon: "✓", text: "効果が制御可能（mRNAは自然分解）" },
    { icon: "✓", text: "化学前処置（lymphodepletion）が不要" },
    { icon: "✓", text: "CRS・ICANSの重篤副作用リスクが低い" },
    { icon: "✓", text: "外来で投与可能（入院不要）" },
  ];

  advantages.forEach((adv, i) => {
    const y = 1.15 + i * 0.95;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y, w: 8.8, h: 0.8,
      fill: { color: i % 2 === 0 ? C.lightBg : C.white },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y, w: 0.06, h: 0.8,
      fill: { color: C.blue },
    });
    s.addText(adv.text, {
      x: 0.9, y, w: 8.3, h: 0.8,
      fontSize: 24, fontFace: FJ, color: C.black, bold: true,
      valign: "middle",
    });
  });

  footer(s, 10);
}

// ============================================================
// SLIDE 11: 従来CAR-T vs Descartes-08
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "従来CAR-T vs Descartes-08");

  const rows = [
    [
      { text: "項目", options: { bold: true, color: "FFFFFF", fill: { color: C.navy }, fontSize: 20, fontFace: FJ } },
      { text: "従来CAR-T", options: { bold: true, color: "FFFFFF", fill: { color: "555555" }, fontSize: 20, fontFace: FJ } },
      { text: "Descartes-08", options: { bold: true, color: "FFFFFF", fill: { color: C.blue }, fontSize: 20, fontFace: FJ } },
    ],
    [
      { text: "遺伝子導入", options: { bold: true, fontSize: 20, fontFace: FJ } },
      { text: "DNA（永久）", options: { fontSize: 20, fontFace: FJ } },
      { text: "mRNA（一時的）", options: { fontSize: 20, fontFace: FJ, bold: true, color: C.blue } },
    ],
    [
      { text: "化学前処置", options: { bold: true, fontSize: 20, fontFace: FJ } },
      { text: "必要", options: { fontSize: 20, fontFace: FJ, color: C.red } },
      { text: "不要", options: { fontSize: 20, fontFace: FJ, bold: true, color: C.green } },
    ],
    [
      { text: "入院", options: { bold: true, fontSize: 20, fontFace: FJ } },
      { text: "必要", options: { fontSize: 20, fontFace: FJ } },
      { text: "外来で可能", options: { fontSize: 20, fontFace: FJ, bold: true, color: C.green } },
    ],
    [
      { text: "CRSリスク", options: { bold: true, fontSize: 20, fontFace: FJ } },
      { text: "高い", options: { fontSize: 20, fontFace: FJ, color: C.red } },
      { text: "低い", options: { fontSize: 20, fontFace: FJ, bold: true, color: C.green } },
    ],
  ];

  s.addTable(rows, {
    x: 0.6, y: 1.1, w: 8.8,
    colW: [2.4, 3.2, 3.2],
    border: { pt: 1, color: "DEE2E6" },
    rowH: [0.65, 0.7, 0.7, 0.7, 0.7],
    autoPage: false,
  });

  footer(s, 11);
}

// ============================================================
// SLIDE 12: 第2相試験 ― 有効性
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "第2相試験の結果 ― 有効性");

  s.addText([
    { text: "MG-ADLスコア", options: { fontSize: 28, fontFace: FJ, color: C.blue, bold: true } },
    { text: "の有意な改善", options: { fontSize: 28, fontFace: FJ, color: C.black, bold: true } },
  ], { x: 0.6, y: 1.3, w: 8.8, h: 0.8 });

  s.addText([
    { text: "QMGスコアも改善を確認", options: { fontSize: 24, fontFace: FJ, color: C.black, breakLine: true } },
    { text: " ", options: { fontSize: 18, breakLine: true } },
    { text: "効果は1年以上持続する例もあり", options: { fontSize: 24, fontFace: FJ, color: C.black, breakLine: true } },
    { text: " ", options: { fontSize: 18, breakLine: true } },
    { text: "難治性MG患者を対象に実施", options: { fontSize: 24, fontFace: FJ, color: C.black } },
  ], { x: 0.6, y: 2.3, w: 8.8, h: 2.5, lineSpacingMultiple: 1.3 });

  footer(s, 12);
}

// ============================================================
// SLIDE 13: 第2相試験 ― 安全性
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "第2相試験の結果 ― 安全性");

  // CRS
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.2, w: 8.8, h: 1.3,
    fill: { color: "F0FFF0" },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.2, w: 0.06, h: 1.3,
    fill: { color: C.green },
  });
  s.addText([
    { text: "CRS（サイトカイン放出症候群）", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 軽度またはなし", options: { fontSize: 28, fontFace: FJ, color: C.green, bold: true } },
  ], { x: 0.9, y: 1.3, w: 8.3, h: 1.1, valign: "middle" });

  // ICANS
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.8, w: 8.8, h: 1.3,
    fill: { color: "F0FFF0" },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 2.8, w: 0.06, h: 1.3,
    fill: { color: C.green },
  });
  s.addText([
    { text: "ICANS（神経毒性症候群）", options: { fontSize: 22, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 軽度またはなし", options: { fontSize: 28, fontFace: FJ, color: C.green, bold: true } },
  ], { x: 0.9, y: 2.9, w: 8.3, h: 1.1, valign: "middle" });

  s.addText("RNA-CAR-Tの設計により重篤な副作用リスクが大幅に低減", {
    x: 0.6, y: 4.4, w: 8.8, h: 0.7,
    fontSize: 22, fontFace: FJ, color: C.blue, bold: true, valign: "middle",
  });

  footer(s, 13);
}

// ============================================================
// SLIDE 14: AURORA試験
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "第3相 AURORA試験");

  const items = [
    ["対象", "AChR抗体陽性の全身型MG"],
    ["デザイン", "無作為化二重盲検プラセボ対照"],
    ["患者数", "約100名"],
    ["主要評価項目", "MG-ADLスコアの改善（4か月時点）"],
    ["FDA指定", "Special Protocol Assessment承認済み"],
  ];

  items.forEach((item, i) => {
    const y = 1.1 + i * 0.75;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y, w: 8.8, h: 0.65,
      fill: { color: i % 2 === 0 ? C.lightBg : C.white },
    });
    s.addText(item[0], {
      x: 0.8, y, w: 2.8, h: 0.65,
      fontSize: 20, fontFace: FJ, color: C.blue, bold: true, valign: "middle",
    });
    s.addText(item[1], {
      x: 3.6, y, w: 5.6, h: 0.65,
      fontSize: 22, fontFace: FJ, color: C.black, bold: true, valign: "middle",
    });
  });

  // Bottom note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.85, w: 8.8, h: 0.5,
    fill: { color: C.lightBg },
  });
  s.addText("化学前処置なしCAR-Tの自己免疫疾患における初の大規模RCT", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FJ, color: C.blue, bold: true, valign: "middle",
  });

  footer(s, 14);
}

// ============================================================
// SLIDE 15: パラダイムシフト
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "治療パラダイムの転換");

  // Before
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.2, w: 4.0, h: 2.5,
    fill: { color: C.lightBg },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.2, w: 4.0, h: 0.06,
    fill: { color: C.muted },
  });
  s.addText("従来", {
    x: 0.8, y: 1.35, w: 3.6, h: 0.5,
    fontSize: 22, fontFace: FJ, color: C.muted, bold: true,
  });
  s.addText([
    { text: "免疫抑制を継続", options: { fontSize: 24, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 症状コントロール", options: { fontSize: 22, fontFace: FJ, color: C.muted, breakLine: true } },
    { text: "→ 治療中止で再燃", options: { fontSize: 22, fontFace: FJ, color: C.red } },
  ], { x: 0.8, y: 1.95, w: 3.6, h: 1.6, lineSpacingMultiple: 1.3 });

  // Arrow
  s.addText("→", {
    x: 4.6, y: 1.2, w: 0.8, h: 2.5,
    fontSize: 36, fontFace: FE, color: C.blue, bold: true,
    align: "center", valign: "middle",
  });

  // After
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.4, y: 1.2, w: 4.0, h: 2.5,
    fill: { color: "E8F5E9" },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.4, y: 1.2, w: 4.0, h: 0.06,
    fill: { color: C.green },
  });
  s.addText("CAR-T", {
    x: 5.6, y: 1.35, w: 3.6, h: 0.5,
    fontSize: 22, fontFace: FJ, color: C.green, bold: true,
  });
  s.addText([
    { text: "原因B細胞を除去", options: { fontSize: 24, fontFace: FJ, color: C.black, breakLine: true } },
    { text: "→ 機能的治癒へ", options: { fontSize: 22, fontFace: FJ, color: C.green, bold: true, breakLine: true } },
    { text: "→ 免疫抑制からの解放", options: { fontSize: 22, fontFace: FJ, color: C.green, bold: true } },
  ], { x: 5.6, y: 1.95, w: 3.6, h: 1.6, lineSpacingMultiple: 1.3 });

  // Other diseases
  s.addText("他の自己免疫疾患でも研究が進行中", {
    x: 0.6, y: 4.0, w: 8.8, h: 0.5,
    fontSize: 22, fontFace: FJ, color: C.black, bold: true,
  });
  s.addText("SLE（著効例報告済み）/ 多発性硬化症 / 全身性強皮症 / 炎症性筋疾患", {
    x: 0.6, y: 4.5, w: 8.8, h: 0.6,
    fontSize: 20, fontFace: FJ, color: C.blue,
  });

  footer(s, 15);
}

// ============================================================
// SLIDE 16: 残る課題
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  header(s, "残る課題");

  const challenges = [
    { num: "1", title: "コスト", desc: "がん領域では1回数千万円\nRNA-CAR-Tはコスト削減の可能性" },
    { num: "2", title: "長期安全性", desc: "B細胞除去による感染リスク\n免疫グロブリン低下への注意" },
    { num: "3", title: "再投与", desc: "効果は一時的 → 繰り返し投与の\n安全性・有効性の確認が必要" },
  ];

  challenges.forEach((ch, i) => {
    const y = 1.1 + i * 1.3;
    // Accent bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y, w: 0.06, h: 1.1,
      fill: { color: C.blue },
    });
    // Number
    s.addShape(pres.shapes.OVAL, {
      x: 0.85, y: y + 0.2, w: 0.55, h: 0.55,
      fill: { color: C.blue },
    });
    s.addText(ch.num, {
      x: 0.85, y: y + 0.2, w: 0.55, h: 0.55,
      fontSize: 24, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    // Title
    s.addText(ch.title, {
      x: 1.6, y: y + 0.05, w: 2.0, h: 0.9,
      fontSize: 24, fontFace: FJ, color: C.black, bold: true, valign: "middle",
    });
    // Description
    s.addText(ch.desc, {
      x: 3.6, y: y + 0.05, w: 5.8, h: 0.9,
      fontSize: 20, fontFace: FJ, color: C.black, valign: "middle",
      lineSpacingMultiple: 1.3,
    });
  });

  s.addText("AURORA試験の結果が上記すべての課題に重要データを提供", {
    x: 0.6, y: 4.85, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FJ, color: C.blue, bold: true,
  });

  footer(s, 16);
}

// ============================================================
// SLIDE 17: まとめ (dark)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.navy };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.gold },
  });

  s.addText("まとめ", {
    x: 0.6, y: 0.3, w: 8.8, h: 0.7,
    fontSize: 32, fontFace: FJ, color: C.white, bold: true,
  });

  const points = [
    "MGは自己抗体によるAChR障害が原因",
    "病態の本質は自己反応性B細胞",
    "CAR-TはB細胞を直接攻撃する新アプローチ",
    "Descartes-08はmRNA型「可逆的」CAR-T",
    "第3相AURORA試験が進行中",
  ];

  points.forEach((p, i) => {
    const y = 1.2 + i * 0.7;
    s.addShape(pres.shapes.OVAL, {
      x: 0.8, y: y + 0.12, w: 0.18, h: 0.18,
      fill: { color: C.gold },
    });
    s.addText(p, {
      x: 1.2, y, w: 8.2, h: 0.55,
      fontSize: 24, fontFace: FJ, color: C.white, valign: "middle",
    });
  });

  // Closing
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.7, w: 9.2, h: 0.7,
    fill: { color: "243B5A" },
  });
  s.addText("「免疫を抑える」時代から「原因細胞を除去する」時代へ", {
    x: 0.6, y: 4.7, w: 8.8, h: 0.7,
    fontSize: 22, fontFace: FJ, color: C.gold, bold: true,
    align: "center", valign: "middle",
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.45, w: 10, h: 0.18, fill: { color: C.gold },
  });
  s.addText(`${TOTAL} / ${TOTAL}`, {
    x: 8.5, y: 5.15, w: 1.2, h: 0.3,
    fontSize: 18, fontFace: FE, color: C.muted, align: "right",
  });
}

// ── Write ──────────────────────────────────────────────────
const outPath = process.argv[2] || "C:/Users/jsber/OneDrive/Documents/Claude_task/descartes08_slides.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("PPTX created: " + outPath);
}).catch(err => {
  console.error("Error:", err);
});
