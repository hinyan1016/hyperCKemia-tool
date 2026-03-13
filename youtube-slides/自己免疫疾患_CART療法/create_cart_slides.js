const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "医知創造ラボ";
pres.title = "自己免疫疾患CAR-Tの全体像";

// ========== WIKIMEDIA IMAGES (base64 embedded) ==========
const fs = require("fs");
const IMG_DIR = __dirname;
function loadImgB64(name) {
  const b64 = fs.readFileSync(`${IMG_DIR}/img_${name}_b64.txt`, "utf-8").trim();
  return "image/png;base64," + b64;
}
const IMG = {
  cart_therapy: loadImgB64("cart_therapy"),
  bcell_activation: loadImgB64("bcell_activation"),
  neuromuscular: loadImgB64("neuromuscular"),
  myelin: loadImgB64("myelin"),
};

// ========== DESIGN CONSTANTS ==========
const C = {
  navy: "1B3A5C",
  blue: "2C5AA0",
  lightBlue: "E8F4FD",
  paleBlue: "F0F7FF",
  white: "FFFFFF",
  offWhite: "F8F9FA",
  lightGray: "DEE2E6",
  gray: "666666",
  darkGray: "333333",
  red: "DC3545",
  yellow: "F5C518",
  green: "28A745",
  orange: "E8850C",
  amber: "FFC107",
  warmBg: "FFF3CD",
  redBg: "F8D7DA",
};
const FONT_JP = "游ゴシック";
const FONT_EN = "Segoe UI";
const W = 10;       // slide width
const H = 5.625;    // slide height

// ========== HELPER FUNCTIONS ==========
const makeShadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 });

function addFooter(slide, num, total) {
  slide.addText(`${num} / ${total}`, {
    x: 8.5, y: 5.2, w: 1.2, h: 0.3, fontSize: 10, fontFace: FONT_EN,
    color: "999999", align: "right"
  });
}

function addSectionLabel(slide, text) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.45, fill: { color: C.navy }
  });
  slide.addText(text, {
    x: 0.4, y: 0.02, w: 9, h: 0.4, fontSize: 12, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
}

function titleSlideStyle(slide) {
  slide.background = { color: C.navy };
}

function sectionDividerStyle(slide) {
  slide.background = { color: C.blue };
}

function contentSlideStyle(slide) {
  slide.background = { color: C.white };
}

// Card helper
function addCard(slide, x, y, w, h, opts = {}) {
  const { accentColor, accentSide } = opts;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: opts.fill || C.white },
    shadow: makeShadow(),
    line: opts.borderColor ? { color: opts.borderColor, width: 1 } : undefined
  });
  if (accentColor) {
    const side = accentSide || "left";
    if (side === "left") {
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w: 0.06, h, fill: { color: accentColor }
      });
    } else if (side === "top") {
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w, h: 0.06, fill: { color: accentColor }
      });
    }
  }
}

const TOTAL = 42;
let slideNum = 0;

// ========================================================
// SLIDE 1: TITLE
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  titleSlideStyle(s);
  // Decorative top line
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.08, fill: { color: C.blue }
  });
  // Main title
  s.addText("自己免疫疾患CAR-Tの全体像", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.2, fontSize: 40, fontFace: FONT_JP,
    color: C.white, bold: true, align: "center"
  });
  // Subtitle
  s.addText("神経内科医のための、できるだけ分かりやすい整理", {
    x: 0.8, y: 2.5, w: 8.4, h: 0.6, fontSize: 20, fontFace: FONT_JP,
    color: "CADCFC", align: "center"
  });
  // Bottom decorative box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 3, y: 3.6, w: 4, h: 0.03, fill: { color: C.blue }
  });
  s.addText("医知創造ラボ", {
    x: 0.8, y: 3.9, w: 8.4, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "8899BB", align: "center"
  });
  s.addText("YouTube公開用プレゼンテーション", {
    x: 0.8, y: 4.4, w: 8.4, h: 0.4, fontSize: 14, fontFace: FONT_JP,
    color: "667799", align: "center"
  });
})();

// ========================================================
// SLIDE 2: INTRODUCTION
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "イントロダクション");
  s.addText("CAR-T療法とは何か？", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 30, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Left card: 従来の治療
  addCard(s, 0.5, 1.5, 4.2, 3.5, { accentColor: C.gray });
  s.addText("従来の自己免疫治療", {
    x: 0.7, y: 1.65, w: 3.8, h: 0.4, fontSize: 18, fontFace: FONT_JP,
    color: C.navy, bold: true
  });
  s.addText([
    { text: "免疫を「弱め続ける」治療", options: { bullet: true, breakLine: true, fontSize: 16, color: C.darkGray } },
    { text: "ステロイド・免疫抑制薬の長期使用", options: { bullet: true, breakLine: true, fontSize: 16, color: C.darkGray } },
    { text: "抗体の「出荷量」を減らす", options: { bullet: true, breakLine: true, fontSize: 16, color: C.darkGray } },
    { text: "再燃のリスクが常に残る", options: { bullet: true, fontSize: 16, color: C.darkGray } },
  ], { x: 0.7, y: 2.2, w: 3.8, h: 2.5, fontFace: FONT_JP, paraSpaceAfter: 8 });

  // Right card: CAR-T
  addCard(s, 5.3, 1.5, 4.2, 3.5, { accentColor: C.blue });
  s.addText("CAR-T療法", {
    x: 5.5, y: 1.65, w: 3.8, h: 0.4, fontSize: 18, fontFace: FONT_JP,
    color: C.blue, bold: true
  });
  s.addText([
    { text: "病気を作るB細胞系を「深く掃除」", options: { bullet: true, breakLine: true, fontSize: 16, color: C.darkGray } },
    { text: "免疫の設定を「組み直す」", options: { bullet: true, breakLine: true, fontSize: 16, color: C.darkGray } },
    { text: "設計部門・工場そのものを壊す", options: { bullet: true, breakLine: true, fontSize: 16, color: C.darkGray } },
    { text: "免疫の「誤学習」をリセット", options: { bullet: true, fontSize: 16, color: C.blue, bold: true } },
  ], { x: 5.5, y: 2.2, w: 3.8, h: 2.5, fontFace: FONT_JP, paraSpaceAfter: 8 });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 3: 目次
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "全体構成");
  s.addText("本日の内容", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 30, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  const items = [
    ["1", "まず結論", "3つのポイント"],
    ["2", "なぜCAR-Tが自己免疫に効くのか", "機序と臨床データ"],
    ["3", "標的の違いを整理する", "CD19 / BCMA / CAAR-T"],
    ["4", "rituximabとの違い", "組織レベルの深部除去"],
    ["5", "神経内科の各論", "MG・NMOSD・MS・CIDP・その他"],
    ["6", "4つの臨床的問い", "実践的な判断枠組み"],
    ["7", "自己免疫CAR-T特有の副作用", "LICATS"],
    ["8", "次に来るもの", "同種CAR-T・CAAR-T"],
    ["9", "まとめ", "免疫抑制から免疫リセットへ"],
  ];

  items.forEach((item, i) => {
    const yy = 1.4 + i * 0.44;
    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.6, y: yy, w: 0.35, h: 0.35, fill: { color: i === 0 ? C.blue : C.navy }
    });
    s.addText(item[0], {
      x: 0.6, y: yy, w: 0.35, h: 0.35, fontSize: 14, fontFace: FONT_EN,
      color: C.white, bold: true, align: "center", valign: "middle"
    });
    s.addText(item[1], {
      x: 1.1, y: yy, w: 4, h: 0.35, fontSize: 16, fontFace: FONT_JP,
      color: C.navy, bold: true, valign: "middle", margin: 0
    });
    s.addText(item[2], {
      x: 5.2, y: yy, w: 4, h: 0.35, fontSize: 14, fontFace: FONT_JP,
      color: C.gray, valign: "middle", margin: 0
    });
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 4: SECTION DIVIDER - 結論
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("01", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("まず結論", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 36, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("3つのポイントで理解する自己免疫CAR-T", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
})();

// ========================================================
// SLIDE 5: 結論 - 3つのポイント
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "1. まず結論");
  s.addText("自己免疫CAR-Tの3つの核心", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 28, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Card 1
  addCard(s, 0.4, 1.4, 2.9, 2.2, { accentColor: C.blue, accentSide: "top" });
  s.addText("第1", {
    x: 0.5, y: 1.55, w: 2.7, h: 0.35, fontSize: 20, fontFace: FONT_JP,
    color: C.blue, bold: true
  });
  s.addText("標的は「抗体」そのものではなく\n「抗体を作る系統の細胞」", {
    x: 0.5, y: 2.0, w: 2.7, h: 1.3, fontSize: 16, fontFace: FONT_JP,
    color: C.darkGray, valign: "top"
  });

  // Card 2
  addCard(s, 3.55, 1.4, 2.9, 2.2, { accentColor: C.green, accentSide: "top" });
  s.addText("第2", {
    x: 3.65, y: 1.55, w: 2.7, h: 0.35, fontSize: 20, fontFace: FONT_JP,
    color: C.green, bold: true
  });
  s.addText("rituximabの延長線ではなく\nもっと深い場所まで\nB細胞系を消せる治療", {
    x: 3.65, y: 2.0, w: 2.7, h: 1.3, fontSize: 16, fontFace: FONT_JP,
    color: C.darkGray, valign: "top"
  });

  // Card 3
  addCard(s, 6.7, 1.4, 2.9, 2.2, { accentColor: C.orange, accentSide: "top" });
  s.addText("第3", {
    x: 6.8, y: 1.55, w: 2.7, h: 0.35, fontSize: 20, fontFace: FONT_JP,
    color: C.orange, bold: true
  });
  s.addText("自己抗体・B細胞の関与が強く\n臓器障害がまだ可逆的な\n段階で相性がよい", {
    x: 6.8, y: 2.0, w: 2.7, h: 1.3, fontSize: 16, fontFace: FONT_JP,
    color: C.darkGray, valign: "top"
  });

  // Bottom summary
  addCard(s, 0.4, 3.85, 9.2, 1.3, { fill: C.lightBlue, accentColor: C.blue });
  s.addText([
    { text: "最も開発が進んでいるのは ", options: { fontSize: 16, color: C.darkGray } },
    { text: "MG", options: { fontSize: 16, color: C.blue, bold: true } },
    { text: "。NMOSD は生物学的に相性◎。MS は中枢到達性が核心。CIDP・GAD関連疾患・自己抗体性脳炎は「効きうる疾患群」が見え始めている。", options: { fontSize: 16, color: C.darkGray } },
  ], { x: 0.6, y: 3.95, w: 8.8, h: 1.0, fontFace: FONT_JP });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 6: SECTION DIVIDER - メカニズム
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("02", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("なぜCAR-Tが自己免疫に効くのか", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 34, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("メカニズムとSLEでの臨床証拠", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
  // Wikimedia: CAR-T therapy diagram (960x300, ratio 3.2:1) — white card for visibility
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.1, y: 3.45, w: 3.6, h: 1.65, fill: { color: "FFFFFF" }, rectRadius: 0.1,
    shadow: { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.2 }
  });
  s.addImage({ data: IMG.cart_therapy, x: 6.25, y: 3.6, w: 3.3, h: 1.03 });
  s.addText("Wikimedia Commons (CC BY-SA 4.0)", {
    x: 6.1, y: 4.7, w: 3.6, h: 0.2, fontSize: 8, fontFace: FONT_EN, color: "999999", align: "center"
  });
})();

// ========================================================
// SLIDE 7: メカニズム概念
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "2. なぜCAR-Tが自己免疫に効くのか");
  s.addText("自己免疫疾患の本質とCAR-Tの発想", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 26, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Flow diagram with shapes
  // Step 1: 誤学習B細胞
  addCard(s, 0.4, 1.5, 2.5, 1.4, { accentColor: C.red, accentSide: "top" });
  s.addText("① 誤学習したB細胞", {
    x: 0.5, y: 1.6, w: 2.3, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true
  });
  s.addText("間違った設計図を持つ\nB細胞が存在", {
    x: 0.5, y: 2.05, w: 2.3, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.darkGray
  });

  // Arrow
  s.addText("→", { x: 2.95, y: 1.85, w: 0.5, h: 0.5, fontSize: 30, fontFace: FONT_EN, color: C.blue, align: "center", valign: "middle" });

  // Step 2: 形質細胞・自己抗体
  addCard(s, 3.5, 1.5, 2.5, 1.4, { accentColor: C.orange, accentSide: "top" });
  s.addText("② 抗体工場が稼働", {
    x: 3.6, y: 1.6, w: 2.3, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true
  });
  s.addText("形質芽細胞・形質細胞が\n自己抗体を作り続ける", {
    x: 3.6, y: 2.05, w: 2.3, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.darkGray
  });

  // Arrow
  s.addText("→", { x: 6.05, y: 1.85, w: 0.5, h: 0.5, fontSize: 30, fontFace: FONT_EN, color: C.blue, align: "center", valign: "middle" });

  // Step 3: 自己免疫疾患
  addCard(s, 6.6, 1.5, 2.9, 1.4, { accentColor: C.red, accentSide: "top" });
  s.addText("③ 自己免疫疾患", {
    x: 6.7, y: 1.6, w: 2.7, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.red, bold: true
  });
  s.addText("自己抗体による\n臓器障害が持続", {
    x: 6.7, y: 2.05, w: 2.7, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.darkGray
  });

  // 従来 vs CAR-T comparison
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.2, w: 4.3, h: 0.4, fill: { color: C.gray }
  });
  s.addText("従来の治療：出荷量を減らす", {
    x: 0.5, y: 3.2, w: 4.1, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle"
  });
  s.addText("ステロイド・免疫抑制薬で工場の生産を抑えるが、設計部門（B細胞）・工場（形質細胞）は残存する。", {
    x: 0.5, y: 3.7, w: 4.1, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.darkGray
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 3.2, w: 4.3, h: 0.4, fill: { color: C.blue }
  });
  s.addText("CAR-T：工場ごと壊す", {
    x: 5.4, y: 3.2, w: 4.1, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle"
  });
  s.addText("設計部門（B細胞）や工場（形質細胞）そのものを破壊。単なる鎮静ではなく、免疫の「誤学習」をリセット。", {
    x: 5.4, y: 3.7, w: 4.1, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.darkGray
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 8: SLEデータ (2022)
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "2. なぜCAR-Tが自己免疫に効くのか");
  s.addText("SLEデータが示した「免疫リセット」の証拠", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 26, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Mackensen 2022 card
  addCard(s, 0.4, 1.4, 9.2, 1.6, { accentColor: C.blue });
  s.addText("Mackensen et al. Nat Med 2022", {
    x: 0.6, y: 1.5, w: 5, h: 0.3, fontSize: 14, fontFace: FONT_EN, color: C.blue, bold: true
  });
  s.addText("難治性SLE 5例 → CD19 CAR-T", {
    x: 0.6, y: 1.85, w: 4, h: 0.35, fontSize: 18, fontFace: FONT_JP, color: C.navy, bold: true
  });

  // Key numbers
  const nums = [
    { val: "5/5", label: "DORIS寛解\n（3か月）", color: C.green },
    { val: "100%", label: "薬剤フリー\n寛解持続", color: C.blue },
    { val: "Naive寄り", label: "再出現B細胞\nの性質", color: C.orange },
  ];
  nums.forEach((n, i) => {
    const xx = 5.5 + i * 1.55;
    s.addText(n.val, {
      x: xx, y: 1.5, w: 1.4, h: 0.55, fontSize: 22, fontFace: FONT_EN,
      color: n.color, bold: true, align: "center", valign: "middle"
    });
    s.addText(n.label, {
      x: xx, y: 2.05, w: 1.4, h: 0.7, fontSize: 12, fontFace: FONT_JP,
      color: C.gray, align: "center"
    });
  });

  // Key insight
  addCard(s, 0.4, 3.2, 9.2, 1.0, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("戻ってきたB細胞は未熟・naive寄り → 病気を覚えたB細胞がそのまま戻ったわけではない → 免疫の再起動が起きている", {
    x: 0.6, y: 3.3, w: 8.8, h: 0.8, fontSize: 16, fontFace: FONT_JP, color: C.navy, bold: true
  });

  // NEJM / CASTLE
  addCard(s, 0.4, 4.4, 4.3, 0.9, { accentColor: C.green });
  s.addText([
    { text: "NEJM 15例 ", options: { bold: true, color: C.navy, fontSize: 15 } },
    { text: "(SLE・筋炎・強皮症)\n多くで免疫抑制薬を止めても寛解持続", options: { color: C.darkGray, fontSize: 14 } },
  ], { x: 0.6, y: 4.45, w: 3.9, h: 0.8, fontFace: FONT_JP });

  addCard(s, 5.3, 4.4, 4.3, 0.9, { accentColor: C.green });
  s.addText([
    { text: "CASTLE試験 24例 ", options: { bold: true, color: C.navy, fontSize: 15 } },
    { text: "(phase 1/2)\n寛解・改善が長期持続することを確認", options: { color: C.darkGray, fontSize: 14 } },
  ], { x: 5.5, y: 4.45, w: 3.9, h: 0.8, fontFace: FONT_JP });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 9: SECTION DIVIDER - 標的の違い
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("03", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("標的の違いを1枚で整理する", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 34, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("CD19 / BCMA / CAAR-T — どこを狙うかで治療が変わる", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
  // Wikimedia: B cell activation (960x1475, ratio 0.65:1, tall) — white card
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.8, y: 1.8, w: 1.9, h: 3.35, fill: { color: "FFFFFF" }, rectRadius: 0.1,
    shadow: { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.2 }
  });
  s.addImage({ data: IMG.bcell_activation, x: 7.95, y: 1.95, w: 1.6, h: 2.46 });
  s.addText("Wikimedia Commons\n(CC BY-SA 4.0)", {
    x: 7.8, y: 4.5, w: 1.9, h: 0.35, fontSize: 7, fontFace: FONT_EN, color: "999999", align: "center"
  });
})();

// ========================================================
// SLIDE 10: B細胞分化の流れ
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "3. 標的の違いを整理する");
  s.addText("B細胞系の「川上→川下」と各標的", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 26, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Flow: naive → memory → plasmablast → plasma cell
  const cells = [
    { name: "Naive\nB cell", x: 0.5, color: "4A90D9" },
    { name: "Memory\nB cell", x: 2.8, color: "2C5AA0" },
    { name: "Plasma-\nblast", x: 5.1, color: "1B3A5C" },
    { name: "Plasma\ncell", x: 7.4, color: "0F2440" },
  ];
  cells.forEach(c => {
    s.addShape(pres.shapes.OVAL, {
      x: c.x, y: 1.5, w: 2.0, h: 1.0, fill: { color: c.color }
    });
    s.addText(c.name, {
      x: c.x, y: 1.5, w: 2.0, h: 1.0, fontSize: 15, fontFace: FONT_EN,
      color: C.white, bold: true, align: "center", valign: "middle"
    });
  });
  // Arrows between cells
  [2.5, 4.8, 7.1].forEach(ax => {
    s.addText("→", {
      x: ax, y: 1.7, w: 0.4, h: 0.5, fontSize: 24, fontFace: FONT_EN,
      color: C.blue, align: "center", valign: "middle"
    });
  });

  // Brackets showing coverage
  // CD20 coverage
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 2.7, w: 4.0, h: 0.06, fill: { color: C.gray }
  });
  s.addText("CD20 — 成熟B細胞", {
    x: 0.8, y: 2.8, w: 4.0, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.gray, align: "center"
  });

  // CD19 coverage (wider)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.2, w: 7.0, h: 0.06, fill: { color: C.blue }
  });
  s.addText("CD19 — B細胞系を広くカバー", {
    x: 0.5, y: 3.3, w: 7.0, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.blue, align: "center", bold: true
  });

  // BCMA coverage
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 3.7, w: 4.3, h: 0.06, fill: { color: C.orange }
  });
  s.addText("BCMA — 形質芽細胞・形質細胞", {
    x: 5.1, y: 3.8, w: 4.3, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.orange, align: "center", bold: true
  });

  // Bottom summary
  addCard(s, 0.4, 4.3, 9.2, 1.0, { fill: C.paleBlue, accentColor: C.blue });
  s.addText("CD19 は「広く深く」、BCMA は「抗体工場を直撃」、CD19+BCMA は「両方」、CAAR-T は「犯人だけ」", {
    x: 0.6, y: 4.35, w: 8.8, h: 0.85, fontSize: 16, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle"
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 11: 標的比較テーブル
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "3. 標的の違いを整理する");
  s.addText("標的別CAR-Tの比較", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 26, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  const headerOpts = { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 14, fontFace: FONT_JP, align: "center", valign: "middle" };
  const cellOpts = { fontSize: 14, fontFace: FONT_JP, color: C.darkGray, valign: "middle" };
  const cellBoldOpts = { ...cellOpts, bold: true, color: C.navy };

  const rows = [
    [
      { text: "標的", options: headerOpts },
      { text: "たとえ", options: headerOpts },
      { text: "主に狙う場所", options: headerOpts },
      { text: "何が得意か", options: headerOpts },
    ],
    [
      { text: "CD20", options: cellBoldOpts },
      { text: "現場の作業員", options: cellOpts },
      { text: "成熟B細胞", options: cellOpts },
      { text: "ふつうのB細胞を減らす", options: cellOpts },
    ],
    [
      { text: "CD19", options: { ...cellBoldOpts, color: C.blue } },
      { text: "会社全体＋記憶倉庫", options: cellOpts },
      { text: "B細胞系を広く", options: cellOpts },
      { text: "免疫記憶ごと深くリセット", options: { ...cellOpts, bold: true, color: C.blue } },
    ],
    [
      { text: "BCMA", options: { ...cellBoldOpts, color: C.orange } },
      { text: "抗体工場", options: cellOpts },
      { text: "plasmablast/\nplasma cell", options: cellOpts },
      { text: "いま抗体を作る細胞を叩く", options: { ...cellOpts, bold: true, color: C.orange } },
    ],
    [
      { text: "CD19+BCMA", options: { ...cellBoldOpts, color: C.green } },
      { text: "倉庫＋工場の両方", options: cellOpts },
      { text: "上流と下流", options: cellOpts },
      { text: "より深い除去", options: { ...cellOpts, bold: true, color: C.green } },
    ],
    [
      { text: "CAAR-T", options: { ...cellBoldOpts, color: C.red } },
      { text: "指名手配犯だけ狙う\n狙撃手", options: cellOpts },
      { text: "自己抗原特異的\nB細胞", options: cellOpts },
      { text: "正常免疫をなるべく残す", options: { ...cellOpts, bold: true, color: C.red } },
    ],
  ];

  s.addTable(rows, {
    x: 0.4, y: 1.4, w: 9.2,
    colW: [1.6, 2.4, 2.2, 3.0],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.45, 0.5, 0.5, 0.55, 0.5, 0.55],
    autoPage: false,
  });

  // Highlight box
  addCard(s, 0.4, 4.55, 9.2, 0.8, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("現在、神経内科で最も現実味があるのは CD19、BCMA、CD19+BCMA。CAAR-Tは次世代。", {
    x: 0.6, y: 4.6, w: 8.8, h: 0.7, fontSize: 16, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle"
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 12: SECTION DIVIDER - rituximab
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("04", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("rituximabと何が決定的に違うのか", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 34, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("組織レベルでの深部B細胞除去", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
})();

// ========================================================
// SLIDE 13: rituximab vs CAR-T
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "4. rituximabとの違い");
  s.addText("rituximab vs CAR-T：何が違うのか", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 26, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Left: rituximab
  addCard(s, 0.4, 1.4, 4.3, 2.8, { accentColor: C.gray });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.4, w: 4.3, h: 0.45, fill: { color: C.gray }
  });
  s.addText("rituximab（CD20標的）", {
    x: 0.5, y: 1.4, w: 4.1, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle"
  });
  s.addText([
    { text: "末梢血のB細胞は消える", options: { bullet: true, breakLine: true, fontSize: 15, color: C.darkGray } },
    { text: "形質細胞はCD20陰性 → 残存", options: { bullet: true, breakLine: true, fontSize: 15, color: C.red, bold: true } },
    { text: "リンパ節・骨髄・病変組織の\n病的B細胞系が残ることがある", options: { bullet: true, breakLine: true, fontSize: 15, color: C.darkGray } },
    { text: "\"病気の根っこ\" が残る場合", options: { bullet: true, fontSize: 15, color: C.red } },
  ], { x: 0.6, y: 2.0, w: 3.9, h: 2.0, fontFace: FONT_JP, paraSpaceAfter: 6 });

  // Right: CAR-T
  addCard(s, 5.3, 1.4, 4.3, 2.8, { accentColor: C.blue });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.4, w: 4.3, h: 0.45, fill: { color: C.blue }
  });
  s.addText("CAR-T（CD19標的）", {
    x: 5.4, y: 1.4, w: 4.1, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle"
  });
  s.addText([
    { text: "末梢血のB細胞も消える", options: { bullet: true, breakLine: true, fontSize: 15, color: C.darkGray } },
    { text: "リンパ節のB細胞も深く除去", options: { bullet: true, breakLine: true, fontSize: 15, color: C.blue, bold: true } },
    { text: "非リンパ組織でもB細胞が\n大きく減少", options: { bullet: true, breakLine: true, fontSize: 15, color: C.darkGray } },
    { text: "組織の奥の病的B細胞まで掘り起こす", options: { bullet: true, fontSize: 15, color: C.blue, bold: true } },
  ], { x: 5.5, y: 2.0, w: 3.9, h: 2.0, fontFace: FONT_JP, paraSpaceAfter: 6 });

  // Bottom key message
  addCard(s, 0.4, 4.45, 9.2, 0.9, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("CAR-TはrituximabよりCT「強い薬」というより、「別の相の治療」。組織比較研究（Tur et al. 2025）で直接証明された。", {
    x: 0.6, y: 4.5, w: 8.8, h: 0.8, fontSize: 16, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle"
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 14: SECTION DIVIDER - 神経内科各論
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("05", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("神経内科の各論", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 36, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("MG・NMOSD・MS・CIDP・その他", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
  // Wikimedia: Neuromuscular junction (960x705, ratio 1.36:1) — white card
  // w:2.6 → h = 2.6/1.36 = 1.91
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.4, y: 3.25, w: 3.0, h: 2.2, fill: { color: "FFFFFF" }, rectRadius: 0.1,
    shadow: { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.2 }
  });
  s.addImage({ data: IMG.neuromuscular, x: 6.6, y: 3.35, w: 2.6, h: 1.91 });
  s.addText("Wikimedia Commons (CC BY-SA 4.0)", {
    x: 6.4, y: 5.1, w: 3.0, h: 0.2, fontSize: 8, fontFace: FONT_EN, color: "999999", align: "center"
  });
})();

// ========================================================
// SLIDE 15: MG overview
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "5-1. MG（重症筋無力症）");
  s.addText("MGは、いま最も面白い", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 28, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  s.addText("MGでは2つのCAR-T思想が並走している", {
    x: 0.5, y: 1.3, w: 9, h: 0.4, fontSize: 18, fontFace: FONT_JP,
    color: C.darkGray, margin: 0
  });

  // Left: Descartes-08 (やさしい型)
  addCard(s, 0.4, 1.9, 4.3, 3.0, { accentColor: C.green, accentSide: "top" });
  s.addText("やさしく叩く型", {
    x: 0.5, y: 2.05, w: 4.1, h: 0.35, fontSize: 18, fontFace: FONT_JP,
    color: C.green, bold: true
  });
  s.addText("Descartes-08 (BCMA mRNA CAR-T)", {
    x: 0.5, y: 2.4, w: 4.1, h: 0.3, fontSize: 14, fontFace: FONT_EN,
    color: C.blue, bold: true
  });
  s.addText([
    { text: "外来投与が可能", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "リンパ球除去前処置なし", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "mRNA型で一過性（安全）", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "再投与で再び改善可能", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "→ 抗体工場を外来で安全に間引く", options: { fontSize: 15, bold: true, color: C.green } },
  ], { x: 0.6, y: 2.8, w: 3.9, h: 2.0, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 4 });

  // Right: 深いリセット型
  addCard(s, 5.3, 1.9, 4.3, 3.0, { accentColor: C.blue, accentSide: "top" });
  s.addText("深くリセットする型", {
    x: 5.4, y: 2.05, w: 4.1, h: 0.35, fontSize: 18, fontFace: FONT_JP,
    color: C.blue, bold: true
  });
  s.addText("BCMA/CD19二重標的 CAR-T", {
    x: 5.4, y: 2.4, w: 4.1, h: 0.3, fontSize: 14, fontFace: FONT_EN,
    color: C.blue, bold: true
  });
  s.addText([
    { text: "単回投与", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "前処置あり", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "より深いB細胞除去", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "ステロイド・免疫抑制薬中止可能", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "→ 深い免疫リセットを目指す", options: { fontSize: 15, bold: true, color: C.blue } },
  ], { x: 5.5, y: 2.8, w: 3.9, h: 2.0, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 4 });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 16: MG - Descartes-08 詳細
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "5-1. MG — Descartes-08");
  s.addText("Descartes-08：外来型mRNA CAR-Tの臨床データ", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 24, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // MG-001
  addCard(s, 0.4, 1.4, 4.3, 2.0, { accentColor: C.green });
  s.addText("MG-001試験 (Phase 1b/2a)", {
    x: 0.6, y: 1.5, w: 4, h: 0.35, fontSize: 16, fontFace: FONT_EN, color: C.green, bold: true
  });
  s.addText([
    { text: "重篤なCRS/ICANSなし", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "症状スケール改善が最大9か月持続", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "6週投与群：7例全例が3か月で改善", options: { bullet: true, breakLine: true, fontSize: 15, bold: true } },
    { text: "5/7例が12か月時点でも反応維持", options: { bullet: true, fontSize: 15, bold: true } },
  ], { x: 0.6, y: 1.9, w: 3.9, h: 1.4, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 4 });

  // Phase 2b
  addCard(s, 5.3, 1.4, 4.3, 2.0, { accentColor: C.blue });
  s.addText("Phase 2b (2026 Nat Med)", {
    x: 5.5, y: 1.5, w: 4, h: 0.35, fontSize: 16, fontFace: FONT_EN, color: C.blue, bold: true
  });
  s.addText([
    { text: "無作為化・二重盲検・プラセボ対照", options: { bullet: true, breakLine: true, fontSize: 15, bold: true } },
    { text: "Phase 3 (AURORA試験) も進行中", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "再増悪3例中2例は再投与で", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "速やかに再改善", options: { fontSize: 15, bold: true, color: C.blue } },
  ], { x: 5.5, y: 1.9, w: 3.9, h: 1.4, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 4 });

  // Key interpretation
  addCard(s, 0.4, 3.6, 9.2, 1.6, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("臨床的な解釈", {
    x: 0.6, y: 3.7, w: 3, h: 0.35, fontSize: 16, fontFace: FONT_JP, color: C.blue, bold: true
  });
  s.addText("Descartes-08は、「1回で深く免疫をリセットする古典的CAR-T」というより、「抗体工場を外来で安全に間引く、再投与可能な細胞治療」に近い。mRNA型で一過性、前処置不要、再投与で再反応が得られていることがこの見方を支持。", {
    x: 0.6, y: 4.05, w: 8.8, h: 1.0, fontSize: 15, fontFace: FONT_JP, color: C.darkGray
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 17: MG - BCMA/CD19二重標的
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "5-1. MG — BCMA/CD19二重標的");
  s.addText("BCMA/CD19二重標的CAR-T：Phase 1試験", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 24, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  s.addText("難治性全身型MG 18例・単回投与 (Zhang et al. 2025)", {
    x: 0.5, y: 1.25, w: 9, h: 0.35, fontSize: 16, fontFace: FONT_JP, color: C.gray
  });

  // Key numbers in big cards
  const stats = [
    { val: "0%", label: "ICANS", sub: "神経毒性なし", color: C.green },
    { val: "39%", label: "CRS (Gr1)", sub: "軽微のみ", color: C.green },
    { val: "14/17", label: "MM status", sub: "minimal manifestation", color: C.blue },
    { val: "15/17", label: "ステロイド", sub: "中止達成", color: C.blue },
  ];

  stats.forEach((st, i) => {
    const xx = 0.4 + i * 2.35;
    addCard(s, xx, 1.8, 2.15, 1.8, { accentColor: st.color, accentSide: "top" });
    s.addText(st.val, {
      x: xx, y: 1.95, w: 2.15, h: 0.7, fontSize: 32, fontFace: FONT_EN,
      color: st.color, bold: true, align: "center", valign: "middle"
    });
    s.addText(st.label, {
      x: xx, y: 2.7, w: 2.15, h: 0.35, fontSize: 14, fontFace: FONT_JP,
      color: C.navy, bold: true, align: "center"
    });
    s.addText(st.sub, {
      x: xx, y: 3.05, w: 2.15, h: 0.35, fontSize: 13, fontFace: FONT_JP,
      color: C.gray, align: "center"
    });
  });

  // Additional findings
  addCard(s, 0.4, 3.85, 9.2, 0.7, { fill: C.paleBlue });
  s.addText("全例が非ステロイド系免疫抑制薬を中止。AChR抗体陰性化も一部で確認。→ Descartes-08よりも「深いリセット」に近い。", {
    x: 0.6, y: 3.9, w: 8.8, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle"
  });

  // Warning box
  addCard(s, 0.4, 4.7, 9.2, 0.7, { fill: C.warmBg, accentColor: C.amber });
  s.addText("⚠ MGの本質的な面白さ：同じCAR-Tでも「外来でやさしく叩く型」と「深くリセットする型」が共存。他の神経免疫疾患にも広がる可能性。", {
    x: 0.6, y: 4.75, w: 8.8, h: 0.6, fontSize: 15, fontFace: FONT_JP, color: C.darkGray, valign: "middle"
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 18: NMOSD
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "5-2. NMOSD");
  s.addText("NMOSD：生物学的にはかなり相性がよい", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 26, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Why it fits
  addCard(s, 0.4, 1.4, 4.3, 1.6, { accentColor: C.green, accentSide: "top" });
  s.addText("なぜ相性がよいか", {
    x: 0.6, y: 1.55, w: 4, h: 0.35, fontSize: 18, fontFace: FONT_JP, color: C.green, bold: true
  });
  s.addText([
    { text: "AQP4抗体の病原性が明確", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "自己抗体が病態の中心", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "抗体産生源を狙う意義が大きい", options: { bullet: true, fontSize: 15, bold: true } },
  ], { x: 0.6, y: 1.95, w: 3.9, h: 1.0, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 4 });

  // Clinical data
  addCard(s, 5.3, 1.4, 4.3, 1.6, { accentColor: C.blue, accentSide: "top" });
  s.addText("Phase 1試験 (BCMA CAR-T)", {
    x: 5.5, y: 1.55, w: 4, h: 0.35, fontSize: 16, fontFace: FONT_EN, color: C.blue, bold: true
  });
  s.addText("AQP4-IgG陽性NMOSD 12例", {
    x: 5.5, y: 1.95, w: 4, h: 0.3, fontSize: 15, fontFace: FONT_JP, color: C.navy, bold: true
  });

  // Numbers
  const nmoNums = [
    { val: "11/12", label: "無再発", x: 5.5 },
    { val: "Gr1-2", label: "CRS", x: 7.0 },
    { val: "↓", label: "AQP4抗体", x: 8.5 },
  ];
  nmoNums.forEach(n => {
    s.addText(n.val, {
      x: n.x, y: 2.3, w: 1.3, h: 0.45, fontSize: 22, fontFace: FONT_EN,
      color: C.blue, bold: true, align: "center"
    });
    s.addText(n.label, {
      x: n.x, y: 2.75, w: 1.3, h: 0.2, fontSize: 12, fontFace: FONT_JP,
      color: C.gray, align: "center"
    });
  });

  // Limitations
  addCard(s, 0.4, 3.3, 9.2, 1.3, { fill: C.warmBg, accentColor: C.amber });
  s.addText("⚠ 留意点", {
    x: 0.6, y: 3.4, w: 2, h: 0.3, fontSize: 16, fontFace: FONT_JP, color: C.orange, bold: true
  });
  s.addText([
    { text: "追跡期間はまだ短い（中央値5.5か月）", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "長期再発率・感染リスク・再投与戦略は未確立", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "それでも、自己免疫CAR-Tの適応候補として上位", options: { bullet: true, fontSize: 15, bold: true, color: C.navy } },
  ], { x: 0.6, y: 3.75, w: 8.8, h: 0.8, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 2 });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 19: MS
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "5-3. MS（多発性硬化症）");
  s.addText("MS：「中枢に届くか」が勝負", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 28, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Challenge
  addCard(s, 0.4, 1.4, 4.3, 1.5, { accentColor: C.red });
  s.addText("核心的な課題", {
    x: 0.6, y: 1.5, w: 4, h: 0.35, fontSize: 18, fontFace: FONT_JP, color: C.red, bold: true
  });
  s.addText("進行型MSでは末梢だけでなく中枢神経内に閉じこもったB細胞炎症（CNSコンパートメント）が問題。CAR-Tがここに届くなら意味は非常に大きい。", {
    x: 0.6, y: 1.9, w: 3.9, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.darkGray
  });

  // Initial data
  addCard(s, 5.3, 1.4, 4.3, 1.5, { accentColor: C.green });
  s.addText("初期データ (CD19 CAR-T)", {
    x: 5.5, y: 1.5, w: 4, h: 0.35, fontSize: 16, fontFace: FONT_EN, color: C.green, bold: true
  });
  s.addText("進行型MS 2例：\n• CAR-Tが髄液内で検出\n• ICANSは起きず\n• 1例で髄液内抗体産生が\n  day 64まで低下", {
    x: 5.5, y: 1.9, w: 3.9, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.darkGray
  });

  // Key insight
  addCard(s, 0.4, 3.15, 9.2, 0.9, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("「CAR-TがCNSコンパートメントに入り、少なくとも一部では機能している」ことを示す初期シグナル", {
    x: 0.6, y: 3.2, w: 8.8, h: 0.8, fontSize: 16, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle"
  });

  // Ongoing trials
  addCard(s, 0.4, 4.3, 6.6, 1.0, { accentColor: C.blue });
  s.addText("進行中の試験", {
    x: 0.6, y: 4.35, w: 2, h: 0.3, fontSize: 16, fontFace: FONT_JP, color: C.blue, bold: true
  });
  s.addText([
    { text: "KYV-101試験 ", options: { bold: true, color: C.navy } },
    { text: "(progressive MS対象)  ", options: { color: C.darkGray } },
    { text: "AZD0120 ", options: { bold: true, color: C.navy } },
    { text: "(CD19/BCMA二重標的)", options: { color: C.darkGray } },
  ], { x: 0.6, y: 4.7, w: 6.2, h: 0.5, fontFace: FONT_JP, fontSize: 14 });

  // Myelin illustration (960x1136, ratio 0.85:1, tall) — correct aspect ratio
  // h:1.4 → w = 1.4 * 0.845 = 1.18
  s.addImage({ data: IMG.myelin, x: 7.5, y: 3.6, w: 1.18, h: 1.4 });
  s.addText("Wikimedia Commons\n(CC BY-SA 4.0)", {
    x: 7.2, y: 4.95, w: 1.8, h: 0.3, fontSize: 7, fontFace: FONT_EN, color: "7B8FA8", align: "center"
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 20: CIDP
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "5-4. CIDP");
  s.addText("CIDP：「全員」ではなく「合うサブタイプ」がありそう", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 24, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  s.addText("CIDPは異質性が大きい疾患群。CAR-Tが全体に効くというより、自己抗体・形質細胞依存性の強いサブタイプに効くと考えるのが自然。", {
    x: 0.5, y: 1.3, w: 9, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.darkGray
  });

  // Case reports
  addCard(s, 0.4, 2.0, 4.3, 2.2, { accentColor: C.blue });
  s.addText("BCMA CAR-T（2例）", {
    x: 0.6, y: 2.1, w: 4, h: 0.35, fontSize: 16, fontFace: FONT_EN, color: C.blue, bold: true
  });
  s.addText([
    { text: "2例とも6か月以内に薬剤フリー寛解", options: { bullet: true, breakLine: true, fontSize: 15, bold: true } },
    { text: "1例は24か月以上寛解を維持", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "1例は重症COVID-19後に12か月で再発", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "再発時に病的B細胞の再活性化と\n自己抗体の再出現を確認", options: { bullet: true, fontSize: 15 } },
  ], { x: 0.6, y: 2.5, w: 3.9, h: 1.6, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 4 });

  // Implication
  addCard(s, 5.3, 2.0, 4.3, 2.2, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("臨床的示唆", {
    x: 5.5, y: 2.1, w: 4, h: 0.35, fontSize: 18, fontFace: FONT_JP, color: C.blue, bold: true
  });
  s.addText([
    { text: "CIDPは1つではない", options: { breakLine: true, fontSize: 16, bold: true, color: C.navy } },
    { text: "\nB細胞・抗体依存性の強い患者\nでは、CAR-Tが「根治的に近い寛解」\nを狙えるかもしれない", options: { fontSize: 15, color: C.darkGray } },
    { text: "\n\n軸索障害が前景で自己抗体の\n証拠が弱い例には当てはめにくい", options: { fontSize: 14, color: C.gray } },
  ], { x: 5.5, y: 2.5, w: 3.9, h: 1.6, fontFace: FONT_JP });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 21: SPS/GAD/脳炎
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "5-5. SPS・GAD関連・自己抗体性脳炎");
  s.addText("SPS、GAD関連疾患、自己抗体性脳炎", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 24, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  s.addText("まだ症例報告中心だが、内容はかなり刺激的", {
    x: 0.5, y: 1.2, w: 9, h: 0.35, fontSize: 16, fontFace: FONT_JP, color: C.gray
  });

  // SPS
  addCard(s, 0.4, 1.7, 2.9, 2.2, { accentColor: C.blue, accentSide: "top" });
  s.addText("Stiff-Person\nSyndrome", {
    x: 0.5, y: 1.85, w: 2.7, h: 0.55, fontSize: 15, fontFace: FONT_EN, color: C.blue, bold: true
  });
  s.addText("CD19 CAR-T後：\n• 歩行速度 100%以上改善\n• 歩行距離 50m未満→6km超\n• ベンゾジアゼピン40%減\n• 毒性は低グレードCRS", {
    x: 0.5, y: 2.4, w: 2.7, h: 1.4, fontSize: 13, fontFace: FONT_JP, color: C.darkGray
  });

  // DAGLA脳炎
  addCard(s, 3.55, 1.7, 2.9, 2.2, { accentColor: C.green, accentSide: "top" });
  s.addText("DAGLA抗体\n関連脳炎", {
    x: 3.65, y: 1.85, w: 2.7, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.green, bold: true
  });
  s.addText("CD19 CAR-T後：\n• 臨床スコア改善\n• 血清・髄液抗体が低下\n• 髄液OCBが\n  陽性→陰性化", {
    x: 3.65, y: 2.4, w: 2.7, h: 1.4, fontSize: 13, fontFace: FONT_JP, color: C.darkGray
  });

  // GAD65小脳失調
  addCard(s, 6.7, 1.7, 2.9, 2.2, { accentColor: C.orange, accentSide: "top" });
  s.addText("GAD65抗体\n小脳失調症", {
    x: 6.8, y: 1.85, w: 2.7, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true
  });
  s.addText("CD19 CAR-T後：\n• Day 90 血清GAD65抗体\n  95%低下\n• Day 30から失調改善\n• 毒性はGr1 CRSのみ", {
    x: 6.8, y: 2.4, w: 2.7, h: 1.4, fontSize: 13, fontFace: FONT_JP, color: C.darkGray
  });

  // Warning
  addCard(s, 0.4, 4.15, 9.2, 1.1, { fill: C.warmBg, accentColor: C.amber });
  s.addText("⚠ まだ「有望」であって「確立」ではない", {
    x: 0.6, y: 4.25, w: 8.8, h: 0.35, fontSize: 16, fontFace: FONT_JP, color: C.orange, bold: true
  });
  s.addText("症例報告は強い仮説を生むが、適応選択・長期予後・再投与・感染管理はこれから。", {
    x: 0.6, y: 4.65, w: 8.8, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.darkGray
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 22: SECTION DIVIDER - 4つの問い
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("06", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("神経内科医が考えるべき4つの問い", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 34, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("臨床的な判断枠組み", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
})();

// ========================================================
// SLIDE 23: 問い1
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "6. 4つの問い");

  // Question number
  s.addShape(pres.shapes.OVAL, {
    x: 0.5, y: 0.7, w: 0.6, h: 0.6, fill: { color: C.blue }
  });
  s.addText("Q1", {
    x: 0.5, y: 0.7, w: 0.6, h: 0.6, fontSize: 18, fontFace: FONT_EN,
    color: C.white, bold: true, align: "center", valign: "middle"
  });
  s.addText("その病気は、本当にB細胞/自己抗体が主役か？", {
    x: 1.3, y: 0.7, w: 8, h: 0.6, fontSize: 24, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0, valign: "middle"
  });

  s.addText("CAR-Tは万能の免疫抑制ではない。B細胞が病態の中心であるほど、相性が良い。", {
    x: 0.5, y: 1.5, w: 9, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.darkGray
  });

  // Suitability spectrum
  const diseases = [
    { name: "NMOSD\n(AQP4抗体)", fit: "◎", color: C.green, desc: "抗体の病原性が明確" },
    { name: "AChR抗体\n陽性MG", fit: "◎", color: C.green, desc: "B細胞/抗体依存" },
    { name: "自己抗体性\n脳炎（一部）", fit: "○", color: C.blue, desc: "抗体標的が明確な型" },
    { name: "進行型MS", fit: "△", color: C.orange, desc: "病型の見極めが重要" },
    { name: "CIDP", fit: "△", color: C.orange, desc: "サブタイプ依存" },
  ];

  diseases.forEach((d, i) => {
    const yy = 2.1 + i * 0.6;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: yy, w: 9, h: 0.5,
      fill: { color: i % 2 === 0 ? C.offWhite : C.white }
    });
    s.addText(d.fit, {
      x: 0.6, y: yy, w: 0.5, h: 0.5, fontSize: 20, fontFace: FONT_JP,
      color: d.color, bold: true, align: "center", valign: "middle"
    });
    s.addText(d.name, {
      x: 1.2, y: yy, w: 2.5, h: 0.5, fontSize: 15, fontFace: FONT_JP,
      color: C.navy, bold: true, valign: "middle"
    });
    s.addText(d.desc, {
      x: 4.0, y: yy, w: 5.5, h: 0.5, fontSize: 15, fontFace: FONT_JP,
      color: C.darkGray, valign: "middle"
    });
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 24: 問い2
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "6. 4つの問い");

  s.addShape(pres.shapes.OVAL, {
    x: 0.5, y: 0.7, w: 0.6, h: 0.6, fill: { color: C.blue }
  });
  s.addText("Q2", {
    x: 0.5, y: 0.7, w: 0.6, h: 0.6, fontSize: 18, fontFace: FONT_EN,
    color: C.white, bold: true, align: "center", valign: "middle"
  });
  s.addText("障害はまだ可逆的か？", {
    x: 1.3, y: 0.7, w: 8, h: 0.6, fontSize: 26, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0, valign: "middle"
  });

  // Diagram: reversible vs irreversible
  // Left: good timing
  addCard(s, 0.4, 1.6, 4.3, 2.5, { accentColor: C.green, accentSide: "top" });
  s.addText("CAR-Tが効きやすいタイミング", {
    x: 0.6, y: 1.75, w: 4, h: 0.4, fontSize: 17, fontFace: FONT_JP, color: C.green, bold: true
  });
  s.addText([
    { text: "まだ機能改善余地のある時期", options: { bullet: true, breakLine: true, fontSize: 16, bold: true } },
    { text: "炎症が活動性で進行中", options: { bullet: true, breakLine: true, fontSize: 16 } },
    { text: "組織障害が可逆的な段階", options: { bullet: true, breakLine: true, fontSize: 16 } },
    { text: "自己抗体が病態を駆動中", options: { bullet: true, fontSize: 16 } },
  ], { x: 0.6, y: 2.2, w: 3.9, h: 1.7, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 6 });

  // Right: bad timing
  addCard(s, 5.3, 1.6, 4.3, 2.5, { accentColor: C.red, accentSide: "top" });
  s.addText("効果が限定的な可能性", {
    x: 5.5, y: 1.75, w: 4, h: 0.4, fontSize: 17, fontFace: FONT_JP, color: C.red, bold: true
  });
  s.addText([
    { text: "固定化した不可逆障害が前景", options: { bullet: true, breakLine: true, fontSize: 16, bold: true } },
    { text: "瘢痕化・変性が進行済み", options: { bullet: true, breakLine: true, fontSize: 16 } },
    { text: "自己抗体を止めても\n壊れた組織は戻らない", options: { bullet: true, breakLine: true, fontSize: 16 } },
    { text: "神経再生そのものではない", options: { bullet: true, fontSize: 16 } },
  ], { x: 5.5, y: 2.2, w: 3.9, h: 1.7, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 6 });

  // Key message
  addCard(s, 0.4, 4.35, 9.2, 0.9, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("CAR-Tは「炎症や病的免疫のドライバーを止める治療」であって、神経再生ではない。タイミングが重要。", {
    x: 0.6, y: 4.4, w: 8.8, h: 0.8, fontSize: 16, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle"
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 25: 問い3
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "6. 4つの問い");

  s.addShape(pres.shapes.OVAL, {
    x: 0.5, y: 0.7, w: 0.6, h: 0.6, fill: { color: C.blue }
  });
  s.addText("Q3", {
    x: 0.5, y: 0.7, w: 0.6, h: 0.6, fontSize: 18, fontFace: FONT_EN,
    color: C.white, bold: true, align: "center", valign: "middle"
  });
  s.addText("CD19か、BCMAか、両方か？", {
    x: 1.3, y: 0.7, w: 8, h: 0.6, fontSize: 26, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0, valign: "middle"
  });

  // Three columns
  const targets = [
    { name: "CD19", motto: "広く深くリセットしたい", desc: "B細胞系全体を広くカバー。免疫記憶ごとリセットを狙う。", color: C.blue },
    { name: "BCMA", motto: "抗体工場を叩きたい", desc: "形質芽細胞・形質細胞を直撃。いま抗体を産生する細胞を標的。", color: C.orange },
    { name: "CD19+BCMA", motto: "供給源も工場も両方叩く", desc: "上流と下流を同時に標的。より深い除去を目指す。", color: C.green },
  ];

  targets.forEach((t, i) => {
    const xx = 0.4 + i * 3.15;
    addCard(s, xx, 1.5, 2.95, 2.2, { accentColor: t.color, accentSide: "top" });
    s.addText(t.name, {
      x: xx + 0.1, y: 1.65, w: 2.75, h: 0.45, fontSize: 22, fontFace: FONT_EN,
      color: t.color, bold: true, align: "center"
    });
    s.addText(t.motto, {
      x: xx + 0.1, y: 2.1, w: 2.75, h: 0.4, fontSize: 15, fontFace: FONT_JP,
      color: C.navy, bold: true, align: "center"
    });
    s.addText(t.desc, {
      x: xx + 0.1, y: 2.55, w: 2.75, h: 0.9, fontSize: 14, fontFace: FONT_JP,
      color: C.darkGray, align: "center"
    });
  });

  // CAAR-T future
  addCard(s, 0.4, 3.95, 9.2, 1.3, { accentColor: C.red });
  s.addText("将来：CAAR-T（犯人だけを狙う）", {
    x: 0.6, y: 4.05, w: 4, h: 0.35, fontSize: 17, fontFace: FONT_JP, color: C.red, bold: true
  });
  s.addText([
    { text: "MuSK-CAAR-T：", options: { bold: true, color: C.navy } },
    { text: "抗MuSK B細胞を選択除去、総B細胞・IgGは保持  ", options: { color: C.darkGray } },
    { text: "NMDAR-CAAR-T：", options: { bold: true, color: C.navy } },
    { text: "抗NMDAR B細胞を選択除去 → 正常免疫を温存しながら犯人クローンだけを消す方向", options: { color: C.darkGray } },
  ], { x: 0.6, y: 4.45, w: 8.8, h: 0.7, fontFace: FONT_JP, fontSize: 14 });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 26: 問い4
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "6. 4つの問い");

  s.addShape(pres.shapes.OVAL, {
    x: 0.5, y: 0.7, w: 0.6, h: 0.6, fill: { color: C.blue }
  });
  s.addText("Q4", {
    x: 0.5, y: 0.7, w: 0.6, h: 0.6, fontSize: 18, fontFace: FONT_EN,
    color: C.white, bold: true, align: "center", valign: "middle"
  });
  s.addText("患者はCAR-Tの物流と毒性に耐えられるか？", {
    x: 1.3, y: 0.7, w: 8, h: 0.6, fontSize: 22, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0, valign: "middle"
  });

  // Process flow
  const steps = [
    { num: "1", label: "アフェレーシス", desc: "T細胞採取" },
    { num: "2", label: "製造待機", desc: "数週間" },
    { num: "3", label: "リンパ球除去\n化学療法", desc: "前処置" },
    { num: "4", label: "CAR-T投与", desc: "単回注入" },
    { num: "5", label: "CRS/ICANS\n監視", desc: "入院管理" },
  ];

  steps.forEach((st, i) => {
    const xx = 0.3 + i * 1.95;
    s.addShape(pres.shapes.RECTANGLE, {
      x: xx, y: 1.5, w: 1.75, h: 1.4,
      fill: { color: C.offWhite }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: xx, y: 1.5, w: 1.75, h: 0.06, fill: { color: C.navy }
    });
    s.addText(st.num, {
      x: xx, y: 1.6, w: 0.35, h: 0.35, fontSize: 14, fontFace: FONT_EN,
      color: C.white, bold: true, align: "center", valign: "middle"
    });
    s.addShape(pres.shapes.OVAL, {
      x: xx + 0.05, y: 1.62, w: 0.3, h: 0.3, fill: { color: C.navy }
    });
    s.addText(st.num, {
      x: xx + 0.05, y: 1.62, w: 0.3, h: 0.3, fontSize: 12, fontFace: FONT_EN,
      color: C.white, bold: true, align: "center", valign: "middle"
    });
    s.addText(st.label, {
      x: xx + 0.05, y: 2.0, w: 1.65, h: 0.5, fontSize: 14, fontFace: FONT_JP,
      color: C.navy, bold: true, align: "center"
    });
    s.addText(st.desc, {
      x: xx + 0.05, y: 2.5, w: 1.65, h: 0.3, fontSize: 12, fontFace: FONT_JP,
      color: C.gray, align: "center"
    });
  });

  // Key message
  addCard(s, 0.4, 3.2, 9.2, 1.0, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("古典的CAR-Tは「良い薬」だけでなく「良いセンター」が必要な治療。アフェレーシスから入院管理まで一連の物流体制が求められる。", {
    x: 0.6, y: 3.25, w: 8.8, h: 0.9, fontSize: 16, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle"
  });

  // Descartes-08 exception
  addCard(s, 0.4, 4.45, 9.2, 0.9, { fill: C.warmBg, accentColor: C.green });
  s.addText([
    { text: "Descartes-08が特別な理由：", options: { bold: true, color: C.green, fontSize: 16 } },
    { text: "この物流を大きく簡略化した。外来投与、前処置なし、mRNA一過性 → 患者負担が劇的に軽減。", options: { color: C.darkGray, fontSize: 15 } },
  ], { x: 0.6, y: 4.5, w: 8.8, h: 0.8, fontFace: FONT_JP, valign: "middle" });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 27: SECTION DIVIDER - LICATS
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("07", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("自己免疫CAR-T特有の副作用", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 34, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("LICATS — 局所免疫細胞関連毒性症候群", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
})();

// ========================================================
// SLIDE 28: LICATS 詳細
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "7. LICATS副作用");
  s.addText("LICATS：自己免疫CAR-T特有の新しい毒性概念", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 24, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // What is LICATS
  addCard(s, 0.4, 1.4, 4.3, 2.3, { accentColor: C.orange });
  s.addText("LICATSとは", {
    x: 0.6, y: 1.5, w: 4, h: 0.35, fontSize: 18, fontFace: FONT_JP, color: C.orange, bold: true
  });
  s.addText("Local Immune Cell-Associated\nToxicity Syndrome", {
    x: 0.6, y: 1.9, w: 4, h: 0.45, fontSize: 14, fontFace: FONT_EN, color: C.blue, italic: true
  });
  s.addText([
    { text: "既存病変臓器に一致した局所反応", options: { bullet: true, breakLine: true, fontSize: 15, bold: true } },
    { text: "77%の患者に発生（39例解析）", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "多くはgrade 1–2", options: { bullet: true, breakLine: true, fontSize: 15 } },
    { text: "中央値10日で出現、11日で軽快", options: { bullet: true, fontSize: 15 } },
  ], { x: 0.6, y: 2.4, w: 3.9, h: 1.2, fontFace: FONT_JP, color: C.darkGray, paraSpaceAfter: 4 });

  // Key numbers
  addCard(s, 5.3, 1.4, 4.3, 2.3, { accentColor: C.blue });
  s.addText("臨床データ（39例）", {
    x: 5.5, y: 1.5, w: 4, h: 0.35, fontSize: 18, fontFace: FONT_JP, color: C.blue, bold: true
  });

  const licNums = [
    { val: "77%", label: "発生率" },
    { val: "Gr1-2", label: "重症度" },
    { val: "10日", label: "出現時期" },
    { val: "11日", label: "軽快まで" },
  ];
  licNums.forEach((n, i) => {
    const xx = 5.5 + (i % 2) * 2.1;
    const yy = 1.95 + Math.floor(i / 2) * 0.85;
    s.addText(n.val, {
      x: xx, y: yy, w: 1.9, h: 0.45, fontSize: 24, fontFace: FONT_EN,
      color: C.blue, bold: true, align: "center"
    });
    s.addText(n.label, {
      x: xx, y: yy + 0.45, w: 1.9, h: 0.25, fontSize: 12, fontFace: FONT_JP,
      color: C.gray, align: "center"
    });
  });

  // Warning for neurology
  addCard(s, 0.4, 3.95, 9.2, 1.3, { fill: C.redBg, accentColor: C.red });
  s.addText("⚠ 神経内科への重要な示唆", {
    x: 0.6, y: 4.05, w: 4, h: 0.35, fontSize: 16, fontFace: FONT_JP, color: C.red, bold: true
  });
  s.addText("神経根、筋、髄膜、脳辺縁系など既存病変部で同様の反応が起これば、再燃との区別が極めて難しい可能性がある。\n病変臓器から免疫細胞が掃除される過程に伴う一過性反応と解釈されているが、神経内科では直接データがまだ乏しい。", {
    x: 0.6, y: 4.4, w: 8.8, h: 0.8, fontSize: 14, fontFace: FONT_JP, color: C.darkGray
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 29: SECTION DIVIDER - 次に来るもの
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("08", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("次に来るもの", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 36, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("同種CAR-TとPrecision CAR-T（CAAR-T）", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
})();

// ========================================================
// SLIDE 30: Allogeneic / CAAR-T
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "8. 次に来るもの");
  s.addText("未来の2つの方向性", {
    x: 0.5, y: 0.6, w: 9, h: 0.6, fontSize: 28, fontFace: FONT_JP,
    color: C.navy, bold: true, margin: 0
  });

  // Left: Allogeneic
  addCard(s, 0.4, 1.4, 4.3, 3.5, { accentColor: C.green, accentSide: "top" });
  s.addText("Allogeneic / Off-the-shelf化", {
    x: 0.6, y: 1.55, w: 4, h: 0.4, fontSize: 18, fontFace: FONT_JP, color: C.green, bold: true
  });
  s.addText("健康ドナー由来のCD19 CAR-T", {
    x: 0.6, y: 2.0, w: 4, h: 0.3, fontSize: 14, fontFace: FONT_JP, color: C.gray
  });
  s.addText([
    { text: "筋炎・強皮症3例：", options: { bold: true, fontSize: 15, color: C.navy, breakLine: true } },
    { text: "2週以内に完全B細胞枯渇\n6か月で深い寛解、重篤AEなし", options: { fontSize: 14, color: C.darkGray, breakLine: true } },
    { text: "\nSLE 3例：", options: { bold: true, fontSize: 15, color: C.navy, breakLine: true } },
    { text: "GvHD・CRS・ICANSなしに寛解", options: { fontSize: 14, color: C.darkGray, breakLine: true } },
    { text: "\n→ 待機時間と物流の壁が大きく下がる", options: { bold: true, fontSize: 15, color: C.green } },
  ], { x: 0.6, y: 2.3, w: 3.9, h: 2.4, fontFace: FONT_JP, paraSpaceAfter: 2 });

  // Right: CAAR-T
  addCard(s, 5.3, 1.4, 4.3, 3.5, { accentColor: C.red, accentSide: "top" });
  s.addText("Precision CAR-T（CAAR-T）", {
    x: 5.5, y: 1.55, w: 4, h: 0.4, fontSize: 18, fontFace: FONT_JP, color: C.red, bold: true
  });
  s.addText("自己抗原特異的B細胞だけを狙う", {
    x: 5.5, y: 2.0, w: 4, h: 0.3, fontSize: 14, fontFace: FONT_JP, color: C.gray
  });
  s.addText([
    { text: "MuSK-CAAR-T：", options: { bold: true, fontSize: 15, color: C.navy, breakLine: true } },
    { text: "抗MuSK B細胞を選択的に除去\n総B細胞・IgGは保持", options: { fontSize: 14, color: C.darkGray, breakLine: true } },
    { text: "\nNMDAR-CAAR-T：", options: { bold: true, fontSize: 15, color: C.navy, breakLine: true } },
    { text: "抗NMDAR B細胞を選択的に除去", options: { fontSize: 14, color: C.darkGray, breakLine: true } },
    { text: "\n→ 正常免疫を温存しながら\n犯人クローンだけを消す", options: { bold: true, fontSize: 15, color: C.red } },
  ], { x: 5.5, y: 2.3, w: 3.9, h: 2.4, fontFace: FONT_JP, paraSpaceAfter: 2 });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 31: SECTION DIVIDER - まとめ
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  sectionDividerStyle(s);
  s.addText("09", {
    x: 0.8, y: 1.0, w: 2, h: 1.0, fontSize: 60, fontFace: FONT_EN,
    color: "7B8FA8", bold: true, margin: 0
  });
  s.addText("まとめ", {
    x: 0.8, y: 2.0, w: 8, h: 1.0, fontSize: 36, fontFace: FONT_JP,
    color: C.white, bold: true, margin: 0
  });
  s.addText("「免疫抑制」から「免疫リセット」への転換", {
    x: 0.8, y: 3.0, w: 8, h: 0.5, fontSize: 18, fontFace: FONT_JP,
    color: "CADCFC", margin: 0
  });
})();

// ========================================================
// SLIDE 32: まとめ - Key Message
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "9. まとめ");

  // Big key message
  addCard(s, 0.4, 0.7, 9.2, 1.0, { fill: C.lightBlue, accentColor: C.blue });
  s.addText("自己免疫CAR-Tを一言で言うと：「免疫抑制」から「免疫リセット」への転換", {
    x: 0.6, y: 0.75, w: 8.8, h: 0.9, fontSize: 20, fontFace: FONT_JP, color: C.navy, bold: true, valign: "middle"
  });

  // Disease summary
  const diseases = [
    { name: "MG", desc: "最前線。「やさしいBCMA mRNA型」と「深いCD19/BCMA型」が並走", color: C.blue },
    { name: "NMOSD", desc: "生物学的に非常に相性がよい", color: C.green },
    { name: "MS", desc: "CNSコンパートメント攻略が核心", color: C.orange },
    { name: "CIDP", desc: "合うサブタイプを見つける時代に", color: C.navy },
    { name: "GAD/脳炎", desc: "合う患者を見つける時代に入りつつある", color: C.red },
  ];

  diseases.forEach((d, i) => {
    const yy = 1.95 + i * 0.55;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: yy, w: 1.5, h: 0.45, fill: { color: d.color }
    });
    s.addText(d.name, {
      x: 0.5, y: yy, w: 1.5, h: 0.45, fontSize: 16, fontFace: FONT_JP,
      color: C.white, bold: true, align: "center", valign: "middle"
    });
    s.addText(d.desc, {
      x: 2.2, y: yy, w: 7.3, h: 0.45, fontSize: 16, fontFace: FONT_JP,
      color: C.darkGray, valign: "middle"
    });
  });

  // Final message
  addCard(s, 0.4, 4.7, 9.2, 0.7, { fill: C.navy });
  s.addText("CAR-Tはrituximabの強化版ではなく、病的B細胞系の深部除去という別の治療概念", {
    x: 0.6, y: 4.73, w: 8.8, h: 0.65, fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle"
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 33: References 1
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "参考文献");
  s.addText("References (1/2)", {
    x: 0.5, y: 0.6, w: 9, h: 0.5, fontSize: 24, fontFace: FONT_EN,
    color: C.navy, bold: true, margin: 0
  });

  const refs1 = [
    "[1] Schett G, et al. Nat Rev Rheumatol. 2024;20(9):531-544.",
    "[2] Mackensen A, et al. Nat Med. 2022;28(10):2124-2132.",
    "[3] Müller F, et al. N Engl J Med. 2024;390(8):687-700.",
    "[4] Müller F, et al. Nat Med. 2026 Jan 7. CASTLE trial.",
    "[5] Tur C, et al. Ann Rheum Dis. 2025;84(1):106-114.",
    "[6] Granit V, et al. Lancet Neurol. 2023;22(7):578-590.",
    "[7] Chahin N, et al. Ann Clin Transl Neurol. 2025;12(11):2358-2366.",
    "[8] Vu T, et al. Nat Med. 2026 Jan 9. Phase 2b trial.",
    "[9] Zhang Y, et al. EClinicalMedicine. 2025;90:103621.",
    "[10] Qin C, et al. Signal Transduct Target Ther. 2023;8(1):5.",
  ];

  refs1.forEach((r, i) => {
    s.addText(r, {
      x: 0.5, y: 1.2 + i * 0.4, w: 9, h: 0.35, fontSize: 12, fontFace: FONT_EN,
      color: C.darkGray
    });
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 34: References 2
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  contentSlideStyle(s);
  addSectionLabel(s, "参考文献");
  s.addText("References (2/2)", {
    x: 0.5, y: 0.6, w: 9, h: 0.5, fontSize: 24, fontFace: FONT_EN,
    color: C.navy, bold: true, margin: 0
  });

  const refs2 = [
    "[11] Fischbach F, et al. Med. 2024;5(6):550-558.e2.",
    "[12] Dong MH, et al. Med. 2025;6(9):100704.",
    "[13] Faissner S, et al. Proc Natl Acad Sci USA. 2024;121(26):e2403227121.",
    "[14] Hegelmaier T, et al. Med. 2025;6(9):100776.",
    "[15] Vaisvilas M, et al. Front Immunol. 2026;17:1755797.",
    "[16] Oh S, et al. Nat Biotechnol. 2023;41(9):1229-1238.",
    "[17] Momsen Reincke S, et al. Cell. 2023;186(23):5084-5097.e18.",
    "[18] Hagen M, et al. Lancet Rheumatol. 2025;7(6):e424-e433.",
    "[19] Wang X, et al. Cell. 2024;187(18):4890-4904.e9.",
    "[20] Wang D, et al. Med. 2025;6(10):100749.",
  ];

  refs2.forEach((r, i) => {
    s.addText(r, {
      x: 0.5, y: 1.2 + i * 0.4, w: 9, h: 0.35, fontSize: 12, fontFace: FONT_EN,
      color: C.darkGray
    });
  });

  addFooter(s, slideNum, TOTAL);
})();

// ========================================================
// SLIDE 35: END SLIDE
// ========================================================
(() => {
  slideNum++;
  let s = pres.addSlide();
  titleSlideStyle(s);
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.08, fill: { color: C.blue }
  });
  s.addText("ご視聴ありがとうございました", {
    x: 0.8, y: 1.5, w: 8.4, h: 1.0, fontSize: 36, fontFace: FONT_JP,
    color: C.white, bold: true, align: "center"
  });
  s.addText("自己免疫疾患CAR-Tの全体像", {
    x: 0.8, y: 2.6, w: 8.4, h: 0.5, fontSize: 20, fontFace: FONT_JP,
    color: "CADCFC", align: "center"
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 3, y: 3.3, w: 4, h: 0.03, fill: { color: C.blue }
  });
  s.addText("医知創造ラボ", {
    x: 0.8, y: 3.6, w: 8.4, h: 0.5, fontSize: 20, fontFace: FONT_JP,
    color: "8899BB", align: "center"
  });
  s.addText("チャンネル登録・高評価よろしくお願いします", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.4, fontSize: 16, fontFace: FONT_JP,
    color: "667799", align: "center"
  });
})();

// Update TOTAL
// We've created 35 slides total, update footer constant
// (Already set in advance - adjust if needed)

// ========================================================
// SAVE
// ========================================================
const outPath = "C:/Users/jsber/OneDrive/Documents/Claude_task/autoimmune_cart_slides.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("PPTX saved to: " + outPath);
  console.log("Total slides: " + slideNum);
}).catch(err => {
  console.error("Error:", err);
});
