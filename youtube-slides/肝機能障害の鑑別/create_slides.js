var pptxgen = require("pptxgenjs");
var pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "医知創造ラボ";
pres.title = "肝機能障害の鑑別診断";

var C = {
  navy: "1B3A5C", blue: "2C5AA0", ltBlue: "E8F0FA", bg: "F0F4F8",
  white: "FFFFFF", text: "333333", red: "DC3545", yellow: "F5C518",
  green: "28A745", orange: "E8850C", ltGray: "F8F9FA", gray: "555555", subText: "666666",
  purple: "7B2D8E", ltPurple: "F3E8F9", ltRed: "FDE8EA", ltGreen: "E8F5E9",
  ltOrange: "FFF3E0", ltYellow: "FFFDE7", darkRed: "8B1A1A",
};
var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var mkShadow = function() { return { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 }; };

function addFooter(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.15, w: 10, h: 0.475, fill: { color: C.navy } });
  slide.addText("肝機能障害の鑑別診断", { x: 0.5, y: 5.15, w: 6, h: 0.475, fontSize: 18, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
}

function addSectionTitle(slide, title) {
  slide.addText(title, { x: 0.6, y: 0.35, w: 8.8, h: 0.65, fontSize: 28, fontFace: FONT_JP, color: C.navy, bold: true, margin: 0 });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.05, w: 1.2, h: 0.06, fill: { color: C.blue } });
}

// カード型ボックスヘルパー
function addCard(slide, x, y, w, h, opts) {
  var bgColor = opts.bgColor || C.white;
  var borderColor = opts.borderColor || C.blue;
  slide.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: bgColor }, rectRadius: 0.1, shadow: mkShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.06, h: h, fill: { color: borderColor }, rectRadius: 0.03 });
}

// ========================================
// Slide 1: タイトルスライド
// ========================================
var s1 = pres.addSlide();
s1.background = { fill: C.navy };
// 装飾円
s1.addShape(pres.shapes.OVAL, { x: 7.5, y: -1.5, w: 4, h: 4, fill: { color: C.blue, transparency: 70 } });
s1.addShape(pres.shapes.OVAL, { x: 8.5, y: 3.5, w: 3, h: 3, fill: { color: C.blue, transparency: 80 } });
// アクセントバー
s1.addShape(pres.shapes.RECTANGLE, { x: 1.2, y: 1.0, w: 0.08, h: 2.8, fill: { color: C.yellow } });
// タイトル
s1.addText([
  { text: "肝機能障害\n", options: { fontSize: 42, fontFace: FONT_JP, color: C.white, bold: true } },
  { text: "（トランスアミナーゼ上昇）\n", options: { fontSize: 30, fontFace: FONT_JP, color: C.ltBlue } },
  { text: "の鑑別診断", options: { fontSize: 42, fontFace: FONT_JP, color: C.white, bold: true } }
], { x: 1.6, y: 1.0, w: 7, h: 2.8, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
// サブタイトル
s1.addText("― R値 × 7ステップで系統的に迫る ―", { x: 1.6, y: 4.0, w: 7, h: 0.6, fontSize: 22, fontFace: FONT_JP, color: C.yellow, margin: 0 });
// 著者
s1.addText("医知創造ラボ", { x: 1.6, y: 4.8, w: 5, h: 0.4, fontSize: 16, fontFace: FONT_JP, color: C.ltBlue, margin: 0 });

// ========================================
// Slide 2: 「肝機能」の落とし穴
// ========================================
var s2 = pres.addSlide();
s2.background = { fill: C.bg };
addSectionTitle(s2, "「肝機能」の落とし穴");
addFooter(s2);

// 左カラム：障害の指標
addCard(s2, 0.5, 1.3, 4.3, 3.5, { bgColor: C.white, borderColor: C.red });
s2.addText("障害 (Injury) の指標", { x: 0.7, y: 1.4, w: 3.9, h: 0.5, fontSize: 20, fontFace: FONT_JP, color: C.red, bold: true, margin: 0 });
s2.addText([
  { text: "AST (GOT)\n", options: { fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true } },
  { text: "→ 肝以外（骨格筋・心筋）にも存在\n\n", options: { fontSize: 14, fontFace: FONT_JP, color: C.gray } },
  { text: "ALT (GPT)\n", options: { fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true } },
  { text: "→ 比較的肝特異性が高い\n\n", options: { fontSize: 14, fontFace: FONT_JP, color: C.gray } },
  { text: "ALP\n", options: { fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true } },
  { text: "→ 胆管系障害・骨疾患\n\n", options: { fontSize: 14, fontFace: FONT_JP, color: C.gray } },
  { text: "ビリルビン\n", options: { fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true } },
  { text: "→ 代謝・排泄障害の指標", options: { fontSize: 14, fontFace: FONT_JP, color: C.gray } }
], { x: 0.8, y: 2.0, w: 3.8, h: 2.6, valign: "top", margin: 0, lineSpacingMultiple: 1.0 });

// 右カラム：合成能の指標
addCard(s2, 5.2, 1.3, 4.3, 3.5, { bgColor: C.white, borderColor: C.green });
s2.addText("合成能 (Function) の指標", { x: 5.4, y: 1.4, w: 3.9, h: 0.5, fontSize: 20, fontFace: FONT_JP, color: C.green, bold: true, margin: 0 });
s2.addText([
  { text: "PT / INR\n", options: { fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true } },
  { text: "→ 凝固因子合成能\n\n", options: { fontSize: 14, fontFace: FONT_JP, color: C.gray } },
  { text: "アルブミン\n", options: { fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true } },
  { text: "→ 蛋白合成能\n\n", options: { fontSize: 14, fontFace: FONT_JP, color: C.gray } },
  { text: "コリンエステラーゼ\n", options: { fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true } },
  { text: "→ 肝予備能の指標", options: { fontSize: 14, fontFace: FONT_JP, color: C.gray } }
], { x: 5.4, y: 2.0, w: 3.8, h: 2.6, valign: "top", margin: 0, lineSpacingMultiple: 1.0 });

// 注意バナー
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.85, w: 9, h: 0.25, fill: { color: C.ltOrange }, rectRadius: 0.05 });
s2.addText("AST優位の上昇 → CKも確認！（筋障害の除外が重要）", { x: 0.7, y: 4.85, w: 8.6, h: 0.25, fontSize: 13, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0, valign: "middle" });

// ========================================
// Slide 3: R値とパターン分類
// ========================================
var s3 = pres.addSlide();
s3.background = { fill: C.bg };
addSectionTitle(s3, "R値とパターン分類");
addFooter(s3);

// R値の式
s3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 9, h: 0.8, fill: { color: C.ltBlue }, rectRadius: 0.1, shadow: mkShadow() });
s3.addText([
  { text: "R = ", options: { fontSize: 22, fontFace: FONT_EN, color: C.navy, bold: true } },
  { text: "(ALT / ALT ULN)", options: { fontSize: 20, fontFace: FONT_EN, color: C.blue } },
  { text: " ÷ ", options: { fontSize: 22, fontFace: FONT_EN, color: C.navy, bold: true } },
  { text: "(ALP / ALP ULN)", options: { fontSize: 20, fontFace: FONT_EN, color: C.blue } }
], { x: 0.8, y: 1.3, w: 8.4, h: 0.8, valign: "middle", align: "center", margin: 0 });

// 3つのパターンカード
// 肝細胞障害型
addCard(s3, 0.5, 2.4, 2.8, 2.4, { bgColor: C.ltRed, borderColor: C.red });
s3.addText("R ≥ 5", { x: 0.7, y: 2.5, w: 2.4, h: 0.5, fontSize: 26, fontFace: FONT_EN, color: C.red, bold: true, align: "center", margin: 0 });
s3.addText("肝細胞障害型", { x: 0.7, y: 3.0, w: 2.4, h: 0.45, fontSize: 20, fontFace: FONT_JP, color: C.text, bold: true, align: "center", margin: 0 });
s3.addText("ウイルス性肝炎\n虚血性肝炎\nアセトアミノフェン中毒\n自己免疫性肝炎", { x: 0.7, y: 3.5, w: 2.4, h: 1.2, fontSize: 13, fontFace: FONT_JP, color: C.gray, align: "center", margin: 0, lineSpacingMultiple: 1.2 });

// 混合型
addCard(s3, 3.6, 2.4, 2.8, 2.4, { bgColor: C.ltPurple, borderColor: C.purple });
s3.addText("2 < R < 5", { x: 3.8, y: 2.5, w: 2.4, h: 0.5, fontSize: 26, fontFace: FONT_EN, color: C.purple, bold: true, align: "center", margin: 0 });
s3.addText("混合型", { x: 3.8, y: 3.0, w: 2.4, h: 0.45, fontSize: 20, fontFace: FONT_JP, color: C.text, bold: true, align: "center", margin: 0 });
s3.addText("薬剤性肝障害(DILI)\nEBV/CMV肝炎\nうっ血肝", { x: 3.8, y: 3.5, w: 2.4, h: 1.2, fontSize: 13, fontFace: FONT_JP, color: C.gray, align: "center", margin: 0, lineSpacingMultiple: 1.2 });

// 胆汁うっ滞型
addCard(s3, 6.7, 2.4, 2.8, 2.4, { bgColor: C.ltBlue, borderColor: C.blue });
s3.addText("R ≤ 2", { x: 6.9, y: 2.5, w: 2.4, h: 0.5, fontSize: 26, fontFace: FONT_EN, color: C.blue, bold: true, align: "center", margin: 0 });
s3.addText("胆汁うっ滞型", { x: 6.9, y: 3.0, w: 2.4, h: 0.45, fontSize: 20, fontFace: FONT_JP, color: C.text, bold: true, align: "center", margin: 0 });
s3.addText("総胆管結石\n胆管炎\nPBC / PSC\nIgG4関連硬化性胆管炎", { x: 6.9, y: 3.5, w: 2.4, h: 1.2, fontSize: 13, fontFace: FONT_JP, color: C.gray, align: "center", margin: 0, lineSpacingMultiple: 1.2 });

// ========================================
// Slide 4: 肝細胞障害型の鑑別（前半 D1-D5）
// ========================================
var s4 = pres.addSlide();
s4.background = { fill: C.bg };
addSectionTitle(s4, "肝細胞障害型の鑑別（前半）");
addFooter(s4);

var d1_5 = [
  { id: "D1", name: "急性ウイルス性肝炎(A/B/E)", key: "ALT優位、曝露歴・旅行歴", color: C.red },
  { id: "D2", name: "慢性B/C型肝炎", key: "無症候性、線維化進行", color: C.orange },
  { id: "D3", name: "NAFLD / NASH", key: "代謝異常、ALT軽度上昇", color: C.green },
  { id: "D4", name: "アルコール関連肝障害", key: "AST/ALT比 ≥ 2", color: C.blue },
  { id: "D5", name: "自己免疫性肝炎(AIH)", key: "女性、IgG↑、ANA/ASMA", color: C.purple },
];
d1_5.forEach(function(d, i) {
  var yPos = 1.3 + i * 0.72;
  addCard(s4, 0.5, yPos, 9, 0.65, { bgColor: C.white, borderColor: d.color });
  s4.addText(d.id, { x: 0.7, y: yPos + 0.05, w: 0.6, h: 0.55, fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", fill: { color: d.color }, rectRadius: 0.08, margin: 0 });
  s4.addText(d.name, { x: 1.4, y: yPos + 0.02, w: 3.5, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
  s4.addText(d.key, { x: 5.0, y: yPos + 0.02, w: 4.3, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.gray, valign: "middle", margin: 0 });
});

// ========================================
// Slide 5: 肝細胞障害型の鑑別（後半 D6-D10）
// ========================================
var s5 = pres.addSlide();
s5.background = { fill: C.bg };
addSectionTitle(s5, "肝細胞障害型の鑑別（後半）");
addFooter(s5);

var d6_10 = [
  { id: "D6", name: "アセトアミノフェン中毒", key: "AST/ALT数千、NAC投与", color: C.darkRed },
  { id: "D7", name: "虚血性肝炎", key: "ショック後、LDH著増", color: C.red },
  { id: "D8", name: "Budd-Chiari症候群", key: "肝静脈閉塞、腹水", color: C.orange },
  { id: "D9", name: "Wilson病", key: "若年、銅代謝異常、KF ring", color: C.green },
  { id: "D10", name: "ヘモクロマトーシス", key: "TSAT ≥ 45%、フェリチン↑", color: C.blue },
];
d6_10.forEach(function(d, i) {
  var yPos = 1.3 + i * 0.72;
  addCard(s5, 0.5, yPos, 9, 0.65, { bgColor: C.white, borderColor: d.color });
  s5.addText(d.id, { x: 0.7, y: yPos + 0.05, w: 0.6, h: 0.55, fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", fill: { color: d.color }, rectRadius: 0.08, margin: 0 });
  s5.addText(d.name, { x: 1.4, y: yPos + 0.02, w: 3.5, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
  s5.addText(d.key, { x: 5.0, y: yPos + 0.02, w: 4.3, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.gray, valign: "middle", margin: 0 });
});

// ========================================
// Slide 6: 胆汁うっ滞型の鑑別（D11-D15）
// ========================================
var s6 = pres.addSlide();
s6.background = { fill: C.bg };
addSectionTitle(s6, "胆汁うっ滞型の鑑別");
addFooter(s6);

var d11_15 = [
  { id: "D11", name: "総胆管結石・胆道閉塞", key: "胆管拡張、黄疸、腹痛", color: C.blue },
  { id: "D12", name: "急性胆管炎", key: "3徴/5徴、緊急ERCP適応", color: C.darkRed },
  { id: "D13", name: "原発性胆汁性胆管炎(PBC)", key: "AMA陽性、中年女性", color: C.purple },
  { id: "D14", name: "原発性硬化性胆管炎(PSC)", key: "IBD合併、MRCP多発狭窄", color: C.green },
  { id: "D15", name: "IgG4関連硬化性胆管炎", key: "IgG4高値、ステロイド反応性", color: C.orange },
];
d11_15.forEach(function(d, i) {
  var yPos = 1.3 + i * 0.72;
  addCard(s6, 0.5, yPos, 9, 0.65, { bgColor: C.white, borderColor: d.color });
  s6.addText(d.id, { x: 0.7, y: yPos + 0.05, w: 0.6, h: 0.55, fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", fill: { color: d.color }, rectRadius: 0.08, margin: 0 });
  s6.addText(d.name, { x: 1.4, y: yPos + 0.02, w: 3.8, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
  s6.addText(d.key, { x: 5.3, y: yPos + 0.02, w: 4.0, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.gray, valign: "middle", margin: 0 });
});

// ========================================
// Slide 7: 混合型・浸潤性・その他（D16-D20）
// ========================================
var s7 = pres.addSlide();
s7.background = { fill: C.bg };
addSectionTitle(s7, "混合型・浸潤性・その他");
addFooter(s7);

var d16_20 = [
  { id: "D16", name: "薬剤性肝障害(DILI)", key: "除外診断、RUCAMスコア", color: C.purple },
  { id: "D17", name: "EBV / CMV肝炎", key: "異型リンパ球、肝脾腫", color: C.orange },
  { id: "D18", name: "うっ血肝", key: "右心不全、肝静脈拡張", color: C.blue },
  { id: "D19", name: "悪性腫瘍・浸潤性疾患", key: "画像異常、ALP優位上昇", color: C.darkRed },
  { id: "D20", name: "筋障害（肝外AST上昇）", key: "CK高値がkey、ALT正常〜軽度↑", color: C.green },
];
d16_20.forEach(function(d, i) {
  var yPos = 1.3 + i * 0.72;
  addCard(s7, 0.5, yPos, 9, 0.65, { bgColor: C.white, borderColor: d.color });
  s7.addText(d.id, { x: 0.7, y: yPos + 0.05, w: 0.6, h: 0.55, fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", fill: { color: d.color }, rectRadius: 0.08, margin: 0 });
  s7.addText(d.name, { x: 1.4, y: yPos + 0.02, w: 3.8, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
  s7.addText(d.key, { x: 5.3, y: yPos + 0.02, w: 4.0, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.gray, valign: "middle", margin: 0 });
});

// ========================================
// Slide 8: ALT>1000の3本柱
// ========================================
var s8 = pres.addSlide();
s8.background = { fill: C.bg };
addSectionTitle(s8, "ALT > 1000 の3本柱");
addFooter(s8);

var pillars = [
  { emoji: "⚡", name: "虚血性肝炎\n（ショック肝）", desc: "ショック・低灌流\nLDH著増が特徴", color: C.red, bgColor: C.ltRed },
  { emoji: "💊", name: "アセトアミノフェン\n中毒", desc: "過量服薬歴\nNAC投与で救命", color: C.orange, bgColor: C.ltOrange },
  { emoji: "🦠", name: "急性ウイルス性\n肝炎", desc: "A/B/E型\n曝露歴・旅行歴", color: C.green, bgColor: C.ltGreen },
];
pillars.forEach(function(p, i) {
  var xPos = 0.5 + i * 3.15;
  addCard(s8, xPos, 1.3, 2.9, 3.0, { bgColor: p.bgColor, borderColor: p.color });
  s8.addText(p.emoji, { x: xPos, y: 1.4, w: 2.9, h: 0.6, fontSize: 36, align: "center", valign: "middle", margin: 0 });
  s8.addText(p.name, { x: xPos + 0.15, y: 2.1, w: 2.6, h: 0.8, fontSize: 17, fontFace: FONT_JP, color: p.color, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  s8.addText(p.desc, { x: xPos + 0.15, y: 2.95, w: 2.6, h: 0.9, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
});

// 注意バナー
s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 9, h: 0.5, fill: { color: C.ltYellow }, rectRadius: 0.08, line: { color: C.yellow, width: 1.5 } });
s8.addText("⚠ 胆道疾患でも一過性に ALT > 1000 となることあり（胆石嵌頓など）", { x: 0.8, y: 4.5, w: 8.4, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });

// ========================================
// Slide 9: 7ステップアルゴリズム概要
// ========================================
var s9 = pres.addSlide();
s9.background = { fill: C.bg };
addSectionTitle(s9, "7ステップアルゴリズム概要");
addFooter(s9);

var steps = [
  { num: "1", text: "パターン分類（R値）", color: C.red },
  { num: "2", text: "Red flagの確認", color: C.darkRed },
  { num: "3", text: "急性 vs 慢性", color: C.orange },
  { num: "4", text: "上昇の程度", color: C.yellow },
  { num: "5", text: "原因に直結する問診", color: C.green },
  { num: "6", text: "スクリーニング検査", color: C.blue },
  { num: "7", text: "画像所見", color: C.purple },
];
steps.forEach(function(st, i) {
  var yPos = 1.25 + i * 0.52;
  // 番号バッジ
  s9.addText(st.num, { x: 0.8, y: yPos, w: 0.5, h: 0.45, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", fill: { color: st.color }, rectRadius: 0.23, margin: 0 });
  // テキスト
  s9.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: yPos, w: 7.8, h: 0.45, fill: { color: C.white }, rectRadius: 0.08, shadow: mkShadow() });
  s9.addText(st.text, { x: 1.7, y: yPos, w: 7.4, h: 0.45, fontSize: 18, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  // 矢印（最後以外）
  if (i < 6) {
    s9.addText("▼", { x: 0.8, y: yPos + 0.42, w: 0.5, h: 0.13, fontSize: 8, color: C.gray, align: "center", margin: 0 });
  }
});

// ========================================
// Slide 10: Step 1-2 パターン分類とRed flag
// ========================================
var s10 = pres.addSlide();
s10.background = { fill: C.bg };
addSectionTitle(s10, "Step 1-2: パターン分類と Red Flag");
addFooter(s10);

// 左カラム: パターン分類
addCard(s10, 0.5, 1.3, 4.3, 3.6, { bgColor: C.white, borderColor: C.blue });
s10.addText("Step 1: パターン分類", { x: 0.7, y: 1.4, w: 3.9, h: 0.45, fontSize: 18, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0 });
// R値ミニカード
var rPatterns = [
  { label: "R ≥ 5  肝細胞障害型", color: C.red, bgColor: C.ltRed },
  { label: "2 < R < 5  混合型", color: C.purple, bgColor: C.ltPurple },
  { label: "R ≤ 2  胆汁うっ滞型", color: C.blue, bgColor: C.ltBlue },
];
rPatterns.forEach(function(rp, i) {
  var yP = 2.0 + i * 0.85;
  s10.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yP, w: 3.7, h: 0.7, fill: { color: rp.bgColor }, rectRadius: 0.08, line: { color: rp.color, width: 1.5 } });
  s10.addText(rp.label, { x: 1.0, y: yP, w: 3.3, h: 0.7, fontSize: 16, fontFace: FONT_JP, color: rp.color, bold: true, valign: "middle", margin: 0 });
});

// 右カラム: Red Flags
addCard(s10, 5.2, 1.3, 4.3, 3.6, { bgColor: C.white, borderColor: C.darkRed });
s10.addText("Step 2: Red Flag", { x: 5.4, y: 1.4, w: 3.9, h: 0.45, fontSize: 18, fontFace: FONT_JP, color: C.darkRed, bold: true, margin: 0 });
var redFlags = [
  "🚨 意識障害（肝性脳症）",
  "🚨 PT-INR ≥ 1.5",
  "🚨 発熱 + 黄疸（胆管炎）",
  "🚨 過量服薬",
  "🚨 ショック / 低灌流",
];
redFlags.forEach(function(rf, i) {
  s10.addText(rf, { x: 5.5, y: 2.0 + i * 0.5, w: 3.8, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
});

// ========================================
// Slide 11: Step 3-4 急性/慢性と重症度
// ========================================
var s11 = pres.addSlide();
s11.background = { fill: C.bg };
addSectionTitle(s11, "Step 3-4: 急性/慢性と重症度");
addFooter(s11);

// 左カラム: 経過別
addCard(s11, 0.5, 1.3, 4.3, 3.6, { bgColor: C.white, borderColor: C.orange });
s11.addText("Step 3: 経過の評価", { x: 0.7, y: 1.4, w: 3.9, h: 0.45, fontSize: 18, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0 });

var durations = [
  { label: "急性", items: "ウイルス性、虚血、中毒、胆道", color: C.red },
  { label: "亜急性", items: "DILI、AIH、悪性腫瘍", color: C.orange },
  { label: "慢性", items: "慢性肝炎、NAFLD、PBC/PSC", color: C.green },
];
durations.forEach(function(dur, i) {
  var yP = 2.0 + i * 0.95;
  s11.addText(dur.label, { x: 0.8, y: yP, w: 1.2, h: 0.35, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", fill: { color: dur.color }, rectRadius: 0.06, margin: 0 });
  s11.addText("→ " + dur.items, { x: 2.1, y: yP, w: 2.5, h: 0.7, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0 });
});

// 右カラム: 重症度スケール
addCard(s11, 5.2, 1.3, 4.3, 3.6, { bgColor: C.white, borderColor: C.blue });
s11.addText("Step 4: 上昇の程度", { x: 5.4, y: 1.4, w: 3.9, h: 0.45, fontSize: 18, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0 });

var severities = [
  { level: "軽度", range: "< 5×ULN", cause: "慢性疾患群", color: C.green },
  { level: "中等度", range: "5-15×ULN", cause: "急性肝炎、AIH", color: C.yellow },
  { level: "高度", range: "> 15×ULN", cause: "虚血、中毒", color: C.orange },
  { level: "著明", range: "ALT>1000", cause: "3本柱", color: C.red },
];
severities.forEach(function(sv, i) {
  var yP = 2.0 + i * 0.7;
  s11.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: yP, w: 3.9, h: 0.58, fill: { color: C.ltGray }, rectRadius: 0.06 });
  s11.addText(sv.level, { x: 5.5, y: yP + 0.02, w: 0.9, h: 0.26, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", fill: { color: sv.color }, rectRadius: 0.05, margin: 0 });
  s11.addText(sv.range, { x: 6.5, y: yP + 0.02, w: 1.3, h: 0.54, fontSize: 14, fontFace: FONT_EN, color: C.text, bold: true, valign: "middle", margin: 0 });
  s11.addText(sv.cause, { x: 7.8, y: yP + 0.02, w: 1.4, h: 0.54, fontSize: 13, fontFace: FONT_JP, color: C.gray, valign: "middle", margin: 0 });
});

// ========================================
// Slide 12: Step 5-7 問診・検査・画像
// ========================================
var s12 = pres.addSlide();
s12.background = { fill: C.bg };
addSectionTitle(s12, "Step 5-7: 問診・検査・画像");
addFooter(s12);

// 問診
addCard(s12, 0.3, 1.3, 3.0, 3.6, { bgColor: C.white, borderColor: C.green });
s12.addText("Step 5: 問診", { x: 0.5, y: 1.4, w: 2.6, h: 0.4, fontSize: 17, fontFace: FONT_JP, color: C.green, bold: true, margin: 0 });
s12.addText("□ 薬剤歴（OTC含む）\n□ 飲酒量\n□ 代謝リスク因子\n□ 若年 → Wilson病\n□ 血栓リスク\n□ 感染曝露歴\n□ 輸血・注射歴", { x: 0.5, y: 1.9, w: 2.6, h: 2.8, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

// 検査
addCard(s12, 3.5, 1.3, 3.0, 3.6, { bgColor: C.white, borderColor: C.blue });
s12.addText("Step 6: 検査", { x: 3.7, y: 1.4, w: 2.6, h: 0.4, fontSize: 17, fontFace: FONT_JP, color: C.blue, bold: true, margin: 0 });
s12.addText("□ HAV-IgM, HBs抗原\n□ HCV抗体\n□ ANA, ASMA, IgG\n□ 鉄代謝(TSAT,フェリチン)\n□ 銅, セルロプラスミン\n□ IgG4\n□ CK", { x: 3.7, y: 1.9, w: 2.6, h: 2.8, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

// 画像
addCard(s12, 6.7, 1.3, 3.0, 3.6, { bgColor: C.white, borderColor: C.purple });
s12.addText("Step 7: 画像", { x: 6.9, y: 1.4, w: 2.6, h: 0.4, fontSize: 17, fontFace: FONT_JP, color: C.purple, bold: true, margin: 0 });
s12.addText("□ 胆管拡張\n□ 脂肪肝\n□ 肝硬変所見\n□ 肝静脈閉塞\n□ 腫瘤性病変\n□ 門脈血栓\n□ 脾腫", { x: 6.9, y: 1.9, w: 2.6, h: 2.8, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

// ========================================
// Slide 13: 代表的カットオフ値
// ========================================
var s13 = pres.addSlide();
s13.background = { fill: C.bg };
addSectionTitle(s13, "代表的カットオフ値");
addFooter(s13);

var tableRows = [
  ["指標", "基準値 / カットオフ", "臨床的意義"],
  ["ALT > 1000", "> 1000 IU/L", "虚血 / アセトアミノフェン / ウイルス"],
  ["急性肝不全", "INR ≥ 1.5 + 脳症", "移植施設への転送を検討"],
  ["胆管炎(TG18)", "A + B + C 基準", "Charcot 3徴 / Reynolds 5徴"],
  ["Hy's law", "ALT ≥ 3×ULN + T-Bil > 2×ULN", "DILI重症化の予測"],
  ["TSAT", "≥ 45%", "ヘモクロマトーシス"],
  ["FIB-4", "スコア算出", "NAFLD線維化リスク評価"],
  ["IgG4", "≥ 135 mg/dL", "IgG4関連硬化性胆管炎"],
];
s13.addTable(tableRows, {
  x: 0.5, y: 1.3, w: 9, h: 3.5,
  fontSize: 13, fontFace: FONT_JP, color: C.text,
  border: { type: "solid", pt: 0.5, color: "CCCCCC" },
  colW: [2.0, 3.0, 4.0],
  rowH: [0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42],
  autoPage: false,
  headerRow: true,
  align: "left",
  valign: "middle",
  margin: [2, 5, 2, 5],
});

// テーブルヘッダーを手動でスタイリング（上から被せる）
s13.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 9, h: 0.42, fill: { color: C.navy } });
s13.addText("指標", { x: 0.5, y: 1.3, w: 2.0, h: 0.42, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: [0, 5, 0, 5] });
s13.addText("基準値 / カットオフ", { x: 2.5, y: 1.3, w: 3.0, h: 0.42, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: [0, 5, 0, 5] });
s13.addText("臨床的意義", { x: 5.5, y: 1.3, w: 4.0, h: 0.42, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: [0, 5, 0, 5] });

// ========================================
// Slide 14: Red Flags まとめ
// ========================================
var s14 = pres.addSlide();
s14.background = { fill: C.bg };
addSectionTitle(s14, "Red Flags まとめ");
addFooter(s14);

// 背景に警告感
s14.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.25, w: 9.2, h: 3.7, fill: { color: C.ltRed }, rectRadius: 0.15 });

var redFlagItems = [
  { num: "1", text: "急性肝不全（INR ≥ 1.5 + 脳症）", action: "→ 移植施設相談", color: C.darkRed },
  { num: "2", text: "胆管炎（発熱+黄疸+腹痛）", action: "→ 緊急ドレナージ", color: C.red },
  { num: "3", text: "DILI重症（Hy's law陽性）", action: "→ 被疑薬中止", color: C.orange },
  { num: "4", text: "悪性腫瘍を示唆する所見", action: "→ 画像+組織診断", color: C.purple },
  { num: "5", text: "「肝臓ではない」可能性", action: "→ CK確認（筋障害）", color: C.blue },
];
redFlagItems.forEach(function(rf, i) {
  var yP = 1.4 + i * 0.68;
  // 番号
  s14.addText(rf.num, { x: 0.7, y: yP, w: 0.5, h: 0.5, fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", fill: { color: rf.color }, rectRadius: 0.25, margin: 0 });
  // カード
  s14.addShape(pres.shapes.RECTANGLE, { x: 1.4, y: yP, w: 7.9, h: 0.55, fill: { color: C.white }, rectRadius: 0.08, shadow: mkShadow() });
  s14.addText(rf.text, { x: 1.6, y: yP, w: 4.5, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
  s14.addText(rf.action, { x: 6.2, y: yP, w: 2.9, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: rf.color, bold: true, valign: "middle", margin: 0 });
});

// ========================================
// Slide 15: Take-home Message
// ========================================
var s15 = pres.addSlide();
s15.background = { fill: C.navy };

// 装飾
s15.addShape(pres.shapes.OVAL, { x: -1, y: -1, w: 3, h: 3, fill: { color: C.blue, transparency: 75 } });
s15.addShape(pres.shapes.OVAL, { x: 8, y: 4, w: 3, h: 3, fill: { color: C.blue, transparency: 80 } });

s15.addText("Take-home Message", { x: 0.6, y: 0.3, w: 8.8, h: 0.7, fontSize: 30, fontFace: FONT_EN, color: C.yellow, bold: true, margin: 0 });
s15.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.0, w: 1.5, h: 0.06, fill: { color: C.yellow } });

var messages = [
  { num: "1", text: "R値でパターン分類（肝細胞障害型 / 胆汁うっ滞型 / 混合型）" },
  { num: "2", text: "ALT > 1000 → 3本柱（虚血・アセトアミノフェン・ウイルス）" },
  { num: "3", text: "Red flag → 迅速対応（急性肝不全・胆管炎・Hy's law）" },
  { num: "4", text: "AST優位の上昇 → CK確認（筋障害の除外）" },
  { num: "5", text: "7ステップで系統的に鑑別し、見逃しを防ぐ" },
];
messages.forEach(function(msg, i) {
  var yP = 1.3 + i * 0.75;
  // 番号バッジ
  s15.addText(msg.num, { x: 0.7, y: yP, w: 0.55, h: 0.55, fontSize: 22, fontFace: FONT_EN, color: C.navy, bold: true, align: "center", valign: "middle", fill: { color: C.yellow }, rectRadius: 0.28, margin: 0 });
  // テキストカード
  s15.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: yP, w: 8.0, h: 0.55, fill: { color: C.white, transparency: 85 }, rectRadius: 0.08 });
  s15.addText(msg.text, { x: 1.7, y: yP, w: 7.6, h: 0.55, fontSize: 17, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
});

// フッター
s15.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.15, w: 10, h: 0.475, fill: { color: "0F2640" } });
s15.addText("肝機能障害の鑑別診断  |  医知創造ラボ", { x: 0.5, y: 5.15, w: 9, h: 0.475, fontSize: 16, fontFace: FONT_JP, color: C.ltBlue, valign: "middle", margin: 0 });

// ========================================
// 保存
// ========================================
pres.writeFile({ fileName: "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/肝機能障害の鑑別/肝機能障害の鑑別.pptx" })
  .then(function() { console.log("Done: 肝機能障害の鑑別.pptx を作成しました"); })
  .catch(function(e) { console.error(e); });
