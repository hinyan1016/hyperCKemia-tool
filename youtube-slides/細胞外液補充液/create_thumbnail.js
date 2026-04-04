var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "サムネイル: 細胞外液補充液の使い分け";

var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "E8850C",
  white: "FFFFFF",
  red: "DC3545",
  green: "28A745",
  yellow: "F5C518",
  lightRed: "F8D7DA",
  lightGreen: "D4EDDA",
  text: "2D3436",
};

var FJ = "BIZ UDPGothic";
var FE = "Segoe UI";

var s = pres.addSlide();

// 背景: ダークネイビー
s.background = { color: C.dark };

// 上部アクセントライン
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

// 左半分: 生食（赤系）
s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.5, w: 4.2, h: 4.7, fill: { color: "1E2D4A" }, rectRadius: 0.15 });
s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.5, w: 4.2, h: 0.08, fill: { color: C.red } });

s.addText("生理食塩水", {
  x: 0.4, y: 0.8, w: 4.0, h: 0.8,
  fontSize: 32, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

s.addText("Cl⁻ 154", {
  x: 0.4, y: 1.7, w: 4.0, h: 0.7,
  fontSize: 40, fontFace: FE, color: C.red, bold: true, align: "center", valign: "middle", margin: 0,
});

s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.6, w: 3.2, h: 0.5, fill: { color: C.red }, rectRadius: 0.05 });
s.addText("高Cl性アシドーシス", {
  x: 0.8, y: 2.6, w: 3.2, h: 0.5,
  fontSize: 20, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// 生食側のポイント
s.addText("K⁺ = 0  緩衝剤なし", {
  x: 0.4, y: 3.3, w: 4.0, h: 0.5,
  fontSize: 18, fontFace: FJ, color: "8899AA", align: "center", valign: "middle", margin: 0,
});

s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 4.0, w: 2.6, h: 0.5, fill: { color: "2A1A1A" }, rectRadius: 0.05 });
s.addText("TBI・高K → 生食", {
  x: 1.0, y: 4.0, w: 2.6, h: 0.5,
  fontSize: 16, fontFace: FJ, color: C.yellow, bold: true, align: "center", valign: "middle", margin: 0,
});

// 右半分: リンゲル液（緑系）
s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 0.5, w: 4.2, h: 4.7, fill: { color: "1E2D4A" }, rectRadius: 0.15 });
s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 0.5, w: 4.2, h: 0.08, fill: { color: C.green } });

s.addText("リンゲル液", {
  x: 5.6, y: 0.8, w: 4.0, h: 0.8,
  fontSize: 32, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

s.addText("Cl⁻ 109", {
  x: 5.6, y: 1.7, w: 4.0, h: 0.7,
  fontSize: 40, fontFace: FE, color: C.green, bold: true, align: "center", valign: "middle", margin: 0,
});

s.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 2.6, w: 3.2, h: 0.5, fill: { color: C.green }, rectRadius: 0.05 });
s.addText("血漿に近い組成", {
  x: 6.0, y: 2.6, w: 3.2, h: 0.5,
  fontSize: 20, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// リンゲル液側のポイント
s.addText("乳酸 / 酢酸 / 重炭酸", {
  x: 5.6, y: 3.3, w: 4.0, h: 0.5,
  fontSize: 18, fontFace: FJ, color: "8899AA", align: "center", valign: "middle", margin: 0,
});

s.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 4.0, w: 2.6, h: 0.5, fill: { color: "1A2A1A" }, rectRadius: 0.05 });
s.addText("敗血症・DKA → 有利", {
  x: 6.2, y: 4.0, w: 2.6, h: 0.5,
  fontSize: 16, fontFace: FJ, color: C.yellow, bold: true, align: "center", valign: "middle", margin: 0,
});

// 中央: VS
s.addShape(pres.shapes.OVAL, { x: 4.15, y: 1.9, w: 1.7, h: 1.7, fill: { color: C.accent } });
s.addText("VS", {
  x: 4.15, y: 1.9, w: 1.7, h: 1.7,
  fontSize: 44, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// 下部: メインフック
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.0, w: 10, h: 0.625, fill: { color: C.accent } });
s.addText("結局どっちを選ぶ？ 最新エビデンスで徹底比較", {
  x: 0.3, y: 5.0, w: 9.4, h: 0.625,
  fontSize: 24, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

pres.writeFile({ fileName: "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/細胞外液補充液/thumbnail.pptx" }).then(function() {
  console.log("Thumbnail PPTX created");
});
