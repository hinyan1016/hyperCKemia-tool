var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "Tolosa-Hunt症候群 サムネイル";

var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "4A90D9",
  light: "E8F0FE",
  white: "FFFFFF",
  red: "DC3545",
  orange: "E8850C",
  yellow: "F5C518",
  green: "28A745",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";

var s = pres.addSlide();
s.background = { color: C.dark };

// 上下のアクセントバー
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });

// 警告バッジ
s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.35, w: 3.8, h: 0.7, fill: { color: C.red }, rectRadius: 0.1 });
s.addText("激しい眼窩痛＋複視", {
  x: 0.3, y: 0.35, w: 3.8, h: 0.7,
  fontSize: 22, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// メインタイトル
s.addText("Tolosa-Hunt", {
  x: 0.5, y: 1.3, w: 9, h: 1.0,
  fontSize: 52, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
});
s.addText("症候群", {
  x: 0.5, y: 2.2, w: 9, h: 0.9,
  fontSize: 48, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
});

// サブタイトル
s.addText("― 有痛性眼筋麻痺の鑑別と治療 ―", {
  x: 0.5, y: 3.1, w: 9, h: 0.6,
  fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
});

// キーポイントボックス
var points = [
  { text: "除外診断", color: C.red },
  { text: "ICHD-3基準", color: C.primary },
  { text: "40-50%再発", color: C.orange },
];
points.forEach(function(p, i) {
  var xPos = 1.0 + i * 2.8;
  s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.9, w: 2.5, h: 0.7, fill: { color: p.color }, rectRadius: 0.1 });
  s.addText(p.text, {
    x: xPos, y: 3.9, w: 2.5, h: 0.7,
    fontSize: 20, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
});

// チャンネル名
s.addText("医知創造ラボ", {
  x: 0.5, y: 4.9, w: 9, h: 0.5,
  fontSize: 18, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
});

var outPath = __dirname + "/thumbnail_slide.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
