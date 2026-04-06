var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";

var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "E8850C",
  white: "FFFFFF",
  red: "DC3545",
};

var FJ = "BIZ UDPGothic";

(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };

  // 上下アクセントライン
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });

  // 左の赤アクセントバー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.8, w: 0.15, h: 4.0, fill: { color: C.red } });

  // メインタイトル
  s.addText("片側性舞踏病", {
    x: 0.7, y: 0.8, w: 8.8, h: 1.5,
    fontSize: 54, fontFace: FJ, color: C.white, bold: true, margin: 0,
  });
  s.addText("の鑑別", {
    x: 0.7, y: 2.1, w: 8.8, h: 1.0,
    fontSize: 48, fontFace: FJ, color: C.accent, bold: true, margin: 0,
  });

  // サブタイトル
  s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 3.3, w: 8.5, h: 0.03, fill: { color: "B0BEC5" } });
  s.addText("脳梗塞 ・ 高血糖 ・ 薬剤性", {
    x: 0.7, y: 3.5, w: 8.8, h: 0.8,
    fontSize: 32, fontFace: FJ, color: C.white, bold: true, margin: 0,
  });
  s.addText("を見逃さない！", {
    x: 0.7, y: 4.2, w: 8.8, h: 0.7,
    fontSize: 30, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });

  // チャンネル名
  s.addText("医知創造ラボ", {
    x: 6.5, y: 5.0, w: 3.2, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "B0BEC5", align: "right", margin: 0,
  });
})();

var outPath = __dirname + "/thumbnail.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
