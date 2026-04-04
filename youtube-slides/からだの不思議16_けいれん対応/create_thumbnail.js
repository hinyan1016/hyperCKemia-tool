var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";

var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "4A90D9",
  light: "E8F0FE",
  white: "FFFFFF",
  red: "DC3545",
  yellow: "F5C518",
  green: "28A745",
};

var s = pres.addSlide();
s.background = { color: C.dark };

// Top accent bar
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
// Bottom accent bar
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });

// Series badge (top-left)
s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.4, w: 4.2, h: 0.65, fill: { color: C.accent } });
s.addText("からだの不思議 #16", {
  x: 0.3, y: 0.4, w: 4.2, h: 0.65,
  fontSize: 22, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// Main title line 1
s.addText("けいれん目撃！", {
  x: 0.3, y: 1.25, w: 9.4, h: 1.15,
  fontSize: 54, fontFace: "BIZ UDPGothic", color: C.yellow, bold: true, align: "center", margin: 0,
});

// Main title line 2
s.addText("正しい対応は？", {
  x: 0.3, y: 2.3, w: 9.4, h: 1.0,
  fontSize: 48, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", margin: 0,
});

// Separator
s.addShape(pres.shapes.LINE, { x: 2, y: 3.45, w: 6, h: 0, line: { color: C.accent, width: 2 } });

// 3 key points
var points = [
  { text: "正しい対応7項目", color: C.green },
  { text: "絶対NGの行動", color: C.red },
  { text: "5分ルールとは", color: C.yellow },
];

points.forEach(function(p, i) {
  var xPos = 0.5 + i * 3.1;
  s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.65, w: 2.8, h: 0.65, fill: { color: p.color } });
  s.addText(p.text, {
    x: xPos, y: 3.65, w: 2.8, h: 0.65,
    fontSize: 17, fontFace: "BIZ UDPGothic", color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });
});

// Channel name
s.addText("医知創造ラボ", {
  x: 0.3, y: 4.85, w: 9.4, h: 0.5,
  fontSize: 18, fontFace: "BIZ UDPGothic", color: C.accent, align: "center", margin: 0,
});

var outPath = __dirname + "/thumbnail_slide.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
