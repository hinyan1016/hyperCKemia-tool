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

// Red warning badge (top-left)
s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.4, w: 3.8, h: 0.7, fill: { color: C.red } });
s.addText("\u26A1 その顔面痛、本当に歯？", {
  x: 0.3, y: 0.4, w: 3.8, h: 0.7,
  fontSize: 22, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// Main title
s.addText("\u4E09\u53C9\u795E\u7D4C\u75DB", {
  x: 0.3, y: 1.3, w: 9.4, h: 1.5,
  fontSize: 56, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", margin: 0,
});
s.addText("\u9451\u5225\u3068\u6CBB\u7642\u6226\u7565", {
  x: 0.3, y: 2.6, w: 9.4, h: 0.8,
  fontSize: 36, fontFace: "BIZ UDPGothic", color: C.accent, bold: true, align: "center", margin: 0,
});

// Subtitle line
s.addShape(pres.shapes.LINE, { x: 2, y: 3.55, w: 6, h: 0, line: { color: C.accent, width: 2 } });

// Key points
var points = [
  { text: "ICHD-3\u5206\u985E\u89E3\u8AAC", color: C.yellow },
  { text: "SUNCT/SUNA\u9451\u5225", color: C.green },
  { text: "CBZ\u306ENNT 1.7\u306E\u5A01\u529B", color: C.red },
];

points.forEach(function(p, i) {
  var xPos = 0.5 + i * 3.1;
  s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.8, w: 2.8, h: 0.65, fill: { color: p.color } });
  s.addText(p.text, {
    x: xPos, y: 3.8, w: 2.8, h: 0.65,
    fontSize: 16, fontFace: "BIZ UDPGothic", color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });
});

// Bottom: channel name
s.addText("\u533B\u77E5\u5275\u9020\u30E9\u30DC", {
  x: 0.3, y: 4.9, w: 9.4, h: 0.5,
  fontSize: 18, fontFace: "BIZ UDPGothic", color: C.accent, align: "center", margin: 0,
});

var outPath = __dirname + "/thumbnail_slide.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
