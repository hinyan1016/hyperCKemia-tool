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
  orange: "E8850C",
};

var s = pres.addSlide();
s.background = { color: C.dark };

// Top accent bar
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
// Bottom accent bar
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });

// Badge
s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.4, w: 4.5, h: 0.65, fill: { color: C.accent } });
s.addText("脳神経内科医が解説", {
  x: 0.3, y: 0.4, w: 4.5, h: 0.65,
  fontSize: 22, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// Main title
s.addText("理想的な病状説明", {
  x: 0.3, y: 1.3, w: 9.4, h: 1.2,
  fontSize: 52, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", margin: 0,
});
s.addText("「伝わる」技術", {
  x: 0.3, y: 2.4, w: 9.4, h: 1.0,
  fontSize: 42, fontFace: "BIZ UDPGothic", color: C.yellow, bold: true, align: "center", margin: 0,
});

// Separator
s.addShape(pres.shapes.LINE, { x: 2, y: 3.55, w: 6, h: 0, line: { color: C.accent, width: 2 } });

// Key points
var points = [
  { text: "SPIKES", color: C.primary },
  { text: "3つの条件", color: C.green },
  { text: "セリフ例付き", color: C.orange },
];

points.forEach(function(p, i) {
  var xPos = 0.5 + i * 3.1;
  s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.8, w: 2.8, h: 0.65, fill: { color: p.color } });
  s.addText(p.text, {
    x: xPos, y: 3.8, w: 2.8, h: 0.65,
    fontSize: 20, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
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
