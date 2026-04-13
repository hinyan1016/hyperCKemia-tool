var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";

var C = {
  dark: "0D1B2A",
  primary: "1B3A5C",
  accent: "C0392B",
  accentLight: "E74C3C",
  light: "F0F4F8",
  white: "FFFFFF",
  yellow: "F5C518",
  green: "27AE60",
  orange: "E67E22",
};

var s = pres.addSlide();
s.background = { color: C.dark };

// Top accent bar
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
// Bottom accent bar
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });

// Series badge (top-left)
s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.4, w: 4.0, h: 0.65, fill: { color: C.accent } });
s.addText("金融リスク解説", {
  x: 0.3, y: 0.4, w: 4.0, h: 0.65,
  fontSize: 22, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// Warning badge (top-right)
s.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 0.4, w: 3.7, h: 0.65, fill: { color: C.orange } });
s.addText("第2のリーマン？", {
  x: 6.0, y: 0.4, w: 3.7, h: 0.65,
  fontSize: 20, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
});

// Main title line 1
s.addText("プライベートクレジット", {
  x: 0.3, y: 1.4, w: 9.4, h: 1.1,
  fontSize: 48, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", margin: 0,
});

// Main title line 2 (accent color)
s.addText("問題とは何か？", {
  x: 0.3, y: 2.4, w: 9.4, h: 1.0,
  fontSize: 46, fontFace: "BIZ UDPGothic", color: C.yellow, bold: true, align: "center", margin: 0,
});

// Separator
s.addShape(pres.shapes.LINE, { x: 2, y: 3.55, w: 6, h: 0, line: { color: C.accent, width: 2 } });

// Key points
var points = [
  { text: "3兆ドル市場", color: C.accent },
  { text: "3大構造リスク", color: C.orange },
  { text: "日本への影響", color: C.primary },
];

points.forEach(function(p, i) {
  var xPos = 0.5 + i * 3.1;
  s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.8, w: 2.8, h: 0.65, fill: { color: p.color } });
  s.addText(p.text, {
    x: xPos, y: 3.8, w: 2.8, h: 0.65,
    fontSize: 17, fontFace: "BIZ UDPGothic", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
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
