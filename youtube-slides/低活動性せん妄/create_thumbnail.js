var pptxgen = require("pptxgenjs");
var pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

var FJ = "BIZ UDPGothic";
var FE = "Segoe UI";

var s = pres.addSlide();

// ── Background: dark navy
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "0D1B2A" } });

// ── Left accent bar (orange gradient feel)
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.25, h: 5.625, fill: { color: "E8850C" } });

// ── Top-right warning badge
s.addShape(pres.shapes.RECTANGLE, {
  x: 5.8, y: 0.15, w: 4.0, h: 0.85,
  fill: { color: "DC3545" },
  shadow: { type: "outer", blur: 6, offset: 3, angle: 135, color: "000000", opacity: 0.3 }
});
s.addText("見逃し厳禁", {
  x: 5.8, y: 0.15, w: 4.0, h: 0.85,
  fontSize: 32, fontFace: FJ, color: "FFFFFF", bold: true,
  align: "center", valign: "middle"
});

// ── Main title area
// "静かな" in accent color
s.addText("「静かな」", {
  x: 0.6, y: 0.9, w: 9.0, h: 1.0,
  fontSize: 36, fontFace: FJ, color: "E8850C", bold: true,
  align: "left", valign: "bottom"
});

// "せん妄" huge
s.addText("せん妄", {
  x: 0.6, y: 1.7, w: 9.0, h: 1.6,
  fontSize: 80, fontFace: FJ, color: "FFFFFF", bold: true,
  align: "left", valign: "middle"
});

// Subtitle line
s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.35, w: 5.5, h: 0.06, fill: { color: "E8850C" } });

s.addText("低活動性せん妄を見逃さない", {
  x: 0.6, y: 3.5, w: 6.0, h: 0.65,
  fontSize: 26, fontFace: FJ, color: "AABBD0", bold: false,
  align: "left", valign: "middle"
});

// ── HYPO-D cards at bottom
var letters = ["H", "Y", "P", "O", "D"];
var colors = ["2C5AA0", "7B2D8E", "DC3545", "E8850C", "28A745"];
var labels = ["過眠", "注意力↓", "無関心", "食欲↓", "昼夜逆転"];

for (var i = 0; i < letters.length; i++) {
  var kx = 0.6 + i * 1.85;
  // Card
  s.addShape(pres.shapes.RECTANGLE, {
    x: kx, y: 4.2, w: 1.6, h: 1.15,
    fill: { color: colors[i] },
    shadow: { type: "outer", blur: 5, offset: 2, angle: 135, color: "000000", opacity: 0.25 }
  });
  // Letter
  s.addText(letters[i], {
    x: kx, y: 4.15, w: 1.6, h: 0.8,
    fontSize: 40, fontFace: FE, color: "FFFFFF", bold: true,
    align: "center", valign: "middle"
  });
  // Label
  s.addText(labels[i], {
    x: kx, y: 4.85, w: 1.6, h: 0.45,
    fontSize: 14, fontFace: FJ, color: "FFFFFF", bold: false,
    align: "center", valign: "middle"
  });
}

// ── Right side: key stat
s.addShape(pres.shapes.RECTANGLE, {
  x: 6.8, y: 1.4, w: 2.8, h: 2.2,
  fill: { color: "1B3A5C" },
  shadow: { type: "outer", blur: 6, offset: 3, angle: 135, color: "000000", opacity: 0.25 },
  line: { color: "E8850C", width: 2 }
});
s.addText("全体の", {
  x: 6.8, y: 1.45, w: 2.8, h: 0.5,
  fontSize: 18, fontFace: FJ, color: "AABBD0",
  align: "center", valign: "bottom"
});
s.addText("50%", {
  x: 6.8, y: 1.9, w: 2.8, h: 1.1,
  fontSize: 60, fontFace: FE, color: "E8850C", bold: true,
  align: "center", valign: "middle"
});
s.addText("が見逃される", {
  x: 6.8, y: 2.95, w: 2.8, h: 0.5,
  fontSize: 18, fontFace: FJ, color: "AABBD0",
  align: "center", valign: "top"
});

var outPath = __dirname + "/thumbnail.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Thumbnail PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
