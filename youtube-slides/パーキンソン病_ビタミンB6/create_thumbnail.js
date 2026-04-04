var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";

var C = {
  dark: "0D1B2A",
  primary: "1B3A5C",
  accent: "4A90D9",
  red: "DC3545",
  white: "FFFFFF",
  light: "E8F0FE",
};

var FONT_JP = "BIZ UDPGothic";

(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };

  // Top accent line
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.red } });

  // Warning badge
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.4, w: 3.2, h: 0.55, fill: { color: C.red } });
  s.addText("\u26A0 2026\u5E74 FDA\u8B66\u544A", {
    x: 0.5, y: 0.4, w: 3.2, h: 0.55,
    fontSize: 20, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  // Main title
  s.addText("\u30D1\u30FC\u30AD\u30F3\u30BD\u30F3\u75C5\u3068", {
    x: 0.5, y: 1.3, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_JP, color: C.white, bold: true, margin: 0,
  });
  s.addText("\u30D3\u30BF\u30DF\u30F3B6", {
    x: 0.5, y: 2.0, w: 9, h: 1.2,
    fontSize: 56, fontFace: FONT_JP, color: C.accent, bold: true, margin: 0,
  });

  // Subtitle
  s.addText("\u300C\u907F\u3051\u308B\u300D\u304B\u3089\u300C\u8A55\u4FA1\u3057\u3066\u88DC\u3046\u300D\u3078", {
    x: 0.5, y: 3.3, w: 9, h: 0.7,
    fontSize: 26, fontFace: FONT_JP, color: C.light, bold: true, margin: 0,
  });

  // Key points
  s.addShape(pres.shapes.LINE, { x: 0.5, y: 4.2, w: 9, h: 0, line: { color: C.accent, width: 1.5 } });

  s.addText("\u767A\u4F5C\u30FB\u672B\u68A2\u795E\u7D4C\u969C\u5BB3\u30FB\u8CA7\u8840 \u2015 B6\u6B20\u4E4F\u304C\u539F\u56E0\u304B\u3082", {
    x: 0.5, y: 4.4, w: 9, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });

  s.addText("\u533B\u77E5\u5275\u9020\u30E9\u30DC", {
    x: 7.0, y: 5.0, w: 2.5, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "right", margin: 0,
  });

  // Bottom accent line
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.red } });
})();

var outPath = __dirname + "/thumbnail.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Saved: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
