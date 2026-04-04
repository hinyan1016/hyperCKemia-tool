var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";

var FONT_JP = "BIZ UDPGothic";

(function() {
  var s = pres.addSlide();
  s.background = { color: "0D1B2A" };

  // Top accent bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: "DC3545" } });

  // Patch icon area
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 0.5, w: 3.2, h: 5.0, fill: { color: "1B3A5C" }, rectRadius: 0.1 });
  s.addText("\uD83E\uDE79", { x: 0.3, y: 1.0, w: 3.2, h: 2.0, fontSize: 80, align: "center", margin: 0 });
  s.addText("\u8CBC\u4ED8\u5264", { x: 0.3, y: 2.8, w: 3.2, h: 0.8, fontSize: 28, fontFace: FONT_JP, color: "4A90D9", bold: true, align: "center", margin: 0 });
  s.addText("\u30A2\u30EA\u30C9\u30CD\u30D1\u30C3\u30C1", { x: 0.3, y: 3.5, w: 3.2, h: 0.6, fontSize: 20, fontFace: FONT_JP, color: "E8F0FE", align: "center", margin: 0 });

  // Main text area
  s.addText("\u98DF\u6B32\u306F", {
    x: 3.8, y: 0.4, w: 6.0, h: 1.0,
    fontSize: 40, fontFace: FONT_JP, color: "FFFFFF", bold: true, margin: 0,
  });
  s.addText("\u6539\u5584\u3059\u308B\u306E\u304B\uFF1F", {
    x: 3.8, y: 1.2, w: 6.0, h: 1.2,
    fontSize: 52, fontFace: FONT_JP, color: "DC3545", bold: true, margin: 0,
  });

  // Verdict area
  s.addShape(pres.shapes.RECTANGLE, { x: 3.8, y: 2.8, w: 5.8, h: 1.2, fill: { color: "DC3545" }, rectRadius: 0.08 });
  s.addText("\u30A8\u30D3\u30C7\u30F3\u30B9\u3092\n\u6279\u5224\u7684\u306B\u5410\u5473\u3057\u307E\u3059", {
    x: 4.0, y: 2.9, w: 5.4, h: 1.0,
    fontSize: 28, fontFace: FONT_JP, color: "FFFFFF", bold: true, align: "center", margin: 0, lineSpacingMultiple: 1.2,
  });

  // Sub points
  s.addText("\u2022 \u30C9\u30CD\u30DA\u30B8\u30EB\u306F\u305D\u3082\u305D\u3082\u98DF\u6B32\u3092\u4E0B\u3052\u308B\u85AC", {
    x: 4.0, y: 4.2, w: 5.5, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: "E8F0FE", margin: 0,
  });
  s.addText("\u2022 \u8CBC\u4ED8\u5264\u3067\u3082\u6D88\u5316\u5668\u526F\u4F5C\u7528\u306F\u6B8B\u308B", {
    x: 4.0, y: 4.6, w: 5.5, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: "E8F0FE", margin: 0,
  });
  s.addText("\u2022 \u5B9F\u81E8\u5E8A\u3067\u306E\u5224\u65AD\u57FA\u6E96\u3092\u89E3\u8AAC", {
    x: 4.0, y: 5.0, w: 5.5, h: 0.4,
    fontSize: 18, fontFace: FONT_JP, color: "E8F0FE", margin: 0,
  });

  // Channel name
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: "DC3545" } });
})();

var outPath = __dirname + "/thumbnail.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Saved: " + outPath);
});
