var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";

var FJ = "BIZ UDPGothic";

(function() {
  var s = pres.addSlide();
  s.background = { color: "0D1B2A" };

  // 上部アクセントライン
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: "E8850C" } });

  // 左サイドバー
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: "C0392B" } });

  // DNAイメージ背景（装飾的な斜め帯）
  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 0.5, w: 5, h: 0.8, fill: { color: "1B3A5C" }, rectRadius: 0.1, rotate: -10 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: 1.5, w: 4.5, h: 0.6, fill: { color: "1A3550" }, rectRadius: 0.1, rotate: -10 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 2.5, w: 4, h: 0.5, fill: { color: "162D45" }, rectRadius: 0.1, rotate: -10 });

  // メインタイトル
  s.addText("日本人は", { x: 0.5, y: 0.4, w: 6, h: 1.0, fontSize: 48, fontFace: FJ, color: "FFFFFF", bold: true, margin: 0 });
  s.addText("どこから来たのか？", { x: 0.5, y: 1.3, w: 7, h: 1.2, fontSize: 52, fontFace: FJ, color: "E8850C", bold: true, margin: 0 });

  // サブタイトル背景
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.8, w: 6.5, h: 0.7, fill: { color: "C0392B" }, rectRadius: 0.05 });
  s.addText("最新ゲノム解析で判明！", { x: 0.7, y: 2.8, w: 6.1, h: 0.7, fontSize: 30, fontFace: FJ, color: "FFFFFF", bold: true, valign: "middle", margin: 0 });

  // 3つのキーワード
  var keywords = [
    { text: "縄文人", color: "C0392B" },
    { text: "弥生人", color: "2980B9" },
    { text: "古墳人", color: "27AE60" },
  ];
  keywords.forEach(function(kw, i) {
    var xPos = 0.5 + i * 2.3;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.8, w: 2.1, h: 0.65, fill: { color: kw.color }, rectRadius: 0.08 });
    s.addText(kw.text, { x: xPos, y: 3.8, w: 2.1, h: 0.65, fontSize: 26, fontFace: FJ, color: "FFFFFF", bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // 三重構造モデルテキスト
  s.addText("三重構造モデル", { x: 0.5, y: 4.7, w: 5, h: 0.6, fontSize: 28, fontFace: FJ, color: "78909C", bold: true, margin: 0 });

  // 下部アクセントライン
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: "E8850C" } });
})();

var fileName = "thumbnail.pptx";
pres.writeFile({ fileName: fileName }).then(function() {
  console.log("Created: " + fileName);
});
