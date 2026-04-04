var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";

var FJ = "BIZ UDPGothic";

(function() {
  var s = pres.addSlide();
  s.background = { color: "0D1B2A" };

  // 上部アクセントライン
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: "DC3545" } });

  // 左サイドバー
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: "E8850C" } });

  // 背景装飾（ウイルスイメージの斜め帯）
  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 0.3, w: 5, h: 0.9, fill: { color: "1B3A5C" }, rectRadius: 0.1, rotate: -8 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 1.4, w: 4.5, h: 0.7, fill: { color: "1A3550" }, rectRadius: 0.1, rotate: -8 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 2.4, w: 4, h: 0.5, fill: { color: "162D45" }, rectRadius: 0.1, rotate: -8 });

  // 警告マーク風の装飾（右上）
  s.addShape(pres.shapes.RECTANGLE, { x: 8.2, y: 0.3, w: 1.5, h: 1.5, fill: { color: "DC3545" }, rectRadius: 0.15, rotate: 5 });
  s.addText("⚠", { x: 8.2, y: 0.3, w: 1.5, h: 1.5, fontSize: 50, align: "center", valign: "middle", color: "FFFFFF", rotate: 5 });

  // メインタイトル
  s.addText("感染症の", { x: 0.5, y: 0.4, w: 7, h: 1.0, fontSize: 50, fontFace: FJ, color: "FFFFFF", bold: true, margin: 0 });
  s.addText("最新動向 2025", { x: 0.5, y: 1.3, w: 8, h: 1.2, fontSize: 56, fontFace: FJ, color: "E8850C", bold: true, margin: 0 });

  // サブタイトル背景
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.7, w: 7.5, h: 0.7, fill: { color: "DC3545" }, rectRadius: 0.05 });
  s.addText("医療従事者が今知るべき8つの脅威", { x: 0.7, y: 2.7, w: 7.1, h: 0.7, fontSize: 28, fontFace: FJ, color: "FFFFFF", bold: true, valign: "middle", margin: 0 });

  // 4つのキーワード（上段）
  var keywords1 = [
    { text: "H5N1", color: "DC3545" },
    { text: "AMR", color: "2C5AA0" },
    { text: "梅毒", color: "8E44AD" },
    { text: "マイコプラズマ", color: "E67E22" },
  ];
  keywords1.forEach(function(kw, i) {
    var xPos = 0.5 + i * 2.2;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.7, w: 2.0, h: 0.55, fill: { color: kw.color }, rectRadius: 0.08 });
    s.addText(kw.text, { x: xPos, y: 3.7, w: 2.0, h: 0.55, fontSize: 20, fontFace: FJ, color: "FFFFFF", bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // 4つのキーワード（下段）
  var keywords2 = [
    { text: "COVID変異株", color: "1A6B3C" },
    { text: "麻疹", color: "C0392B" },
    { text: "mpox", color: "2980B9" },
    { text: "iGAS", color: "D35400" },
  ];
  keywords2.forEach(function(kw, i) {
    var xPos = 0.5 + i * 2.2;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 4.4, w: 2.0, h: 0.55, fill: { color: kw.color }, rectRadius: 0.08 });
    s.addText(kw.text, { x: xPos, y: 4.4, w: 2.0, h: 0.55, fontSize: 20, fontFace: FJ, color: "FFFFFF", bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // 下部テキスト
  s.addText("Immunity Debt × グローバル化", { x: 0.5, y: 5.05, w: 6, h: 0.45, fontSize: 18, fontFace: FJ, color: "78909C", bold: true, margin: 0 });

  // 下部アクセントライン
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: "DC3545" } });
})();

var fileName = "thumbnail.pptx";
pres.writeFile({ fileName: fileName }).then(function() {
  console.log("Created: " + fileName);
});
