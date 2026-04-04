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
  green: "28A745",
  yellow: "F5C518",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";

(function() {
  var s = pres.addSlide();

  // 背景: 濃紺
  s.background = { color: C.dark };

  // 上下アクセントライン
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.accent } });

  // 左側カラーバー
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: 0.12, h: 5.465, fill: { color: C.red } });

  // メインタイトル
  s.addText("小脳梗塞", {
    x: 0.5, y: 0.3, w: 9, h: 1.3,
    fontSize: 64, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addText("リハビリテーション", {
    x: 0.5, y: 1.4, w: 9, h: 1.0,
    fontSize: 52, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", margin: 0,
  });

  // 区切り線
  s.addShape(pres.shapes.LINE, { x: 1.5, y: 2.6, w: 7, h: 0, line: { color: C.accent, width: 3 } });

  // サブタイトル
  s.addText("回復はいつまで続く？最新エビデンス解説", {
    x: 0.5, y: 2.8, w: 9, h: 0.7,
    fontSize: 28, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  // キーワードバッジ
  var badges = [
    { text: "KOSCO", color: C.green },
    { text: "NIBS", color: C.accent },
    { text: "DBS", color: C.red },
    { text: "VR", color: C.yellow },
  ];

  badges.forEach(function(b, i) {
    var xPos = 1.5 + i * 1.9;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 3.7, w: 1.6, h: 0.55,
      fill: { color: b.color }, rectRadius: 0.15,
    });
    s.addText(b.text, {
      x: xPos, y: 3.7, w: 1.6, h: 0.55,
      fontSize: 20, fontFace: FONT_EN, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // フッター
  s.addText("医知創造ラボ", {
    x: 0.5, y: 4.6, w: 9, h: 0.5,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  // 右下にチャンネルロゴ的テキスト
  s.addText("脳神経内科医が解説", {
    x: 6.5, y: 5.0, w: 3, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: "636E72", align: "right", margin: 0,
  });
})();

var outPath = __dirname + "/thumbnail_slide.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Thumbnail PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
