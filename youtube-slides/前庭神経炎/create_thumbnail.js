var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "前庭神経炎サムネイル";

var FONT_H = "Arial Black";
var FONT_B = "Arial";

var s = pres.addSlide();

// === 背景: ダークネイビーグラデーション ===
s.background = { color: "0A1628" };

// === 上部アクセントバー ===
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: "DC3545" } });

// === 左サイド 赤い縦バー（危険を強調） ===
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: "DC3545" } });

// === 大きな背景装飾円（右上） ===
s.addShape(pres.shapes.OVAL, { x: 6.5, y: -1.5, w: 5, h: 5, fill: { color: "122040" } });

// === 大きな背景装飾円（左下） ===
s.addShape(pres.shapes.OVAL, { x: -2, y: 3, w: 4.5, h: 4.5, fill: { color: "0E1A30" } });

// === 警告アイコン装飾（右上） ===
s.addText("⚠️", {
  x: 8.2, y: 0.3, w: 1.5, h: 1.5,
  fontSize: 72, fontFace: FONT_B, align: "center", valign: "middle", margin: 0,
  transparency: 70,
});

// === メインタイトルエリア ===
// 赤い強調ライン
s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.8, w: 0.25, h: 2.0, fill: { color: "DC3545" } });

// メインタイトル「前庭神経炎」
s.addText("前庭神経炎", {
  x: 1.1, y: 0.5, w: 7.5, h: 1.5,
  fontSize: 66, fontFace: FONT_H, color: "FFFFFF", bold: true, margin: 0,
  shadow: { type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.5 },
});

// サブタイトル「の鑑別診断」
s.addText("の鑑別診断", {
  x: 1.1, y: 1.8, w: 7.5, h: 0.9,
  fontSize: 44, fontFace: FONT_H, color: "4A90D9", bold: true, margin: 0,
  shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.4 },
});

// === 赤い警告バナー ===
s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.0, w: 9, h: 1.1,
  fill: { color: "DC3545" },
  shadow: { type: "outer", blur: 10, offset: 4, angle: 135, color: "000000", opacity: 0.4 },
});

// バナー内テキスト
s.addText("🚨 脳卒中を見逃さない！", {
  x: 0.5, y: 3.0, w: 9, h: 1.1,
  fontSize: 46, fontFace: FONT_H, color: "FFFFFF", bold: true, align: "center", valign: "middle", margin: 0,
  shadow: { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.5 },
});

// === 下部キーワードカード群 ===
var keywords = [
  { text: "HINTS", color: "2C5AA0" },
  { text: "眼振パターン", color: "1B6B3A" },
  { text: "鑑別5疾患", color: "8B5A00" },
];
keywords.forEach(function(kw, i) {
  var xPos = 0.8 + i * 3.0;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: xPos, y: 4.3, w: 2.7, h: 0.75,
    fill: { color: kw.color },
    rectRadius: 0.15,
    shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.35 },
  });
  s.addText(kw.text, {
    x: xPos, y: 4.3, w: 2.7, h: 0.75,
    fontSize: 22, fontFace: FONT_H, color: "FFFFFF", bold: true, align: "center", valign: "middle", margin: 0,
  });
});

// === 右下ロゴ ===
s.addText("医知創造ラボ", {
  x: 6.8, y: 5.15, w: 3, h: 0.4,
  fontSize: 16, fontFace: FONT_B, color: "4A90D9", align: "right", margin: 0,
  italic: true,
});

// === 下部アクセントバー ===
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: "4A90D9" } });

// === SAVE ===
var outPath = __dirname + "/thumbnail_slide.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Thumbnail PPTX saved: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
