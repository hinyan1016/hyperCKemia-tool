var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "細胞外液補充液の使い分け｜生食 vs リンゲル液を徹底比較";

// --- Color Palette ---
var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "E8850C",
  light: "E8F4FD",
  warmBg: "F8F9FA",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "DC3545",
  green: "28A745",
  yellow: "F5C518",
  lightRed: "F8D7DA",
  lightYellow: "FFF3CD",
  lightGreen: "D4EDDA",
};

var FJ = "BIZ UDPGothic";
var FE = "Segoe UI";

function shd() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function hdr(s, title, bgColor) {
  bgColor = bgColor || C.primary;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: bgColor } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 26, fontFace: FJ, color: C.white, bold: true, margin: 0 });
}

function card(s, x, y, w, h, fillColor, borderColor) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: fillColor || C.white }, shadow: shd(), rectRadius: 0.1 });
  if (borderColor) {
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.12, h: h, fill: { color: borderColor }, rectRadius: 0.05 });
  }
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("細胞外液補充液の使い分け", { x: 0.6, y: 1.0, w: 8.8, h: 1.2, fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("生食 vs リンゲル液を徹底比較", { x: 0.6, y: 2.8, w: 8.8, h: 0.6, fontSize: 26, fontFace: FJ, color: C.accent, align: "center", margin: 0 });
  s.addText("「とりあえず生食」から卒業する", { x: 0.6, y: 3.8, w: 8.8, h: 0.6, fontSize: 20, fontFace: FJ, color: "B0BEC5", align: "center", italic: true, margin: 0 });
  s.addText("医知創造ラボ", { x: 0.6, y: 4.6, w: 8.8, h: 0.5, fontSize: 18, fontFace: FJ, color: "78909C", align: "center", margin: 0 });
})();

// SLIDE 2: 学習目標
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "本日の学習目標");
  var goals = [
    { num: "1", title: "組成の違いを\n説明できる", desc: "生食・乳酸RL・酢酸RL・\n重炭酸RLの特徴", color: C.primary },
    { num: "2", title: "エビデンスに基づき\n使い分けられる", desc: "SMART/BaSICS/PLUS\n疾患別の選択", color: C.accent },
    { num: "3", title: "Fluid Stewardship\nを実践できる", desc: "種類・量・速度・\nタイミングの管理", color: C.green },
  ];
  goals.forEach(function(g, i) {
    var xPos = 0.4 + i * 3.15;
    card(s, xPos, 1.2, 2.95, 3.6, C.white, g.color);
    s.addShape(pres.shapes.OVAL, { x: xPos + 1.1, y: 1.4, w: 0.75, h: 0.75, fill: { color: g.color } });
    s.addText(g.num, { x: xPos + 1.1, y: 1.4, w: 0.75, h: 0.75, fontSize: 28, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(g.title, { x: xPos + 0.15, y: 2.3, w: 2.65, h: 1.0, fontSize: 20, fontFace: FJ, color: g.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(g.desc, { x: xPos + 0.15, y: 3.4, w: 2.65, h: 1.2, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", valign: "top", margin: 0 });
  });
})();

// SLIDE 3: 体液分布
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "体液分布の基礎 — 60-40-20の法則");
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.2, w: 9.2, h: 4.0, fill: { color: "D6EAF8" }, rectRadius: 0.1 });
  s.addText("体液 ＝ 体重の約 60%", { x: 0.5, y: 1.25, w: 9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.9, w: 5.6, h: 3.1, fill: { color: "AED6F1" }, rectRadius: 0.1 });
  s.addText("細胞内液（ICF）", { x: 0.8, y: 2.0, w: 5.2, h: 0.5, fontSize: 22, fontFace: FJ, color: C.dark, bold: true, margin: 0 });
  s.addText("体重の約 40%（体液の 2/3）", { x: 0.8, y: 2.5, w: 5.2, h: 0.5, fontSize: 18, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("主要電解質：K⁺、リン酸", { x: 0.8, y: 3.0, w: 5.2, h: 0.5, fontSize: 18, fontFace: FJ, color: C.sub, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.4, y: 1.9, w: 3.0, h: 3.1, fill: { color: "5DADE2" }, rectRadius: 0.1 });
  s.addText("細胞外液\n（ECF）", { x: 6.5, y: 2.0, w: 2.8, h: 0.7, fontSize: 22, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText("体重の約 20%\n（体液の 1/3）", { x: 6.5, y: 2.8, w: 2.8, h: 0.6, fontSize: 18, fontFace: FJ, color: C.white, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 3.6, w: 2.8, h: 0.6, fill: { color: "2E86C1" }, rectRadius: 0.05 });
  s.addText("間質液 15%（3/4）", { x: 6.5, y: 3.6, w: 2.8, h: 0.6, fontSize: 16, fontFace: FJ, color: C.white, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 4.3, w: 2.8, h: 0.6, fill: { color: C.dark }, rectRadius: 0.05 });
  s.addText("血漿 5%（1/4）", { x: 6.5, y: 4.3, w: 2.8, h: 0.6, fontSize: 16, fontFace: FJ, color: C.white, align: "center", valign: "middle", margin: 0 });
})();

// SLIDE 4: 1/4ルール
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "輸液はどこに分布する？ — 1/4ルール");
  card(s, 0.4, 1.2, 4.3, 3.2, C.white, C.primary);
  s.addText("細胞外液補充液", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("1,000 mL 投与", { x: 0.6, y: 1.9, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("↓", { x: 0.6, y: 2.4, w: 3.9, h: 0.4, fontSize: 22, fontFace: FE, color: C.sub, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.9, w: 3.7, h: 0.7, fill: { color: C.light }, rectRadius: 0.05 });
  s.addText("血管内に約 250 mL\n（投与量の 1/4）", { x: 0.8, y: 2.9, w: 3.7, h: 0.7, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("残り 750 mL は間質へ", { x: 0.6, y: 3.7, w: 3.9, h: 0.5, fontSize: 18, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  card(s, 5.3, 1.2, 4.3, 3.2, C.white, C.sub);
  s.addText("5% ブドウ糖液", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.sub, bold: true, margin: 0 });
  s.addText("1,000 mL 投与", { x: 5.5, y: 1.9, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("↓", { x: 5.5, y: 2.4, w: 3.9, h: 0.4, fontSize: 22, fontFace: FE, color: C.sub, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.6, y: 2.9, w: 3.7, h: 0.7, fill: { color: C.lightRed }, rectRadius: 0.05 });
  s.addText("血管内に約 83 mL\n（投与量の 1/12）", { x: 5.6, y: 2.9, w: 3.7, h: 0.7, fontSize: 20, fontFace: FJ, color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("全体液に拡散", { x: 5.5, y: 3.7, w: 3.9, h: 0.5, fontSize: 18, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  card(s, 0.5, 4.6, 9, 0.7, C.lightYellow, C.yellow);
  s.addText("出血性ショックで「出血量の3倍の晶質液」と言われる根拠がここにある", { x: 0.7, y: 4.7, w: 8.6, h: 0.5, fontSize: 18, fontFace: FJ, color: C.text, bold: true, valign: "middle", align: "center", margin: 0 });
})();

// SLIDE 5: 分類
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "細胞外液補充液の種類");
  card(s, 0.4, 1.2, 4.3, 1.5, C.white, C.sub);
  s.addText("生理食塩水（0.9% NaCl）", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.sub, bold: true, margin: 0 });
  s.addText("Na⁺ と Cl⁻ のみ。緩衝剤なし", { x: 0.6, y: 1.9, w: 3.9, h: 0.6, fontSize: 18, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 5.3, 1.2, 4.3, 1.5, C.white, C.primary);
  s.addText("バランス液（リンゲル液）", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("血漿に近い電解質組成＋緩衝剤", { x: 5.5, y: 1.9, w: 3.9, h: 0.6, fontSize: 18, fontFace: FJ, color: C.text, margin: 0 });
  var types = [
    { title: "乳酸リンゲル液", sub: "ラクテック\nソルラクト", desc: "肝臓で代謝\n使用実績が最も豊富", color: C.primary },
    { title: "酢酸リンゲル液", sub: "ヴィーンF\nソルアセトF", desc: "全身組織で代謝\n肝不全でも使用可", color: C.accent },
    { title: "重炭酸リンゲル液", sub: "ビカネイト", desc: "代謝不要で直接緩衝\n最も「生理的」", color: C.green },
  ];
  types.forEach(function(t, i) {
    var xPos = 0.4 + i * 3.15;
    card(s, xPos, 3.0, 2.95, 2.3, C.white, t.color);
    s.addText(t.title, { x: xPos + 0.15, y: 3.1, w: 2.65, h: 0.5, fontSize: 18, fontFace: FJ, color: t.color, bold: true, align: "center", margin: 0 });
    s.addText(t.sub, { x: xPos + 0.15, y: 3.6, w: 2.65, h: 0.6, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
    s.addText(t.desc, { x: xPos + 0.15, y: 4.2, w: 2.55, h: 0.8, fontSize: 15, fontFace: FJ, color: C.text, align: "center", margin: 0 });
  });
})();

// SLIDE 6: 組成比較表
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "組成比較表 — 生食の「非生理性」に注目");
  var hO = { fontSize: 14, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center", valign: "middle" };
  var cO = { fontSize: 14, fontFace: FE, color: C.text, align: "center", valign: "middle" };
  var cB = { fontSize: 14, fontFace: FE, color: C.text, bold: true, align: "center", valign: "middle" };
  var cR = { fontSize: 14, fontFace: FE, color: C.red, bold: true, align: "center", valign: "middle", fill: { color: C.lightRed } };
  var cJ = { fontSize: 14, fontFace: FJ, color: C.text, align: "center", valign: "middle" };
  var rows = [
    [{ text: "項目", options: hO },{ text: "血漿", options: hO },{ text: "生食", options: { fontSize: 14, fontFace: FJ, color: C.white, bold: true, fill: { color: C.red }, align: "center", valign: "middle" } },{ text: "乳酸RL", options: hO },{ text: "酢酸RL", options: hO },{ text: "重炭酸RL", options: hO }],
    [{ text: "Na⁺", options: cB },{ text: "140", options: cO },{ text: "154", options: cR },{ text: "130", options: cO },{ text: "130", options: cO },{ text: "130", options: cO }],
    [{ text: "K⁺", options: cB },{ text: "4.5", options: cO },{ text: "0", options: cR },{ text: "4", options: cO },{ text: "4", options: cO },{ text: "4", options: cO }],
    [{ text: "Ca²⁺", options: cB },{ text: "5", options: cO },{ text: "0", options: cR },{ text: "3", options: cO },{ text: "3", options: cO },{ text: "3", options: cO }],
    [{ text: "Cl⁻", options: cB },{ text: "100", options: cO },{ text: "154", options: cR },{ text: "109", options: cO },{ text: "109", options: cO },{ text: "109", options: cO }],
    [{ text: "緩衝剤", options: { fontSize: 13, fontFace: FJ, color: C.text, bold: true, align: "center", valign: "middle" } },{ text: "HCO₃⁻ 24", options: cJ },{ text: "なし", options: cR },{ text: "乳酸 28", options: cJ },{ text: "酢酸 28", options: cJ },{ text: "HCO₃⁻ 28", options: cJ }],
    [{ text: "浸透圧", options: cB },{ text: "285", options: cO },{ text: "308", options: cR },{ text: "273", options: cO },{ text: "280", options: cO },{ text: "285", options: cO }],
    [{ text: "pH", options: cB },{ text: "7.4", options: cO },{ text: "5.5", options: cR },{ text: "6.5", options: cO },{ text: "6.7", options: cO },{ text: "7.4", options: cO }],
  ];
  s.addTable(rows, { x: 0.3, y: 1.1, w: 9.4, colW: [1.3, 1.3, 1.3, 1.6, 1.6, 1.6], rowH: [0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45], border: { pt: 1, color: "DEE2E6" }, autoPage: false });
  card(s, 0.5, 4.8, 9, 0.6, C.lightRed, C.red);
  s.addText("生食の Cl⁻ = 154 は血漿（100）より約 50% 高い → 大量投与で高 Cl 性アシドーシス", { x: 0.7, y: 4.85, w: 8.6, h: 0.5, fontSize: 16, fontFace: FJ, color: C.red, bold: true, valign: "middle", align: "center", margin: 0 });
})();

// SLIDE 7: 高Cl性アシドーシス機序
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "生食の大量投与で何が起こるか？", C.red);
  var flow = [
    { text: "生食を大量投与", y: 1.2, color: C.sub },
    { text: "血中 Cl⁻ 濃度が上昇", y: 1.9, color: C.red },
    { text: "SID（Strong Ion Difference）が低下", y: 2.6, color: C.red },
    { text: "代謝性アシドーシス", y: 3.5, color: C.red },
  ];
  flow.forEach(function(f, i) {
    card(s, 0.5, f.y, 4.0, 0.6, C.white, f.color);
    s.addText(f.text, { x: 0.7, y: f.y + 0.05, w: 3.6, h: 0.5, fontSize: 18, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });
    if (i < 3) {
      s.addShape(pres.shapes.DOWN_ARROW, { x: 2.15, y: f.y + 0.62, w: 0.7, h: 0.25, fill: { color: C.red } });
    }
  });
  card(s, 5.3, 1.2, 4.3, 3.6, C.lightRed, C.red);
  s.addText("腎臓への影響", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText([
    { text: "Cl⁻ 上昇\n", options: { fontSize: 20, fontFace: FJ, color: C.text, bold: true } },
    { text: "↓\n", options: { fontSize: 18, fontFace: FE, color: C.red } },
    { text: "TGF 活性化\n", options: { fontSize: 20, fontFace: FJ, color: C.text, bold: true } },
    { text: "↓\n", options: { fontSize: 18, fontFace: FE, color: C.red } },
    { text: "輸入細動脈 収縮\n", options: { fontSize: 20, fontFace: FJ, color: C.text, bold: true } },
    { text: "↓\n", options: { fontSize: 18, fontFace: FE, color: C.red } },
    { text: "GFR 低下 → AKI ↑", options: { fontSize: 20, fontFace: FJ, color: C.red, bold: true } },
  ], { x: 5.5, y: 1.9, w: 3.9, h: 2.8, valign: "top", align: "center", margin: 0 });
})();

// SLIDE 8: SMART試験
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "SMART 試験（2018）— 最初の大規模試験");
  card(s, 0.4, 1.2, 9.2, 1.2, C.white, C.primary);
  s.addText([
    { text: "単施設・クラスターランダム化（Vanderbilt ICU） ", options: { fontSize: 18, fontFace: FJ, color: C.text } },
    { text: "N = 15,802", options: { fontSize: 20, fontFace: FE, color: C.primary, bold: true } },
  ], { x: 0.6, y: 1.3, w: 8.8, h: 1.0, valign: "middle", margin: 0 });
  card(s, 0.4, 2.7, 4.3, 2.0, C.light, C.primary);
  s.addText("バランス液群", { x: 0.6, y: 2.8, w: 3.9, h: 0.4, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("MAKE30", { x: 0.6, y: 3.2, w: 3.9, h: 0.4, fontSize: 18, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText("14.3%", { x: 0.6, y: 3.6, w: 3.9, h: 0.8, fontSize: 48, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  card(s, 5.3, 2.7, 4.3, 2.0, C.white, C.sub);
  s.addText("生食群", { x: 5.5, y: 2.8, w: 3.9, h: 0.4, fontSize: 20, fontFace: FJ, color: C.sub, bold: true, margin: 0 });
  s.addText("MAKE30", { x: 5.5, y: 3.2, w: 3.9, h: 0.4, fontSize: 18, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText("15.4%", { x: 5.5, y: 3.6, w: 3.9, h: 0.8, fontSize: 48, fontFace: FE, color: C.sub, bold: true, align: "center", valign: "middle", margin: 0 });
  card(s, 0.5, 4.9, 9, 0.6, C.lightGreen, C.green);
  s.addText("P = 0.04  バランス液が有意に優位", { x: 0.7, y: 4.95, w: 8.6, h: 0.5, fontSize: 18, fontFace: FJ, color: C.text, bold: true, valign: "middle", align: "center", margin: 0 });
})();

// SLIDE 9: BaSICS・PLUS
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "BaSICS・PLUS — 全体集団では有意差なし");
  card(s, 0.4, 1.2, 4.3, 2.5, C.white, C.primary);
  s.addText("BaSICS（2021）", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FE, color: C.primary, bold: true, margin: 0 });
  s.addText("75施設  N = 11,052", { x: 0.6, y: 1.8, w: 3.9, h: 0.4, fontSize: 16, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText([
    { text: "90日死亡率\n", options: { fontSize: 18, fontFace: FJ, color: C.sub } },
    { text: "26.4% vs 27.2%\n", options: { fontSize: 28, fontFace: FE, color: C.text, bold: true } },
    { text: "P = 0.47", options: { fontSize: 18, fontFace: FE, color: C.sub } },
  ], { x: 0.6, y: 2.3, w: 3.9, h: 1.2, valign: "middle", align: "center", margin: 0 });
  card(s, 5.3, 1.2, 4.3, 2.5, C.white, C.primary);
  s.addText("PLUS（2022）", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FE, color: C.primary, bold: true, margin: 0 });
  s.addText("53施設 二重盲検  N = 5,037", { x: 5.5, y: 1.8, w: 3.9, h: 0.4, fontSize: 16, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText([
    { text: "90日死亡率\n", options: { fontSize: 18, fontFace: FJ, color: C.sub } },
    { text: "21.8% vs 22.0%\n", options: { fontSize: 28, fontFace: FE, color: C.text, bold: true } },
    { text: "P = 0.90", options: { fontSize: 18, fontFace: FE, color: C.sub } },
  ], { x: 5.5, y: 2.3, w: 3.9, h: 1.2, valign: "middle", align: "center", margin: 0 });
  card(s, 0.4, 4.0, 9.2, 1.3, C.light, C.primary);
  s.addText("BEST-Living IPD メタアナリシス（2023）", { x: 0.6, y: 4.1, w: 8.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "6 RCT・34,685 名 → 有益性の事後確率 89.5%（有意差なし）\n", options: { fontSize: 16, fontFace: FJ, color: C.text } },
    { text: "サブグループでは明確な違いあり → 次スライド", options: { fontSize: 16, fontFace: FJ, color: C.accent, bold: true } },
  ], { x: 0.6, y: 4.5, w: 8.8, h: 0.7, valign: "middle", margin: 0 });
})();

// SLIDE 10: バランス液有利
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "バランス液が有利な病態", C.green);
  var items = [
    { title: "敗血症", desc: "BEST-Living: 有益性の事後確率 89.5%\nSSC 2021: バランス液を弱く推奨", color: C.primary },
    { title: "DKA", desc: "RINSE-DKA: HAGMA 解消が有意に早い\n（adjusted HR 1.325, P < 0.001）", color: C.accent },
    { title: "急性膵炎", desc: "乳酸リンゲル液で中等度〜重度への\n進展リスク低下傾向", color: C.green },
  ];
  items.forEach(function(it, i) {
    var yPos = 1.1 + i * 1.5;
    card(s, 0.4, yPos, 9.2, 1.3, C.white, it.color);
    s.addText(it.title, { x: 0.6, y: yPos + 0.1, w: 2.5, h: 1.1, fontSize: 24, fontFace: FJ, color: it.color, bold: true, valign: "middle", margin: 0 });
    s.addText(it.desc, { x: 3.2, y: yPos + 0.1, w: 6.2, h: 1.1, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
})();

// SLIDE 11: 生食を選ぶべき
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "生食を選ぶべき病態", C.red);
  card(s, 0.4, 1.2, 9.2, 2.0, C.lightRed, C.red);
  s.addText("外傷性脳損傷（TBI）", { x: 0.6, y: 1.3, w: 3.5, h: 0.5, fontSize: 24, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText([
    { text: "BEST-Living: harm 確率 ", options: { fontSize: 20, fontFace: FJ, color: C.text } },
    { text: "97.5%\n", options: { fontSize: 24, fontFace: FE, color: C.red, bold: true } },
    { text: "乳酸 RL 浸透圧 273 mOsm/L → 低張 → 脳浮腫悪化リスク\n", options: { fontSize: 18, fontFace: FJ, color: C.text } },
    { text: "ESICM 2024: TBI には生食を条件付きで推奨", options: { fontSize: 18, fontFace: FJ, color: C.red, bold: true } },
  ], { x: 0.6, y: 1.9, w: 8.8, h: 1.2, valign: "top", margin: 0 });
  card(s, 0.4, 3.5, 9.2, 1.5, C.lightYellow, C.yellow);
  s.addText("高カリウム血症", { x: 0.6, y: 3.6, w: 3.5, h: 0.5, fontSize: 24, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText([
    { text: "バランス液は K⁺ 4 mEq/L を含有 → 重度高 K（> 6.0）では避ける\n", options: { fontSize: 20, fontFace: FJ, color: C.text } },
    { text: "生食は K⁺ を含まないため第一選択", options: { fontSize: 20, fontFace: FJ, color: C.text, bold: true } },
  ], { x: 0.6, y: 4.1, w: 8.8, h: 0.8, valign: "top", margin: 0 });
})();

// SLIDE 12: 緩衝剤
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "緩衝剤の違い — 乳酸 vs 酢酸 vs 重炭酸");
  var bufs = [
    { title: "乳酸", sub: "ラクテック等", lines: [
        { label: "代謝", val: "肝臓" },
        { label: "速度", val: "中等度" },
        { label: "利点", val: "使用実績が最も豊富" },
        { label: "注意", val: "肝不全時に注意" },
      ], color: C.primary },
    { title: "酢酸", sub: "ヴィーンF等", lines: [
        { label: "代謝", val: "全身の組織" },
        { label: "速度", val: "やや速い" },
        { label: "利点", val: "肝不全でも使用可" },
        { label: "注意", val: "急速投与で低血圧" },
      ], color: C.accent },
    { title: "重炭酸", sub: "ビカネイト", lines: [
        { label: "代謝", val: "不要（直接緩衝）" },
        { label: "速度", val: "即時" },
        { label: "利点", val: "最も生理的" },
        { label: "注意", val: "隔壁バッグが必要" },
      ], color: C.green },
  ];
  bufs.forEach(function(b, i) {
    var xPos = 0.4 + i * 3.15;
    card(s, xPos, 1.2, 2.95, 3.2, C.white, b.color);
    s.addText(b.title, { x: xPos + 0.15, y: 1.3, w: 2.65, h: 0.6, fontSize: 24, fontFace: FJ, color: b.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(b.sub, { x: xPos + 0.15, y: 1.9, w: 2.65, h: 0.4, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
    b.lines.forEach(function(ln, j) {
      var ly = 2.4 + j * 0.45;
      s.addText(ln.label, { x: xPos + 0.2, y: ly, w: 0.8, h: 0.4, fontSize: 14, fontFace: FJ, color: b.color, bold: true, valign: "middle", margin: 0 });
      s.addText(ln.val, { x: xPos + 1.0, y: ly, w: 1.75, h: 0.4, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    });
  });
  card(s, 0.5, 4.6, 9, 0.7, C.light, C.primary);
  s.addText("メタアナリシスでは酢酸 RL と乳酸 RL の間で死亡率・AKI に有意差なし", { x: 0.7, y: 4.7, w: 8.6, h: 0.5, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", align: "center", margin: 0 });
})();

// SLIDE 13: 肝不全
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "肝不全時、乳酸リンゲル液は本当にダメ？");
  card(s, 0.4, 1.2, 4.3, 2.5, C.white, C.primary);
  s.addText("輸液由来の乳酸", { x: 0.6, y: 1.3, w: 3.9, h: 0.4, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("28 mEq/L", { x: 0.6, y: 1.8, w: 3.9, h: 1.0, fontSize: 48, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("（乳酸 RL 1L あたり）", { x: 0.6, y: 2.8, w: 3.9, h: 0.5, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  card(s, 5.3, 1.2, 4.3, 2.5, C.white, C.accent);
  s.addText("内因性の乳酸産生", { x: 5.5, y: 1.3, w: 3.9, h: 0.4, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("1,000〜1,500\nmEq/時", { x: 5.5, y: 1.8, w: 3.9, h: 1.0, fontSize: 36, fontFace: FE, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("（体内で毎時産生）", { x: 5.5, y: 2.8, w: 3.9, h: 0.5, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  s.addText("vs", { x: 4.5, y: 2.0, w: 1.0, h: 0.8, fontSize: 28, fontFace: FE, color: C.sub, bold: true, align: "center", valign: "middle", margin: 0 });
  card(s, 0.4, 4.0, 9.2, 1.3, C.lightGreen, C.green);
  s.addText([
    { text: "臨床的に問題になることは稀。", options: { fontSize: 20, fontFace: FJ, color: C.green, bold: true } },
    { text: "輸液由来の乳酸はごくわずか\n", options: { fontSize: 20, fontFace: FJ, color: C.text } },
    { text: "不安なら → 酢酸 RL or 重炭酸 RL に切り替え", options: { fontSize: 18, fontFace: FJ, color: C.text } },
  ], { x: 0.6, y: 4.1, w: 8.8, h: 1.1, valign: "middle", margin: 0 });
})();

// SLIDE 14-17: 疾患別選択（出血/敗血症/DKA/周術期）
// SLIDE 14: 出血性ショック
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "疾患別選択 — 出血性ショック");
  card(s, 0.4, 1.2, 9.2, 1.5, C.white, C.red);
  s.addText("第一選択：バランス液（細胞外液補充液）", { x: 0.6, y: 1.3, w: 8.8, h: 0.5, fontSize: 24, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("出血量 × 3 の晶質液が目安（1/4 ルール）", { x: 0.6, y: 1.9, w: 8.8, h: 0.5, fontSize: 20, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 0.4, 3.0, 9.2, 2.3, C.lightYellow, C.yellow);
  s.addText("⚠ Damage Control Resuscitation", { x: 0.6, y: 3.1, w: 8.8, h: 0.5, fontSize: 22, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText([
    { text: "・大量晶質液 → ", options: { fontSize: 20, fontFace: FJ, color: C.text } },
    { text: "希釈性凝固障害", options: { fontSize: 20, fontFace: FJ, color: C.red, bold: true } },
    { text: "・間質浮腫のリスク\n", options: { fontSize: 20, fontFace: FJ, color: C.text } },
    { text: "・晶質液を控えめにし ", options: { fontSize: 20, fontFace: FJ, color: C.text } },
    { text: "早期に輸血へ移行\n", options: { fontSize: 20, fontFace: FJ, color: C.red, bold: true } },
    { text: "・推定 1,500 mL 以上の出血では MTP 発動を検討", options: { fontSize: 20, fontFace: FJ, color: C.text } },
  ], { x: 0.6, y: 3.7, w: 8.8, h: 1.4, valign: "top", margin: 0 });
})();

// SLIDE 15: 敗血症
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "疾患別選択 — 敗血症");
  var pts = [
    { text: "SSC 2021：バランス液を弱く推奨", b: true },
    { text: "初期蘇生：30 mL/kg 以内（最初の 3 時間）", b: false },
    { text: "輸液反応性の評価（PLR・SVV）に基づき追加", b: false },
    { text: "早期のノルアドレナリン併用が推奨傾向", b: true },
    { text: "アルブミンは大量晶質液投与後に考慮", b: false },
  ];
  pts.forEach(function(p, i) {
    var yPos = 1.2 + i * 0.85;
    card(s, 0.4, yPos, 9.2, 0.7, C.white, (i === 0 || i === 3) ? C.primary : C.sub);
    s.addText(p.text, { x: 0.7, y: yPos + 0.05, w: 8.7, h: 0.6, fontSize: 20, fontFace: FJ, color: p.b ? C.primary : C.text, bold: p.b, valign: "middle", margin: 0 });
  });
})();

// SLIDE 16: DKA
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "疾患別選択 — DKA");
  card(s, 0.4, 1.2, 9.2, 1.3, C.lightGreen, C.green);
  s.addText("第一選択：バランス液（乳酸リンゲル液）", { x: 0.6, y: 1.3, w: 8.8, h: 0.5, fontSize: 24, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("生食よりも病態生理学的に合理的", { x: 0.6, y: 1.9, w: 8.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 0.4, 2.8, 4.3, 2.5, C.white, C.primary);
  s.addText("RINSE-DKA", { x: 0.6, y: 2.9, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "HAGMA 解消\n", options: { fontSize: 18, fontFace: FJ, color: C.sub } },
    { text: "HR 1.325\n", options: { fontSize: 36, fontFace: FE, color: C.primary, bold: true } },
    { text: "(95%CI 1.121-1.566)\nP < 0.001", options: { fontSize: 16, fontFace: FE, color: C.sub } },
  ], { x: 0.6, y: 3.5, w: 3.9, h: 1.6, valign: "middle", align: "center", margin: 0 });
  card(s, 5.3, 2.8, 4.3, 2.5, C.lightRed, C.red);
  s.addText("なぜ生食は不利？", { x: 5.5, y: 2.9, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("生食 大量投与\n↓\n高Cl性 非AG アシドーシス\n↓\nケトアシドーシスの\n解消を遅延", { x: 5.5, y: 3.5, w: 3.9, h: 1.6, fontSize: 18, fontFace: FJ, color: C.text, valign: "top", align: "center", margin: 0 });
})();

// SLIDE 17: 周術期
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "疾患別選択 — 周術期管理");
  var strats = [
    { title: "Restrictive", jp: "制限的", pro: "術後浮腫↓\n肺合併症↓", con: "腎灌流低下\nAKI", color: C.primary },
    { title: "Liberal", jp: "開放的", pro: "循環動態安定", con: "術後浮腫↑\n肺水腫", color: C.accent },
    { title: "GDFT", jp: "個別化", pro: "過不足なく\n最適化", con: "機器・技術\nが必要", color: C.green },
  ];
  strats.forEach(function(st, i) {
    var xPos = 0.4 + i * 3.15;
    card(s, xPos, 1.2, 2.95, 3.6, C.white, st.color);
    s.addText(st.title, { x: xPos + 0.15, y: 1.3, w: 2.65, h: 0.6, fontSize: 24, fontFace: FE, color: st.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.jp, { x: xPos + 0.15, y: 1.9, w: 2.65, h: 0.4, fontSize: 18, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.2, y: 2.5, w: 2.55, h: 0.8, fill: { color: C.lightGreen }, rectRadius: 0.05 });
    s.addText(st.pro, { x: xPos + 0.2, y: 2.5, w: 2.55, h: 0.8, fontSize: 16, fontFace: FJ, color: C.green, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.2, y: 3.4, w: 2.55, h: 0.8, fill: { color: C.lightRed }, rectRadius: 0.05 });
    s.addText(st.con, { x: xPos + 0.2, y: 3.4, w: 2.55, h: 0.8, fontSize: 16, fontFace: FJ, color: C.red, align: "center", valign: "middle", margin: 0 });
  });
  card(s, 0.5, 5.0, 9, 0.5, C.light, C.green);
  s.addText("現在のコンセンサス → GDFT（SVV・PPV で個別化）", { x: 0.7, y: 5.0, w: 8.6, h: 0.5, fontSize: 18, fontFace: FJ, color: C.green, bold: true, valign: "middle", align: "center", margin: 0 });
})();

// SLIDE 18: 細胞外液 vs 維持液
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "細胞外液補充液 vs 維持液");
  var hO = { fontSize: 16, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center", valign: "middle" };
  var cO = { fontSize: 16, fontFace: FJ, color: C.text, align: "center", valign: "middle" };
  var cB = { fontSize: 16, fontFace: FJ, color: C.primary, bold: true, align: "center", valign: "middle" };
  var rows = [
    [{ text: "項目", options: hO },{ text: "細胞外液補充液", options: hO },{ text: "維持液（3号液）", options: hO }],
    [{ text: "目的", options: cB },{ text: "蘇生・補充", options: cO },{ text: "水・電解質の維持", options: cO }],
    [{ text: "Na⁺", options: cB },{ text: "130〜154（等張）", options: cO },{ text: "35〜50（低張）", options: cO }],
    [{ text: "ブドウ糖", options: cB },{ text: "なし", options: cO },{ text: "あり（2.5〜5%）", options: cO }],
    [{ text: "適応", options: cB },{ text: "ショック・出血・術中", options: cO },{ text: "絶食時維持", options: cO }],
    [{ text: "代表製品", options: cB },{ text: "生食・ラクテック・ヴィーンF", options: cO },{ text: "ソリタT3・フィジオ35", options: cO }],
  ];
  s.addTable(rows, { x: 0.3, y: 1.1, w: 9.4, colW: [2.0, 3.7, 3.7], rowH: [0.5, 0.5, 0.5, 0.5, 0.5, 0.5], border: { pt: 1, color: "DEE2E6" }, autoPage: false });
  card(s, 0.5, 4.3, 9, 0.6, C.lightYellow, C.yellow);
  s.addText("迷ったらまず細胞外液補充液。維持液の急速投与は低 Na 血症リスク", { x: 0.7, y: 4.35, w: 8.6, h: 0.5, fontSize: 18, fontFace: FJ, color: C.text, bold: true, valign: "middle", align: "center", margin: 0 });
})();

// SLIDE 19: 晶質液 vs 膠質液
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "晶質液 vs 膠質液");
  card(s, 0.4, 1.2, 2.9, 2.5, C.white, C.primary);
  s.addText("晶質液", { x: 0.6, y: 1.3, w: 2.5, h: 0.5, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("初期蘇生の\n第一選択", { x: 0.6, y: 1.9, w: 2.5, h: 0.7, fontSize: 20, fontFace: FJ, color: C.text, bold: true, align: "center", margin: 0 });
  s.addText("血管内に 1/4\n大量投与で間質浮腫", { x: 0.6, y: 2.7, w: 2.5, h: 0.7, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  card(s, 3.5, 1.2, 2.9, 2.5, C.white, C.accent);
  s.addText("アルブミン", { x: 3.7, y: 1.3, w: 2.5, h: 0.5, fontSize: 22, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("肝硬変の蘇生\n大量晶質液後", { x: 3.7, y: 1.9, w: 2.5, h: 0.7, fontSize: 20, fontFace: FJ, color: C.text, bold: true, align: "center", margin: 0 });
  s.addText("ESICM: 肝硬変で\n条件付き推奨", { x: 3.7, y: 2.7, w: 2.5, h: 0.7, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  card(s, 6.6, 1.2, 3.0, 2.5, C.lightRed, C.red);
  s.addText("HES", { x: 6.8, y: 1.3, w: 2.6, h: 0.5, fontSize: 22, fontFace: FE, color: C.red, bold: true, margin: 0 });
  s.addText("推奨されない", { x: 6.8, y: 1.9, w: 2.6, h: 0.5, fontSize: 22, fontFace: FJ, color: C.red, bold: true, align: "center", margin: 0 });
  s.addText("AKI↑ 死亡率↑\nEMA 使用停止勧告\nSSC 強い推奨で非推奨", { x: 6.8, y: 2.5, w: 2.6, h: 1.0, fontSize: 16, fontFace: FJ, color: C.red, align: "center", margin: 0 });
})();

// SLIDE 20: ROSE model
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "Fluid Stewardship — ROSE model");
  var phases = [
    { letter: "R", name: "Resuscitation", jp: "蘇生", desc: "急速ボーラスで\n臓器灌流を回復", color: C.red },
    { letter: "O", name: "Optimization", jp: "最適化", desc: "輸液反応性に基づき\n酸素運搬を最適化", color: C.accent },
    { letter: "S", name: "Stabilization", jp: "安定化", desc: "維持のみ\n不要な輸液を中止", color: C.yellow },
    { letter: "E", name: "Evacuation", jp: "除去", desc: "利尿薬・限外濾過で\n過剰体液を除去", color: C.green },
  ];
  phases.forEach(function(p, i) {
    var xPos = 0.3 + i * 2.4;
    card(s, xPos, 1.7, 2.2, 3.0, C.white, p.color);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.7, y: 1.85, w: 0.8, h: 0.8, fill: { color: p.color } });
    s.addText(p.letter, { x: xPos + 0.7, y: 1.85, w: 0.8, h: 0.8, fontSize: 32, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.name, { x: xPos + 0.1, y: 2.75, w: 2.0, h: 0.4, fontSize: 15, fontFace: FE, color: p.color, bold: true, align: "center", margin: 0 });
    s.addText(p.jp, { x: xPos + 0.1, y: 3.1, w: 2.0, h: 0.3, fontSize: 18, fontFace: FJ, color: C.text, bold: true, align: "center", margin: 0 });
    s.addText(p.desc, { x: xPos + 0.1, y: 3.5, w: 2.0, h: 0.9, fontSize: 15, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
    if (i < 3) {
      s.addShape(pres.shapes.RIGHT_ARROW, { x: xPos + 2.2, y: 2.3, w: 0.2, h: 0.5, fill: { color: C.sub } });
    }
  });
  card(s, 0.5, 4.9, 9, 0.5, C.light, C.primary);
  s.addText("ESICM 2024-2025: 選択・量・除去の 3 部構成ガイドライン", { x: 0.7, y: 4.9, w: 8.6, h: 0.5, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, valign: "middle", align: "center", margin: 0 });
})();

// SLIDE 21: 4つの問い
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "輸液オーダー時の 4 つの問い");
  var qs = [
    { num: "1", q: "この輸液は本当に必要か？", sub: "経口摂取で代替できないか", color: C.primary },
    { num: "2", q: "どの種類が最適か？", sub: "病態に応じた選択", color: C.accent },
    { num: "3", q: "どのくらいの量が必要か？", sub: "過不足なく", color: C.yellow },
    { num: "4", q: "いつ中止すべきか？", sub: "漫然投与を避ける", color: C.green },
  ];
  qs.forEach(function(q, i) {
    var yPos = 1.1 + i * 1.1;
    card(s, 0.4, yPos, 9.2, 0.9, C.white, q.color);
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.1, w: 0.7, h: 0.7, fill: { color: q.color } });
    s.addText(q.num, { x: 0.7, y: yPos + 0.1, w: 0.7, h: 0.7, fontSize: 24, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(q.q, { x: 1.6, y: yPos + 0.05, w: 5.5, h: 0.5, fontSize: 22, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(q.sub, { x: 1.6, y: yPos + 0.5, w: 5.5, h: 0.3, fontSize: 16, fontFace: FJ, color: C.sub, valign: "middle", margin: 0 });
  });
})();

// SLIDE 22: まとめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ — 輸液選択の判断フレームワーク");
  var steps = [
    { num: "1", title: "目的の明確化", desc: "蘇生・補充 → 細胞外液補充液\n維持 → 維持液", color: C.primary },
    { num: "2", title: "種類の選択", desc: "一般 → バランス液が第一選択\nTBI・高K → 生食", color: C.accent },
    { num: "3", title: "量と速度", desc: "蘇生：ボーラス＋輸液反応性評価\n安定後：維持量へ移行", color: C.yellow },
    { num: "4", title: "再評価と中止", desc: "「この輸液はまだ必要か？」\n漫然投与を避ける", color: C.green },
  ];
  steps.forEach(function(st, i) {
    var yPos = 1.1 + i * 1.1;
    card(s, 0.4, yPos, 9.2, 0.95, C.white, st.color);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: yPos + 0.1, w: 0.75, h: 0.75, fill: { color: st.color }, rectRadius: 0.1 });
    s.addText(st.num, { x: 0.6, y: yPos + 0.1, w: 0.75, h: 0.75, fontSize: 28, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.title, { x: 1.5, y: yPos + 0.05, w: 2.5, h: 0.85, fontSize: 20, fontFace: FJ, color: st.color, bold: true, valign: "middle", margin: 0 });
    s.addText(st.desc, { x: 4.1, y: yPos + 0.05, w: 5.3, h: 0.85, fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
})();

// SLIDE 23: エンディング
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("ご視聴ありがとうございました", { x: 0.6, y: 0.4, w: 8.8, h: 0.8, fontSize: 32, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText("医知創造ラボ", { x: 0.6, y: 1.1, w: 8.8, h: 0.4, fontSize: 18, fontFace: FJ, color: C.accent, align: "center", margin: 0 });
  var refs = [
    "1. Semler MW, et al. SMART Trial. NEJM 2018;378:829-839.",
    "2. Self WH, et al. SALT-ED. NEJM 2018;378:819-828.",
    "3. Zampieri FG, et al. BaSICS Trial. JAMA 2021;326:818-829.",
    "4. Finfer S, et al. PLUS Trial. NEJM 2022;386:815-826.",
    "5. Zampieri FG, et al. BEST-Living IPD MA. Lancet Respir Med 2023.",
    "6. Evans L, et al. SSC Guidelines 2021. Crit Care Med 2021.",
    "7. Zampieri FG, et al. ESICM Fluid GL Part 1. ICM 2024.",
    "8. Malbrain MLNG, et al. Fluid Stewardship. Ann Intensive Care 2023.",
    "9. RINSE-DKA. Pharmacotherapy 2024.",
  ];
  s.addText(refs.join("\n"), { x: 0.6, y: 1.7, w: 8.8, h: 3.6, fontSize: 14, fontFace: FE, color: C.white, valign: "top", margin: 0 });
})();

pres.writeFile({ fileName: "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/細胞外液補充液/細胞外液補充液_slides.pptx" }).then(function() {
  console.log("PPTX created: 23 slides");
});
