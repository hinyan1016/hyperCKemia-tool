var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "ビタミンB12欠乏症【前編】症状・原因・診断アプローチ";

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

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("ビタミンB12欠乏症【前編】", {
    x: 0.6, y: 1.0, w: 8.8, h: 1.2,
    fontSize: 38, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("症状・原因・診断アプローチ", {
    x: 0.6, y: 2.8, w: 8.8, h: 0.6,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("「MCVが正常でも、B12欠乏は否定できない」", {
    x: 0.6, y: 3.8, w: 8.8, h: 0.6,
    fontSize: 18, fontFace: FJ, color: "B0BEC5", align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 見落とさないための3つの鉄則
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "見落とさないための3つの鉄則");

  var rules = [
    { num: "1", title: "貧血なし\nでも発症", desc: "神経症状例の1/4以上で\nHb・MCVは正常範囲", icon: "💧", color: C.primary },
    { num: "2", title: "正常低値でも\n否定できない", desc: "血清B12「正常低値」でも\n機能的欠乏あり\n→ MMA追加測定が必要", icon: "⚠", color: C.yellow },
    { num: "3", title: "治療の遅れ\n＝不可逆的障害", desc: "早期発見・早期介入が\n予後を決定する", icon: "⏱", color: C.red },
  ];
  rules.forEach(function(r, i) {
    var xPos = 0.4 + i * 3.15;
    card(s, xPos, 1.2, 2.95, 3.6, C.white, r.color);
    s.addText(r.icon, { x: xPos, y: 1.3, w: 2.95, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(r.title, { x: xPos + 0.15, y: 2.0, w: 2.65, h: 0.8, fontSize: 18, fontFace: FJ, color: r.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(r.desc, { x: xPos + 0.15, y: 2.9, w: 2.65, h: 1.6, fontSize: 13, fontFace: FJ, color: C.text, align: "center", valign: "top", margin: 0 });
  });
})();

// ============================================================
// SLIDE 3: B12代謝経路（2つのマーカー）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "B12欠乏で上昇する2つのマーカー：ホモシステインとMMA");

  // 左：メチオニン合成酵素
  card(s, 0.3, 1.2, 4.5, 3.2, C.white, C.primary);
  s.addText("細胞質（メチオニン合成酵素）", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0 });

  // フロー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.9, w: 1.6, h: 0.5, fill: { color: C.light }, rectRadius: 0.05 });
  s.addText("ホモシステイン", { x: 0.8, y: 1.9, w: 1.6, h: 0.5, fontSize: 12, fontFace: FJ, color: C.text, align: "center", valign: "middle", margin: 0 });
  s.addText("→ B12 →", { x: 2.4, y: 1.9, w: 0.8, h: 0.5, fontSize: 11, fontFace: FE, color: C.primary, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.2, y: 1.9, w: 1.3, h: 0.5, fill: { color: C.light }, rectRadius: 0.05 });
  s.addText("メチオニン", { x: 3.2, y: 1.9, w: 1.3, h: 0.5, fontSize: 12, fontFace: FJ, color: C.text, align: "center", valign: "middle", margin: 0 });

  // 欠乏時
  s.addText("↓ 欠乏時", { x: 0.8, y: 2.5, w: 1.6, h: 0.3, fontSize: 11, fontFace: FJ, color: C.red, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.9, w: 2.0, h: 0.4, fill: { color: C.lightRed }, rectRadius: 0.05 });
  s.addText("ホモシステイン蓄積", { x: 0.6, y: 2.9, w: 2.0, h: 0.4, fontSize: 11, fontFace: FJ, color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("↓", { x: 0.6, y: 3.3, w: 2.0, h: 0.3, fontSize: 11, fontFace: FJ, color: C.red, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.6, w: 2.0, h: 0.4, fill: { color: C.lightRed }, rectRadius: 0.05 });
  s.addText("DNA合成障害", { x: 0.6, y: 3.6, w: 2.0, h: 0.4, fontSize: 11, fontFace: FJ, color: C.red, align: "center", valign: "middle", margin: 0 });
  s.addText("↓ 大球性貧血", { x: 0.6, y: 4.0, w: 2.0, h: 0.3, fontSize: 11, fontFace: FJ, color: C.red, bold: true, align: "center", margin: 0 });

  // 右：メチルマロニルCoAムターゼ
  card(s, 5.2, 1.2, 4.5, 3.2, C.white, C.accent);
  s.addText("ミトコンドリア（メチルマロニルCoAムターゼ）", { x: 5.4, y: 1.3, w: 4.1, h: 0.4, fontSize: 14, fontFace: FJ, color: C.accent, bold: true, margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.7, y: 1.9, w: 1.2, h: 0.5, fill: { color: C.lightYellow }, rectRadius: 0.05 });
  s.addText("MMA", { x: 5.7, y: 1.9, w: 1.2, h: 0.5, fontSize: 13, fontFace: FE, color: C.text, align: "center", valign: "middle", margin: 0 });
  s.addText("→ B12 →", { x: 6.9, y: 1.9, w: 0.8, h: 0.5, fontSize: 11, fontFace: FE, color: C.accent, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.7, y: 1.9, w: 1.6, h: 0.5, fill: { color: C.lightYellow }, rectRadius: 0.05 });
  s.addText("スクシニルCoA", { x: 7.7, y: 1.9, w: 1.6, h: 0.5, fontSize: 12, fontFace: FJ, color: C.text, align: "center", valign: "middle", margin: 0 });

  s.addText("↓ 欠乏時", { x: 5.7, y: 2.5, w: 1.2, h: 0.3, fontSize: 11, fontFace: FJ, color: C.red, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 2.9, w: 2.0, h: 0.4, fill: { color: C.lightRed }, rectRadius: 0.05 });
  s.addText("MMA蓄積", { x: 5.5, y: 2.9, w: 2.0, h: 0.4, fontSize: 11, fontFace: FJ, color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("↓", { x: 5.5, y: 3.3, w: 2.0, h: 0.3, fontSize: 11, fontFace: FJ, color: C.red, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 3.6, w: 2.0, h: 0.4, fill: { color: C.lightRed }, rectRadius: 0.05 });
  s.addText("神経障害（SCD等）", { x: 5.5, y: 3.6, w: 2.0, h: 0.4, fontSize: 11, fontFace: FJ, color: C.red, align: "center", valign: "middle", margin: 0 });

  // ボトムメッセージ
  card(s, 0.5, 4.6, 9, 0.7, C.light, C.primary);
  s.addText("B12欠乏の確認マーカー：ホモシステイン↑ ＋ MMA↑", {
    x: 0.7, y: 4.7, w: 8.6, h: 0.5, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", align: "center", margin: 0
  });
})();

// ============================================================
// SLIDE 4: 吸収経路と障害ポイント
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "吸収の各ステップが欠乏の原因となる");

  var steps = [
    { title: "食事中B12\n（タンパク結合）", x: 0.3, color: C.sub },
    { title: "胃酸・ペプシン\nで遊離", x: 2.5, color: C.primary, warn: "障害:\n萎縮性胃炎・\nPPI・H2RA" },
    { title: "内因子と結合", x: 4.7, color: C.primary, warn: "障害:\n悪性貧血・\n胃全摘後" },
    { title: "回腸末端で吸収", x: 6.9, color: C.primary, warn: "障害:\nクローン病・\n回腸切除" },
  ];
  steps.forEach(function(st, i) {
    card(s, st.x, 1.2, 2.0, 1.2, C.white, st.color);
    s.addText(st.title, { x: st.x + 0.15, y: 1.3, w: 1.7, h: 1.0, fontSize: 13, fontFace: FJ, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
    if (i < 3) {
      s.addText("→", { x: st.x + 2.0, y: 1.5, w: 0.5, h: 0.6, fontSize: 22, fontFace: FE, color: C.sub, align: "center", valign: "middle", margin: 0 });
    }
    if (st.warn) {
      s.addShape(pres.shapes.RECTANGLE, { x: st.x, y: 2.6, w: 2.0, h: 1.2, fill: { color: C.lightRed }, rectRadius: 0.05 });
      s.addText("⚠ " + st.warn, { x: st.x + 0.1, y: 2.7, w: 1.8, h: 1.0, fontSize: 11, fontFace: FJ, color: C.red, align: "center", valign: "middle", margin: 0 });
    }
  });

  // 貯蔵量
  card(s, 0.5, 4.2, 9, 0.7, C.light, C.primary);
  s.addText("体内貯蔵量：2〜5mg（肝臓）→ 枯渇まで数年 → 診断の遅れの原因", {
    x: 0.7, y: 4.3, w: 8.6, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, bold: true, valign: "middle", align: "center", margin: 0
  });
})();

// ============================================================
// SLIDE 5: SCD：後索＋側索の障害
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "SCD：後索＋側索の障害が特徴的パターンを生む");

  // 左：脊髄横断面の模式図（テキストで表現）
  card(s, 0.3, 1.2, 4.5, 3.0, C.white);
  s.addText("亜急性連合変性（SCD）", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });

  // 脊髄模式図をシェイプで表現
  // 外枠（脊髄断面）
  s.addShape(pres.shapes.OVAL, { x: 1.3, y: 1.9, w: 2.4, h: 2.0, fill: { color: "E0E0E0" }, line: { color: C.sub, width: 1 } });
  // 後索（上部）- 赤でハイライト
  s.addShape(pres.shapes.OVAL, { x: 1.9, y: 2.0, w: 1.2, h: 0.7, fill: { color: "FF6B6B" }, line: { color: C.red, width: 1.5 } });
  s.addText("後索", { x: 1.9, y: 2.05, w: 1.2, h: 0.6, fontSize: 11, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  // 側索（左右）- 赤でハイライト
  s.addShape(pres.shapes.OVAL, { x: 1.35, y: 2.6, w: 0.6, h: 0.8, fill: { color: "FF6B6B" }, line: { color: C.red, width: 1.5 } });
  s.addText("側索", { x: 1.35, y: 2.7, w: 0.6, h: 0.6, fontSize: 9, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.OVAL, { x: 3.05, y: 2.6, w: 0.6, h: 0.8, fill: { color: "FF6B6B" }, line: { color: C.red, width: 1.5 } });
  s.addText("側索", { x: 3.05, y: 2.7, w: 0.6, h: 0.6, fontSize: 9, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // 右：臨床情報
  card(s, 5.2, 1.2, 4.5, 3.0, C.white, C.primary);
  s.addText("進行パターン", { x: 5.4, y: 1.3, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "1. 深部感覚障害\n", options: { fontSize: 15, fontFace: FJ, color: C.text } },
    { text: "2. 感覚性失調\n", options: { fontSize: 15, fontFace: FJ, color: C.text } },
    { text: "3. 痙性対麻痺\n\n", options: { fontSize: 15, fontFace: FJ, color: C.text } },
    { text: "▶ 末梢神経障害\n", options: { fontSize: 14, fontFace: FJ, color: C.sub } },
    { text: "  対称性手袋靴下型しびれ", options: { fontSize: 14, fontFace: FJ, color: C.sub } },
  ], { x: 5.4, y: 1.8, w: 4.1, h: 2.2, valign: "top", margin: [0, 0, 0, 8] });

  // 予後バー
  card(s, 0.3, 4.5, 4.5, 0.7, C.lightGreen, C.green);
  s.addText("独歩可能で治療開始 → 90%完全回復", {
    x: 0.5, y: 4.6, w: 4.1, h: 0.5, fontSize: 14, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0
  });
  card(s, 5.2, 4.5, 4.5, 0.7, C.lightRed, C.red);
  s.addText("車椅子状態で治療開始 → 33〜50%回復", {
    x: 5.4, y: 4.6, w: 4.1, h: 0.5, fontSize: 14, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0
  });
})();

// ============================================================
// SLIDE 6: 貧血がなくても疑うべきケース
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "貧血がなくてもB12欠乏を疑うべきケースがある");

  // 上段左：血液所見
  card(s, 0.3, 1.2, 4.5, 1.8, C.white, C.primary);
  s.addText("血液所見", { x: 0.5, y: 1.3, w: 4.1, h: 0.35, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "• MCV上昇（大球性貧血）\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "• 過分葉好中球\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "• 重症例：汎血球減少", options: { fontSize: 13, fontFace: FJ, color: C.text } },
  ], { x: 0.5, y: 1.7, w: 4.1, h: 1.1, valign: "top", margin: 0 });

  // 上段右：ピットフォール
  card(s, 5.2, 1.2, 4.5, 1.8, C.lightYellow, C.yellow);
  s.addText("⚠ ピットフォール", { x: 5.4, y: 1.3, w: 4.1, h: 0.35, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText([
    { text: "LDH著増＋間接Bil上昇\n", options: { fontSize: 14, fontFace: FJ, color: C.text, bold: true } },
    { text: "→ 溶血性貧血・TTPと\n  誤診されうる", options: { fontSize: 14, fontFace: FJ, color: C.red } },
  ], { x: 5.4, y: 1.7, w: 4.1, h: 1.1, valign: "top", margin: 0 });

  // 中段：神経精神症状
  card(s, 0.3, 3.2, 9.4, 0.7, C.white, C.sub);
  s.addText("神経精神症状：うつ・集中困難・記銘力低下・認知機能低下（「治療可能な認知症」）", {
    x: 0.5, y: 3.3, w: 9, h: 0.5, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0
  });

  // 下段：貧血なしでも疑う
  card(s, 0.3, 4.1, 9.4, 1.2, C.lightRed, C.red);
  s.addText([
    { text: "以下では貧血なしでもB12欠乏を疑う：", options: { fontSize: 14, fontFace: FJ, color: C.red, bold: true } },
    { text: " 原因不明の対称性しびれ、原因不明の歩行障害、\n鉄欠乏/サラセミア合併でMCVが正球性、高齢者の原因不明の認知機能低下", options: { fontSize: 13, fontFace: FJ, color: C.text } },
  ], { x: 0.5, y: 4.2, w: 9, h: 1.0, valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 7: 欠乏の3大カテゴリ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "欠乏の3大カテゴリ：摂取不足・吸収障害・薬剤性");

  // 摂取不足
  card(s, 0.3, 1.2, 2.8, 4.0, C.white, C.sub);
  s.addText("摂取不足", { x: 0.5, y: 1.3, w: 2.4, h: 0.4, fontSize: 16, fontFace: FJ, color: C.sub, bold: true, margin: 0 });
  s.addText([
    { text: "• ヴィーガン・\n  低栄養・高齢者\n\n", options: { fontSize: 12, fontFace: FJ, color: C.text } },
    { text: "• B12は動物性食品\n  のみに含有\n\n", options: { fontSize: 12, fontFace: FJ, color: C.text } },
    { text: "• 日本DRI 2025:\n  成人 ", options: { fontSize: 12, fontFace: FJ, color: C.text } },
    { text: "4.0μg/日", options: { fontSize: 13, fontFace: FE, color: C.primary, bold: true } },
  ], { x: 0.5, y: 1.8, w: 2.4, h: 3.2, valign: "top", margin: 0 });

  // 吸収障害（最多原因）
  card(s, 3.3, 1.2, 3.2, 4.0, C.white, C.primary);
  s.addText("吸収障害（最多原因）", { x: 3.5, y: 1.3, w: 2.8, h: 0.4, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "食物結合B12吸収障害\n", options: { fontSize: 12, fontFace: FJ, color: C.text, bold: true } },
    { text: "(胃酸低下、現在の最多原因・高齢者)\n\n", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
    { text: "悪性貧血/AIG\n", options: { fontSize: 12, fontFace: FJ, color: C.text, bold: true } },
    { text: "(内因子欠乏、甲状腺疾患合併)\n\n", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
    { text: "胃全摘後\n", options: { fontSize: 12, fontFace: FJ, color: C.text, bold: true } },
    { text: "(内因子消失、術後4-5年で必発)\n\n", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
    { text: "回腸末端病変\n", options: { fontSize: 12, fontFace: FJ, color: C.text, bold: true } },
    { text: "(吸収部位喪失、クローン病等)", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
  ], { x: 3.5, y: 1.8, w: 2.8, h: 3.2, valign: "top", margin: 0 });

  // 薬剤性
  card(s, 6.7, 1.2, 3.0, 4.0, C.white, C.accent);
  s.addText("薬剤性", { x: 6.9, y: 1.3, w: 2.6, h: 0.4, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText([
    { text: "メトホルミン\n", options: { fontSize: 12, fontFace: FJ, color: C.text, bold: true } },
    { text: "(Ca依存性吸収阻害、\n4ヶ月以上)\n\n", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
    { text: "PPI/H2RA\n", options: { fontSize: 12, fontFace: FJ, color: C.text, bold: true } },
    { text: "(胃酸抑制、12ヶ月以上)\n\n", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
    { text: "亜酸化窒素\n", options: { fontSize: 12, fontFace: FJ, color: C.text, bold: true } },
    { text: "(Co酸化・不活化、\n潜在性欠乏を急速顕在化)", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
  ], { x: 6.9, y: 1.8, w: 2.6, h: 3.2, valign: "top", margin: 0 });
})();

// ============================================================
// SLIDE 8: 血清B12値＋二次マーカー
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "血清B12値は「正常」でも機能的欠乏を見逃す");

  // 3段階バー
  var zones = [
    { label: "<180 pg/mL", action: "欠乏確定 → 治療開始", color: C.red, fill: C.lightRed, x: 0.3, w: 3.0 },
    { label: "180〜350 pg/mL", action: "境界域 → MMA追加測定", color: C.yellow, fill: C.lightYellow, x: 3.5, w: 3.0 },
    { label: ">350 pg/mL", action: "欠乏の可能性低い", color: C.green, fill: C.lightGreen, x: 6.7, w: 3.0 },
  ];
  zones.forEach(function(z) {
    s.addShape(pres.shapes.RECTANGLE, { x: z.x, y: 1.15, w: z.w, h: 0.45, fill: { color: z.color }, rectRadius: 0.05 });
    s.addText(z.label, { x: z.x, y: 1.15, w: z.w, h: 0.45, fontSize: 14, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(z.action, { x: z.x, y: 1.65, w: z.w, h: 0.35, fontSize: 11, fontFace: FJ, color: C.text, align: "center", valign: "middle", margin: 0 });
  });

  // 注記
  s.addText("⚠ 感度・特異度は不十分。臨床症状の認識を最優先（Delphi 2024）", {
    x: 0.3, y: 2.1, w: 9.4, h: 0.35, fontSize: 12, fontFace: FJ, color: C.red, bold: true, margin: 0
  });

  // 二次マーカー比較表
  s.addText("二次的マーカーの比較（Secondary Markers Comparison）", {
    x: 0.3, y: 2.6, w: 9.4, h: 0.35, fontSize: 14, fontFace: FJ, color: C.primary, bold: true, margin: 0
  });

  var tOpts = { x: 0.3, y: 3.0, w: 9.4, h: 0.01, colW: [2.2, 2.4, 4.8], rowH: [0.4, 0.4, 0.4, 0.4], margin: [4,6,4,6], autoPage: false };
  var rows = [
    [
      { text: "マーカー", options: { fontSize: 13, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary } } },
      { text: "変化・特異性", options: { fontSize: 13, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary } } },
      { text: "偽陽性要因", options: { fontSize: 13, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary } } },
    ],
    [
      { text: "MMA", options: { fontSize: 13, fontFace: FE, bold: true, fill: { color: C.white } } },
      { text: "↑  特異性高い", options: { fontSize: 12, fontFace: FJ, fill: { color: C.white } } },
      { text: "偽陽性: 腎機能低下", options: { fontSize: 12, fontFace: FJ, fill: { color: C.white } } },
    ],
    [
      { text: "ホモシステイン", options: { fontSize: 12, fontFace: FJ, bold: true, fill: { color: C.warmBg } } },
      { text: "↑  特異性低い", options: { fontSize: 12, fontFace: FJ, fill: { color: C.warmBg } } },
      { text: "偽陽性: 葉酸欠乏・腎機能低下・加齢", options: { fontSize: 12, fontFace: FJ, fill: { color: C.warmBg } } },
    ],
    [
      { text: "ホロTC", options: { fontSize: 13, fontFace: FE, bold: true, fill: { color: C.white } } },
      { text: "↓  特異性高い", options: { fontSize: 12, fontFace: FJ, fill: { color: C.white } } },
      { text: "偽陽性: 妊娠・OC", options: { fontSize: 12, fontFace: FJ, fill: { color: C.white } } },
    ],
  ];
  s.addTable(rows, tOpts);

  // 日本の制限
  card(s, 0.3, 4.7, 9.4, 0.6, C.lightYellow, C.yellow);
  s.addText("日本ではMMA測定に保険適用制限あり → 診断的治療も選択肢", {
    x: 0.5, y: 4.8, w: 9, h: 0.4, fontSize: 13, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0
  });
})();

// ============================================================
// SLIDE 9: 診断アルゴリズム
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "臨床的疑いから治療開始までのフロー");

  // Step 1
  card(s, 0.3, 1.2, 5.5, 0.6, C.light);
  s.addText("1. 臨床的にB12欠乏を疑う（しびれ・歩行障害・認知機能低下・大球性貧血）", {
    x: 0.5, y: 1.25, w: 5.1, h: 0.5, fontSize: 12, fontFace: FJ, color: C.text, valign: "middle", margin: 0
  });
  s.addText("↓", { x: 2.5, y: 1.8, w: 1, h: 0.3, fontSize: 14, fontFace: FE, color: C.sub, align: "center", margin: 0 });

  // Step 2
  card(s, 0.3, 2.1, 5.5, 0.5, C.light);
  s.addText("2. 血清B12測定 ＋ CBC ＋ 末梢血塗抹 ＋ LDH・間接Bil", {
    x: 0.5, y: 2.15, w: 5.1, h: 0.4, fontSize: 12, fontFace: FJ, color: C.text, valign: "middle", margin: 0
  });
  s.addText("↓", { x: 2.5, y: 2.6, w: 1, h: 0.3, fontSize: 14, fontFace: FE, color: C.sub, align: "center", margin: 0 });

  // 3分岐
  var branches = [
    { label: "<180 pg/mL:\n欠乏確定", color: C.red, fill: C.lightRed, x: 0.3 },
    { label: "180-350 pg/mL:\n境界域", color: C.yellow, fill: C.lightYellow, x: 2.1 },
    { label: ">350 pg/mL:\n可能性低い", color: C.green, fill: C.lightGreen, x: 3.9 },
  ];
  branches.forEach(function(b) {
    s.addShape(pres.shapes.RECTANGLE, { x: b.x, y: 2.9, w: 1.6, h: 0.7, fill: { color: b.fill }, rectRadius: 0.05, line: { color: b.color, width: 1.5 } });
    s.addText(b.label, { x: b.x, y: 2.95, w: 1.6, h: 0.6, fontSize: 10, fontFace: FJ, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // MMA測定
  s.addText("↓", { x: 2.7, y: 3.6, w: 0.5, h: 0.3, fontSize: 12, fontFace: FE, color: C.sub, align: "center", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.1, y: 3.9, w: 1.6, h: 0.5, fill: { color: C.lightYellow }, rectRadius: 0.05 });
  s.addText("MMA測定", { x: 2.1, y: 3.9, w: 1.6, h: 0.5, fontSize: 12, fontFace: FJ, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("↓ MMA↑", { x: 2.5, y: 4.4, w: 0.8, h: 0.3, fontSize: 10, fontFace: FE, color: C.red, align: "center", margin: 0 });

  // Step 4
  card(s, 0.3, 4.7, 5.5, 0.5, C.light);
  s.addText("4. 原因検索：食事歴・薬剤歴・抗IF抗体・H.pylori・上部内視鏡", {
    x: 0.5, y: 4.75, w: 5.1, h: 0.4, fontSize: 12, fontFace: FJ, color: C.text, valign: "middle", margin: 0
  });

  // 右側：原因検索の詳細
  card(s, 6.2, 1.2, 3.5, 4.0, C.white, C.primary);
  s.addText("原因検索", { x: 6.4, y: 1.3, w: 3.1, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "抗内因子抗体\n", options: { fontSize: 13, fontFace: FJ, color: C.text, bold: true } },
    { text: "(特異度高・感度<50%)\n\n", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
    { text: "H.pylori陽性 → 除菌\n\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "上部消化管内視鏡：\n", options: { fontSize: 13, fontFace: FJ, color: C.text, bold: true } },
    { text: "AIG/悪性貧血・胃NET\nスクリーニング", options: { fontSize: 11, fontFace: FJ, color: C.sub } },
  ], { x: 6.4, y: 1.8, w: 3.1, h: 3.2, valign: "top", margin: 0 });
})();

// ============================================================
// SLIDE 10: Take-home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("4つのTake-home Message", { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 26, fontFace: FJ, color: C.white, bold: true, margin: 0 });

  // 警告バー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 9, h: 0.6, fill: { color: "3D1F1F" }, rectRadius: 0.05, line: { color: C.red, width: 1.5 } });
  s.addText("⚠ 葉酸単独補充は危険（masking効果）→ 必ずB12＋葉酸の両方を測定", {
    x: 0.7, y: 1.15, w: 8.6, h: 0.5, fontSize: 14, fontFace: FJ, color: C.white, bold: true, valign: "middle", margin: 0
  });

  // 鑑別
  card(s, 0.5, 1.9, 4.2, 1.2, "24466E");
  s.addText([
    { text: "• 大球性貧血の鑑別：\n", options: { fontSize: 13, fontFace: FJ, color: C.white, bold: true } },
    { text: "  葉酸欠乏・MDS・甲状腺機能低下・\n  アルコール・薬剤性", options: { fontSize: 12, fontFace: FJ, color: "B0BEC5" } },
  ], { x: 0.7, y: 2.0, w: 3.8, h: 1.0, valign: "top", margin: 0 });

  card(s, 5.3, 1.9, 4.2, 1.2, "24466E");
  s.addText([
    { text: "• 神経症状の鑑別：\n", options: { fontSize: 13, fontFace: FJ, color: C.white, bold: true } },
    { text: "  銅欠乏・糖尿病性ニューロパチー・\n  頚椎症性脊髄症", options: { fontSize: 12, fontFace: FJ, color: "B0BEC5" } },
  ], { x: 5.5, y: 2.0, w: 3.8, h: 1.0, valign: "top", margin: 0 });

  // 4つのメッセージ
  var msgs = [
    { num: "1", title: "貧血なしでも疑う", desc: "MCVが正常でも\n鑑別から外さない", color: C.primary },
    { num: "2", title: "境界域B12では\nMMAを追加", desc: "機能的欠乏を\n見逃さない", color: C.yellow },
    { num: "3", title: "最多原因は食物\n結合B12吸収障害", desc: "メトホルミン+PPI\n併用に注意", color: C.accent },
    { num: "4", title: "葉酸単独補充の\nmasking効果に注意", desc: "必ず両方測定", color: C.red },
  ];
  msgs.forEach(function(m, i) {
    var xPos = 0.3 + i * 2.45;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.4, w: 2.25, h: 1.8, fill: { color: "24466E" }, shadow: shd(), rectRadius: 0.1 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.4, w: 2.25, h: 0.06, fill: { color: m.color } });
    s.addText(m.num + ".", { x: xPos, y: 3.5, w: 2.25, h: 0.35, fontSize: 18, fontFace: FE, color: m.color, bold: true, align: "center", margin: 0 });
    s.addText(m.title, { x: xPos + 0.1, y: 3.85, w: 2.05, h: 0.6, fontSize: 12, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.desc, { x: xPos + 0.1, y: 4.5, w: 2.05, h: 0.6, fontSize: 11, fontFace: FJ, color: "B0BEC5", align: "center", valign: "top", margin: 0 });
  });

  s.addText("後編：治療・フォローアップ・特殊集団", {
    x: 0, y: 5.1, w: 10, h: 0.4, fontSize: 14, fontFace: FJ, color: C.accent, align: "center", margin: 0
  });
})();

// ============================================================
// SAVE
// ============================================================
var outPath = "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/ビタミンB12欠乏症/B12欠乏症_前編_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX saved: " + outPath);
}).catch(function(e) {
  console.error("Error: " + e);
});
