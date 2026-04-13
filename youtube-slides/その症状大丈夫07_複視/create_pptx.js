var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#07】物が二重に見えたら要注意";

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

function ftr(s) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.38, fill: { color: C.dark } });
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #07", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("👁️", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("物が二重に見えたら\n要注意", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.9, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("複視の原因と受診すべき科を\n脳神経内科医が解説", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.9,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #07", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.7, w: 8.8, h: 0.4,
    fontSize: 12, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 2: こんな経験ありませんか？
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな経験、ありませんか？");

  var items = [
    { emoji: "👀", text: "文字や景色が二重に見える" },
    { emoji: "🤔", text: "片目をつぶると1つに見える" },
    { emoji: "🚗", text: "運転中に信号が2つに見えてヒヤッとした" },
    { emoji: "😩", text: "夕方になると物がダブって見える" },
  ];

  items.forEach(function(item, i) {
    var y = 1.2 + i * 0.95;
    card(s, 0.6, y, 8.8, 0.8, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.1, w: 0.7, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.1, w: 7.5, h: 0.6, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.5, 5.0, 7.0, 0.2, C.accent);
  ftr(s);
})();

// ============================================================
// SLIDE 3: 複視とは？ ― 2つのタイプ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「複視」とは？ ― 2つのタイプ");

  card(s, 0.5, 1.1, 9.0, 0.7, C.light, C.primary);
  s.addText("複視（ふくし）＝ 物が二重に見える症状", {
    x: 0.8, y: 1.15, w: 8.5, h: 0.6, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  // 単眼性 vs 両眼性
  card(s, 0.4, 2.1, 4.4, 2.7, C.lightGreen, C.green);
  s.addText("👁️ 単眼性複視", { x: 0.7, y: 2.2, w: 3.8, h: 0.5, fontSize: 21, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("片目だけで見ても二重", { x: 0.7, y: 2.7, w: 3.8, h: 0.35, fontSize: 17, fontFace: FJ, color: C.text, margin: 0 });
  var mono = [
    "目そのものの問題",
    "乱視・白内障・ドライアイ",
    "→ 眼科を受診",
  ];
  mono.forEach(function(t, i) {
    s.addText("・" + t, { x: 0.8, y: 3.2 + i * 0.38, w: 3.8, h: 0.33, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.2, 2.1, 4.4, 2.7, C.lightRed, C.red);
  s.addText("👁️👁️ 両眼性複視", { x: 5.5, y: 2.2, w: 3.8, h: 0.5, fontSize: 21, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("両目で見ると二重、片目だと1つ", { x: 5.5, y: 2.7, w: 3.8, h: 0.35, fontSize: 17, fontFace: FJ, color: C.text, margin: 0 });
  var bino = [
    "目を動かす神経や筋肉の問題",
    "脳の病気が原因のことも",
    "→ 脳神経内科を受診",
  ];
  bino.forEach(function(t, i) {
    s.addText("・" + t, { x: 5.6, y: 3.2 + i * 0.38, w: 3.8, h: 0.33, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  });

  s.addText("⚠️ 両眼性複視は脳や神経の病気が隠れている可能性があります", {
    x: 0.5, y: 4.95, w: 9.0, h: 0.3, fontSize: 14, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: 目を動かす神経 ― 3本の脳神経
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "目を動かす3本の脳神経");

  card(s, 0.5, 1.1, 9.0, 0.7, C.light, C.primary);
  s.addText("眼球は6つの筋肉で動き、3本の脳神経がコントロールしています", {
    x: 0.8, y: 1.15, w: 8.5, h: 0.6, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  var nerves = [
    { name: "動眼神経（第III脳神経）", role: "上・下・内側への眼球運動\nまぶたを上げる・瞳孔を縮める", note: "麻痺→目が外向き＋まぶた下垂", color: C.primary },
    { name: "滑車神経（第IV脳神経）", role: "眼球を内下方に向ける\n（階段を下りるとき等）", note: "麻痺→階段でものが二重に", color: C.accent },
    { name: "外転神経（第VI脳神経）", role: "眼球を外側に向ける", note: "麻痺→目が内側に寄る\n最も多い眼球運動神経麻痺", color: C.red },
  ];

  nerves.forEach(function(n, i) {
    var y = 2.05 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, n.color);
    s.addText(n.name, { x: 0.8, y: y + 0.05, w: 3.5, h: 0.35, fontSize: 17, fontFace: FJ, color: n.color, bold: true, margin: 0 });
    s.addText(n.role, { x: 0.8, y: y + 0.4, w: 4.5, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText(n.note, { x: 5.5, y: y + 0.08, w: 3.8, h: 0.7, fontSize: 14, fontFace: FJ, color: C.sub, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 5: よくある原因（心配度別）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "複視の原因 ― 心配度で分類");

  var causes = [
    { level: "🟢 軽度", name: "疲労・ドライアイ・老眼", desc: "疲れ目や加齢による調節機能の低下。休息で改善", bg: C.lightGreen, color: C.green },
    { level: "🟡 要注意", name: "糖尿病性神経障害", desc: "高血糖が眼球運動神経を障害。糖尿病の合併症で最多", bg: C.lightYellow, color: "856404" },
    { level: "🟡 要注意", name: "甲状腺眼症（バセドウ病眼症）", desc: "甲状腺の異常で眼球周囲の筋肉が腫れる", bg: C.lightYellow, color: "856404" },
    { level: "🟡 要注意", name: "重症筋無力症", desc: "夕方に悪化する複視＋まぶた下垂が特徴", bg: C.lightYellow, color: "856404" },
    { level: "🔴 緊急", name: "脳動脈瘤・脳卒中・脳腫瘍", desc: "突然の複視＋頭痛・瞳孔異常は緊急。脳の血管や腫瘍が原因", bg: C.lightRed, color: C.red },
  ];

  causes.forEach(function(c, i) {
    var y = 1.1 + i * 0.8;
    card(s, 0.5, y, 9.0, 0.67, c.bg, c.color);
    s.addText(c.level, { x: 0.7, y: y + 0.05, w: 1.3, h: 0.55, fontSize: 15, fontFace: FJ, color: c.color, bold: true, valign: "middle", margin: 0 });
    s.addText(c.name, { x: 2.0, y: y + 0.05, w: 3.0, h: 0.55, fontSize: 17, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(c.desc, { x: 5.1, y: y + 0.05, w: 4.2, h: 0.55, fontSize: 14, fontFace: FJ, color: C.sub, valign: "middle", lineSpacingMultiple: 1.1, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 6: 重症筋無力症 ― 見逃されやすい原因
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "見逃されやすい原因 ― 重症筋無力症");

  card(s, 0.5, 1.1, 9.0, 0.8, C.lightYellow, C.accent);
  s.addText("複視＋まぶた下垂が夕方に悪化 → 重症筋無力症を疑う", {
    x: 0.8, y: 1.15, w: 8.5, h: 0.35, fontSize: 19, fontFace: FJ, color: C.accent, bold: true, margin: 0,
  });
  s.addText("神経と筋肉の接合部（神経筋接合部）に自己抗体ができる自己免疫疾患", {
    x: 0.8, y: 1.55, w: 8.5, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0,
  });

  var features = [
    { icon: "🌆", title: "夕方に悪化（日内変動）", desc: "朝は比較的良く、夕方〜夜に\n症状が悪化するのが特徴的" },
    { icon: "😴", title: "休むと改善", desc: "目を閉じて休憩すると\n一時的に症状が軽くなる" },
    { icon: "👁️", title: "まぶた下垂を伴う", desc: "片側または両側のまぶたが\n下がってくる（眼瞼下垂）" },
    { icon: "💪", title: "全身型に進行することも", desc: "眼症状から始まり、嚥下障害や\n四肢の筋力低下に広がることも" },
  ];

  features.forEach(function(f, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 2.15 + row * 1.4;
    card(s, x, y, 4.5, 1.2, C.white, C.primary);
    s.addText(f.icon, { x: x + 0.2, y: y + 0.08, w: 0.6, h: 0.5, fontSize: 22, align: "center", margin: 0 });
    s.addText(f.title, { x: x + 0.8, y: y + 0.08, w: 3.4, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(f.desc, { x: x + 0.8, y: y + 0.5, w: 3.4, h: 0.6, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 7: セルフチェック ― 片目テスト
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "自分でできるチェック ― 片目テスト");

  var steps = [
    { num: "1", title: "片目を手で覆う", desc: "右目を隠して左目だけで見る\n次に左目を隠して右目だけで見る", icon: "🫣" },
    { num: "2", title: "結果を確認", desc: "片目だと1つに見える → 両眼性複視\n片目でも二重に見える → 単眼性複視", icon: "✅" },
    { num: "3", title: "どの方向で二重になるか", desc: "上下左右どの方向を見た時に\n二重が強くなるかを確認する", icon: "↕️" },
  ];

  steps.forEach(function(st, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, C.primary);
    s.addText(st.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.85, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.icon, { x: 1.5, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addText(st.title, { x: 2.3, y: y + 0.1, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(st.desc, { x: 2.3, y: y + 0.55, w: 6.5, h: 0.45, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
  });

  s.addText("※ 結果は受診時に医師に伝えてください。診断の参考になります", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.35, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 8: 危険なサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな場合はすぐ受診 ― 危険なサイン", C.red);

  var signs = [
    { text: "突然の複視＋激しい頭痛（脳動脈瘤破裂の可能性）", icon: "🚨" },
    { text: "瞳孔の大きさが左右で違う（動眼神経の圧迫）", icon: "👁️" },
    { text: "複視＋まぶた下垂＋瞳孔散大（脳動脈瘤を強く疑う）", icon: "⚠️" },
    { text: "複視＋ろれつ障害＋片側の麻痺（脳卒中の可能性）", icon: "🧠" },
    { text: "複視が急速に悪化している（数時間〜数日で進行）", icon: "📈" },
  ];

  signs.forEach(function(sign, i) {
    var y = 1.15 + i * 0.78;
    card(s, 0.5, y, 9.0, 0.65, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.05, w: 0.6, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.55, fontSize: 19, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.0, 5.0, 8.0, 0.2, C.red);
  ftr(s);
})();

// ============================================================
// SLIDE 9: 受診の目安 ― 緊急度3段階
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");

  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "突然の複視＋頭痛 / 瞳孔異常\nろれつ障害・片側麻痺を伴う", action: "脳動脈瘤・脳卒中の\n可能性→119番", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "数日以上続く / 悪化傾向\nまぶた下垂を伴う", action: "脳神経内科を受診\n原因精査が必要", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "疲れた時だけ一時的に見える\n片目でも二重（乱視など）", action: "眼科で視力チェック\n悪化したら再受診", bg: C.lightGreen, color: C.green },
  ];

  levels.forEach(function(lv, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, lv.bg, lv.color);
    s.addText(lv.emoji + " " + lv.label, { x: 0.8, y: y + 0.05, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    s.addText(lv.desc, { x: 0.8, y: y + 0.5, w: 5.0, h: 0.55, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText(lv.action, { x: 6.0, y: y + 0.15, w: 3.3, h: 0.8, fontSize: 16, fontFace: FJ, color: lv.color, bold: true, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 10: 必要な検査
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "複視の原因を調べる検査");

  var tests = [
    { icon: "🧠", name: "頭部MRI / MRA", targets: "脳動脈瘤・脳腫瘍・脳卒中の評価", note: "両眼性複視では必須" },
    { icon: "🩸", name: "血液検査", targets: "糖尿病・甲状腺機能・抗AChR抗体", note: "重症筋無力症の抗体検査" },
    { icon: "👁️", name: "眼科検査", targets: "視力・眼位・眼球運動の評価", note: "単眼性/両眼性の鑑別" },
    { icon: "⚡", name: "反復刺激試験", targets: "神経筋接合部の伝達障害", note: "重症筋無力症の電気生理検査" },
    { icon: "💉", name: "テンシロンテスト", targets: "抗コリンエステラーゼ薬の効果確認", note: "重症筋無力症の確認" },
  ];

  tests.forEach(function(t, i) {
    var y = 1.05 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(t.icon, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.55, fontSize: 22, align: "center", margin: 0 });
    s.addText(t.name, { x: 1.4, y: y + 0.08, w: 2.8, h: 0.55, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(t.targets, { x: 4.3, y: y + 0.08, w: 4.0, h: 0.35, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText(t.note, { x: 4.3, y: y + 0.4, w: 4.0, h: 0.25, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 11: よくある質問
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある質問 Q&A");

  var qa = [
    { q: "疲れると物が二重に見えるのは病気ですか？", a: "疲労による一時的な複視は多くの場合心配不要です。ただし頻繁に繰り返す、夕方に悪化する場合は重症筋無力症の可能性があり受診をおすすめします" },
    { q: "糖尿病で目が二重に見えることがありますか？", a: "はい。糖尿病性の眼球運動神経麻痺は比較的多く、特に外転神経麻痺が起きやすいです。多くは3〜6ヶ月で自然回復しますが、血糖コントロールが重要です" },
    { q: "複視は治りますか？", a: "原因によります。糖尿病性は自然回復が多く、重症筋無力症は薬で改善、脳動脈瘤は手術で治療できます。原因の特定が治療の第一歩です" },
  ];

  qa.forEach(function(item, i) {
    var y = 1.1 + i * 1.35;
    card(s, 0.5, y, 9.0, 1.15, C.white, C.accent);
    s.addText("Q.", { x: 0.8, y: y + 0.05, w: 0.5, h: 0.45, fontSize: 22, fontFace: FE, color: C.accent, bold: true, margin: 0 });
    s.addText(item.q, { x: 1.3, y: y + 0.05, w: 8.0, h: 0.45, fontSize: 18, fontFace: FJ, color: C.accent, bold: true, valign: "middle", margin: 0 });
    s.addText("A.", { x: 0.8, y: y + 0.55, w: 0.5, h: 0.5, fontSize: 22, fontFace: FE, color: C.primary, bold: true, margin: 0 });
    s.addText(item.a, { x: 1.3, y: y + 0.55, w: 8.0, h: 0.55, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.1, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 物が二重に見えたら");

  var points = [
    { num: "1", text: "片目テストで\nタイプを確認", sub: "片目で1つに見える→両眼性\n脳や神経の問題の可能性" },
    { num: "2", text: "突然の複視＋頭痛は\n緊急のサイン", sub: "脳動脈瘤・脳卒中の可能性\nすぐに119番" },
    { num: "3", text: "夕方に悪化する複視は\n重症筋無力症かも", sub: "まぶた下垂を伴うなら\n脳神経内科で精査を" },
  ];

  points.forEach(function(p, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, C.primary);
    s.addText(p.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.9, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.text, { x: 1.5, y: y + 0.05, w: 3.8, h: 0.95, fontSize: 19, fontFace: FJ, color: C.text, bold: true, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText(p.sub, { x: 5.5, y: y + 0.1, w: 3.8, h: 0.9, fontSize: 15, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 13: エンディング
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("ご視聴ありがとうございました", {
    x: 0.6, y: 0.8, w: 8.8, h: 0.7,
    fontSize: 30, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 2.5, y: 1.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("チャンネル登録・高評価よろしくお願いします！", {
    x: 0.6, y: 2.0, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });

  var links = [
    "🔗 ブログ記事で詳しく読む（概要欄にリンク）",
    "🔍 見え方セルフチェックツール（概要欄にリンク）",
    "📺 次回：「むせやすくなった」は老化のせい？ ― 嚥下障害の早期サイン",
  ];

  links.forEach(function(link, i) {
    s.addText(link, {
      x: 1.5, y: 2.9 + i * 0.55, w: 7.0, h: 0.45,
      fontSize: 17, fontFace: FJ, color: "B0BEC5", align: "left", margin: 0,
    });
  });

  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/diplopia_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
