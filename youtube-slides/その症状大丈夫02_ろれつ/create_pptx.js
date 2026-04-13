var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#02】ろれつが回らない…脳のSOS？";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #02", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("🗣️", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("ろれつが回りにくい…\nこれって脳のSOS？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.4,
    fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.9, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("脳神経内科医が教える\n緊急性の見分け方", {
    x: 0.6, y: 3.1, w: 8.8, h: 0.9,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #02", {
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
    { emoji: "🗣️", text: "話していて舌がもつれる感じがする" },
    { emoji: "📞", text: "電話で「酔ってるの？」と聞かれた" },
    { emoji: "🍽️", text: "食事中に口の端から食べ物がこぼれる" },
    { emoji: "💬", text: "言いたい言葉がうまく発音できない" },
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
// SLIDE 3: ろれつが回らないとは？（構音障害）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「ろれつが回らない」とは？");

  card(s, 0.5, 1.1, 9.0, 1.2, C.light, C.primary);
  s.addText("構音障害（こうおんしょうがい）", { x: 0.8, y: 1.15, w: 6, h: 0.45, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("言葉を発する筋肉や神経の問題で\n発音が不明瞭になる状態", {
    x: 0.8, y: 1.55, w: 8.5, h: 0.65, fontSize: 20, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  card(s, 0.5, 2.6, 4.2, 1.5, C.white, C.accent);
  s.addText("失語症との違い", { x: 0.8, y: 2.7, w: 3.8, h: 0.5, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("構音障害: 言葉は浮かぶが\n口がうまく動かない", { x: 0.8, y: 3.2, w: 3.8, h: 0.4, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
  s.addText("失語症: 言葉自体が出てこない", { x: 0.8, y: 3.65, w: 3.8, h: 0.35, fontSize: 16, fontFace: FJ, color: C.sub, margin: 0 });

  card(s, 5.3, 2.6, 4.2, 1.5, C.white, C.red);
  s.addText("最重要ポイント", { x: 5.6, y: 2.7, w: 3.8, h: 0.5, fontSize: 20, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("「突然」か「徐々に」か\nで緊急性が全く違う", {
    x: 5.6, y: 3.2, w: 3.8, h: 0.8, fontSize: 20, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 4: 突然 vs 徐々に
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "緊急性の分類 ― 突然 vs 徐々に");

  // 左カード: 突然
  card(s, 0.4, 1.15, 4.5, 3.8, C.lightRed, C.red);
  s.addText("⚡ 突然発症", { x: 0.7, y: 1.25, w: 4.0, h: 0.5, fontSize: 24, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.7, y: 1.8, w: 3.8, h: 0, line: { color: C.red, width: 1 } });
  s.addText("数分〜数時間で発症", { x: 0.7, y: 1.9, w: 4.0, h: 0.4, fontSize: 18, fontFace: FJ, color: C.text, margin: 0 });
  var sudden = ["脳梗塞", "脳出血", "一過性脳虚血発作（TIA）"];
  sudden.forEach(function(item, i) {
    s.addText("● " + item, { x: 0.7, y: 2.4 + i * 0.45, w: 4.0, h: 0.4, fontSize: 18, fontFace: FJ, color: C.text, margin: 0 });
  });
  s.addText("→ 119番！ 一刻を争う", { x: 0.7, y: 3.85, w: 4.0, h: 0.5, fontSize: 20, fontFace: FJ, color: C.red, bold: true, margin: 0 });

  // 右カード: 徐々に
  card(s, 5.1, 1.15, 4.5, 3.8, C.light, C.primary);
  s.addText("📈 徐々に進行", { x: 5.4, y: 1.25, w: 4.0, h: 0.5, fontSize: 24, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 5.4, y: 1.8, w: 3.8, h: 0, line: { color: C.primary, width: 1 } });
  s.addText("数週間〜数ヶ月で進行", { x: 5.4, y: 1.9, w: 4.0, h: 0.4, fontSize: 18, fontFace: FJ, color: C.text, margin: 0 });
  var gradual = ["ALS（筋萎縮性側索硬化症）", "パーキンソン病", "多発性硬化症", "重症筋無力症"];
  gradual.forEach(function(item, i) {
    s.addText("● " + item, { x: 5.4, y: 2.4 + i * 0.45, w: 4.0, h: 0.4, fontSize: 17, fontFace: FJ, color: C.text, margin: 0 });
  });
  s.addText("→ 早めに脳神経内科へ", { x: 5.4, y: 4.3, w: 4.0, h: 0.5, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });

  ftr(s);
})();

// ============================================================
// SLIDE 5: 突然のろれつ障害 ― 脳卒中
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "突然のろれつ障害 ― 脳卒中を疑え！", C.red);

  card(s, 0.5, 1.1, 9.0, 1.0, C.lightRed, C.red);
  s.addText("「さっきまで普通だったのに急にしゃべりにくくなった」= 脳の緊急事態", {
    x: 0.8, y: 1.2, w: 8.5, h: 0.8, fontSize: 20, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0,
  });

  var facts = [
    "脳梗塞は日本人の死因第4位、要介護の原因第1位",
    "発症から4.5時間以内なら血栓溶解療法（t-PA）が可能",
    "血管内治療は最大24時間以内が適応（条件による）",
    "「Time is Brain」― 1分あたり約190万個の神経細胞が失われる",
  ];

  facts.forEach(function(f, i) {
    var y = 2.4 + i * 0.65;
    s.addText("●", { x: 0.7, y: y, w: 0.4, h: 0.5, fontSize: 14, color: C.red, margin: 0 });
    s.addText(f, { x: 1.1, y: y, w: 8.2, h: 0.5, fontSize: 19, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.0, 4.95, 8.0, 0.25, C.red);

  ftr(s);
})();

// ============================================================
// SLIDE 6: FASTチェック
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "FASTチェック ― 脳卒中を見逃さない", C.red);

  var fast = [
    { letter: "F", word: "Face（顔）", desc: "笑顔を作ると片側が\n下がっていませんか？", color: C.red, bg: C.lightRed },
    { letter: "A", word: "Arms（腕）", desc: "両腕を上げると片方が\n下がってきませんか？", color: C.accent, bg: C.lightYellow },
    { letter: "S", word: "Speech（言葉）", desc: "「今日はいい天気」\nと言えますか？", color: C.primary, bg: C.light },
    { letter: "T", word: "Time（時間）", desc: "1つでも当てはまったら\nすぐ119番！", color: "6C3483", bg: "F5EEF8" },
  ];

  fast.forEach(function(f, i) {
    var x = 0.3 + i * 2.4;
    card(s, x, 1.15, 2.2, 3.7, f.bg, f.color);
    s.addText(f.letter, { x: x + 0.1, y: 1.25, w: 2.0, h: 0.8, fontSize: 48, fontFace: FE, color: f.color, bold: true, align: "center", margin: 0 });
    s.addText(f.word, { x: x + 0.1, y: 2.1, w: 2.0, h: 0.45, fontSize: 16, fontFace: FJ, color: f.color, bold: true, align: "center", margin: 0 });
    s.addShape(pres.shapes.LINE, { x: x + 0.3, y: 2.6, w: 1.6, h: 0, line: { color: f.color, width: 1 } });
    s.addText(f.desc, { x: x + 0.15, y: 2.7, w: 1.9, h: 0.9, fontSize: 14, fontFace: FJ, color: C.text, align: "center", lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 7: 徐々に進行するケース
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "徐々に進行するケース ― 原因疾患");

  var diseases = [
    { name: "ALS", full: "筋萎縮性側索硬化症", features: "舌の萎縮・線維束性収縮\n飲み込みにくさも出現", color: C.red, bg: C.lightRed },
    { name: "パーキンソン病", full: "", features: "声が小さくなる・単調になる\n手の震え・動作緩慢を伴う", color: C.primary, bg: C.light },
    { name: "多発性硬化症", full: "MS", features: "症状が出たり消えたりする\n20〜40代の女性に多い", color: C.accent, bg: C.lightYellow },
    { name: "重症筋無力症", full: "MG", features: "夕方に悪化するのが特徴\nまぶたの下がりを伴うことも", color: C.green, bg: C.lightGreen },
  ];

  diseases.forEach(function(d, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.15 + row * 2.0;
    card(s, x, y, 4.5, 1.75, d.bg, d.color);
    s.addText(d.name, { x: x + 0.25, y: y + 0.1, w: 4.0, h: 0.5, fontSize: 20, fontFace: FJ, color: d.color, bold: true, margin: 0 });
    if (d.full) {
      s.addText("（" + d.full + "）", { x: x + 0.25, y: y + 0.55, w: 4.0, h: 0.3, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0 });
    }
    var descY = d.full ? y + 0.8 : y + 0.6;
    s.addText(d.features, { x: x + 0.25, y: descY, w: 4.0, h: 0.8, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 8: その他の原因
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "その他の原因 ― 病気以外も多い");

  var causes = [
    { title: "疲労・睡眠不足", desc: "筋肉の協調が乱れる。\n休息で改善すれば心配不要", icon: "😴", color: C.primary, bg: C.light },
    { title: "薬の副作用", desc: "抗てんかん薬・精神安定剤・\n抗ヒスタミン薬など", icon: "💊", color: C.red, bg: C.lightRed },
    { title: "飲酒", desc: "アルコールは小脳機能を\n一時的に低下させる", icon: "🍺", color: C.accent, bg: C.lightYellow },
    { title: "義歯の不具合", desc: "入れ歯が合っていないと\n発音に影響する場合も", icon: "🦷", color: C.green, bg: C.lightGreen },
  ];

  causes.forEach(function(c, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = 0.4 + col * 4.8;
    var y = 1.15 + row * 2.0;
    card(s, x, y, 4.5, 1.75, c.bg, c.color);
    s.addText(c.icon, { x: x + 0.2, y: y + 0.15, w: 0.7, h: 0.6, fontSize: 28, margin: 0 });
    s.addText(c.title, { x: x + 0.9, y: y + 0.15, w: 3.3, h: 0.5, fontSize: 20, fontFace: FJ, color: c.color, bold: true, margin: 0 });
    s.addText(c.desc, { x: x + 0.9, y: y + 0.7, w: 3.3, h: 0.9, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 9: 危険なサイン
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "⚠️ こんな時はすぐ119番！", C.red);

  var signs = [
    "突然ろれつが回らなくなった（発症時刻を記録！）",
    "顔の片側が下がっている（口角・目が閉じにくい）",
    "片方の手足に力が入らない・しびれる",
    "激しい頭痛を伴う（経験したことのない痛み）",
    "意識がぼんやりする・呼びかけに反応が鈍い",
  ];

  signs.forEach(function(sign, i) {
    var y = 1.15 + i * 0.78;
    card(s, 0.5, y, 9.0, 0.65, C.lightRed, C.red);
    s.addText((i + 1) + ".", { x: 0.8, y: y + 0.05, w: 0.5, h: 0.55, fontSize: 22, fontFace: FE, color: C.red, bold: true, margin: 0 });
    s.addText(sign, { x: 1.3, y: y + 0.05, w: 7.9, h: 0.55, fontSize: 20, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 1.5, 5.05, 7.0, 0.15, C.red);

  ftr(s);
})();

// ============================================================
// SLIDE 10: 受診の目安 ― 緊急度3段階
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");

  var levels = [
    { emoji: "🔴", label: "すぐ119番（救急）", desc: "突然のろれつ障害\n顔・手足の麻痺を伴う", action: "発症時刻を記録して\n救急搬送", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "徐々にろれつが悪化\n飲み込みにくさも出てきた", action: "1〜2週間以内に\n脳神経内科を受診", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "疲労時のみ・飲酒後のみ\n翌日には回復する", action: "生活習慣を見直し\n悪化したら受診", bg: C.lightGreen, color: C.green },
  ];

  levels.forEach(function(lv, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, lv.bg, lv.color);
    s.addText(lv.emoji + " " + lv.label, { x: 0.8, y: y + 0.05, w: 3.5, h: 0.45, fontSize: 21, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    s.addText(lv.desc, { x: 0.8, y: y + 0.5, w: 5.0, h: 0.55, fontSize: 16, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText("→ " + lv.action, { x: 6.0, y: y + 0.15, w: 3.3, h: 0.8, fontSize: 16, fontFace: FJ, color: lv.color, bold: true, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });

  ftr(s);
})();

// ============================================================
// SLIDE 11: TIAに注意
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "要注意！「一時的に治った」場合 ― TIA");

  card(s, 0.5, 1.1, 9.0, 1.3, C.lightYellow, C.accent);
  s.addText("一過性脳虚血発作（TIA）", { x: 0.8, y: 1.15, w: 8, h: 0.5, fontSize: 22, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("ろれつ障害や麻痺が数分〜数十分で完全に回復する発作。\n「治ったから大丈夫」は大間違い！", {
    x: 0.8, y: 1.65, w: 8.5, h: 0.65, fontSize: 19, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  card(s, 0.5, 2.7, 4.2, 1.6, C.lightRed, C.red);
  s.addText("TIA後の脳梗塞リスク", { x: 0.8, y: 2.8, w: 3.8, h: 0.45, fontSize: 19, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("48時間以内: 約5%\n90日以内: 約10〜15%\nが脳梗塞を発症", {
    x: 0.8, y: 3.3, w: 3.8, h: 0.9, fontSize: 18, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  card(s, 5.3, 2.7, 4.2, 1.6, C.lightGreen, C.green);
  s.addText("早期受診で予防できる", { x: 5.6, y: 2.8, w: 3.8, h: 0.45, fontSize: 19, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("適切な治療で脳梗塞リスクを\n約80%減らせるという\n研究結果があります", {
    x: 5.6, y: 3.3, w: 3.8, h: 0.9, fontSize: 18, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });

  s.addText("※ 症状が消えても、必ずその日のうちに医療機関を受診してください", {
    x: 0.5, y: 4.6, w: 9.0, h: 0.35, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });

  ftr(s);
})();

// ============================================================
// SLIDE 12: まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― ろれつの異常で覚えておきたい3つ");

  var points = [
    { num: "1", text: "「突然」なら\n一刻を争う緊急事態", sub: "脳卒中の可能性\n → 迷わず119番" },
    { num: "2", text: "FASTで\n脳卒中をチェック", sub: "顔・腕・言葉の3つを確認\n1つでも異常 → すぐ救急" },
    { num: "3", text: "「治ったから大丈夫」\nは危険", sub: "TIAは脳梗塞の前触れ\n症状が消えても必ず受診" },
  ];

  points.forEach(function(p, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, C.primary);
    s.addText(p.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.9, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.text, { x: 1.5, y: y + 0.05, w: 3.8, h: 0.95, fontSize: 20, fontFace: FJ, color: C.text, bold: true, lineSpacingMultiple: 1.1, margin: 0 });
    s.addText(p.sub, { x: 5.5, y: y + 0.1, w: 3.8, h: 0.9, fontSize: 16, fontFace: FJ, color: C.sub, lineSpacingMultiple: 1.2, margin: 0 });
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
    "🔍 ろれつ緊急度セルフチェックツール（概要欄にリンク）",
    "📺 次回：突然手に力が入らない ― 原因の見分け方",
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
var outPath = __dirname + "/slurred_speech_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
