var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#14】味がしない・味が変わった ― 味覚障害の原因";

var C = {
  dark: "1B3A5C", primary: "2C5AA0", accent: "E8850C", light: "E8F4FD",
  warmBg: "F8F9FA", white: "FFFFFF", text: "2D3436", sub: "636E72",
  red: "DC3545", green: "28A745", yellow: "F5C518",
  lightRed: "F8D7DA", lightYellow: "FFF3CD", lightGreen: "D4EDDA",
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
  if (borderColor) { s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.12, h: h, fill: { color: borderColor }, rectRadius: 0.05 }); }
}
function ftr(s) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.38, fill: { color: C.dark } });
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #14", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("👅", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("味がしない・味が変わった\n味覚障害の原因", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.15, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("コロナだけじゃない！\n脳神経内科医が解説", {
    x: 0.6, y: 2.9, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #14", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.7, w: 8.8, h: 0.4,
    fontSize: 12, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// SLIDE 2: こんな経験ありませんか？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな経験、ありませんか？");
  var items = [
    { emoji: "🍽️", text: "食事の味がしなくなった" },
    { emoji: "🫥", text: "何を食べても味が薄く感じる" },
    { emoji: "😖", text: "口の中がずっと苦い・金属のような味がする" },
    { emoji: "🍰", text: "好きだった食べ物の味が変わった" },
    { emoji: "😞", text: "味がわからないので食事が楽しくない" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 味覚のしくみ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "味覚のしくみ ― 舌から脳まで");
  var steps = [
    { icon: "👅", name: "味蕾が感知", desc: "舌の表面にある約5000〜1万個のセンサーが\n甘味・塩味・酸味・苦味・うま味を感知", color: C.primary },
    { icon: "🔌", name: "味覚神経が伝達", desc: "舌の前方＝顔面神経、後方＝舌咽神経\nのどの奥＝迷走神経の3本が脳に情報を送る", color: C.accent },
    { icon: "🧠", name: "大脳が処理", desc: "味覚野で信号を処理し「味」として認識\n嗅覚の情報も統合され「風味」を感じる", color: "7B2D8E" },
    { icon: "⚠️", name: "どこかに問題が", desc: "味蕾・神経・脳のどこかに問題が起こると\n味覚障害が発生する", color: C.red },
  ];
  steps.forEach(function(st, i) {
    var y = 1.1 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, st.color);
    s.addText(st.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 28, align: "center", margin: 0 });
    s.addText(st.name, { x: 1.4, y: y + 0.05, w: 2.5, h: 0.4, fontSize: 19, fontFace: FJ, color: st.color, bold: true, margin: 0 });
    s.addText(st.desc, { x: 4.0, y: y + 0.05, w: 5.3, h: 0.75, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 4: 味覚障害の種類
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "味覚障害の種類");
  var types = [
    { icon: "📉", name: "味覚低下", desc: "味が薄く感じる・味がわかりにくくなった状態。最も多いタイプ", color: C.primary },
    { icon: "🚫", name: "味覚消失", desc: "まったく味がしなくなった状態。コロナ後遺症で注目された", color: C.red },
    { icon: "🔄", name: "異味症", desc: "本来の味とは違う味に感じる。甘いものが苦く感じるなど", color: C.accent },
    { icon: "😣", name: "自発性異常味覚", desc: "何も食べていないのに口の中に苦味・金属味を感じる", color: "7B2D8E" },
    { icon: "🎯", name: "解離性味覚障害", desc: "5つの基本味のうち、特定の味だけがわからなくなる", color: C.green },
  ];
  types.forEach(function(t, i) {
    var y = 1.0 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, t.color);
    s.addText(t.icon, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(t.name, { x: 1.4, y: y + 0.05, w: 2.8, h: 0.6, fontSize: 18, fontFace: FJ, color: t.color, bold: true, valign: "middle", margin: 0 });
    s.addText(t.desc, { x: 4.3, y: y + 0.05, w: 5.0, h: 0.6, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 亜鉛不足
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "味覚障害の原因 ― 意外と多い亜鉛不足");
  // メインメッセージ
  card(s, 0.5, 1.1, 9.0, 0.7, C.light, C.primary);
  s.addText("味覚障害全体の約3割が亜鉛欠乏に関連", { x: 0.8, y: 1.15, w: 8.5, h: 0.3, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("味蕾の細胞は約10日で入れ替わる → 亜鉛が不足すると新しい味覚細胞が作れない", { x: 0.8, y: 1.5, w: 8.5, h: 0.25, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  // 亜鉛不足になりやすい人
  s.addText("亜鉛不足になりやすい方", { x: 0.5, y: 2.1, w: 9.0, h: 0.35, fontSize: 17, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var risks = [
    { emoji: "🍔", text: "偏った食事をしている方" },
    { emoji: "👴", text: "高齢で食事量が減った方" },
    { emoji: "💊", text: "降圧薬・利尿薬など特定の薬を服用中の方" },
    { emoji: "🍺", text: "過度のアルコール摂取がある方" },
  ];
  risks.forEach(function(r, i) {
    var y = 2.55 + i * 0.65;
    card(s, 0.5, y, 9.0, 0.55, C.white, C.accent);
    s.addText(r.emoji, { x: 0.7, y: y + 0.05, w: 0.6, h: 0.45, fontSize: 22, align: "center", margin: 0 });
    s.addText(r.text, { x: 1.4, y: y + 0.05, w: 7.8, h: 0.45, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  s.addText("※ 血液検査で亜鉛の値を調べることができます", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.3, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0,
  });
  ftr(s);
})();

// SLIDE 6: 薬剤性味覚障害
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "薬剤性味覚障害 ― あの薬が原因？");
  card(s, 0.5, 1.1, 9.0, 0.5, C.lightYellow, C.yellow);
  s.addText("⚠️ 薬剤性味覚障害は全体の約2割を占める", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 18, fontFace: FJ, color: "856404", bold: true, valign: "middle", margin: 0 });
  // 原因薬剤
  var drugs = [
    { name: "降圧薬", desc: "ACE阻害薬・ARB" },
    { name: "抗生物質", desc: "クラリスロマイシン・メトロニダゾール" },
    { name: "抗がん剤", desc: "各種化学療法薬" },
    { name: "抗うつ薬", desc: "SSRI・三環系" },
    { name: "利尿薬", desc: "チアジド系など" },
    { name: "抗アレルギー薬", desc: "抗ヒスタミン薬" },
  ];
  drugs.forEach(function(d, i) {
    var col = i < 3 ? 0 : 1;
    var row = i % 3;
    var x = 0.5 + col * 4.7;
    var y = 1.8 + row * 0.65;
    card(s, x, y, 4.4, 0.55, C.white, C.primary);
    s.addText(d.name, { x: x + 0.2, y: y + 0.05, w: 1.8, h: 0.45, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(d.desc, { x: x + 2.0, y: y + 0.05, w: 2.2, h: 0.45, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  // 注意メッセージ
  card(s, 0.5, 3.85, 9.0, 0.8, C.lightRed, C.red);
  s.addText("新しい薬を飲み始めてから味がおかしいと感じたら", { x: 0.8, y: 3.9, w: 8.5, h: 0.35, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("自己判断で中止せず、必ず処方した医師に相談してください。\n多くの場合、薬を開始して数週間〜数ヶ月で症状が出ます。", {
    x: 0.8, y: 4.25, w: 8.5, h: 0.35, fontSize: 13, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
  });
  ftr(s);
})();

// SLIDE 7: コロナ後遺症と味覚障害
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "コロナ後遺症と味覚障害");
  // 統計
  card(s, 0.5, 1.1, 9.0, 0.6, C.light, C.primary);
  s.addText("コロナ感染者の約40〜60%に味覚・嗅覚障害が報告", { x: 0.8, y: 1.15, w: 8.5, h: 0.5, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
  // メカニズム
  var mech = [
    { icon: "🦠", name: "ウイルスが鼻の奥に感染", desc: "嗅覚神経の支持細胞がダメージを受ける", color: C.red },
    { icon: "👃", name: "嗅覚が低下", desc: "味覚と嗅覚は密接に関連しているため\n「味がしない」と感じることが多い", color: C.accent },
    { icon: "⏳", name: "1〜3ヶ月で自然回復が多い", desc: "一部では半年以上続く場合も\n回復が遅い場合は嗅覚トレーニングが有効", color: C.green },
  ];
  mech.forEach(function(m, i) {
    var y = 1.95 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, m.color);
    s.addText(m.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 28, align: "center", margin: 0 });
    s.addText(m.name, { x: 1.4, y: y + 0.05, w: 3.2, h: 0.4, fontSize: 18, fontFace: FJ, color: m.color, bold: true, margin: 0 });
    s.addText(m.desc, { x: 4.7, y: y + 0.05, w: 4.6, h: 0.75, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 8: その他の原因
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "その他の原因 ― 神経・全身疾患・加齢");
  // 3カラム
  var cols = [
    { title: "口腔の問題", color: C.accent, items: ["口腔乾燥症", "舌炎", "口腔カンジダ症", "唾液減少で味物質が\n味蕾に届きにくい"] },
    { title: "神経の問題", color: "7B2D8E", items: ["顔面神経麻痺", "脳卒中", "脳腫瘍", "味覚野の障害"] },
    { title: "全身疾患・加齢", color: C.primary, items: ["糖尿病", "腎不全・肝障害", "甲状腺機能低下症", "70代で味蕾は\n若い頃の約1/3に"] },
  ];
  cols.forEach(function(col, ci) {
    var x = 0.3 + ci * 3.2;
    card(s, x, 1.1, 3.0, 0.45, col.color);
    s.addText(col.title, { x: x + 0.1, y: 1.13, w: 2.8, h: 0.4, fontSize: 17, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    col.items.forEach(function(item, ii) {
      var y = 1.7 + ii * 0.75;
      card(s, x, y, 3.0, 0.65, C.white, col.color);
      s.addText(item, { x: x + 0.2, y: y + 0.05, w: 2.6, h: 0.55, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
    });
  });
  ftr(s);
})();

// SLIDE 9: 危険なサイン
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "危険なサイン ― こんな場合はすぐ受診", C.red);
  var signs = [
    { text: "突然の味覚消失＋嗅覚消失\n（感染症の可能性）", icon: "🦠" },
    { text: "味覚障害＋顔の片側が動かない\n（顔面神経麻痺 → 早期治療が重要）", icon: "😐" },
    { text: "味覚障害＋ろれつ困難＋手足のしびれ\n（脳卒中の可能性 → 119番）", icon: "🚑" },
    { text: "体重減少を伴う味覚障害\n（栄養不良のおそれ）", icon: "⚖️" },
  ];
  signs.forEach(function(sign, i) {
    var y = 1.1 + i * 0.95;
    card(s, 0.5, y, 9.0, 0.8, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.1, w: 0.6, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.7, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 10: 受診の目安
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "味覚障害＋ろれつ困難・手足のしびれ\n味覚障害＋顔面麻痺", action: "脳卒中の可能性\n→119番", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "味覚障害が2週間以上続く\n新しい薬の開始後に味がおかしい\n食事量減少・体重減少", action: "耳鼻咽喉科を\n受診してください", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "風邪の後に味が薄いが徐々に改善\n口が乾いている時だけ味がわかりにくい", action: "水分をしっかりとり\n1〜2週間様子を見る", bg: C.lightGreen, color: C.green },
  ];
  levels.forEach(function(lv, i) {
    var y = 1.15 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, lv.bg, lv.color);
    s.addText(lv.emoji + " " + lv.label, { x: 0.8, y: y + 0.05, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    s.addText(lv.desc, { x: 0.8, y: y + 0.5, w: 5.0, h: 0.55, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText(lv.action, { x: 6.0, y: y + 0.15, w: 3.3, h: 0.8, fontSize: 16, fontFace: FJ, color: lv.color, bold: true, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 11: 検査と治療
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "味覚障害の検査と治療");
  // 検査
  s.addText("検査", { x: 0.5, y: 1.05, w: 4.3, h: 0.35, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var tests = [
    { name: "電気味覚検査・ろ紙ディスク法", desc: "味覚の感度を客観的に測定" },
    { name: "血液検査", desc: "亜鉛値の確認" },
    { name: "頭部MRI", desc: "神経・脳の異常を調べる（必要時）" },
  ];
  tests.forEach(function(t, i) {
    var y = 1.5 + i * 0.65;
    card(s, 0.5, y, 4.3, 0.55, C.white, C.primary);
    s.addText(t.name, { x: 0.7, y: y + 0.03, w: 3.8, h: 0.25, fontSize: 14, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(t.desc, { x: 0.7, y: y + 0.28, w: 3.8, h: 0.22, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 治療
  s.addText("治療", { x: 5.2, y: 1.05, w: 4.3, h: 0.35, fontSize: 17, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  var treatments = [
    { name: "亜鉛製剤の内服", desc: "2〜6ヶ月の補充で約60〜70%が改善" },
    { name: "原因薬の変更", desc: "薬剤性の場合は処方を見直す" },
    { name: "嗅覚トレーニング", desc: "コロナ後遺症に有効" },
    { name: "口腔乾燥の治療", desc: "唾液分泌を促す薬・ケア" },
  ];
  treatments.forEach(function(t, i) {
    var y = 1.5 + i * 0.65;
    card(s, 5.2, y, 4.3, 0.55, C.white, C.accent);
    s.addText(t.name, { x: 5.4, y: y + 0.03, w: 3.8, h: 0.25, fontSize: 14, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
    s.addText(t.desc, { x: 5.4, y: y + 0.28, w: 3.8, h: 0.22, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  // 亜鉛補充の強調
  card(s, 0.5, 4.25, 9.0, 0.65, C.lightGreen, C.green);
  s.addText("亜鉛製剤の補充が基本治療 ― 原因を特定し適切な治療につなげることが大切です", {
    x: 0.8, y: 4.3, w: 8.5, h: 0.55, fontSize: 16, fontFace: FJ, color: C.green, bold: true, valign: "middle", margin: 0,
  });
  ftr(s);
})();

// SLIDE 12: まとめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 味覚障害の原因と対処");
  var points = [
    { num: "1", text: "味覚障害は\n「コロナだけ」ではない", sub: "亜鉛不足・薬の副作用・口腔の問題\n神経疾患など原因は多岐にわたる" },
    { num: "2", text: "最も多い原因は\n亜鉛不足", sub: "血液検査で確認でき\n亜鉛製剤の補充で多くの方が改善" },
    { num: "3", text: "2週間以上続けば\n耳鼻咽喉科へ", sub: "原因を特定し\n適切な治療につなげることが大切" },
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

// SLIDE 13: エンディング
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("ご視聴ありがとうございました", {
    x: 0.6, y: 0.6, w: 8.8, h: 0.7,
    fontSize: 30, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 1.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("チャンネル登録・高評価よろしくお願いします！", {
    x: 0.6, y: 1.7, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("🔗 ブログ記事で詳しく読む（概要欄にリンク）", {
    x: 1.5, y: 2.5, w: 7.0, h: 0.4, fontSize: 16, fontFace: FJ, color: "B0BEC5", margin: 0,
  });
  s.addText("次回予告", {
    x: 0.6, y: 3.3, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: C.accent, bold: true, align: "center", margin: 0,
  });
  s.addText("#15 眠れないのは脳のせい？ ― 不眠のメカニズムと正しい対処法", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/taste_disorder_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
