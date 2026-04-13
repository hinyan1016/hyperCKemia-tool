var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#12】頭をぶつけた後、病院に行くべき？";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #12", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🤕", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("頭をぶつけた後\n病院に行くべき？", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.15, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("経過観察のチェックポイントを\n脳神経内科医が解説", {
    x: 0.6, y: 2.9, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #12", {
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
    { emoji: "🚪", text: "棚の角に頭をぶつけた" },
    { emoji: "🧒", text: "子どもが転んで頭を打った" },
    { emoji: "🧓", text: "高齢の親が転倒して頭をぶつけた" },
    { emoji: "🤔", text: "病院に行くべきか迷った" },
    { emoji: "😰", text: "「たんこぶができたから大丈夫」は本当？" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 頭部打撲のメカニズム
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "頭部打撲で何が起こるか");
  var layers = [
    { icon: "🦴", name: "頭蓋骨", role: "脳を守る硬い殻", injury: "骨折 → 出血・感染のリスク", color: C.primary },
    { icon: "🧠", name: "脳", role: "頭蓋骨の中で浮いている", injury: "脳震盪 → 一時的な脳機能障害", color: C.accent },
    { icon: "🩸", name: "硬膜と血管", role: "脳を包む膜と栄養血管", injury: "硬膜下血腫・硬膜外血腫", color: C.red },
    { icon: "⚡", name: "脳の表面", role: "衝撃で反対側にもぶつかる", injury: "脳挫傷（対側損傷）", color: "7B2D8E" },
  ];
  layers.forEach(function(l, i) {
    var y = 1.1 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, l.color);
    s.addText(l.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 28, align: "center", margin: 0 });
    s.addText(l.name, { x: 1.4, y: y + 0.05, w: 1.8, h: 0.4, fontSize: 19, fontFace: FJ, color: l.color, bold: true, margin: 0 });
    s.addText(l.role, { x: 3.3, y: y + 0.05, w: 3.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    s.addText(l.injury, { x: 3.3, y: y + 0.45, w: 5.5, h: 0.35, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 4: すぐ病院に行くべきケース
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "すぐ病院に行くべきケース", C.red);
  var signs = [
    { text: "意識を失った（一瞬でも）", icon: "😵" },
    { text: "激しい頭痛が続く・どんどん悪化する", icon: "🤯" },
    { text: "繰り返し吐く（2回以上）", icon: "🤢" },
    { text: "けいれんが起きた", icon: "⚡" },
    { text: "手足に力が入らない・しびれる", icon: "🖐️" },
    { text: "耳や鼻から透明な液体や血が出る", icon: "👂" },
  ];
  signs.forEach(function(sign, i) {
    var y = 1.05 + i * 0.68;
    card(s, 0.5, y, 9.0, 0.56, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.03, w: 0.6, h: 0.5, fontSize: 22, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.03, w: 7.7, h: 0.5, fontSize: 19, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  s.addText("※ 1つでも当てはまれば → 119番 or すぐ救急外来へ", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.2, fontSize: 13, fontFace: FJ, color: C.red, bold: true, margin: 0,
  });
  ftr(s);
})();

// SLIDE 5: 経過観察のチェックポイント（48時間ルール）
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "経過観察のチェックポイント ― 48時間ルール");
  s.addText("頭をぶつけた後、特に最初の48時間は注意深い観察が必要です", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.35, fontSize: 15, fontFace: FJ, color: C.accent, bold: true, margin: 0,
  });
  var checks = [
    { time: "直後〜6時間", items: "意識レベル・頭痛・嘔吐・瞳孔の左右差\n最もリスクが高い時間帯", color: C.red },
    { time: "6〜24時間", items: "頭痛の変化・眠気の程度・言動の異常\n数時間おきに声をかけて確認", color: C.accent },
    { time: "24〜48時間", items: "症状が悪化しないか確認\n新たな症状が出ていないか", color: C.primary },
    { time: "48時間以降", items: "急性期リスクは大幅に低下\nただし慢性硬膜下血腫に注意（数週間後）", color: C.green },
  ];
  checks.forEach(function(ch, i) {
    var y = 1.55 + i * 0.88;
    card(s, 0.5, y, 9.0, 0.75, C.white, ch.color);
    s.addText(ch.time, { x: 0.8, y: y + 0.08, w: 2.2, h: 0.6, fontSize: 17, fontFace: FJ, color: ch.color, bold: true, valign: "middle", margin: 0 });
    s.addText(ch.items, { x: 3.1, y: y + 0.08, w: 6.2, h: 0.6, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 6: たんこぶの誤解
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「たんこぶ」の誤解と真実");
  var myths = [
    { myth: "たんこぶができたから大丈夫", truth: "たんこぶは皮下の出血。脳内の出血とは無関係。\nたんこぶがあっても脳内出血の可能性はある", icon: "❌" },
    { myth: "たんこぶができないほうが危ない", truth: "必ずしも正しくない。衝撃の方向や部位による。\nたんこぶの有無だけで判断してはいけない", icon: "❌" },
    { myth: "眠らせてはいけない", truth: "眠ること自体は問題ない。ただし定期的に\n声をかけて意識を確認することが重要", icon: "⭕" },
  ];
  myths.forEach(function(m, i) {
    var y = 1.1 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, C.accent);
    s.addText(m.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.4, fontSize: 26, align: "center", margin: 0 });
    s.addText(m.myth, { x: 1.4, y: y + 0.05, w: 7.8, h: 0.45, fontSize: 18, fontFace: FJ, color: C.accent, bold: true, valign: "middle", margin: 0 });
    s.addText(m.truth, { x: 1.4, y: y + 0.5, w: 7.8, h: 0.55, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 7: 高齢者の注意点 ― 慢性硬膜下血腫
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "高齢者の注意点 ― 慢性硬膜下血腫");
  s.addText("軽い頭部打撲の数週間〜数ヶ月後に症状が出る ― 手術で治る病気", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.35, fontSize: 15, fontFace: FJ, color: C.accent, bold: true, margin: 0,
  });
  var points = [
    { icon: "🧓", title: "高齢者に多い", desc: "脳の萎縮で血管が伸展され\n軽い衝撃でも破れやすい", color: C.primary },
    { icon: "📅", title: "数週間〜数ヶ月後に発症", desc: "頭をぶつけたことを忘れている\nことも多い", color: C.accent },
    { icon: "🧠", title: "認知症と間違われやすい", desc: "物忘れ・意欲低下・歩行障害\n→ 認知症やうつ病と誤診される", color: C.red },
    { icon: "🏥", title: "手術で改善可能", desc: "穿頭血腫洗浄術で血腫を除去\n多くの場合は速やかに回復", color: C.green },
  ];
  points.forEach(function(p, i) {
    var y = 1.55 + i * 0.88;
    card(s, 0.5, y, 9.0, 0.75, C.white, p.color);
    s.addText(p.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(p.title, { x: 1.4, y: y + 0.05, w: 2.5, h: 0.65, fontSize: 18, fontFace: FJ, color: p.color, bold: true, valign: "middle", margin: 0 });
    s.addText(p.desc, { x: 4.0, y: y + 0.08, w: 5.3, h: 0.6, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 8: 子どもの頭部打撲
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "子どもの頭部打撲 ― 親が知るべきポイント");
  var checks = [
    { num: "1", title: "受傷の状況を確認", desc: "どこから落ちた？高さは？何にぶつけた？\n高さ90cm以上（2歳未満）/ 150cm以上（2歳以上）は要注意", icon: "📝" },
    { num: "2", title: "年齢で対応が変わる", desc: "2歳未満は症状を訴えられない\n機嫌・哺乳・活動性の変化を観察", icon: "👶" },
    { num: "3", title: "嘔吐の回数がカギ", desc: "1回の嘔吐は経過観察可\n2回以上の繰り返す嘔吐は受診が必要", icon: "🤢" },
  ];
  checks.forEach(function(ch, i) {
    var y = 1.1 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.05, C.white, C.primary);
    s.addText(ch.num, { x: 0.7, y: y + 0.1, w: 0.7, h: 0.85, fontSize: 36, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(ch.icon, { x: 1.5, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 30, align: "center", margin: 0 });
    s.addText(ch.title, { x: 2.3, y: y + 0.1, w: 3.0, h: 0.45, fontSize: 21, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });
    s.addText(ch.desc, { x: 2.3, y: y + 0.55, w: 6.5, h: 0.45, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.1, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 9: 危険なサイン（時間経過別）
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "時間が経ってから出る危険なサイン", C.red);
  var signs = [
    { text: "頭痛がどんどん強くなる", icon: "📈" },
    { text: "ぼーっとして会話がかみ合わない", icon: "💭" },
    { text: "片側の手足が動かしにくくなった", icon: "🖐️" },
    { text: "いつもと様子が違う（興奮・無気力）", icon: "😶" },
    { text: "数週間後に物忘れ・歩行障害が出てきた", icon: "🧓" },
  ];
  signs.forEach(function(sign, i) {
    var y = 1.15 + i * 0.78;
    card(s, 0.5, y, 9.0, 0.65, C.lightRed, C.red);
    s.addText(sign.icon, { x: 0.8, y: y + 0.05, w: 0.6, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(sign.text, { x: 1.5, y: y + 0.05, w: 7.7, h: 0.55, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  card(s, 1.0, 5.0, 8.0, 0.2, C.red);
  ftr(s);
})();

// SLIDE 10: 受診の目安
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "意識消失・激しい頭痛・繰り返す嘔吐\nけいれん・手足の麻痺・耳鼻からの液体", action: "119番で救急搬送\n→ 頭部CT", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "頭痛が続く・軽い吐き気がある\n一時的にぼーっとした\n大きなたんこぶがある", action: "当日中に\n脳神経外科・救急へ", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "軽くぶつけただけ・症状なし\n泣いた後すぐに普段通り（子ども）", action: "48時間は経過観察\n症状出現時は受診", bg: C.lightGreen, color: C.green },
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

// SLIDE 11: Q&A
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "よくある質問 Q&A");
  var qa = [
    { q: "頭をぶつけたら必ずCTを撮るべき？", a: "全員が必要なわけではありません。医師が症状や受傷の状況から判断します。特に小児では不要な被ばくを避けるため、観察で十分な場合も多いです" },
    { q: "お酒を飲んで転んだ場合は？", a: "飲酒時の転倒は頭部外傷の大きなリスク因子です。酔いの症状と脳の症状が区別しにくいため、迷ったら受診することをおすすめします" },
    { q: "血液サラサラの薬を飲んでいますが？", a: "抗凝固薬（ワーファリン等）を服用中の方は出血リスクが高く、軽い打撲でも頭蓋内出血のリスクがあります。必ず受診してください" },
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

// SLIDE 12: まとめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ ― 頭をぶつけた後のチェック");
  var points = [
    { num: "1", text: "最初の48時間が\n最も重要", sub: "意識・頭痛・嘔吐・瞳孔を\n定期的にチェック" },
    { num: "2", text: "高齢者は数週間後にも\n注意が必要", sub: "慢性硬膜下血腫は\n手術で治る ― 見逃さない" },
    { num: "3", text: "迷ったら受診\n特に抗凝固薬服用中", sub: "たんこぶの有無では\n脳内の状態はわからない" },
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
  s.addText("🔍 頭部打撲セルフチェックツール（概要欄にリンク）", {
    x: 1.5, y: 2.95, w: 7.0, h: 0.4, fontSize: 16, fontFace: FJ, color: "B0BEC5", margin: 0,
  });
  s.addText("次回予告", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: C.accent, bold: true, align: "center", margin: 0,
  });
  s.addText("#13 雨の日に頭が痛くなる科学的理由", {
    x: 0.6, y: 4.1, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/head_injury_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
