var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#15】眠れないのは脳のせい？ ― 不眠のメカニズムと正しい対処法";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #15", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🌙", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("眠れないのは脳のせい？\n不眠のメカニズムと正しい対処法", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 38, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.15, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("脳神経内科医が解説する\n正しい睡眠の知識", {
    x: 0.6, y: 2.9, w: 8.8, h: 0.9,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #15", {
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
    { emoji: "🛏️", text: "布団に入っても30分以上眠れない" },
    { emoji: "⏰", text: "夜中に何度も目が覚めてしまう" },
    { emoji: "🌅", text: "朝早く目が覚めて二度寝できない" },
    { emoji: "😴", text: "眠れたはずなのに疲れがとれない" },
    { emoji: "🥱", text: "日中に強い眠気や集中力低下がある" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 睡眠のしくみ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "睡眠のしくみ ― 脳の2つのスイッチ");
  // 体内時計
  card(s, 0.5, 1.1, 4.3, 2.6, C.white, C.primary);
  s.addText("🕐 体内時計", { x: 0.7, y: 1.15, w: 3.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("脳の視交叉上核にある\n約24時間のリズム", { x: 0.7, y: 1.6, w: 3.8, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  s.addText("☀️ 朝の光 → 体内時計リセット\n🌙 夜 → メラトニン分泌\n　  → 眠気が訪れる", { x: 0.7, y: 2.2, w: 3.8, h: 0.9, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.25, margin: 0 });
  s.addText("※ メラトニン = 睡眠ホルモン", { x: 0.7, y: 3.2, w: 3.8, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  // 睡眠圧
  card(s, 5.2, 1.1, 4.3, 2.6, C.white, C.accent);
  s.addText("📊 睡眠圧", { x: 5.4, y: 1.15, w: 3.8, h: 0.4, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("起きている時間が長いほど\n「眠りたい」圧力がたまる", { x: 5.4, y: 1.6, w: 3.8, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  s.addText("アデノシンの蓄積が原因\n\n☕ カフェインはアデノシンの\n　 作用をブロックして\n　 眠気を抑える", { x: 5.4, y: 2.2, w: 3.8, h: 1.1, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  // まとめ
  card(s, 0.5, 3.9, 9.0, 0.6, C.light, C.primary);
  s.addText("この2つのバランスが崩れると不眠が起こる", { x: 0.8, y: 3.95, w: 8.5, h: 0.5, fontSize: 19, fontFace: FJ, color: C.primary, bold: true, valign: "middle", align: "center", margin: 0 });
  ftr(s);
})();

// SLIDE 4: 不眠症の4つのタイプ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "不眠症の4つのタイプ");
  var types = [
    { icon: "🌃", name: "入眠障害", desc: "布団に入っても30分以上眠れない\n最も多いタイプ。ストレス・不安が主な原因", color: C.primary },
    { icon: "🔄", name: "中途覚醒", desc: "夜中に何度も目が覚めてしまう\n加齢やアルコールの影響が大きい", color: C.accent },
    { icon: "🌅", name: "早朝覚醒", desc: "予定より2時間以上早く目が覚め二度寝できない\nうつ病との関連が指摘されている", color: "7B2D8E" },
    { icon: "😩", name: "熟眠障害", desc: "睡眠時間は十分なのに眠りが浅く疲れがとれない\n睡眠時無呼吸症候群が隠れていることも", color: C.red },
  ];
  types.forEach(function(t, i) {
    var y = 1.1 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, t.color);
    s.addText(t.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 28, align: "center", margin: 0 });
    s.addText(t.name, { x: 1.4, y: y + 0.05, w: 2.5, h: 0.4, fontSize: 19, fontFace: FJ, color: t.color, bold: true, margin: 0 });
    s.addText(t.desc, { x: 4.0, y: y + 0.05, w: 5.3, h: 0.75, fontSize: 15, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 不眠の主な原因
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "不眠の主な原因");
  var causes = [
    { title: "心理的", color: C.red, items: ["ストレス・不安", "うつ病", "「眠らなければ」の\nプレッシャー"] },
    { title: "環境", color: C.accent, items: ["騒音・光・室温", "スマホのブルーライト", "メラトニン分泌を抑制"] },
    { title: "身体", color: "7B2D8E", items: ["痛み・かゆみ・頻尿", "むずむず脚症候群", "睡眠時無呼吸症候群"] },
    { title: "生活習慣", color: C.primary, items: ["不規則な生活", "カフェイン・アルコール", "運動不足・過度の昼寝"] },
  ];
  causes.forEach(function(col, ci) {
    var x = 0.3 + ci * 2.4;
    card(s, x, 1.1, 2.2, 0.45, col.color);
    s.addText(col.title, { x: x + 0.1, y: 1.13, w: 2.0, h: 0.4, fontSize: 16, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    col.items.forEach(function(item, ii) {
      var y = 1.7 + ii * 0.85;
      card(s, x, y, 2.2, 0.75, C.white, col.color);
      s.addText(item, { x: x + 0.15, y: y + 0.05, w: 1.9, h: 0.65, fontSize: 13, fontFace: FJ, color: C.text, valign: "middle", lineSpacingMultiple: 1.15, margin: 0 });
    });
  });
  ftr(s);
})();

// SLIDE 6: 脳の病気が隠れていることも
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳の病気が隠れていることも", C.red);
  var diseases = [
    { icon: "🦵", name: "むずむず脚症候群", desc: "夜に脚がむずむず・ピリピリして動かさずにはいられない", detail: "ドーパミン・鉄の代謝異常 → 脳神経内科で治療可能", color: C.accent },
    { icon: "🤜", name: "レム睡眠行動障害", desc: "夢の中の行動を実際に体で再現（大声・殴る蹴る）", detail: "パーキンソン病・レビー小体型認知症の前兆として重要", color: "7B2D8E" },
    { icon: "😮‍💨", name: "睡眠時無呼吸症候群", desc: "睡眠中に呼吸が何度も止まり深い睡眠がとれない", detail: "日中の強い眠気 → いびきが大きい方は要注意", color: C.primary },
  ];
  diseases.forEach(function(d, i) {
    var y = 1.1 + i * 1.3;
    card(s, 0.5, y, 9.0, 1.1, C.white, d.color);
    s.addText(d.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.9, fontSize: 30, align: "center", valign: "middle", margin: 0 });
    s.addText(d.name, { x: 1.4, y: y + 0.05, w: 3.5, h: 0.4, fontSize: 19, fontFace: FJ, color: d.color, bold: true, margin: 0 });
    s.addText(d.desc, { x: 1.4, y: y + 0.45, w: 7.8, h: 0.3, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
    s.addText(d.detail, { x: 1.4, y: y + 0.75, w: 7.8, h: 0.3, fontSize: 14, fontFace: FJ, color: C.sub, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 7: 睡眠薬の正しい知識
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "睡眠薬の正しい知識");
  var drugs = [
    { name: "メラトニン受容体作動薬", desc: "体内時計に働きかけ自然な眠気を促す\n依存性がほとんどなく高齢者にも安全", safety: "依存性◎低い", color: C.green },
    { name: "オレキシン受容体拮抗薬", desc: "覚醒を維持するオレキシンの働きを抑える\n比較的新しいタイプの睡眠薬", safety: "依存性◎低い", color: C.primary },
    { name: "ベンゾジアゼピン系・非BZ系", desc: "脳の活動を抑えて眠りを促す\n効果は確実だが依存性・翌朝のふらつきに注意", safety: "依存性△注意", color: C.accent },
  ];
  drugs.forEach(function(d, i) {
    var y = 1.1 + i * 1.15;
    card(s, 0.5, y, 9.0, 1.0, C.white, d.color);
    s.addText(d.name, { x: 0.7, y: y + 0.05, w: 5.0, h: 0.4, fontSize: 18, fontFace: FJ, color: d.color, bold: true, margin: 0 });
    s.addText(d.desc, { x: 0.7, y: y + 0.45, w: 6.5, h: 0.5, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
    s.addText(d.safety, { x: 7.5, y: y + 0.15, w: 1.8, h: 0.7, fontSize: 14, fontFace: FJ, color: d.color, bold: true, align: "center", valign: "middle", margin: 0 });
  });
  // 注意メッセージ
  card(s, 0.5, 4.55, 9.0, 0.55, C.lightYellow, C.yellow);
  s.addText("⚠️ 睡眠薬は医師の指示で正しく使えば安全です。自己判断での増量や急な中止は避けてください。", {
    x: 0.8, y: 4.58, w: 8.5, h: 0.5, fontSize: 14, fontFace: FJ, color: "856404", valign: "middle", margin: 0,
  });
  ftr(s);
})();

// SLIDE 8: 睡眠衛生
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "睡眠衛生 ― 薬に頼らない対策");
  var tips = [
    { num: "1", title: "朝の光を浴びる", desc: "起床後30分以内に太陽の光で体内時計リセット\n曇りの日でも屋外の光は室内より十分明るい", icon: "☀️" },
    { num: "2", title: "就寝前のスマホを控える", desc: "ブルーライトがメラトニン分泌を抑制\n就寝1時間前にはスマホ・PCを見るのをやめる", icon: "📵" },
    { num: "3", title: "カフェインは午後2時まで", desc: "カフェインの半減期は5〜7時間\n午後遅くのコーヒーや緑茶は睡眠に影響する", icon: "☕" },
    { num: "4", title: "寝室は暗く涼しく", desc: "室温は18〜22℃が理想的\n遮光カーテンの使用も効果的", icon: "🌡️" },
  ];
  tips.forEach(function(t, i) {
    var y = 1.05 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.white, C.primary);
    s.addText(t.num, { x: 0.7, y: y + 0.05, w: 0.5, h: 0.75, fontSize: 30, fontFace: FE, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.icon, { x: 1.3, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 26, align: "center", margin: 0 });
    s.addText(t.title, { x: 2.0, y: y + 0.05, w: 3.0, h: 0.35, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(t.desc, { x: 2.0, y: y + 0.42, w: 7.3, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 9: やってはいけないこと
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "やってはいけないこと ― 逆効果な行動", C.red);
  var dont = [
    { icon: "🛏️", title: "眠れないのに布団にいる", desc: "脳が「ベッド＝眠れない場所」と学習\n→ 20分眠れなければ一度布団を出る" },
    { icon: "🍺", title: "アルコールで眠ろうとする", desc: "寝つきはよくなるが睡眠の質を大幅に下げる\n→ 中途覚醒が増え浅い眠りばかりに" },
    { icon: "😴", title: "休日に寝だめする", desc: "2時間以上の寝坊は体内時計を狂わせる\n→ 起床時間は毎日なるべく一定に" },
    { icon: "💊", title: "市販の睡眠薬に頼り続ける", desc: "市販品は抗ヒスタミン薬で連用すると効果減弱\n→ 2週間以上続く不眠は受診を" },
  ];
  dont.forEach(function(d, i) {
    var y = 1.05 + i * 1.0;
    card(s, 0.5, y, 9.0, 0.85, C.lightRed, C.red);
    s.addText(d.icon, { x: 0.7, y: y + 0.1, w: 0.6, h: 0.65, fontSize: 26, align: "center", margin: 0 });
    s.addText(d.title, { x: 1.4, y: y + 0.05, w: 3.5, h: 0.35, fontSize: 18, fontFace: FJ, color: C.red, bold: true, margin: 0 });
    s.addText(d.desc, { x: 1.4, y: y + 0.42, w: 7.8, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.15, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 10: 受診の目安
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");
  var levels = [
    { emoji: "🔴", label: "すぐ受診", desc: "不眠＋強いうつ症状・希死念慮\n日中の眠気で事故を起こしかけた", action: "精神科・心療内科\nを受診", bg: C.lightRed, color: C.red },
    { emoji: "🟡", label: "早めに受診", desc: "不眠が1ヶ月以上続いている\n日中の生活に支障が出ている\n呼吸停止・大声を指摘された", action: "脳神経内科・\n睡眠外来を受診", bg: C.lightYellow, color: "856404" },
    { emoji: "🟢", label: "様子見OK", desc: "ストレスの多い時期だけ眠りにくい\n生活習慣の改善で対処できている", action: "睡眠衛生の\n見直しから", bg: C.lightGreen, color: C.green },
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
    { q: "理想の睡眠時間は何時間ですか？", a: "個人差が大きく「何時間眠るべき」という正解はありません。日中に眠気なく過ごせていれば十分です。高齢者は6時間程度で自然なことが多いです" },
    { q: "睡眠薬は依存しますか？", a: "メラトニン受容体作動薬やオレキシン受容体拮抗薬は依存性がほとんどありません。ベンゾジアゼピン系も医師の指示で使用すれば安全ですが、急な中止は避けてください" },
    { q: "サプリメントは効きますか？", a: "メラトニンサプリは時差ボケには有効ですが慢性的な不眠への効果は限定的です。グリシンやテアニンにはリラックス効果がありますが治療としては不十分です" },
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
  hdr(s, "まとめ ― 不眠の正しい対処法");
  var points = [
    { num: "1", text: "不眠は脳の2つの\nスイッチの乱れ", sub: "体内時計と睡眠圧のバランス\nしくみを理解すれば対策がわかる" },
    { num: "2", text: "まず睡眠衛生を\n見直す", sub: "朝の光・就寝前スマホ控え\nカフェイン制限が3大ポイント" },
    { num: "3", text: "1ヶ月以上続けば\n受診を", sub: "現在の睡眠薬は安全性が高い\n治療可能な病気が見つかることも" },
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
  s.addText("#16 スマホの使いすぎで手がしびれる？ ― デジタル時代の神経トラブル", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 18, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ", {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

var outPath = __dirname + "/insomnia_brain_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
});
