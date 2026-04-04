var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "日本人はどこから来たのか？最新ゲノム解析が明かす三重構造モデル";

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
  brown: "8B5E3C",
  jomon: "C0392B",
  yayoi: "2980B9",
  kofun: "27AE60",
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

function footer(s, text) {
  s.addText(text, { x: 0.3, y: 5.2, w: 9.4, h: 0.35, fontSize: 12, fontFace: FJ, color: C.sub, italic: true, margin: 0 });
}

// ===== SLIDE 1: タイトル =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("日本人はどこから来たのか？", { x: 0.6, y: 1.0, w: 8.8, h: 1.2, fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("最新ゲノム解析が明かす三重構造モデル", { x: 0.6, y: 2.8, w: 8.8, h: 0.6, fontSize: 26, fontFace: FJ, color: C.accent, align: "center", margin: 0 });
  s.addText("縄文人・弥生人・古墳人 — DNAに刻まれた3つのルーツ", { x: 0.6, y: 3.8, w: 8.8, h: 0.6, fontSize: 20, fontFace: FJ, color: "B0BEC5", align: "center", italic: true, margin: 0 });
  s.addText("医知創造ラボ", { x: 0.6, y: 4.6, w: 8.8, h: 0.5, fontSize: 18, fontFace: FJ, color: "78909C", align: "center", margin: 0 });
})();

// ===== SLIDE 2: 本日の内容 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "本日の内容");
  var items = [
    { num: "1", title: "二重構造モデル\nとは何か", desc: "埴原和郎の仮説と\nその30年", color: C.primary },
    { num: "2", title: "三重構造モデル\nへの転換", desc: "古代DNA解析が\n明かした第3の祖先", color: C.accent },
    { num: "3", title: "DNAに残る\n旧人類の痕跡", desc: "ネアンデルタール人・\nデニソワ人の遺伝子", color: C.green },
  ];
  items.forEach(function(g, i) {
    var xPos = 0.4 + i * 3.15;
    card(s, xPos, 1.2, 2.95, 3.6, C.white, g.color);
    s.addShape(pres.shapes.OVAL, { x: xPos + 1.1, y: 1.4, w: 0.75, h: 0.75, fill: { color: g.color } });
    s.addText(g.num, { x: xPos + 1.1, y: 1.4, w: 0.75, h: 0.75, fontSize: 28, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(g.title, { x: xPos + 0.15, y: 2.3, w: 2.65, h: 1.0, fontSize: 20, fontFace: FJ, color: g.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(g.desc, { x: xPos + 0.15, y: 3.4, w: 2.65, h: 1.2, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", valign: "top", margin: 0 });
  });
})();

// ===== SLIDE 3: 二重構造モデル =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "埴原和郎の「二重構造モデル」（1991年）");
  // 第1層：縄文人
  card(s, 0.4, 1.2, 4.3, 2.8, C.white, C.jomon);
  s.addText("第1層：縄文人", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.jomon, bold: true, margin: 0 });
  s.addText("旧石器時代〜に日本列島へ到達\n狩猟・採集・漁撈の生活\n東南アジア方面から渡来（当時の推定）", { x: 0.6, y: 1.9, w: 3.9, h: 1.8, fontSize: 17, fontFace: FJ, color: C.text, margin: 0 });
  // 第2層：弥生人
  card(s, 5.3, 1.2, 4.3, 2.8, C.white, C.yayoi);
  s.addText("第2層：弥生人", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 22, fontFace: FJ, color: C.yayoi, bold: true, margin: 0 });
  s.addText("約2,300年前に渡来\n水稲農耕を持ち込む\n朝鮮半島経由で北東アジアから", { x: 5.5, y: 1.9, w: 3.9, h: 1.8, fontSize: 17, fontFace: FJ, color: C.text, margin: 0 });
  // 矢印
  s.addText("＋  混血  →", { x: 4.0, y: 4.2, w: 2, h: 0.5, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, align: "center", margin: 0 });
  card(s, 5.8, 4.1, 3.8, 0.7, C.light, C.primary);
  s.addText("現代日本人", { x: 5.9, y: 4.15, w: 3.6, h: 0.6, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  footer(s, "Hudson MJ, et al. Evol Hum Sci. 2020;2:e6.");
})();

// ===== SLIDE 4: 二重構造の証拠と限界 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "二重構造モデルの検証と残された疑問");
  card(s, 0.4, 1.2, 4.3, 2.6, C.white, C.green);
  s.addText("支持された点", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("・アイヌと琉球が遺伝的に近い\n・本土日本人は韓国と遺伝的に近い\n・混血モデルが置換モデルより\n  29〜63倍尤度が高い", { x: 0.6, y: 1.9, w: 3.9, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 5.3, 1.2, 4.3, 2.6, C.white, C.red);
  s.addText("残された疑問", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("・古墳時代の急激な文化変容は\n  弥生の波だけで説明できるか？\n・古墳人の形態が弥生人と異なる\n・縄文人の東南アジア起源は否定", { x: 5.5, y: 1.9, w: 3.9, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 0.4, 4.1, 9.2, 0.8, C.lightYellow, C.yellow);
  s.addText("答えを出すには「古代人のDNAを直接読み解く」パレオゲノミクスの進歩が必要だった", { x: 0.7, y: 4.2, w: 8.7, h: 0.6, fontSize: 18, fontFace: FJ, color: C.text, bold: true, valign: "middle", align: "center", margin: 0 });
  footer(s, "Omoto K, Saitou N. Am J Phys Anthropol. 1997 / Nakagome S, et al. Mol Biol Evol. 2015");
})();

// ===== SLIDE 5: 三重構造モデル概要 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "三重構造モデル — 3つの祖先集団", C.dark);
  var anc = [
    { title: "縄文人", period: "2万〜1.5万年前\n分岐", origin: "大陸基層集団から\n島嶼化で隔離", trait: "有効集団サイズ\n約1,000人", color: C.jomon },
    { title: "弥生渡来人", period: "約3,000年前\n渡来", origin: "北東アジア起源\n（朝鮮半島経由）", trait: "水稲農耕を\n持ち込む", color: C.yayoi },
    { title: "古墳渡来人", period: "約1,700年前\n渡来", origin: "東アジア起源\n（弥生人とは異なる）", trait: "国家形成・\n文字・仏教", color: C.kofun },
  ];
  anc.forEach(function(a, i) {
    var xPos = 0.3 + i * 3.2;
    card(s, xPos, 1.1, 3.0, 3.8, C.white, a.color);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 3.0, h: 0.7, fill: { color: a.color }, rectRadius: 0.1 });
    s.addText(a.title, { x: xPos, y: 1.1, w: 3.0, h: 0.7, fontSize: 24, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(a.period, { x: xPos + 0.15, y: 2.0, w: 2.7, h: 0.8, fontSize: 17, fontFace: FJ, color: a.color, bold: true, align: "center", margin: 0 });
    s.addText(a.origin, { x: xPos + 0.15, y: 2.8, w: 2.7, h: 0.8, fontSize: 16, fontFace: FJ, color: C.text, align: "center", margin: 0 });
    s.addText(a.trait, { x: xPos + 0.15, y: 3.7, w: 2.7, h: 0.8, fontSize: 16, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  });
  footer(s, "Cooke NP, et al. Sci Adv. 2021;7(38):eabh2419.");
})();

// ===== SLIDE 6: 縄文人の遺伝学的特徴 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "縄文人 — 1万年以上孤立した独自の集団", C.jomon);
  var facts = [
    { icon: "🏝️", title: "海面上昇で隔離", desc: "最終氷期の終わりに日本列島が島嶼化\n大陸から遺伝的に隔離された" },
    { icon: "👥", title: "わずか約1,000人", desc: "有効集団サイズが極めて小さく\n数千年にわたり維持された" },
    { icon: "🧬", title: "独自の遺伝プロファイル", desc: "大陸のどの集団とも大きく異なる\n2万〜1.5万年前に分岐" },
    { icon: "⏳", title: "1万年以上の独自進化", desc: "狩猟採集・漁撈文化のなかで\n環境に適応した遺伝的特徴を獲得" },
  ];
  facts.forEach(function(f, i) {
    var yPos = 1.15 + i * 1.05;
    card(s, 0.4, yPos, 9.2, 0.9, C.white, C.jomon);
    s.addText(f.icon, { x: 0.6, y: yPos + 0.05, w: 0.8, h: 0.8, fontSize: 28, align: "center", valign: "middle", margin: 0 });
    s.addText(f.title, { x: 1.5, y: yPos + 0.05, w: 2.5, h: 0.8, fontSize: 19, fontFace: FJ, color: C.jomon, bold: true, valign: "middle", margin: 0 });
    s.addText(f.desc, { x: 4.2, y: yPos + 0.05, w: 5.2, h: 0.8, fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  footer(s, "Cooke NP, et al. Sci Adv. 2021;7(38):eabh2419.");
})();

// ===== SLIDE 7: 弥生人のルーツ =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "弥生人 — 北東アジアから水稲農耕とともに", C.yayoi);
  card(s, 0.4, 1.2, 4.3, 3.5, C.white, C.yayoi);
  s.addText("弥生人の遺伝的起源", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.yayoi, bold: true, margin: 0 });
  s.addText("・北東アジアに遺伝的起源\n・現在の韓国・中国北部の集団と\n  遺伝的に近い\n・約3,000年前に北九州から渡来\n・水稲農耕技術をもたらす\n・土井ヶ浜遺跡人骨で直接確認", { x: 0.6, y: 1.9, w: 3.9, h: 2.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 5.3, 1.2, 4.3, 3.5, C.white, C.accent);
  s.addText("双方向の遺伝子流動", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("韓国・金海市の三国時代（4〜5世紀）\nの古代ゲノムにも\n「縄文関連の祖先成分」を検出\n\n→ 日本列島 ⇄ 朝鮮半島の\n  双方向的な遺伝子の移動を示唆", { x: 5.5, y: 1.9, w: 3.9, h: 2.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  footer(s, "Gelabert P, et al. Curr Biol. 2022;32(15):3232-3244.");
})();

// ===== SLIDE 8: 古墳人 — 第3の波 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "古墳人 — 東アジアからの「第3の波」", C.kofun);
  card(s, 0.4, 1.2, 9.2, 1.5, C.lightYellow, C.accent);
  s.addText("最も注目すべき発見", { x: 0.7, y: 1.3, w: 8.6, h: 0.5, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("古墳時代（3世紀後半〜7世紀）に、弥生人とは異なる「東アジア起源」の第3の祖先集団が日本列島に流入した", { x: 0.7, y: 1.85, w: 8.6, h: 0.7, fontSize: 18, fontFace: FJ, color: C.text, margin: 0 });
  var events = [
    { title: "大和政権の成立", color: C.kofun },
    { title: "文字・仏教の導入", color: C.primary },
    { title: "巨大前方後円墳の築造", color: C.dark },
  ];
  events.forEach(function(e, i) {
    var xPos = 0.4 + i * 3.15;
    card(s, xPos, 3.0, 2.95, 1.0, C.white, e.color);
    s.addText(e.title, { x: xPos + 0.15, y: 3.05, w: 2.65, h: 0.9, fontSize: 20, fontFace: FJ, color: e.color, bold: true, align: "center", valign: "middle", margin: 0 });
  });
  card(s, 0.4, 4.3, 9.2, 0.8, C.lightGreen, C.kofun);
  s.addText("古墳時代に形成された三重構造は、1,500年以上経った今も私たちのDNAに刻まれている", { x: 0.7, y: 4.4, w: 8.6, h: 0.6, fontSize: 18, fontFace: FJ, color: C.kofun, bold: true, valign: "middle", align: "center", margin: 0 });
  footer(s, "Cooke NP, et al. Sci Adv. 2021;7(38):eabh2419.");
})();

// ===== SLIDE 9: JEWEL研究 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "JEWEL — 3,256人の全ゲノム解析（2024年 理研）");
  card(s, 0.4, 1.2, 4.3, 1.5, C.white, C.primary);
  s.addText("研究規模", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("バイオバンク・ジャパンの\n3,256人を高深度全ゲノム配列決定", { x: 0.6, y: 1.85, w: 3.9, h: 0.7, fontSize: 17, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 5.3, 1.2, 4.3, 1.5, C.white, C.accent);
  s.addText("従来との違い", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("マイクロアレイでは見えなかった\n稀少な遺伝的変異まで検出可能", { x: 5.5, y: 1.85, w: 3.9, h: 0.7, fontSize: 17, fontFace: FJ, color: C.text, margin: 0 });
  var findings = [
    { num: "1", text: "三重祖先構造を\n大規模データで確認", color: C.primary },
    { num: "2", text: "前例のない微細な\n遺伝的構造を検出", color: C.accent },
    { num: "3", text: "44のアーカイック\nセグメントを同定", color: C.green },
    { num: "4", text: "自然選択を受けた\n候補遺伝子座を特定", color: C.jomon },
  ];
  findings.forEach(function(f, i) {
    var xPos = 0.3 + i * 2.4;
    card(s, xPos, 3.0, 2.2, 1.8, C.white, f.color);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.8, y: 3.15, w: 0.6, h: 0.6, fill: { color: f.color } });
    s.addText(f.num, { x: xPos + 0.8, y: 3.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.text, { x: xPos + 0.1, y: 3.9, w: 2.0, h: 0.8, fontSize: 15, fontFace: FJ, color: C.text, align: "center", valign: "top", margin: 0 });
  });
  footer(s, "Liu X, et al. Sci Adv. 2024;10(16):eadi8419.");
})();

// ===== SLIDE 10: 地域別の縄文祖先比率 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "日本人の遺伝的多様性 — 地域差マップ");
  var hO = { fontSize: 15, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center", valign: "middle" };
  var cO = { fontSize: 15, fontFace: FJ, color: C.text, align: "center", valign: "middle" };
  var cB = { fontSize: 15, fontFace: FJ, color: C.primary, bold: true, align: "center", valign: "middle" };
  var rows = [
    [{ text: "地域", options: hO }, { text: "縄文祖先の比率", options: hO }, { text: "特徴", options: hO }],
    [{ text: "沖縄", options: cB }, { text: "約28.5%（最高）", options: { fontSize: 15, fontFace: FE, color: C.jomon, bold: true, align: "center", valign: "middle", fill: { color: C.lightRed } } }, { text: "縄文人との遺伝的親和性が最も高い", options: cO }],
    [{ text: "東北", options: cB }, { text: "約19%", options: { fontSize: 15, fontFace: FE, color: C.accent, bold: true, align: "center", valign: "middle", fill: { color: C.lightYellow } } }, { text: "青森・秋田・岩手で縄文系変異が多い", options: cO }],
    [{ text: "関東・中部", options: cB }, { text: "中間", options: cO }, { text: "日本の平均的な遺伝構成", options: cO }],
    [{ text: "近畿・四国", options: cB }, { text: "約12%（低い）", options: { fontSize: 15, fontFace: FE, color: C.yayoi, bold: true, align: "center", valign: "middle", fill: { color: C.light } } }, { text: "漢民族と遺伝的親和性が高い", options: cO }],
  ];
  s.addTable(rows, { x: 0.4, y: 1.2, w: 9.2, colW: [1.8, 2.5, 4.9], rowH: [0.5, 0.6, 0.6, 0.6, 0.6], border: { pt: 1, color: "DEE2E6" }, autoPage: false });
  card(s, 0.4, 4.2, 9.2, 0.8, C.light, C.primary);
  s.addText("弥生人は北九州から東進 → 近畿に大きな影響 → 東北・沖縄には縄文の特徴がより残存", { x: 0.7, y: 4.3, w: 8.6, h: 0.6, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", align: "center", margin: 0 });
  footer(s, "Liu X, et al. Sci Adv. 2024 / Kawai Y, et al. PLoS Genet. 2023");
})();

// ===== SLIDE 11: ネアンデルタール人・デニソワ人 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "日本人のDNAに残る旧人類の痕跡", C.dark);
  card(s, 0.4, 1.2, 4.3, 3.5, C.white, C.primary);
  s.addText("ネアンデルタール人", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("・アフリカ以外の現代人ゲノムに\n  約2%のDNAが残存\n・日本人も同様に保持\n・東アジア人では精神・認知形質\n  との関連が報告", { x: 0.6, y: 1.9, w: 3.9, h: 2.4, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 5.3, 1.2, 4.3, 3.5, C.white, C.accent);
  s.addText("デニソワ人", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 20, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("・2010年にシベリアで発見された\n  「第3の人類」\n・東アジア人に特有の\n  第3のデニソワ人系統が存在\n・NKX6-1遺伝子と2型糖尿病の関連\n・冠動脈疾患の遺伝率に1.69倍濃縮", { x: 5.5, y: 1.9, w: 3.9, h: 2.4, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  footer(s, "Liu X, et al. Sci Adv. 2024 / Koller D, et al. BMC Biol. 2022 / Jacobs GS, et al. Cell. 2019");
})();

// ===== SLIDE 12: アーカイックセグメント =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "44のアーカイックセグメントと疾患リスク");
  var segs = [
    { gene: "NKX6-1", source: "デニソワ人", trait: "2型糖尿病", pop: "東アジア人特異的", color: C.accent },
    { gene: "MHC領域", source: "デニソワ人", trait: "アルブミン/\nグロブリン比", pop: "東アジア人", color: C.accent },
    { gene: "複数遺伝子座", source: "デニソワ人", trait: "冠動脈疾患\n(1.69倍濃縮)", pop: "東アジア人", color: C.accent },
    { gene: "複数遺伝子座", source: "ネアンデルタール", trait: "精神・認知形質", pop: "東アジア人", color: C.primary },
  ];
  var hO = { fontSize: 14, fontFace: FJ, color: C.white, bold: true, fill: { color: C.dark }, align: "center", valign: "middle" };
  var rows = [
    [{ text: "遺伝子/領域", options: hO }, { text: "由来", options: hO }, { text: "関連形質", options: hO }, { text: "集団特異性", options: hO }]
  ];
  segs.forEach(function(seg) {
    rows.push([
      { text: seg.gene, options: { fontSize: 14, fontFace: FE, color: seg.color, bold: true, align: "center", valign: "middle" } },
      { text: seg.source, options: { fontSize: 14, fontFace: FJ, color: C.text, align: "center", valign: "middle" } },
      { text: seg.trait, options: { fontSize: 14, fontFace: FJ, color: C.text, align: "center", valign: "middle" } },
      { text: seg.pop, options: { fontSize: 14, fontFace: FJ, color: C.sub, align: "center", valign: "middle" } }
    ]);
  });
  s.addTable(rows, { x: 0.4, y: 1.2, w: 9.2, colW: [2.3, 2.0, 2.5, 2.4], rowH: [0.5, 0.6, 0.6, 0.6, 0.6], border: { pt: 1, color: "DEE2E6" }, autoPage: false });
  card(s, 0.4, 4.2, 9.2, 0.8, C.lightYellow, C.yellow);
  s.addText("これらのセグメントの多くは東アジア人に特異的で、ヨーロッパ人には見られない", { x: 0.7, y: 4.3, w: 8.6, h: 0.6, fontSize: 17, fontFace: FJ, color: C.text, bold: true, valign: "middle", align: "center", margin: 0 });
  footer(s, "Liu X, et al. Sci Adv. 2024 / Koller D, et al. BMC Biol. 2022");
})();

// ===== SLIDE 13: Y染色体ハプログループ =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "Y染色体ハプログループから見る日本人男性の系譜");
  var hO = { fontSize: 15, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center", valign: "middle" };
  var rows = [
    [{ text: "ハプログループ", options: hO }, { text: "頻度", options: hO }, { text: "推定起源", options: hO }, { text: "拡大時期", options: hO }],
    [
      { text: "D（D-M174）", options: { fontSize: 15, fontFace: FE, color: C.jomon, bold: true, align: "center", valign: "middle" } },
      { text: "約34.7%", options: { fontSize: 15, fontFace: FE, color: C.jomon, bold: true, align: "center", valign: "middle", fill: { color: C.lightRed } } },
      { text: "アフリカ→南アジア→東アジア", options: { fontSize: 14, fontFace: FJ, color: C.text, align: "center", valign: "middle" } },
      { text: "約2万年前", options: { fontSize: 15, fontFace: FJ, color: C.text, align: "center", valign: "middle" } }
    ],
    [
      { text: "O（O-M175）", options: { fontSize: 15, fontFace: FE, color: C.yayoi, bold: true, align: "center", valign: "middle" } },
      { text: "約51.8%", options: { fontSize: 15, fontFace: FE, color: C.yayoi, bold: true, align: "center", valign: "middle", fill: { color: C.light } } },
      { text: "東南アジア→東アジア", options: { fontSize: 14, fontFace: FJ, color: C.text, align: "center", valign: "middle" } },
      { text: "約4,000年前", options: { fontSize: 15, fontFace: FJ, color: C.text, align: "center", valign: "middle" } }
    ],
    [
      { text: "C（C-M130）", options: { fontSize: 15, fontFace: FE, color: C.sub, bold: true, align: "center", valign: "middle" } },
      { text: "約8.5%", options: { fontSize: 15, fontFace: FE, color: C.sub, bold: true, align: "center", valign: "middle" } },
      { text: "中央〜北東アジア", options: { fontSize: 14, fontFace: FJ, color: C.text, align: "center", valign: "middle" } },
      { text: "約1.2万年前", options: { fontSize: 15, fontFace: FJ, color: C.text, align: "center", valign: "middle" } }
    ],
  ];
  s.addTable(rows, { x: 0.3, y: 1.2, w: 9.4, colW: [2.2, 1.5, 3.2, 2.5], rowH: [0.5, 0.65, 0.65, 0.65], border: { pt: 1, color: "DEE2E6" }, autoPage: false });
  card(s, 0.4, 3.9, 9.2, 1.2, C.white, C.jomon);
  s.addText("ハプログループD — U字型分布", { x: 0.6, y: 3.95, w: 4.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.jomon, bold: true, margin: 0 });
  s.addText("アイヌ（約88%）・琉球で高頻度、近畿で低頻度", { x: 0.6, y: 4.45, w: 4.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("ハプログループO — 逆U字型分布", { x: 5.3, y: 3.95, w: 4.0, h: 0.5, fontSize: 18, fontFace: FJ, color: C.yayoi, bold: true, margin: 0 });
  s.addText("九州で最高頻度、アイヌで低頻度", { x: 5.3, y: 4.45, w: 4.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.text, margin: 0 });
  footer(s, "Hammer MF, et al. J Hum Genet. 2006;51(1):47-58.");
})();

// ===== SLIDE 14: ハプログループDの世界分布 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "ハプログループD — 6万年の旅路");
  card(s, 0.4, 1.2, 9.2, 1.5, C.light, C.primary);
  s.addText("D-M174の分布の謎", { x: 0.6, y: 1.3, w: 8.8, h: 0.5, fontSize: 20, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("チベット・日本・アンダマン諸島に高頻度で見られるが、その間の広大な地域ではほとんど検出されない「飛び石分布」", { x: 0.6, y: 1.85, w: 8.8, h: 0.7, fontSize: 17, fontFace: FJ, color: C.text, margin: 0 });
  var points = [
    { title: "約6万年前", desc: "アフリカから出発\n南ルートでユーラシアを東進", y: 3.0 },
    { title: "漢文化の拡大", desc: "新石器時代の人口爆発により\n中間地域からD系統が駆逐", y: 3.8 },
    { title: "現在の分布", desc: "チベット高原と日本列島に\n「避難所」として残存", y: 4.6 },
  ];
  points.forEach(function(p) {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: p.y, w: 0.1, h: 0.6, fill: { color: C.jomon } });
    s.addText(p.title, { x: 0.7, y: p.y, w: 2.3, h: 0.6, fontSize: 17, fontFace: FJ, color: C.jomon, bold: true, valign: "middle", margin: 0 });
    s.addText(p.desc, { x: 3.2, y: p.y, w: 6.2, h: 0.6, fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  footer(s, "Shi H, et al. BMC Biol. 2008;6:45 / Mondal M, et al. Hum Genet. 2017");
})();

// ===== SLIDE 15: アイヌ・琉球・本土の遺伝的関係 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「日本人」の遺伝的多様性 — 均一ではない");
  card(s, 0.4, 1.2, 2.85, 3.0, C.white, C.jomon);
  s.addText("アイヌ", { x: 0.55, y: 1.3, w: 2.6, h: 0.5, fontSize: 22, fontFace: FJ, color: C.jomon, bold: true, align: "center", margin: 0 });
  s.addText("縄文系D: 約88%\n＋北方（ニヴフ・\nコリャーク）の成分", { x: 0.55, y: 1.9, w: 2.6, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, align: "center", margin: 0 });
  card(s, 3.55, 1.2, 2.95, 3.0, C.white, C.yayoi);
  s.addText("本土日本人", { x: 3.7, y: 1.3, w: 2.6, h: 0.5, fontSize: 22, fontFace: FJ, color: C.yayoi, bold: true, align: "center", margin: 0 });
  s.addText("渡来系O: 約52%\n縄文系D: 約35%\n近畿で渡来系最多", { x: 3.7, y: 1.9, w: 2.6, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, align: "center", margin: 0 });
  card(s, 6.8, 1.2, 2.85, 3.0, C.white, C.kofun);
  s.addText("琉球", { x: 6.95, y: 1.3, w: 2.6, h: 0.5, fontSize: 22, fontFace: FJ, color: C.kofun, bold: true, align: "center", margin: 0 });
  s.addText("縄文系の比率:\n約28.5%（最高）\nアイヌと遺伝的に\n比較的近い", { x: 6.95, y: 1.9, w: 2.6, h: 1.8, fontSize: 16, fontFace: FJ, color: C.text, align: "center", margin: 0 });
  // 矢印でアイヌと琉球の近さを示す
  s.addShape(pres.shapes.LINE, { x: 1.8, y: 4.4, w: 6.4, h: 0, line: { color: C.jomon, width: 2, dashType: "dash" } });
  s.addText("遺伝的に比較的近い（共に縄文成分が多い）", { x: 2.0, y: 4.5, w: 6.0, h: 0.4, fontSize: 15, fontFace: FJ, color: C.jomon, italic: true, align: "center", margin: 0 });
  footer(s, "Tajima A, et al. J Hum Genet. 2004 / Hammer MF, et al. J Hum Genet. 2006");
})();

// ===== SLIDE 16: 自然選択の痕跡 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "農耕がゲノムを変えた — 自然選択の痕跡");
  card(s, 0.4, 1.2, 4.3, 2.5, C.white, C.accent);
  s.addText("ALDH2 — お酒に弱い体質", { x: 0.6, y: 1.3, w: 3.9, h: 0.5, fontSize: 19, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("・東アジア人に特徴的な変異\n・アルコール代謝関連遺伝子に\n  正の自然選択の痕跡\n・農耕開始（約12,500年前）と\n  関連する可能性", { x: 0.6, y: 1.9, w: 3.9, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 5.3, 1.2, 4.3, 2.5, C.white, C.green);
  s.addText("脂肪酸代謝", { x: 5.5, y: 1.3, w: 3.9, h: 0.5, fontSize: 19, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("・脂肪酸代謝関連遺伝子にも\n  正の選択圧\n・農耕による食生活の変化が\n  日本人のゲノムを形作った\n・縄文→弥生の食文化転換が\n  選択圧として作用", { x: 5.5, y: 1.9, w: 3.9, h: 1.6, fontSize: 16, fontFace: FJ, color: C.text, margin: 0 });
  card(s, 0.4, 4.0, 9.2, 1.0, C.lightGreen, C.green);
  s.addText("縄文時代の人口減少期と江戸時代の人口減少期 — ゲノムに刻まれた2つのボトルネック", { x: 0.7, y: 4.1, w: 8.6, h: 0.8, fontSize: 17, fontFace: FJ, color: C.text, bold: true, valign: "middle", align: "center", margin: 0 });
  footer(s, "Kawai Y, et al. PLoS Genet. 2023;19(12):e1010625.");
})();

// ===== SLIDE 17: ゲノム医療への応用 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "祖先の遺伝子と未来の医療");
  var apps = [
    { icon: "💊", title: "薬理ゲノミクス", desc: "祖先背景による薬の効き方の違い\n→ 日本人に最適化した用量設定", color: C.primary },
    { icon: "🔬", title: "疾患リスク予測", desc: "デニソワ人由来遺伝子と2型糖尿病\n→ 集団特異的なリスクモデル", color: C.accent },
    { icon: "🏥", title: "精密医療の実現", desc: "JEWELなどの日本人ゲノムDB\n→ 個人最適化医療の基盤", color: C.green },
  ];
  apps.forEach(function(a, i) {
    var yPos = 1.15 + i * 1.35;
    card(s, 0.4, yPos, 9.2, 1.2, C.white, a.color);
    s.addText(a.icon, { x: 0.6, y: yPos + 0.1, w: 0.8, h: 1.0, fontSize: 32, align: "center", valign: "middle", margin: 0 });
    s.addText(a.title, { x: 1.5, y: yPos + 0.1, w: 2.5, h: 1.0, fontSize: 20, fontFace: FJ, color: a.color, bold: true, valign: "middle", margin: 0 });
    s.addText(a.desc, { x: 4.2, y: yPos + 0.1, w: 5.2, h: 1.0, fontSize: 16, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  card(s, 0.4, 4.6, 9.2, 0.5, C.lightYellow, C.yellow);
  s.addText("祖先を知ることは、過去を知るだけでなく、未来の医療を変える力を持つ", { x: 0.7, y: 4.65, w: 8.6, h: 0.4, fontSize: 17, fontFace: FJ, color: C.text, bold: true, valign: "middle", align: "center", margin: 0 });
})();

// ===== SLIDE 18: まとめ =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ — 私たちはどこから来たのか", C.dark);
  var takeaways = [
    { text: "二重構造から三重構造へ — 縄文人・弥生人・古墳人の3つのルーツ", color: C.jomon },
    { text: "縄文人は2万年前に分岐、約1,000人で1万年以上独自に進化した", color: C.accent },
    { text: "ネアンデルタール人・デニソワ人のDNAも私たちのゲノムに残る", color: C.primary },
    { text: "日本人は遺伝的に均一ではなく、都道府県レベルの地域差がある", color: C.kofun },
    { text: "多様な起源を持つ混血集団 — その多様性こそが私たちの強さ", color: C.dark },
  ];
  takeaways.forEach(function(t, i) {
    var yPos = 1.15 + i * 0.85;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 0.12, h: 0.7, fill: { color: t.color } });
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.1, w: 0.5, h: 0.5, fill: { color: t.color } });
    s.addText(String(i + 1), { x: 0.7, y: yPos + 0.1, w: 0.5, h: 0.5, fontSize: 18, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.text, { x: 1.4, y: yPos, w: 8.2, h: 0.7, fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  card(s, 0.4, 5.0, 9.2, 0.5, C.dark);
  s.addText("私たちのDNAには、6万年の壮大な旅路が今も刻まれている", { x: 0.7, y: 5.05, w: 8.6, h: 0.4, fontSize: 18, fontFace: FJ, color: C.white, bold: true, italic: true, valign: "middle", align: "center", margin: 0 });
})();

// ===== SLIDE 19: 参考文献 =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "参考文献");
  var refs = [
    "1. Cooke NP, et al. Sci Adv. 2021;7(38):eabh2419.",
    "2. Liu X, et al. Sci Adv. 2024;10(16):eadi8419.",
    "3. Hudson MJ, et al. Evol Hum Sci. 2020;2:e6.",
    "4. Omoto K, Saitou N. Am J Phys Anthropol. 1997;102(4):437-46.",
    "5. Nakagome S, et al. Mol Biol Evol. 2015;32(6):1533-43.",
    "6. Gelabert P, et al. Curr Biol. 2022;32(15):3232-3244.",
    "7. Koller D, et al. BMC Biol. 2022;20(1):249.",
    "8. Jacobs GS, et al. Cell. 2019;177(4):1010-1021.",
    "9. Hammer MF, et al. J Hum Genet. 2006;51(1):47-58.",
    "10. Shi H, et al. BMC Biol. 2008;6:45.",
    "11. Tajima A, et al. J Hum Genet. 2004;49(4):187-193.",
    "12. Kawai Y, et al. PLoS Genet. 2023;19(12):e1010625.",
  ];
  var refText = refs.join("\n");
  s.addText(refText, { x: 0.5, y: 1.1, w: 9, h: 4.2, fontSize: 13, fontFace: FE, color: C.sub, lineSpacingMultiple: 1.3, margin: 0 });
})();

// ===== SLIDE 20: エンドカード =====
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("ご視聴ありがとうございました", { x: 0.6, y: 1.5, w: 8.8, h: 1.0, fontSize: 36, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.8, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("チャンネル登録・高評価お願いします！", { x: 0.6, y: 3.2, w: 8.8, h: 0.6, fontSize: 24, fontFace: FJ, color: C.accent, align: "center", margin: 0 });
  s.addText("医知創造ラボ", { x: 0.6, y: 4.2, w: 8.8, h: 0.5, fontSize: 20, fontFace: FJ, color: "78909C", align: "center", margin: 0 });
})();

// ===== 出力 =====
var fileName = "japanese_origin_genetics_slides.pptx";
pres.writeFile({ fileName: fileName }).then(function() {
  console.log("Created: " + fileName);
});
