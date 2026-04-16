var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#20】脳ドックは受けるべき？ ― 検査の意味と結果の活かし方";

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
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #20", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🧠", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("脳ドックは受けるべき？", {
    x: 0.6, y: 1.35, w: 8.8, h: 0.85,
    fontSize: 38, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.3, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("検査の意味と結果の活かし方", {
    x: 0.6, y: 2.45, w: 8.8, h: 0.65,
    fontSize: 26, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("その症状、大丈夫？ #20（最終回）", {
    x: 0.6, y: 3.85, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.35, w: 8.8, h: 0.4,
    fontSize: 12, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// SLIDE 2: こんな経験ありませんか？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな経験、ありませんか？");
  var items = [
    { emoji: "🧠", text: "脳ドックを受けてみたいが、本当に必要か迷っている" },
    { emoji: "😰", text: "脳梗塞や脳腫瘍が心配で、「何か見つかったら…」と怖い" },
    { emoji: "👨‍👩‍👦", text: "親が脳梗塞になった。自分も受けるべき？" },
    { emoji: "💰", text: "費用が高そうで、どこで受ければいいかわからない" },
    { emoji: "📋", text: "「異常あり」と言われたら、どうすればいいの？" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 脳ドックとは ― 何を調べる？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳ドックとは ― 何を調べる？");
  card(s, 0.5, 1.1, 9.0, 0.55, C.light, C.primary);
  s.addText("症状がなくても脳・血管の異常を画像で早期発見する検査です", { x: 0.8, y: 1.18, w: 8.5, h: 0.42, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var items = [
    { icon: "🔬", title: "MRI（脳）", body: "脳の断面を磁気で撮影。脳梗塞・腫瘍・萎縮・白質病変を検出" },
    { icon: "🩸", title: "MRA（脳血管）", body: "脳の動脈を3Dで描出。脳動脈瘤・狭窄・閉塞を確認" },
    { icon: "🫀", title: "頸動脈超音波", body: "頸部の動脈の狭窄・動脈硬化プラークをリアルタイムで評価" },
    { icon: "📊", title: "血液検査・血圧測定", body: "脳卒中リスクとなる糖尿病・脂質異常・高血圧を評価" },
  ];
  items.forEach(function(r, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var x = col === 0 ? 0.5 : 5.3;
    var y = 1.9 + row * 1.55;
    card(s, x, y, 4.6, 1.25, C.white, C.primary);
    s.addText(r.icon, { x: x + 0.2, y: y + 0.1, w: 0.7, h: 0.95, fontSize: 28, align: "center", margin: 0 });
    s.addText(r.title, { x: x + 0.95, y: y + 0.1, w: 3.5, h: 0.38, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(r.body, { x: x + 0.95, y: y + 0.5, w: 3.5, h: 0.65, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 4: 脳ドックでわかること
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳ドックでわかること");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("無症状でも発見されることがある主な異常", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var findings = [
    { emoji: "💥", title: "未破裂脳動脈瘤", freq: "2〜3%", body: "くも膜下出血の原因。大きさや形で経過観察か治療かを判断" },
    { emoji: "🔴", title: "無症候性脳梗塞", freq: "約10%", body: "気づかないうちに起きた小さな脳梗塞。将来の脳卒中リスク" },
    { emoji: "⬜", title: "白質病変", freq: "加齢で増加", body: "脳の白い部分の変化。高血圧・動脈硬化と関連" },
    { emoji: "📉", title: "脳萎縮", freq: "加齢で進行", body: "認知症の早期サインになることもある" },
    { emoji: "🫀", title: "頸動脈狭窄", freq: "リスクで変動", body: "脳への血流障害のリスク。動脈硬化の指標にもなる" },
  ];
  findings.forEach(function(f, i) {
    var y = 1.82 + i * 0.67;
    card(s, 0.5, y, 9.0, 0.57, C.white, C.accent);
    s.addText(f.emoji, { x: 0.7, y: y + 0.05, w: 0.55, h: 0.47, fontSize: 22, align: "center", margin: 0 });
    s.addText(f.title, { x: 1.35, y: y + 0.04, w: 2.8, h: 0.26, fontSize: 15, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
    s.addText("発見率：" + f.freq, { x: 1.35, y: y + 0.3, w: 2.8, h: 0.22, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
    s.addText(f.body, { x: 4.3, y: y + 0.1, w: 5.0, h: 0.38, fontSize: 13, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 脳ドックを受けるべき人
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳ドックを受けるべき人");
  card(s, 0.5, 1.1, 9.0, 0.55, C.lightYellow, C.yellow);
  s.addText("特にリスクが高い方は積極的に検討しましょう", { x: 0.8, y: 1.18, w: 8.5, h: 0.42, fontSize: 17, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });

  var highRisk = [
    { emoji: "🎂", text: "40歳以上（特に50〜60代）" },
    { emoji: "👨‍👩‍👦", text: "家族に脳卒中・脳動脈瘤の人がいる" },
    { emoji: "💊", text: "高血圧・糖尿病・脂質異常症がある" },
    { emoji: "🚬", text: "喫煙習慣がある、または長年あった" },
  ];
  var lowRisk = [
    { emoji: "✅", text: "定期的な健康診断で異常なし" },
    { emoji: "🏃", text: "生活習慣が整っている若年者" },
    { emoji: "😌", text: "「安心したい」だけの場合も選択肢" },
  ];

  s.addText("受けるべきリスク因子（当てはまる方は受診推奨）", { x: 0.5, y: 1.82, w: 9.0, h: 0.32, fontSize: 14, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  highRisk.forEach(function(item, i) {
    var y = 2.2 + i * 0.55;
    card(s, 0.5, y, 9.0, 0.46, C.lightRed, C.red);
    s.addText(item.emoji, { x: 0.7, y: y + 0.04, w: 0.55, h: 0.38, fontSize: 22, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.35, y: y + 0.04, w: 7.9, h: 0.38, fontSize: 17, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  s.addText("受診時期の目安（リスクが低い方）", { x: 0.5, y: 4.48, w: 9.0, h: 0.3, fontSize: 13, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  lowRisk.forEach(function(item, i) {
    s.addText(item.emoji + " " + item.text, { x: 1.0, y: 4.82 + i * 0.0, w: 8.5, h: 0.28, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 6: 脳ドックの費用と頻度
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳ドックの費用と受診頻度");

  var cols = [
    {
      title: "💴 費用の目安",
      color: C.accent,
      items: [
        "基本コース：3〜5万円（自費）",
        "簡易コース：1〜2万円のところも",
        "健康保険は原則適用されない",
        "自治体の助成制度を活用しよう",
        "企業健診の一環として実施も増加",
      ]
    },
    {
      title: "📅 受診頻度の目安",
      color: C.primary,
      items: [
        "初回：まず1回受けて基準値を把握",
        "リスクなし：2〜3年に1回が目安",
        "リスクあり：1〜2年に1回を推奨",
        "異常あり後：主治医の指示に従う",
        "毎年は不要なことが多い",
      ]
    }
  ];
  cols.forEach(function(col, ci) {
    var x = ci === 0 ? 0.5 : 5.2;
    card(s, x, 1.1, 4.6, 0.5, col.color === C.accent ? C.lightYellow : C.light, col.color);
    s.addText(col.title, { x: x + 0.2, y: 1.15, w: 4.1, h: 0.4, fontSize: 16, fontFace: FJ, color: col.color, bold: true, valign: "middle", margin: 0 });
    col.items.forEach(function(item, i) {
      var y = 1.78 + i * 0.65;
      card(s, x, y, 4.6, 0.55, C.white, col.color);
      s.addText("• " + item, { x: x + 0.25, y: y + 0.05, w: 4.2, h: 0.44, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    });
  });
  card(s, 0.5, 5.0, 9.0, 0.3, C.lightGreen, C.green);
  s.addText("💡 自治体によっては補助金・クーポンあり。まず市区町村の窓口に問い合わせを！", { x: 0.72, y: 5.03, w: 8.6, h: 0.25, fontSize: 12, fontFace: FJ, color: C.text, margin: 0 });
  ftr(s);
})();

// SLIDE 7: 「異常あり」と言われたら
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "「異常あり」と言われたら");
  card(s, 0.5, 1.1, 9.0, 0.55, C.lightYellow, C.yellow);
  s.addText("「異常あり」＝「すぐ手術・すぐ入院」ではありません。まず主治医に相談を。", { x: 0.8, y: 1.18, w: 8.5, h: 0.42, fontSize: 16, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });

  var findings = [
    { label: "未破裂脳動脈瘤", color: C.red, bgColor: C.lightRed, action: "サイズ・形・部位により経過観察 or 治療。専門医受診が必須" },
    { label: "無症候性脳梗塞", color: C.accent, bgColor: C.lightYellow, action: "危険因子（血圧・血糖・脂質）の管理が最優先。抗血小板薬を検討" },
    { label: "白質病変", color: C.yellow, bgColor: C.lightYellow, action: "軽度なら経過観察。血圧・血糖を適切にコントロールすることが重要" },
    { label: "頸動脈狭窄", color: C.primary, bgColor: C.light, action: "狭窄度により内科治療 or 外科的治療。専門科（脳神経内科・外科）へ" },
    { label: "脳萎縮", color: C.sub, bgColor: "EEEEEE", action: "程度・年齢によって異なる。認知機能低下を伴う場合は専門医へ" },
  ];
  findings.forEach(function(f, i) {
    var y = 1.85 + i * 0.67;
    card(s, 0.5, y, 9.0, 0.57, f.bgColor, f.color);
    s.addText(f.label, { x: 0.75, y: y + 0.04, w: 2.7, h: 0.5, fontSize: 14, fontFace: FJ, color: f.color, bold: true, valign: "middle", margin: 0 });
    s.addText("→ " + f.action, { x: 3.55, y: y + 0.04, w: 5.7, h: 0.5, fontSize: 13, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 8: 未破裂脳動脈瘤と言われたら
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "未破裂脳動脈瘤と言われたら");
  card(s, 0.5, 1.1, 9.0, 0.55, C.light, C.primary);
  s.addText("発見された＝すぐ破裂するわけではない。「大きさ」と「部位」が最重要です", { x: 0.8, y: 1.18, w: 8.5, h: 0.42, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  card(s, 0.5, 1.85, 4.4, 1.6, C.white, C.green);
  s.addText("✅ 経過観察が基本", { x: 0.7, y: 1.9, w: 3.9, h: 0.38, fontSize: 15, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  var obs = ["直径 5mm 未満", "症状なし", "後方循環以外の部位", "破裂リスク年 0.5〜1%"];
  obs.forEach(function(item, i) {
    s.addText("• " + item, { x: 0.7, y: 2.36 + i * 0.28, w: 3.9, h: 0.28, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.1, 1.85, 4.4, 1.6, C.white, C.red);
  s.addText("⚠️ 治療を検討", { x: 5.3, y: 1.9, w: 3.9, h: 0.38, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var treat = ["直径 5〜7mm 以上", "不整形・多房性", "後交通動脈・前交通動脈", "増大が確認された"];
  treat.forEach(function(item, i) {
    s.addText("• " + item, { x: 5.3, y: 2.36 + i * 0.28, w: 3.9, h: 0.28, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });

  var methods = [
    { title: "開頭クリッピング術", body: "頭を開けて動脈瘤の根元を金属クリップで止める手術" },
    { title: "コイル塞栓術（血管内治療）", body: "カテーテルを使って動脈瘤内にコイルを詰めて閉塞" },
  ];
  methods.forEach(function(m, i) {
    var y = 3.65 + i * 0.68;
    card(s, 0.5, y, 9.0, 0.58, C.lightRed, C.red);
    s.addText("🔧 " + m.title, { x: 0.72, y: y + 0.04, w: 3.5, h: 0.5, fontSize: 14, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0 });
    s.addText(m.body, { x: 4.3, y: y + 0.04, w: 5.0, h: 0.5, fontSize: 13, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 9: 脳ドックの限界
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳ドックの限界 ― 万能ではありません");
  card(s, 0.5, 1.1, 9.0, 0.55, C.lightRed, C.red);
  s.addText("脳ドックは「安心のための全額保証」ではありません", { x: 0.8, y: 1.18, w: 8.5, h: 0.42, fontSize: 17, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0 });

  var limits = [
    { emoji: "❌", title: "すべての異常を検出できるわけではない", body: "極小の動脈瘤・初期のがん・脳炎などは見逃すことも" },
    { emoji: "😰", title: "偽陽性による不必要な不安", body: "「加齢変化」を過剰解釈して心配しすぎる可能性" },
    { emoji: "🔁", title: "受けたら終わりではない", body: "脳ドックはスクリーニング。その後の生活習慣改善がなければ意味は薄い" },
    { emoji: "💸", title: "費用対効果は人によって異なる", body: "リスク因子のない若年者では費用に見合わない場合も" },
  ];
  limits.forEach(function(item, i) {
    var y = 1.88 + i * 0.79;
    card(s, 0.5, y, 9.0, 0.68, C.white, C.red);
    s.addText(item.emoji, { x: 0.7, y: y + 0.08, w: 0.6, h: 0.5, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.title, { x: 1.42, y: y + 0.06, w: 7.8, h: 0.3, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0 });
    s.addText(item.body, { x: 1.42, y: y + 0.36, w: 7.8, h: 0.3, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 10: 受診の目安 ― 脳ドック vs 症状受診
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 脳ドック vs 症状受診");

  card(s, 0.5, 1.1, 4.4, 4.0, C.white, C.primary);
  s.addText("🏥 脳ドックを検討", { x: 0.7, y: 1.18, w: 3.8, h: 0.38, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  var dockItems = [
    "症状は特にないが心配",
    "40歳以上で初めて受ける",
    "家族に脳卒中・動脈瘤歴",
    "高血圧・糖尿病・喫煙あり",
    "定期的な健康チェックとして",
  ];
  dockItems.forEach(function(item, i) {
    card(s, 0.5, 1.66 + i * 0.67, 4.4, 0.57, C.light, C.primary);
    s.addText("✓ " + item, { x: 0.72, y: 1.7 + i * 0.67, w: 4.0, h: 0.48, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });

  card(s, 5.1, 1.1, 4.4, 4.0, C.white, C.red);
  s.addText("🚨 すぐに受診（症状あり）", { x: 5.3, y: 1.18, w: 3.8, h: 0.38, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var urgentItems = [
    "突然の激しい頭痛",
    "手足のしびれ・脱力",
    "ろれつが回らない",
    "片側の視野が欠ける",
    "意識がもうろうとする",
  ];
  urgentItems.forEach(function(item, i) {
    card(s, 5.1, 1.66 + i * 0.67, 4.4, 0.57, C.lightRed, C.red);
    s.addText("！ " + item, { x: 5.32, y: 1.7 + i * 0.67, w: 4.0, h: 0.48, fontSize: 14, fontFace: FJ, color: C.red, valign: "middle", bold: true, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 11: 脳を守る日常習慣
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "脳を守る日常習慣");
  card(s, 0.5, 1.1, 9.0, 0.55, C.lightGreen, C.green);
  s.addText("脳ドックより大切なのは、脳卒中リスクを下げる毎日の習慣です", { x: 0.8, y: 1.18, w: 8.5, h: 0.42, fontSize: 16, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0 });

  var habits = [
    { emoji: "💊", title: "血圧管理", body: "収縮期血圧 130mmHg 未満が目標。自宅での血圧測定が重要" },
    { emoji: "🏃", title: "有酸素運動", body: "週150分以上の中強度運動（ウォーキング等）で脳卒中リスクが低下" },
    { emoji: "🚭", title: "禁煙", body: "喫煙は脳卒中リスクを2〜4倍に。禁煙外来の活用を" },
    { emoji: "🥗", title: "食事の見直し", body: "減塩・野菜・魚中心。DASH食・地中海食を参考に" },
    { emoji: "🧩", title: "認知活動・社会参加", body: "読書・会話・趣味。脳への刺激が認知症予防につながる" },
  ];
  habits.forEach(function(h, i) {
    var y = 1.85 + i * 0.67;
    card(s, 0.5, y, 9.0, 0.57, C.white, C.green);
    s.addText(h.emoji, { x: 0.7, y: y + 0.05, w: 0.55, h: 0.47, fontSize: 22, align: "center", margin: 0 });
    s.addText(h.title, { x: 1.35, y: y + 0.04, w: 2.2, h: 0.26, fontSize: 15, fontFace: FJ, color: C.green, bold: true, margin: 0 });
    s.addText(h.body, { x: 3.7, y: y + 0.08, w: 5.6, h: 0.42, fontSize: 13, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 12: まとめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("今日の3つのポイント", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var points = [
    {
      num: "1",
      color: C.primary,
      bgColor: C.light,
      title: "脳ドックは「リスクのある人」が効果的",
      body: "40歳以上・家族歴・高血圧・糖尿病・喫煙がある方は積極的に受検を。リスクが低い方も「基準値を知る」ために1回受けるのは有益。",
    },
    {
      num: "2",
      color: C.accent,
      bgColor: C.lightYellow,
      title: "「異常あり」でも慌てない",
      body: "未破裂脳動脈瘤も無症候性脳梗塞も、すぐに命に直結するものは少ない。専門医と相談してから治療方針を決めるのが基本。",
    },
    {
      num: "3",
      color: C.green,
      bgColor: C.lightGreen,
      title: "最大の脳卒中予防は毎日の習慣",
      body: "血圧管理・禁煙・運動・食事改善が脳卒中リスクを大幅に下げる。脳ドックはあくまで現状把握ツール。その後の行動が大切。",
    },
  ];
  points.forEach(function(p, i) {
    var y = 1.82 + i * 1.1;
    card(s, 0.5, y, 9.0, 0.98, p.bgColor, p.color);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 0.72, h: 0.98, fill: { color: p.color }, rectRadius: 0.05 });
    s.addText(p.num, { x: 0.5, y: y, w: 0.72, h: 0.98, fontSize: 28, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.title, { x: 1.35, y: y + 0.06, w: 7.9, h: 0.36, fontSize: 16, fontFace: FJ, color: p.color, bold: true, margin: 0 });
    s.addText(p.body, { x: 1.35, y: y + 0.44, w: 7.9, h: 0.48, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 13: エンディング（最終回）
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("🧠", { x: 0.6, y: 0.25, w: 8.8, h: 0.8, fontSize: 42, align: "center", margin: 0 });
  s.addText("シリーズ全20回\nご視聴ありがとうございました", {
    x: 0.6, y: 1.1, w: 8.8, h: 1.1,
    fontSize: 34, fontFace: FJ, color: C.white, bold: true, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 1.5, y: 2.3, w: 7, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("「その症状、大丈夫？」シリーズ完結！", {
    x: 0.6, y: 2.45, w: 8.8, h: 0.5,
    fontSize: 20, fontFace: FJ, color: C.accent, align: "center", bold: true, margin: 0,
  });
  s.addText("#01〜#20まで、脳や神経にまつわる症状・病気を\n脳神経内科医がわかりやすく解説してきました。", {
    x: 0.8, y: 3.05, w: 8.4, h: 0.75,
    fontSize: 16, fontFace: FJ, color: "B0BEC5", align: "center", lineSpacingMultiple: 1.3, margin: 0,
  });
  s.addText("チャンネル登録・高評価をよろしくお願いします！\n引き続き医知創造ラボをよろしくお願いします。", {
    x: 0.8, y: 3.88, w: 8.4, h: 0.75,
    fontSize: 15, fontFace: FJ, color: "90A4AE", align: "center", lineSpacingMultiple: 1.3, margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.75, w: 8.8, h: 0.35,
    fontSize: 12, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

pres.writeFile({ fileName: __dirname + "/brain_dock_slides.pptx" })
  .then(function() { console.log("Done: brain_dock_slides.pptx"); })
  .catch(function(err) { console.error(err); });
