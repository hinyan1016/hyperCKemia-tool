var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "片頭痛持ちが知っておくべき5つのこと ― 脳神経内科医が伝えたい「我慢しない」ための知識";

// --- Color Palette ---
var C = {
  dark: "1B3A5C",
  primary: "2C5AA0",
  accent: "4A90D9",
  light: "E8F0FE",
  warmBg: "F5F7FA",
  white: "FFFFFF",
  text: "2D3436",
  sub: "636E72",
  red: "DC3545",
  orange: "E8850C",
  yellow: "F5C518",
  green: "28A745",
};

var FONT_JP = "BIZ UDPGothic";
var FONT_EN = "Segoe UI";
var FONT_H = FONT_JP;
var FONT_B = FONT_JP;

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, margin: 0 });
}

function addCard(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: C.white }, shadow: makeShadow() });
}

function addPageNum(s, num) {
  s.addText(num, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_B, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("脳神経内科医が答える からだの不思議 #02", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("片頭痛持ちが知っておくべき\n5つのこと", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.8,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 「我慢しない」ための知識 ―", {
    x: 0.8, y: 3.0, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.75, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.0, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 片頭痛とは ― どれくらいいるの？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "片頭痛とは ― どれくらいいるの？");

  // 3つの統計カード
  var stats = [
    { num: "約840万人", label: "日本の片頭痛患者数", sub: "（約15人に1人）", color: C.primary },
    { num: "約3割", label: "実際に受診している割合", sub: "（約600万人が我慢中）", color: C.orange },
    { num: "2位", label: "日常生活への支障度", sub: "（WHO・全疾患中）", color: C.red },
  ];
  stats.forEach(function(st, i) {
    var xPos = 0.4 + i * 3.1;
    addCard(s, xPos, 1.1, 2.9, 3.0);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 2.9, h: 0.55, fill: { color: st.color } });
    s.addText(st.label, { x: xPos, y: 1.1, w: 2.9, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.num, { x: xPos, y: 1.75, w: 2.9, h: 1.3, fontSize: 38, fontFace: FONT_JP, color: st.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.sub, { x: xPos, y: 3.05, w: 2.9, h: 0.9, fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", valign: "top", margin: 0 });
  });

  // メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.35, w: 9.2, h: 0.65, fill: { color: C.light } });
  s.addText("「たかが頭痛」ではなく、日常生活に大きな支障をきたす疾患", {
    x: 0.4, y: 4.35, w: 9.2, h: 0.65,
    fontSize: 17, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: ポイント1 ― 脳の過敏体質
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ポイント 1 ｜ 片頭痛は「脳の過敏体質」");

  // メカニズム説明カード
  addCard(s, 0.5, 1.1, 9.0, 1.4);
  s.addText("三叉神経血管系が活性化 → CGRP大量放出 → 血管周囲の炎症 → ズキズキした拍動性の痛み", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.1,
    fontSize: 16, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    lineSpacingMultiple: 1.3,
  });

  // 症状バッジ
  var symptoms = [
    { icon: "☀", text: "光がまぶしい", color: C.yellow },
    { icon: "🔊", text: "音がうるさい", color: C.orange },
    { icon: "👃", text: "においで気分不良", color: C.green },
  ];
  symptoms.forEach(function(sy, i) {
    var xPos = 0.5 + i * 3.1;
    addCard(s, xPos, 2.75, 2.9, 1.3);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.75, w: 2.9, h: 0.45, fill: { color: sy.color } });
    s.addText("脳の過剰反応", { x: xPos, y: 2.75, w: 2.9, h: 0.45, fontSize: 13, fontFace: FONT_JP, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(sy.icon + " " + sy.text, { x: xPos, y: 3.25, w: 2.9, h: 0.75, fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // 遺伝ポイント
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.25, w: 9.0, h: 0.85, fill: { color: C.light } });
  s.addText("遺伝的要因が大きい：親が片頭痛の場合、子どもも片頭痛になる確率は 50〜75%", {
    x: 0.7, y: 4.3, w: 8.6, h: 0.75,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: ポイント2 ― 薬物乱用頭痛
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ポイント 2 ｜ 鎮痛薬の飲みすぎが頭痛を悪化させる");

  // 悪循環の矢印フロー
  var steps = ["頭痛が怖くて\n早めに薬を飲む", "効きにくく\nなってくる", "もっと頻繁に\n飲むようになる", "さらに頭痛が\n増えていく"];
  steps.forEach(function(st, i) {
    var xPos = 0.25 + i * 2.4;
    addCard(s, xPos, 1.1, 2.1, 1.1);
    s.addText(st, { x: xPos, y: 1.1, w: 2.1, h: 1.1, fontSize: 14, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    if (i < 3) {
      s.addText("→", { x: xPos + 2.1, y: 1.3, w: 0.3, h: 0.7, fontSize: 22, fontFace: FONT_EN, color: C.red, align: "center", valign: "middle", margin: 0 });
    }
  });

  // 表 ヘッダー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.5, w: 3.2, h: 0.45, fill: { color: C.primary } });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.6, y: 2.5, w: 2.7, h: 0.45, fill: { color: C.primary } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 2.5, w: 3.1, h: 0.45, fill: { color: C.primary } });
  s.addText("薬の種類", { x: 0.4, y: 2.5, w: 3.2, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("MOH の目安", { x: 3.6, y: 2.5, w: 2.7, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("注意点", { x: 6.3, y: 2.5, w: 3.1, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // 表 行データ
  var rows = [
    ["市販複合鎮痛薬", "月10日以上", "カフェイン配合は特にリスク大"],
    ["NSAIDs（ロキソプロフェン等）", "月15日以上", "単一成分でも頻用は注意"],
    ["トリプタン製剤", "月10日以上", "片頭痛専用薬でも過剰使用あり"],
  ];
  rows.forEach(function(row, i) {
    var yPos = 2.95 + i * 0.6;
    var bg = (i % 2 === 1) ? "F8F9FA" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 9.0, h: 0.55, fill: { color: bg } });
    s.addText(row[0], { x: 0.5, y: yPos, w: 3.0, h: 0.55, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
    s.addText(row[1], { x: 3.6, y: yPos, w: 2.7, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(row[2], { x: 6.3, y: yPos, w: 3.1, h: 0.55, fontSize: 12, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0 });
  });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: ポイント3 ― 予防薬
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ポイント 3 ｜ 「予防薬」という選択肢がある");

  // 対象目安バー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.05, w: 9.0, h: 0.55, fill: { color: C.accent } });
  s.addText("月に 2回以上 片頭痛がある方 → 予防療法の対象！", {
    x: 0.5, y: 1.05, w: 9.0, h: 0.55,
    fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 2カラム：従来薬 / 新薬
  // 従来薬
  addCard(s, 0.4, 1.8, 4.2, 2.9);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.8, w: 4.2, h: 0.5, fill: { color: C.primary } });
  s.addText("従来の予防薬（保険適用）", { x: 0.4, y: 1.8, w: 4.2, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  var conventional = ["バルプロ酸（抗てんかん薬）", "プロプラノロール（β遮断薬）", "アミトリプチリン（抗うつ薬）", "ロメリジン（Ca拮抗薬）"];
  conventional.forEach(function(drug, i) {
    s.addText("・" + drug, { x: 0.6, y: 2.4 + i * 0.5, w: 3.8, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  // 新薬
  addCard(s, 5.0, 1.8, 4.6, 2.9);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y: 1.8, w: 4.6, h: 0.5, fill: { color: C.orange } });
  s.addText("新しい予防薬（CGRP関連抗体薬）", { x: 5.0, y: 1.8, w: 4.6, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("2021年〜 日本でも使用可能", { x: 5.2, y: 2.4, w: 4.2, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0 });
  var newDrugs = ["エレヌマブ", "フレマネズマブ", "ガルカネズマブ"];
  newDrugs.forEach(function(drug, i) {
    s.addText("・" + drug, { x: 5.2, y: 2.85 + i * 0.45, w: 4.2, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });
  s.addText("月1回皮下注射で片頭痛日数を\n平均 約50% 減少", { x: 5.2, y: 4.25, w: 4.2, h: 0.85, fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: ポイント4 ― トリガーを知る
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ポイント 4 ｜ 「自分のトリガー」を知る");

  // トリガー表
  var triggers = [
    { cat: "ストレス", detail: "強い精神的ストレス、解放直後の「週末頭痛」" },
    { cat: "睡眠", detail: "睡眠不足、寝すぎ（休日の朝寝坊）、時差ボケ" },
    { cat: "ホルモン", detail: "月経前〜月経中（エストロゲンの急激な低下）" },
    { cat: "食事", detail: "空腹・欠食、赤ワイン、カフェインの急な中断" },
    { cat: "環境", detail: "気圧低下、強い光・騒音・強いにおい" },
    { cat: "身体", detail: "肩こり・首こり、激しい運動、脱水" },
  ];
  triggers.forEach(function(tr, i) {
    var yPos = 1.05 + i * 0.57;
    var bg = (i % 2 === 0) ? C.light : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 9.2, h: 0.52, fill: { color: bg } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 1.5, h: 0.52, fill: { color: C.primary } });
    s.addText(tr.cat, { x: 0.4, y: yPos, w: 1.5, h: 0.52, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(tr.detail, { x: 2.0, y: yPos, w: 7.4, h: 0.52, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  // ポイントバー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.55, w: 9.2, h: 0.65, fill: { color: C.yellow } });
  s.addText("複数のトリガーが重なったとき「閾値」を超えて発作が起こる", {
    x: 0.4, y: 4.55, w: 9.2, h: 0.65,
    fontSize: 16, fontFace: FONT_JP, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: ポイント5 ― いつもと違う頭痛
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ポイント 5 ｜ 「いつもと違う頭痛」は別の病気のサイン");

  // 緊急ラベル
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.05, w: 9.2, h: 0.5, fill: { color: C.red } });
  s.addText("すぐに救急受診すべきサイン", {
    x: 0.4, y: 1.05, w: 9.2, h: 0.5,
    fontSize: 17, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var emergency = [
    "「人生最悪の頭痛」突然の激しい頭痛 → くも膜下出血の可能性",
    "発熱 ＋ 頭痛 ＋ 首のこわばり → 髄膜炎の可能性",
    "手足の麻痺・しびれ、言葉が出にくい → 脳卒中の可能性",
    "日に日に悪化する頭痛（特に朝が強い）→ 脳腫瘍等の可能性",
  ];
  emergency.forEach(function(em, i) {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.65 + i * 0.55, w: 9.2, h: 0.5, fill: { color: "FFF5F5" } });
    s.addText("⚠ " + em, {
      x: 0.6, y: 1.65 + i * 0.55, w: 8.8, h: 0.5,
      fontSize: 14, fontFace: FONT_JP, color: C.red, valign: "middle", margin: 0,
    });
  });

  // 相談すべきサイン
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.95, w: 9.2, h: 0.45, fill: { color: C.orange } });
  s.addText("かかりつけ医に相談すべきサイン", {
    x: 0.4, y: 3.95, w: 9.2, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  var caution = ["頭痛のパターン（痛み方・場所・持続時間）が明らかに変わった", "いつもの薬がまったく効かない　　50歳以降に初めて頭痛が出現した"];
  caution.forEach(function(ca, i) {
    s.addText("・" + ca, { x: 0.6, y: 4.5 + i * 0.45, w: 8.8, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: 前兆のある片頭痛 ― 閃輝暗点
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "補足 ｜ 前兆（閃輝暗点）のある片頭痛");

  // 割合表示
  addCard(s, 0.5, 1.1, 4.0, 1.5);
  s.addText("前兆のある片頭痛", { x: 0.5, y: 1.1, w: 4.0, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("全患者の 約2〜3割", { x: 0.5, y: 1.6, w: 4.0, h: 0.9, fontSize: 24, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });

  // 閃輝暗点の説明カード
  addCard(s, 4.8, 1.1, 4.8, 1.5);
  s.addText("閃輝暗点（せんきあんてん）", { x: 4.8, y: 1.1, w: 4.8, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("視野にキラキラした光の輪・ギザギザ模様が広がり視野が欠ける", {
    x: 5.0, y: 1.65, w: 4.4, h: 0.9,
    fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
  });

  // 経過
  addCard(s, 0.5, 2.85, 9.0, 1.0);
  s.addText("通常 20〜30分で消失 → 頭痛発作がつづく", {
    x: 0.7, y: 2.85, w: 8.6, h: 1.0,
    fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 注意点
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 9.0, h: 1.15, fill: { color: "FFF3CD" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 0.2, h: 1.15, fill: { color: C.orange } });
  s.addText("視野障害が片側だけで突然始まり、しびれ・麻痺を伴う場合は\n脳梗塞との鑑別が必要 → 医療機関を受診してください", {
    x: 0.85, y: 4.05, w: 8.5, h: 1.05,
    fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
  });

  addPageNum(s, "8/9");
})();

// ============================================================
// SLIDE 9: まとめ ― 5つのポイント
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("まとめ ― 5つのポイント", {
    x: 0.5, y: 0.15, w: 9.0, h: 0.65,
    fontSize: 24, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  var points = [
    { num: "1", text: "片頭痛は「脳の神経が過敏に反応する体質的な疾患」。「気のせい」ではない", color: C.accent },
    { num: "2", text: "鎮痛薬の月10〜15日以上の使用は「薬物乱用頭痛」のリスク。飲みすぎ注意", color: C.yellow },
    { num: "3", text: "「予防薬」で頭痛の頻度を減らす治療がある。月2回以上なら相談を", color: C.green },
    { num: "4", text: "頭痛ダイアリーで自分のトリガーを把握しコントロール", color: C.orange },
    { num: "5", text: "「いつもと違う頭痛」は別の危険な病気のサイン。迷ったら受診", color: C.red },
  ];

  points.forEach(function(pt, i) {
    var yPos = 0.98 + i * 0.87;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 0.7, h: 0.72, fill: { color: pt.color } });
    s.addText(pt.num, { x: 0.4, y: yPos, w: 0.7, h: 0.72, fontSize: 22, fontFace: FONT_JP, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 1.2, y: yPos, w: 8.4, h: 0.72, fill: { color: "1E3A5C" } });
    s.addText(pt.text, { x: 1.35, y: yPos, w: 8.1, h: 0.72, fontSize: 15, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0 });
  });

  s.addText("医知創造ラボ", {
    x: 0.4, y: 5.3, w: 9.2, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "right", margin: 0,
  });
})();

// ============================================================
// OUTPUT
// ============================================================
var outPath = __dirname + "/片頭痛_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
