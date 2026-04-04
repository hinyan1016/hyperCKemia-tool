var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【2025-2026年】感染症の最新動向まとめ";

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
  s.addText("医知創造ラボ", { x: 0.4, y: 5.27, w: 3, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("🦠", { x: 0.6, y: 0.5, w: 8.8, h: 0.8, fontSize: 48, align: "center", margin: 0 });
  s.addText("感染症の最新動向 2025-2026", {
    x: 0.6, y: 1.3, w: 8.8, h: 1.2,
    fontSize: 38, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.7, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("医療従事者が押さえるべき8つのトピック", {
    x: 0.6, y: 3.0, w: 8.8, h: 0.6,
    fontSize: 24, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.6,
    fontSize: 14, fontFace: FJ, color: "B0BEC5", align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 全体像 - 8つのトピック
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "2025-2026年の感染症動向 ― 概観");

  var topics = [
    { icon: "🐔", name: "鳥インフルエンザ\nH5N1", color: C.red },
    { icon: "🦠", name: "薬剤耐性菌\nAMR", color: C.primary },
    { icon: "🔬", name: "梅毒の\n急増", color: "9B59B6" },
    { icon: "🫁", name: "マイコプラズマ\n肺炎", color: C.accent },
    { icon: "😷", name: "COVID-19\n変異株・後遺症", color: C.green },
    { icon: "🔴", name: "麻疹の\n再流行", color: "E74C3C" },
    { icon: "💥", name: "mpox\nサル痘", color: "8E44AD" },
    { icon: "⚡", name: "iGAS\n劇症型溶連菌", color: "C0392B" },
  ];
  topics.forEach(function(t, i) {
    var col = i % 4;
    var row = Math.floor(i / 4);
    var xPos = 0.3 + col * 2.4;
    var yPos = 1.2 + row * 2.0;
    card(s, xPos, yPos, 2.2, 1.8, C.white, t.color);
    s.addText(t.icon, { x: xPos, y: yPos + 0.1, w: 2.2, h: 0.5, fontSize: 24, align: "center", margin: 0 });
    s.addText(t.name, { x: xPos + 0.15, y: yPos + 0.7, w: 1.9, h: 0.9, fontSize: 14, fontFace: FJ, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  s.addText("パンデミック後の「immunity debt」＋グローバル化 → 多面的な脅威", {
    x: 0.3, y: 5.0, w: 9.4, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, italic: true, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 3: H5N1 - 疫学
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "🐔 鳥インフルエンザ H5N1 ― 疫学", C.red);

  card(s, 0.3, 1.2, 4.5, 3.8, C.white, C.red);
  s.addText("哺乳類への越境感染", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 18, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  var items = [
    "H5N1 clade 2.3.4.4b が世界的panzootic",
    "南米・南極に初到達（史上最大規模）",
    "欧州ミンク農場、南米アシカで持続感染",
    "米国乳牛で初の大規模アウトブレイク",
    "乳腺に複製 → 乳汁中に高ウイルス量",
    "酪農作業者で複数のヒト感染確認",
  ];
  items.forEach(function(item, i) {
    s.addText("• " + item, { x: 0.6, y: 1.8 + i * 0.5, w: 4.0, h: 0.45, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });

  card(s, 5.2, 1.2, 4.5, 3.8, C.white, C.accent);
  s.addText("現時点のリスク評価", { x: 5.4, y: 1.3, w: 4.1, h: 0.4, fontSize: 18, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("✅ ヒト-ヒト持続的感染：未確認", { x: 5.5, y: 1.9, w: 4.0, h: 0.4, fontSize: 14, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("⚠ PB2にE627K等の哺乳類\n　適応変異が出現", { x: 5.5, y: 2.4, w: 4.0, h: 0.6, fontSize: 13, fontFace: FJ, color: C.red, margin: 0 });
  s.addText("⚠ ヒト型受容体結合変化は\n　未確認だが変異蓄積中", { x: 5.5, y: 3.1, w: 4.0, h: 0.6, fontSize: 13, fontFace: FJ, color: C.red, margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 3.9, w: 4.1, h: 0.9, fill: { color: C.lightYellow }, rectRadius: 0.08 });
  s.addText("備蓄ワクチン（AS03アジュバント）が\nclade 2.3.4.4bに交差中和活性あり\n→ bridging vaccineとして期待", {
    x: 5.5, y: 3.9, w: 3.9, h: 0.9, fontSize: 12, fontFace: FJ, color: C.text, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 4: AMR - 世界的負荷
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "🦠 薬剤耐性菌（AMR）― 深刻化する脅威");

  // 数値カード
  var stats = [
    { label: "2021年\nAMR関連死", value: "471万人", sub: "起因死114万人", color: C.red },
    { label: "2050年\n予測", value: "822万人", sub: "起因死191万人", color: "8B0000" },
    { label: "回避可能\n死亡数", value: "9,200万人", sub: "適切な介入で累積", color: C.green },
  ];
  stats.forEach(function(st, i) {
    var xPos = 0.3 + i * 3.2;
    card(s, xPos, 1.2, 3.0, 2.0, C.white, st.color);
    s.addText(st.label, { x: xPos + 0.2, y: 1.3, w: 2.6, h: 0.6, fontSize: 13, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
    s.addText(st.value, { x: xPos + 0.2, y: 1.9, w: 2.6, h: 0.6, fontSize: 28, fontFace: FE, color: st.color, bold: true, align: "center", margin: 0 });
    s.addText(st.sub, { x: xPos + 0.2, y: 2.5, w: 2.6, h: 0.4, fontSize: 11, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  });

  // 注目ポイント
  card(s, 0.3, 3.5, 9.4, 1.5, C.white, C.primary);
  s.addText("注目すべき変化", { x: 0.6, y: 3.6, w: 4.0, h: 0.3, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("• MRSA関連死：26.1万人（1990）→ 55万人（2021）に倍増\n• WHO 2024年版：カルバペネム耐性 K. pneumoniae が最高スコア（84%）で最上位\n• 5歳未満は50%以上減少、一方70歳以上は80%以上増加", {
    x: 0.6, y: 4.0, w: 8.8, h: 0.9, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 5: AMR - 新規抗菌薬
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "🦠 新規BL/BLI配合抗菌薬 ― 治療選択肢の拡大");

  // テーブル
  var tblH = [
    { text: "薬剤名", options: { fontSize: 14, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
    { text: "主な対象", options: { fontSize: 14, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
    { text: "臨床的位置づけ", options: { fontSize: 14, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
  ];
  var tblR = [
    [
      { text: "Cefepime/\nenmetazobactam", options: { fontSize: 13, fontFace: FE, bold: true } },
      { text: "ESBL産生菌", options: { fontSize: 13, fontFace: FJ } },
      { text: "カルバペネム温存戦略", options: { fontSize: 13, fontFace: FJ } },
    ],
    [
      { text: "Aztreonam/\navibactam", options: { fontSize: 13, fontFace: FE, bold: true } },
      { text: "CRE\n（MBL含む）", options: { fontSize: 13, fontFace: FJ } },
      { text: "MBL産生菌への数少ない選択肢", options: { fontSize: 13, fontFace: FJ } },
    ],
    [
      { text: "Sulbactam/\ndurlobactam", options: { fontSize: 13, fontFace: FE, bold: true } },
      { text: "CRAB", options: { fontSize: 13, fontFace: FJ } },
      { text: "カルバペネム耐性\nA. baumannii の新選択肢", options: { fontSize: 13, fontFace: FJ } },
    ],
  ];
  s.addTable([tblH].concat(tblR), { x: 0.5, y: 1.2, w: 9, colW: [3, 2.5, 3.5], border: { type: "solid", pt: 0.5, color: "DEE2E6" }, rowH: [0.6, 0.8, 0.8, 0.8] });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.2, w: 9, h: 0.8, fill: { color: C.light }, rectRadius: 0.08 });
  s.addText("💡 実臨床では AST活動の推進 と 培養結果に基づくde-escalation の徹底が引き続き重要", {
    x: 0.7, y: 4.2, w: 8.6, h: 0.8, fontSize: 14, fontFace: FJ, color: C.primary, valign: "middle", margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 6: 梅毒 - 疫学
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "🔬 梅毒の急増 ― 世界と日本の疫学", "9B59B6");

  card(s, 0.3, 1.2, 4.5, 3.8, C.white, "9B59B6");
  s.addText("世界的動向", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 18, fontFace: FJ, color: "9B59B6", bold: true, margin: 0 });
  s.addText("• 年間推定800万人が新規感染（JAMA 2025）\n• 米国：2019-2023に61%増加\n　- 女性112%増、先天梅毒106%増\n• MSMが不均衡に影響\n• 高所得国でも先天梅毒が再興", {
    x: 0.6, y: 1.8, w: 4.0, h: 2.8, fontSize: 14, fontFace: FJ, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  card(s, 5.2, 1.2, 4.5, 3.8, C.white, C.red);
  s.addText("日本の状況", { x: 5.4, y: 1.3, w: 4.1, h: 0.4, fontSize: 18, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("2020年を変曲点に急増", { x: 5.5, y: 1.8, w: 4.0, h: 0.4, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("年間変化率 43.8%", { x: 5.5, y: 2.3, w: 4.0, h: 0.5, fontSize: 28, fontFace: FE, color: C.red, bold: true, margin: 0 });
  s.addText("• 非店舗型デリバリーヘルス事業数と\n　強い相関（Mori et al. 2025）\n• 妊婦梅毒有病率：1/1,215\n　→ 早産9%、子宮内胎児死亡2%\n　→ 新生児先天梅毒7%", {
    x: 5.5, y: 2.9, w: 4.0, h: 2.0, fontSize: 13, fontFace: FJ, color: C.text, margin: 0, lineSpacingMultiple: 1.2,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 7: 梅毒 - 診断と治療
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "🔬 梅毒 ― 診断アルゴリズムと治療", "9B59B6");

  // 診断
  card(s, 0.3, 1.2, 4.5, 2.2, C.white, "9B59B6");
  s.addText("診断アルゴリズム", { x: 0.5, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: "9B59B6", bold: true, margin: 0 });
  s.addText("Reverse algorithm\n（海外で普及）", { x: 0.6, y: 1.7, w: 2.0, h: 0.5, fontSize: 13, fontFace: FJ, color: C.text, bold: true, margin: 0 });
  s.addText("TPHA/FTA-ABS\n（特異抗体）", { x: 0.6, y: 2.2, w: 2.0, h: 0.5, fontSize: 13, fontFace: FJ, color: C.primary, margin: 0 });
  s.addText("→", { x: 2.5, y: 2.2, w: 0.3, h: 0.5, fontSize: 16, fontFace: FE, color: C.text, align: "center", valign: "middle", margin: 0 });
  s.addText("RPR\n（活動性評価）", { x: 2.8, y: 2.2, w: 1.8, h: 0.5, fontSize: 13, fontFace: FJ, color: C.accent, margin: 0 });

  s.addText("⚠ 見逃し注意：神経梅毒・眼梅毒・耳梅毒\n　原因不明の視力障害・聴覚障害で鑑別に", {
    x: 0.6, y: 2.8, w: 4.0, h: 0.5, fontSize: 12, fontFace: FJ, color: C.red, margin: 0,
  });

  // 治療
  card(s, 5.2, 1.2, 4.5, 2.2, C.white, C.green);
  s.addText("治療", { x: 5.4, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("第一選択：BPG 240万単位 筋注\n（日本では2022年から使用可能）", {
    x: 5.5, y: 1.7, w: 4.0, h: 0.6, fontSize: 14, fontFace: FJ, color: C.text, bold: true, margin: 0,
  });
  s.addText("PCGアレルギー → Doxycycline\n100mg 1日2回 14日間", {
    x: 5.5, y: 2.4, w: 4.0, h: 0.5, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });
  s.addText("🆕 DoxyPEP：性的曝露後予防投与\n梅毒・クラミジア・淋菌の予防効果", {
    x: 5.5, y: 3.0, w: 4.0, h: 0.4, fontSize: 12, fontFace: FJ, color: C.primary, margin: 0,
  });

  // 妊婦スクリーニング
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.7, w: 9.4, h: 1.2, fill: { color: C.lightRed }, rectRadius: 0.08 });
  s.addText("🚨 先天梅毒予防のために", { x: 0.5, y: 3.75, w: 4.0, h: 0.3, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("• 妊婦スクリーニングの強化（特に妊娠後期の再検査）\n• パートナーの検査・治療も重要\n• 先天梅毒率はWHO目標の8.5倍に達している", {
    x: 0.5, y: 4.1, w: 9.0, h: 0.7, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 8: マイコプラズマ - 耐性と治療
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "🫁 マイコプラズマ肺炎 ― マクロライド耐性時代");

  card(s, 0.3, 1.2, 4.5, 2.0, C.white, C.accent);
  s.addText("再興の背景", { x: 0.5, y: 1.3, w: 4.1, h: 0.3, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("• パンデミック中のNPIにより循環抑制\n• 免疫低下集団で2024-2025年に大規模再興\n• GBD研究：下気道感染の第3位\n　（年間2,530万エピソード）", {
    x: 0.6, y: 1.7, w: 4.0, h: 1.3, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });

  // 耐性テーブル
  var tblH2 = [
    { text: "地域", options: { fontSize: 13, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
    { text: "マクロライド耐性率", options: { fontSize: 13, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
    { text: "第二選択薬", options: { fontSize: 13, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
  ];
  var tblR2 = [
    [
      { text: "中国", options: { fontSize: 12, fontFace: FJ } },
      { text: "80%超", options: { fontSize: 12, fontFace: FE, color: C.red, bold: true } },
      { text: "テトラサイクリン系", options: { fontSize: 12, fontFace: FJ } },
    ],
    [
      { text: "日本・韓国", options: { fontSize: 12, fontFace: FJ } },
      { text: "50-80%", options: { fontSize: 12, fontFace: FE, color: C.accent, bold: true } },
      { text: "テトラサイクリン系\nトスフロキサシン（小児）", options: { fontSize: 11, fontFace: FJ } },
    ],
    [
      { text: "欧米", options: { fontSize: 12, fontFace: FJ } },
      { text: "10-30%", options: { fontSize: 12, fontFace: FE, color: C.green, bold: true } },
      { text: "マクロライドが第一選択", options: { fontSize: 12, fontFace: FJ } },
    ],
  ];
  s.addTable([tblH2].concat(tblR2), { x: 5.2, y: 1.2, w: 4.5, colW: [1.2, 1.4, 1.9], border: { type: "solid", pt: 0.5, color: "DEE2E6" }, rowH: [0.45, 0.45, 0.6, 0.45] });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.5, w: 9.4, h: 1.5, fill: { color: C.light }, rectRadius: 0.08 });
  s.addText("治療戦略（2024年 世界初MRMP小児コンセンサス）", { x: 0.5, y: 3.55, w: 9.0, h: 0.3, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("① マクロライド 48-72時間で初期治療\n② 改善なし → テトラサイクリン系に変更\n③ 難治性 → コルチコステロイド + IVIG 併用を検討", {
    x: 0.6, y: 3.9, w: 9.0, h: 0.9, fontSize: 14, fontFace: FJ, color: C.text, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 9: COVID-19 - 変異株
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "😷 COVID-19 ― 変異株の現在地", C.green);

  card(s, 0.3, 1.2, 9.4, 1.6, C.white, C.primary);
  s.addText("6血清型とT細胞交差免疫", { x: 0.6, y: 1.3, w: 5.0, h: 0.35, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("• SARS-CoV-2は6つの血清型に分類\n• 血清型間でT細胞交差反応性が段階的に低下\n• Serotype VI（BA.2.86/JN.1系統）に対する交差T細胞応答が顕著に弱い\n• → ユニバーサルワクチン開発の障壁", {
    x: 0.6, y: 1.7, w: 8.8, h: 1.0, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });

  card(s, 0.3, 3.1, 9.4, 1.9, C.white, C.accent);
  s.addText("免疫プロファイルとブースター戦略", { x: 0.6, y: 3.2, w: 5.0, h: 0.35, fontSize: 16, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("• XBB.1.5に対する中和抗体は全年齢で著明に低下\n• ワクチン接種は自然感染単独より強い中和活性を誘導\n• 60歳以上で交差中和の改善が最も大きい\n• → 年齢別ブースター戦略の最適化が重要", {
    x: 0.6, y: 3.6, w: 8.8, h: 1.2, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 10: COVID-19 - Long COVID
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "😷 Long COVID ― 治療の新展開", C.green);

  card(s, 0.3, 1.2, 4.5, 2.5, C.white, C.primary);
  s.addText("メトホルミン", { x: 0.5, y: 1.3, w: 4.1, h: 0.4, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("英国62万人超コホート", { x: 0.6, y: 1.7, w: 4.0, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText("PCC発症リスク", { x: 0.6, y: 2.1, w: 2.0, h: 0.3, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("64%低下", { x: 2.5, y: 2.0, w: 2.0, h: 0.5, fontSize: 28, fontFace: FE, color: C.primary, bold: true, margin: 0 });
  s.addText("HR 0.36 (95%CI 0.32-0.41)", { x: 0.6, y: 2.6, w: 4.0, h: 0.3, fontSize: 12, fontFace: FE, color: C.sub, margin: 0 });
  s.addText("対象：BMI≧25 + COVID診断後90日以内に開始", { x: 0.6, y: 3.0, w: 4.0, h: 0.4, fontSize: 11, fontFace: FJ, color: C.text, margin: 0 });

  card(s, 5.2, 1.2, 4.5, 2.5, C.white, C.accent);
  s.addText("メチルプレドニゾロン", { x: 5.4, y: 1.3, w: 4.1, h: 0.4, fontSize: 18, fontFace: FJ, color: C.accent, bold: true, margin: 0 });
  s.addText("post-COVID ILD 229例 RCT", { x: 5.5, y: 1.7, w: 4.0, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText("0.5mg/kg/日投与", { x: 5.5, y: 2.1, w: 4.0, h: 0.3, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("機能改善率", { x: 5.5, y: 2.5, w: 2.0, h: 0.3, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  s.addText("74.2%", { x: 7.3, y: 2.4, w: 1.5, h: 0.5, fontSize: 24, fontFace: FE, color: C.accent, bold: true, margin: 0 });
  s.addText("vs 対照 55.2%（OR 2.33）", { x: 5.5, y: 3.0, w: 4.0, h: 0.3, fontSize: 12, fontFace: FE, color: C.sub, margin: 0 });

  // 病態メカニズム
  card(s, 0.3, 4.0, 9.4, 1.0, C.light, C.primary);
  s.addText("🔬 病態機序：NK細胞の転写リプログラミング、マスト細胞活性化 → 小径線維神経障害・自律神経障害", {
    x: 0.6, y: 4.1, w: 8.8, h: 0.7, fontSize: 13, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 11: 麻疹
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "🔴 麻疹の再流行 ― ワクチンギャップ", "E74C3C");

  card(s, 0.3, 1.2, 4.5, 2.5, C.white, "E74C3C");
  s.addText("世界的な再流行", { x: 0.5, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: "E74C3C", bold: true, margin: 0 });
  s.addText("• 米国を含む複数の高所得国で再流行\n• 主因：MMR接種率の低下＋ワクチン忌避\n• WHOの95%集団免疫閾値を下回る地域が増加\n• イタリア：18-24歳の感受性率 39.8%\n　2回接種済でも69.6%が感受性", {
    x: 0.6, y: 1.7, w: 4.0, h: 2.0, fontSize: 13, fontFace: FJ, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  card(s, 5.2, 1.2, 4.5, 2.5, C.white, C.primary);
  s.addText("免疫ギャップの実態", { x: 5.4, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText("• アイルランド：10-17歳で\n　血清陽性率87-89%（閾値95%未達）\n• 自然感染はワクチンより持続的な免疫\n　→ ワクチンのみ世代で免疫waning\n• 3回目接種の有効性は未証明\n　→ まず2回接種の徹底が最優先", {
    x: 5.5, y: 1.7, w: 4.0, h: 2.0, fontSize: 13, fontFace: FJ, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.0, w: 9.4, h: 1.0, fill: { color: C.lightYellow }, rectRadius: 0.08 });
  s.addText("🏥 医療従事者への推奨：自身の麻疹抗体価確認と追加接種。特に1978-1990年生まれ（接種1回世代）は要確認", {
    x: 0.5, y: 4.1, w: 9.0, h: 0.8, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 12: mpox
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "💥 mpox（サル痘）― クレードIbの拡散", "8E44AD");

  card(s, 0.3, 1.2, 4.5, 2.3, C.white, "8E44AD");
  s.addText("クレードIb：新たなリスク", { x: 0.5, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: "8E44AD", bold: true, margin: 0 });
  s.addText("• 2023年後半にコンゴで検出\n• アフリカ外に拡散（アジア含む）\n• インド：10例確認（90% UAE渡航歴）\n• APOBEC3変異 → 持続的ヒト-ヒト感染示唆\n• アフリカ全体：9,446例\n• COVID後のmpoxリスク約4倍（IRR 3.77）", {
    x: 0.6, y: 1.7, w: 4.0, h: 1.7, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });

  card(s, 5.2, 1.2, 4.5, 2.3, C.white, C.green);
  s.addText("治療・ワクチン", { x: 5.4, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("前臨床で有効な3剤：\n• テコビリマト\n• シドフォビル\n• ブリンシドフォビル\n\nワクチン：MVA-BN（JYNNEOS）\nACAM2000より安全・免疫原性に優れる\n→ 優先使用が推奨", {
    x: 5.5, y: 1.7, w: 4.0, h: 1.7, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.8, w: 9.4, h: 0.6, fill: { color: "F0E6F6" }, rectRadius: 0.08 });
  s.addText("感染症法分類：4類 ｜ 水疱膿疱性皮疹＋リンパ節腫脹（渡航歴・接触歴）で疑う ｜ PCR（皮膚病変部）で確定", {
    x: 0.5, y: 3.85, w: 9.0, h: 0.5, fontSize: 12, fontFace: FJ, color: "8E44AD", valign: "middle", margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 13: iGAS
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "⚡ 侵襲性A群溶連菌（iGAS）― 劇症型の急増", "C0392B");

  card(s, 0.3, 1.2, 4.5, 2.3, C.white, "C0392B");
  s.addText("疫学", { x: 0.5, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: "C0392B", bold: true, margin: 0 });
  s.addText("日本 STSS（2024年6月時点）", { x: 0.6, y: 1.7, w: 4.0, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText("977例", { x: 0.6, y: 2.0, w: 4.0, h: 0.6, fontSize: 36, fontFace: FE, color: "C0392B", bold: true, margin: 0 });
  s.addText("急増の要因：\n• NPI解除による免疫低下\n• 呼吸器ウイルスとの相互作用\n• M1UK等の超毒性クローン拡大\n　（SpeAスーパー抗原の発現増加）", {
    x: 0.6, y: 2.6, w: 4.0, h: 1.0, fontSize: 12, fontFace: FJ, color: C.text, margin: 0,
  });

  card(s, 5.2, 1.2, 4.5, 2.3, C.white, C.red);
  s.addText("臨床像と予後", { x: 5.4, y: 1.3, w: 4.1, h: 0.35, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("emm1 iGAS 30日死亡率", { x: 5.5, y: 1.7, w: 4.0, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, margin: 0 });
  s.addText("24%", { x: 5.5, y: 2.0, w: 4.0, h: 0.6, fontSize: 36, fontFace: FE, color: C.red, bold: true, margin: 0 });
  s.addText("• 死亡の63.7%が検体採取1日以内\n• 15歳未満：致死例の56.3%が\n　検体採取前に死亡\n• IVIG：iGAS全体への有効性は不確定\n　（日本全国調査 481例）", {
    x: 5.5, y: 2.6, w: 4.0, h: 1.0, fontSize: 12, fontFace: FJ, color: C.text, margin: 0,
  });

  // Red flags
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.8, w: 9.4, h: 1.2, fill: { color: C.lightRed }, rectRadius: 0.08 });
  s.addText("🚨 劇症型を疑うべき所見", { x: 0.5, y: 3.85, w: 4.0, h: 0.3, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  s.addText("• 急速に進行する軟部組織の疼痛（壊死性筋膜炎）\n• 低血圧・頻脈・意識障害（敗血症性ショック）\n• 先行する軽微な外傷・水痘・インフルエンザ後の急激な悪化", {
    x: 0.5, y: 4.2, w: 9.0, h: 0.7, fontSize: 13, fontFace: FJ, color: C.text, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// SLIDE 14: チェックリスト
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "📋 医療従事者のための実践チェックリスト");

  var tblH3 = [
    { text: "感染症", options: { fontSize: 12, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
    { text: "いつ疑う", options: { fontSize: 12, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
    { text: "主な検査", options: { fontSize: 12, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
    { text: "届出", options: { fontSize: 12, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary }, align: "center" } },
  ];
  var tblR3 = [
    [
      { text: "H5N1", options: { fontSize: 11, fontFace: FE, bold: true } },
      { text: "鳥・乳牛接触+呼吸器/結膜症状", options: { fontSize: 10, fontFace: FJ } },
      { text: "PCR（鼻咽頭・結膜）", options: { fontSize: 10, fontFace: FJ } },
      { text: "2類", options: { fontSize: 10, fontFace: FJ, align: "center" } },
    ],
    [
      { text: "梅毒", options: { fontSize: 11, fontFace: FJ, bold: true } },
      { text: "無痛性潰瘍,皮疹,視力・聴力障害", options: { fontSize: 10, fontFace: FJ } },
      { text: "RPR+TPHA", options: { fontSize: 10, fontFace: FJ } },
      { text: "5類全数", options: { fontSize: 10, fontFace: FJ, align: "center" } },
    ],
    [
      { text: "マイコ\nプラズマ", options: { fontSize: 10, fontFace: FJ, bold: true } },
      { text: "乾性咳嗽が遷延,ML 48-72h無効", options: { fontSize: 10, fontFace: FJ } },
      { text: "PCR, 迅速抗原", options: { fontSize: 10, fontFace: FJ } },
      { text: "5類定点", options: { fontSize: 10, fontFace: FJ, align: "center" } },
    ],
    [
      { text: "麻疹", options: { fontSize: 11, fontFace: FJ, bold: true } },
      { text: "発熱+発疹+カタル（渡航歴）", options: { fontSize: 10, fontFace: FJ } },
      { text: "IgM, PCR", options: { fontSize: 10, fontFace: FJ } },
      { text: "5類全数\n（直ちに）", options: { fontSize: 10, fontFace: FJ, align: "center" } },
    ],
    [
      { text: "mpox", options: { fontSize: 11, fontFace: FE, bold: true } },
      { text: "水疱膿疱性皮疹+リンパ節腫脹", options: { fontSize: 10, fontFace: FJ } },
      { text: "PCR（皮膚病変）", options: { fontSize: 10, fontFace: FJ } },
      { text: "4類", options: { fontSize: 10, fontFace: FJ, align: "center" } },
    ],
    [
      { text: "iGAS\n/STSS", options: { fontSize: 10, fontFace: FJ, bold: true } },
      { text: "急速進行の軟部組織痛,ショック", options: { fontSize: 10, fontFace: FJ } },
      { text: "血培, 迅速抗原", options: { fontSize: 10, fontFace: FJ } },
      { text: "5類全数\n（STSS）", options: { fontSize: 10, fontFace: FJ, align: "center" } },
    ],
  ];
  s.addTable([tblH3].concat(tblR3), { x: 0.3, y: 1.1, w: 9.4, colW: [1.2, 3.2, 2.5, 1.2], border: { type: "solid", pt: 0.5, color: "DEE2E6" }, rowH: [0.5, 0.5, 0.5, 0.6, 0.5, 0.5, 0.5] });
  ftr(s);
})();

// ============================================================
// SLIDE 15: Take Home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.6, y: 0.4, w: 8.8, h: 0.7,
    fontSize: 30, fontFace: FE, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var msgs = [
    { icon: "🐔", text: "H5N1：乳牛→ヒト感染の新経路。ヒト-ヒト感染は未確認だが監視継続" },
    { icon: "🦠", text: "AMR：2050年に822万人。新規BL/BLI配合薬で治療選択肢拡大" },
    { icon: "🔬", text: "梅毒：日本で年43.8%増。妊婦スクリーニングとDoxyPEPに注目" },
    { icon: "🫁", text: "マイコプラズマ：東アジアではML耐性率50-80%。テトラサイクリン系が代替" },
    { icon: "😷", text: "Long COVID：メトホルミンでPCCリスク64%低下（BMI≧25）" },
    { icon: "🔴", text: "麻疹：ワクチンギャップ。1978-1990年生まれは抗体価確認を" },
    { icon: "💥", text: "mpox：クレードIbが世界拡散。MVA-BNワクチンを優先" },
    { icon: "⚡", text: "iGAS：STSS 30日死亡率24%。培養前に作用する早期治療が鍵" },
  ];
  msgs.forEach(function(m, i) {
    var yPos = 1.2 + i * 0.52;
    s.addText(m.icon + " " + m.text, {
      x: 0.6, y: yPos, w: 8.8, h: 0.48,
      fontSize: 13, fontFace: FJ, color: C.white, valign: "middle", margin: 0,
    });
  });

  s.addText("ご視聴ありがとうございました", {
    x: 0.6, y: 5.0, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FJ, color: "B0BEC5", align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 16: 参考文献
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "📚 主要参考文献");

  var refs = [
    "[1] GBD 2021 LRI Collaborators. Lancet Infect Dis 2024",
    "[2] Peacock TP, et al. Nature 2024",
    "[7] GBD 2021 AMR Collaborators. Lancet 2024",
    "[12] Chevalier FJ, et al. JAMA 2025",
    "[17] Wang YS, et al. World J Pediatr 2024",
    "[19] Guo Y, et al. Emerg Microbes Infect 2026",
    "[21] Chaichana U, et al. Clin Infect Dis 2026",
    "[25] Edward M. Global Public Health 2026",
    "[30] Judson SD, et al. Lancet Infect Dis 2026",
    "[35] Flamant A, et al. Lancet Infect Dis 2025",
    "[38] Li HK, et al. Clin Infect Dis 2025",
  ];
  refs.forEach(function(r, i) {
    s.addText(r, { x: 0.5, y: 1.1 + i * 0.35, w: 9, h: 0.3, fontSize: 11, fontFace: FE, color: C.text, margin: 0 });
  });
  s.addText("全39文献の詳細はブログ記事をご参照ください", {
    x: 0.5, y: 5.0, w: 9, h: 0.3, fontSize: 12, fontFace: FJ, color: C.sub, italic: true, margin: 0,
  });
  ftr(s);
})();

// ============================================================
// 出力
// ============================================================
var outFile = "感染症最新動向2025_slides.pptx";
pres.writeFile({ fileName: outFile }).then(function() {
  console.log("Created: " + outFile);
}).catch(function(err) {
  console.error("Error:", err);
});
