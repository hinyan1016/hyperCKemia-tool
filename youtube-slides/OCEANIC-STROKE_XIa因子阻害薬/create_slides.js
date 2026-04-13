var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "OCEANIC-STROKE試験の衝撃 ― XIa因子阻害薬が脳梗塞二次予防を変える";

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
  s.addText(num, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("脳梗塞二次予防の新時代", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("OCEANIC-STROKE試験の衝撃", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.0,
    fontSize: 38, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addText("XIa因子阻害薬 asundexian が二次予防を変える", {
    x: 0.8, y: 2.3, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.15, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText([
    { text: "ISC 2026  |  ", options: { color: C.light, fontSize: 14 } },
    { text: "Sharma M, et al. Abstract LB009", options: { color: C.accent, fontSize: 14, italic: true } },
  ], { x: 0.8, y: 3.4, w: 8.4, h: 0.4, fontFace: FONT_EN, align: "center", margin: 0 });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 非心原性脳梗塞の二次予防 ― 現状の課題
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "非心原性脳梗塞の二次予防 ― 現状の課題");

  // 現状カード
  addCard(s, 0.4, 1.1, 9.2, 1.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 9.2, h: 0.45, fill: { color: C.primary } });
  s.addText("現在の標準治療", { x: 0.7, y: 1.1, w: 8.6, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "抗血小板薬", options: { bold: true, color: C.primary, fontSize: 16 } },
    { text: "（アスピリン、クロピドグレルなど）が標準。", options: { color: C.text, fontSize: 15 } },
    { text: "\n急性期はDAPT（二剤併用）→ 慢性期は単剤が一般的", options: { color: C.sub, fontSize: 14 } },
  ], { x: 0.7, y: 1.6, w: 8.6, h: 0.8, fontFace: FONT_JP, valign: "middle", margin: 0 });

  // 3つの課題
  var issues = [
    { icon: "1", title: "再発リスクが残る", desc: "抗血小板薬にもかかわらず\n年間4〜8%が脳卒中再発", color: C.red },
    { icon: "2", title: "DOACはESUSに無効", desc: "NAVIGATE ESUS・RE-SPECT\nESUS ともに失敗", color: C.orange },
    { icon: "3", title: "出血と効果のジレンマ", desc: "抗凝固薬を強化すると\n出血リスクが増大", color: C.red },
  ];
  issues.forEach(function(p, i) {
    var xPos = 0.4 + i * 3.13;
    addCard(s, xPos, 2.75, 2.93, 2.4);
    s.addShape(pres.shapes.OVAL, { x: xPos + 1.1, y: 2.85, w: 0.7, h: 0.7, fill: { color: p.color } });
    s.addText(p.icon, { x: xPos + 1.1, y: 2.85, w: 0.7, h: 0.7, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.title, { x: xPos + 0.15, y: 3.65, w: 2.63, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.desc, { x: xPos + 0.15, y: 4.1, w: 2.63, h: 0.95, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "2/10");
})();

// ============================================================
// SLIDE 3: 凝固カスケード ― 外因系 vs 内因系
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "凝固カスケード ― 外因系 vs 内因系");

  // 外因系カード（左）
  addCard(s, 0.3, 1.15, 4.5, 3.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.15, w: 4.5, h: 0.5, fill: { color: C.red } });
  s.addText("外因系（組織因子経路）", { x: 0.6, y: 1.15, w: 3.9, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "組織因子 → VIIa → ", options: { fontSize: 14, color: C.text } },
    { text: "Xa → IIa", options: { fontSize: 16, color: C.red, bold: true } },
  ], { x: 0.5, y: 1.8, w: 4.1, h: 0.4, fontFace: FONT_EN, margin: 0 });

  var extItems = [
    "生理的止血に必須",
    "組織損傷 → 出血を止める防御機構",
    "DOACの標的（Xa・IIa）がここにある",
    "→ DOAC使用で出血リスク増大",
  ];
  extItems.forEach(function(item, i) {
    var yy = 2.35 + i * 0.5;
    s.addText(item, { x: 0.7, y: yy, w: 3.9, h: 0.45, fontSize: 12, fontFace: FONT_JP, color: i < 2 ? C.text : C.red, margin: 0, lineSpacingMultiple: 1.1 });
  });

  // 内因系カード（右）
  addCard(s, 5.2, 1.15, 4.5, 3.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.15, w: 4.5, h: 0.5, fill: { color: C.green } });
  s.addText("内因系（接触活性化経路）", { x: 5.5, y: 1.15, w: 3.9, h: 0.5, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "XII → ", options: { fontSize: 14, color: C.text } },
    { text: "XIa", options: { fontSize: 16, color: C.green, bold: true } },
    { text: " → IXa → Xa → IIa", options: { fontSize: 14, color: C.text } },
  ], { x: 5.4, y: 1.8, w: 4.1, h: 0.4, fontFace: FONT_EN, margin: 0 });

  var intItems = [
    "病的血栓形成に関与",
    "プラーク破綻・炎症・乱流で活性化",
    "Asundexianの標的（XIa）がここにある",
    "→ 出血リスクを増やさず血栓を予防",
  ];
  intItems.forEach(function(item, i) {
    var yy = 2.35 + i * 0.5;
    s.addText(item, { x: 5.6, y: yy, w: 3.9, h: 0.45, fontSize: 12, fontFace: FONT_JP, color: i < 2 ? C.text : C.green, margin: 0, lineSpacingMultiple: 1.1 });
  });

  // 下部メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.55, fill: { color: C.light } });
  s.addText([
    { text: "KEY: ", options: { bold: true, color: C.primary, fontSize: 13 } },
    { text: "XIa因子は病的血栓にのみ関与 → 阻害しても生理的止血は温存される", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.6, y: 5.0, w: 8.8, h: 0.55, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, "3/10");
})();

// ============================================================
// SLIDE 4: XI因子欠損症 ― 自然界のエビデンス
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "XI因子欠損症 ― 自然界のエビデンス");

  addCard(s, 0.4, 1.15, 9.2, 1.6);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 9.2, h: 0.45, fill: { color: C.primary } });
  s.addText("先天性XI因子欠損症の疫学データ", { x: 0.7, y: 1.15, w: 8.6, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "XI因子が生まれつき欠損している患者では、", options: { color: C.text, fontSize: 14 } },
    { text: "虚血性脳卒中・深部静脈血栓症のリスクが低下", options: { color: C.green, fontSize: 14, bold: true } },
    { text: "する。\n一方で、", options: { color: C.text, fontSize: 14 } },
    { text: "出血性脳卒中のリスク増大は認められない", options: { color: C.primary, fontSize: 14, bold: true } },
    { text: "。", options: { color: C.text, fontSize: 14 } },
  ], { x: 0.7, y: 1.65, w: 8.6, h: 1.0, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });

  // 2つの論文カード
  var papers = [
    { author: "Salomon O, et al.", journal: "Thromb Haemost 2011", finding: "重度XI因子欠損症で\n深部静脈血栓症リスクが有意に低下" },
    { author: "Zucker M, et al.", journal: "Blood 2017", finding: "XI因子欠損が心血管イベント\nおよびVTEリスク低下と関連" },
  ];
  papers.forEach(function(p, i) {
    var xPos = 0.4 + i * 4.7;
    addCard(s, xPos, 3.0, 4.5, 1.7);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.0, w: 4.5, h: 0.4, fill: { color: C.accent } });
    s.addText(p.author, { x: xPos + 0.15, y: 3.0, w: 4.2, h: 0.4, fontSize: 12, fontFace: FONT_EN, color: C.white, bold: true, valign: "middle", margin: 0 });
    s.addText(p.journal, { x: xPos + 0.15, y: 3.45, w: 4.2, h: 0.35, fontSize: 11, fontFace: FONT_EN, color: C.sub, italic: true, margin: 0 });
    s.addText(p.finding, { x: xPos + 0.15, y: 3.85, w: 4.2, h: 0.75, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2 });
  });

  // メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.9, w: 9.2, h: 0.65, fill: { color: C.light } });
  s.addText([
    { text: "結論: ", options: { bold: true, color: C.primary, fontSize: 14 } },
    { text: "「血栓ができにくく、出血もしにくい」 ― この自然実験がXIa因子阻害薬の合理性を裏付ける", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.7, y: 4.9, w: 8.6, h: 0.65, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, "4/10");
})();

// ============================================================
// SLIDE 5: Asundexianの薬理プロファイル
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "Asundexianの薬理プロファイル");

  s.addText("Heitmeier S, et al. J Thromb Haemost 2022", { x: 0.5, y: 0.92, w: 6, h: 0.3, fontSize: 10, fontFace: FONT_EN, color: C.sub, italic: true, margin: 0 });

  // テーブル形式
  var rows = [
    ["作用機序", "XIa因子活性中心への直接的・可逆的阻害（経口低分子）"],
    ["阻害活性（IC50）", "1.0 nM（緩衝液）/ 0.14 \u00B5M（ヒト血漿中）"],
    ["選択性", "他のプロテアーゼに対し1000倍以上の選択性"],
    ["出血時間", "耳・歯肉・肝損傷モデルで延長なし"],
    ["抗血小板薬併用時", "ASA+チカグレロル併用でも出血時間の追加延長なし"],
  ];

  rows.forEach(function(row, i) {
    var yy = 1.3 + i * 0.75;
    addCard(s, 0.4, yy, 9.2, 0.65);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yy, w: 2.6, h: 0.65, fill: { color: C.primary } });
    s.addText(row[0], { x: 0.5, y: yy, w: 2.4, h: 0.65, fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
    s.addText(row[1], { x: 3.2, y: yy, w: 6.2, h: 0.65, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 8, 0, 8] });
  });

  // 比較ボックス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.15, w: 9.2, h: 0.55, fill: { color: "FFF3CD" } });
  s.addText([
    { text: "比較: ", options: { bold: true, color: C.orange, fontSize: 12 } },
    { text: "アピキサバン併用時は出血時間が2.8〜6.6倍延長 → XIa因子阻害薬の安全性上の優位性", options: { color: C.text, fontSize: 12 } },
  ], { x: 0.7, y: 5.15, w: 8.6, h: 0.55, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, "5/10");
})();

// ============================================================
// SLIDE 6: OCEANIC-STROKE ― 試験デザイン
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "OCEANIC-STROKE ― 試験デザイン");

  // 左列: 基本情報
  addCard(s, 0.3, 1.15, 5.2, 4.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.15, w: 5.2, h: 0.45, fill: { color: C.primary } });
  s.addText("試験概要", { x: 0.6, y: 1.15, w: 4.6, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var items = [
    { label: "デザイン", val: "第3相・RCT・二重盲検・プラセボ対照" },
    { label: "対象", val: "非心原性虚血性脳卒中（NIHSS\u226415）\nまたは高リスクTIA（ABCD2 6-7）" },
    { label: "発症〜登録", val: "72時間以内" },
    { label: "登録数", val: "12,237例（37か国・702施設）" },
    { label: "治療", val: "asundexian 50mg/日 vs プラセボ\n（抗血小板薬への上乗せ）" },
    { label: "観察期間", val: "中央値約2年（3〜31か月）" },
  ];
  items.forEach(function(item, i) {
    var yy = 1.75 + i * 0.6;
    s.addText([
      { text: item.label + ": ", options: { bold: true, color: C.primary, fontSize: 12 } },
      { text: item.val, options: { color: C.text, fontSize: 12 } },
    ], { x: 0.5, y: yy, w: 4.8, h: 0.55, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.1 });
  });

  // 右列: 患者背景・病型
  addCard(s, 5.7, 1.15, 4.0, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.7, y: 1.15, w: 4.0, h: 0.45, fill: { color: C.accent } });
  s.addText("患者背景", { x: 5.9, y: 1.15, w: 3.6, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("平均年齢: 68歳\n女性: 34%\n血栓溶解/回収療法: 27.4%\nNIHSS中央値: 2", {
    x: 5.9, y: 1.65, w: 3.6, h: 1.4, fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  addCard(s, 5.7, 3.35, 4.0, 2.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.7, y: 3.35, w: 4.0, h: 0.45, fill: { color: C.accent } });
  s.addText("病型内訳", { x: 5.9, y: 3.35, w: 3.6, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var subtypes = [
    { name: "大血管アテローム硬化症", pct: "41%" },
    { name: "小血管閉塞（ラクナ）", pct: "22%" },
    { name: "分類不能（ESUS相当）", pct: "29%" },
  ];
  subtypes.forEach(function(st, i) {
    var yy = 3.9 + i * 0.45;
    s.addText([
      { text: st.name + "  ", options: { color: C.text, fontSize: 12 } },
      { text: st.pct, options: { color: C.primary, fontSize: 14, bold: true } },
    ], { x: 5.9, y: yy, w: 3.6, h: 0.4, fontFace: FONT_JP, margin: 0 });
  });

  addPageNum(s, "6/10");
})();

// ============================================================
// SLIDE 7: OCEANIC-STROKE ― 主要結果
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "OCEANIC-STROKE ― 主要結果");

  // 有効性: 大きな数字カード
  addCard(s, 0.3, 1.15, 4.55, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.15, w: 4.55, h: 0.45, fill: { color: C.green } });
  s.addText("有効性: 虚血性脳卒中の再発", { x: 0.5, y: 1.15, w: 4.15, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "Asundexian  ", options: { color: C.text, fontSize: 14 } },
    { text: "6.2%", options: { color: C.green, fontSize: 32, bold: true } },
    { text: "  vs  ", options: { color: C.sub, fontSize: 14 } },
    { text: "プラセボ  ", options: { color: C.text, fontSize: 14 } },
    { text: "8.4%", options: { color: C.red, fontSize: 32, bold: true } },
  ], { x: 0.4, y: 1.7, w: 4.35, h: 0.9, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  s.addText([
    { text: "HR 0.74", options: { color: C.primary, fontSize: 18, bold: true } },
    { text: "（95%CI 0.65-0.84、P<0.001）", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.4, y: 2.65, w: 4.35, h: 0.35, fontFace: FONT_EN, align: "center", margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.1, w: 3.55, h: 0.45, fill: { color: C.green } });
  s.addText("26%のリスク減少", { x: 0.8, y: 3.1, w: 3.55, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // 安全性カード
  addCard(s, 5.15, 1.15, 4.55, 2.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.15, w: 4.55, h: 0.45, fill: { color: C.primary } });
  s.addText("安全性: ISTH基準の大出血", { x: 5.35, y: 1.15, w: 4.15, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "Asundexian  ", options: { color: C.text, fontSize: 14 } },
    { text: "1.9%", options: { color: C.primary, fontSize: 32, bold: true } },
    { text: "  vs  ", options: { color: C.sub, fontSize: 14 } },
    { text: "プラセボ  ", options: { color: C.text, fontSize: 14 } },
    { text: "1.7%", options: { color: C.primary, fontSize: 32, bold: true } },
  ], { x: 5.25, y: 1.7, w: 4.35, h: 0.9, fontFace: FONT_JP, align: "center", valign: "middle", margin: 0 });

  s.addText([
    { text: "HR 1.10", options: { color: C.primary, fontSize: 18, bold: true } },
    { text: "（90%CI 0.85-1.44、P=0.46）", options: { color: C.text, fontSize: 13 } },
  ], { x: 5.25, y: 2.65, w: 4.35, h: 0.35, fontFace: FONT_EN, align: "center", margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.65, y: 3.1, w: 3.55, h: 0.45, fill: { color: C.primary } });
  s.addText("大出血の増加なし", { x: 5.65, y: 3.1, w: 3.55, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  // 副次評価項目テーブル
  addCard(s, 0.3, 3.85, 9.4, 1.7);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.85, w: 9.4, h: 0.4, fill: { color: C.dark } });
  s.addText("副次評価項目", { x: 0.5, y: 3.85, w: 3, h: 0.4, fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var secondary = [
    ["全脳卒中", "HR 0.74 (0.65-0.84)", "P<0.001"],
    ["障害残存/致死的脳卒中", "HR 0.69 (0.55-0.87)", ""],
    ["CV死+MI+脳卒中", "HR 0.83 (0.74-0.92)", ""],
  ];
  secondary.forEach(function(row, i) {
    var yy = 4.3 + i * 0.4;
    s.addText(row[0], { x: 0.5, y: yy, w: 3.5, h: 0.38, fontSize: 11, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
    s.addText(row[1], { x: 4.0, y: yy, w: 3.5, h: 0.38, fontSize: 12, fontFace: FONT_EN, color: C.green, bold: true, valign: "middle", margin: 0 });
    s.addText(row[2], { x: 7.5, y: yy, w: 2, h: 0.38, fontSize: 11, fontFace: FONT_EN, color: C.sub, valign: "middle", margin: 0 });
  });

  addPageNum(s, "7/10");
})();

// ============================================================
// SLIDE 8: 病型別サブグループ・ESUS ― 47%リスク減少
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "病型別サブグループ解析 ― ESUSへの希望");

  s.addText("Shoamanesh A, et al. ISC 2026 (事前規定の副次解析)", { x: 0.5, y: 0.92, w: 7, h: 0.3, fontSize: 10, fontFace: FONT_EN, color: C.sub, italic: true, margin: 0 });

  // 3つの病型カード
  var subtypes = [
    { name: "大血管\nアテローム硬化症", n: "約5,000例", result: "有効\n（全体と一貫）", color: C.green, bgColor: "D4EDDA" },
    { name: "小血管閉塞\n（ラクナ）", n: "約2,600例", result: "有効\n（全体と一貫）", color: C.green, bgColor: "D4EDDA" },
    { name: "分類不能\n（ESUS相当）", n: "約3,500例", result: "47%\nリスク減少", color: C.primary, bgColor: C.light },
  ];
  subtypes.forEach(function(st, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.4, 3.0, 2.3);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.4, w: 3.0, h: 0.75, fill: { color: C.primary } });
    s.addText(st.name, { x: xPos + 0.1, y: 1.4, w: 2.8, h: 0.75, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.n, { x: xPos + 0.1, y: 2.2, w: 2.8, h: 0.35, fontSize: 11, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0 });

    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.3, y: 2.65, w: 2.4, h: 0.85, fill: { color: st.bgColor } });
    s.addText(st.result, { x: xPos + 0.3, y: 2.65, w: 2.4, h: 0.85, fontSize: i === 2 ? 18 : 14, fontFace: FONT_JP, color: st.color, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // 注釈
  s.addText("交互作用P値 = 0.07（探索的解析）", { x: 6.3, y: 3.55, w: 3.5, h: 0.3, fontSize: 10, fontFace: FONT_EN, color: C.sub, italic: true, margin: 0 });

  // ESUS解説ボックス
  addCard(s, 0.3, 4.05, 9.4, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.05, w: 9.4, h: 0.45, fill: { color: C.orange } });
  s.addText("ESUSでの意義", { x: 0.6, y: 4.05, w: 8.8, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "ESUS（塞栓源不明の脳塞栓症）は全脳梗塞の約17%。", options: { color: C.text, fontSize: 12 } },
    { text: "\nDOAC（リバーロキサバン・ダビガトラン）は2つの大規模RCTでESUSに無効 + 出血増加。", options: { color: C.red, fontSize: 12 } },
    { text: "\nAsundexianはESUS相当群で全体を上回る効果（47% vs 26%）を示した初めての薬剤。", options: { color: C.green, fontSize: 12, bold: true } },
  ], { x: 0.6, y: 4.55, w: 8.8, h: 0.9, fontFace: FONT_JP, margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "8/10");
})();

// ============================================================
// SLIDE 9: DOACとの比較 ― なぜDOACは失敗したのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "DOACとの比較 ― なぜ従来薬は失敗したのか");

  // 3カラム比較
  var trials = [
    { name: "NAVIGATE ESUS", drug: "リバーロキサバン 15mg", eff: "アスピリンと差なし\n（5.1% vs 4.8%/年）", saf: "大出血 約3倍増加\n頭蓋内出血 4倍増加", color: C.red, icon: "X" },
    { name: "RE-SPECT ESUS", drug: "ダビガトラン 150mg", eff: "アスピリンと差なし\n（4.1% vs 4.8%/年）", saf: "大出血は差なし\n非大出血が増加", color: C.orange, icon: "X" },
    { name: "OCEANIC-STROKE", drug: "asundexian 50mg", eff: "ESUS相当群で\n47%リスク減少", saf: "大出血の増加なし\n（HR 1.10, NS）", color: C.green, icon: "O" },
  ];

  trials.forEach(function(t, i) {
    var xPos = 0.2 + i * 3.27;
    addCard(s, xPos, 1.15, 3.07, 4.35);

    // タイトル
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.15, w: 3.07, h: 0.7, fill: { color: t.color } });
    s.addText(t.name, { x: xPos + 0.1, y: 1.15, w: 2.87, h: 0.4, fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(t.drug, { x: xPos + 0.1, y: 1.55, w: 2.87, h: 0.3, fontSize: 11, fontFace: FONT_JP, color: C.white, align: "center", valign: "middle", margin: 0 });

    // 有効性
    s.addText("有効性", { x: xPos + 0.1, y: 2.0, w: 2.87, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
    s.addText(t.eff, { x: xPos + 0.1, y: 2.35, w: 2.87, h: 0.85, fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2 });

    // 安全性
    s.addText("安全性", { x: xPos + 0.1, y: 3.3, w: 2.87, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0 });
    s.addText(t.saf, { x: xPos + 0.1, y: 3.65, w: 2.87, h: 0.85, fontSize: 12, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.2 });

    // 判定
    s.addShape(pres.shapes.OVAL, { x: xPos + 1.13, y: 4.6, w: 0.8, h: 0.8, fill: { color: t.color } });
    s.addText(t.icon, { x: xPos + 1.13, y: 4.6, w: 0.8, h: 0.8, fontSize: 24, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  addPageNum(s, "9/10");
})();

// ============================================================
// SLIDE 10: Take Home Message
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "Asundexianは非心原性脳梗塞の再発を26%減少させ、大出血は増加しなかった", highlight: "有効性と安全性の両立" },
    { num: "2", text: "内因系のXIa因子のみを阻害し、外因系（生理的止血）を温存する新しい機序", highlight: "選択的な凝固阻害" },
    { num: "3", text: "アテローム血栓性・ラクナ梗塞・ESUS相当の全病型で一貫した効果", highlight: "病型横断的な有効性" },
    { num: "4", text: "DOACが失敗したESUS領域で初めて有望な結果を示した薬剤", highlight: "ESUSへの新たな選択肢" },
    { num: "5", text: "論文未発表・最初の90日間は有意差なし・心房細動は対象外", highlight: "今後の課題" },
  ];

  messages.forEach(function(m, i) {
    var yy = 1.05 + i * 0.9;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yy, w: 9, h: 0.8, fill: { color: "243B5A" } });
    s.addShape(pres.shapes.OVAL, { x: 0.65, y: yy + 0.13, w: 0.55, h: 0.55, fill: { color: C.accent } });
    s.addText(m.num, { x: 0.65, y: yy + 0.13, w: 0.55, h: 0.55, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

    s.addText(m.highlight, { x: 1.4, y: yy + 0.05, w: 3.2, h: 0.3, fontSize: 12, fontFace: FONT_JP, color: C.accent, bold: true, margin: 0 });
    s.addText(m.text, { x: 1.4, y: yy + 0.35, w: 7.9, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.light, margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Sharma M, et al. ISC 2026; Abstract LB009", {
    x: 0.5, y: 5.15, w: 9, h: 0.35,
    fontSize: 10, fontFace: FONT_EN, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SAVE
// ============================================================
var outPath = __dirname + "/OCEANIC-STROKE_XIa因子阻害薬_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX saved: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
