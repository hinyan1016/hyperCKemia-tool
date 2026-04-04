var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "ストレスで本当に体は動かなくなるのか ― 機能性神経障害（FND）の真実";

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

  s.addText("脳神経内科医が答える からだの不思議 #05", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("ストレスで本当に\n体は動かなくなるのか", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.8,
    fontSize: 38, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 機能性神経障害（FND）の真実 ―", {
    x: 0.8, y: 3.1, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.85, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.1, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.65, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: こんな経験・こんな話を聞いたことはありませんか？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんな話を聞いたことはありませんか？");
  addPageNum(s, "2");

  var quotes = [
    { text: "「急に足が動かなくなった。でもMRIでは\n  何も映っていなかった」", color: C.primary },
    { text: "「大事な場面で突然声が出なくなった。\n  精神的なものと言われた」", color: C.primary },
    { text: "「手が震えて止まらない。\n  検査では異常がないと言われた」", color: C.primary },
  ];

  quotes.forEach(function(q, i) {
    var yPos = 1.1 + i * 1.35;
    addCard(s, 0.8, yPos, 8.4, 1.1);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.18, h: 1.1, fill: { color: q.color } });
    s.addText(q.text, {
      x: 1.2, y: yPos + 0.08, w: 7.8, h: 0.95,
      fontSize: 18, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  s.addText("これ、「気のせい」でも「仮病」でもありません", {
    x: 0.8, y: 5.05, w: 8.4, h: 0.4,
    fontSize: 16, fontFace: FONT_JP, color: C.red, bold: true, align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 3: 機能性神経障害（FND）とは何か
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "機能性神経障害（FND）とは");
  addPageNum(s, "3");

  // 定義カード
  addCard(s, 0.4, 1.0, 9.2, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.0, w: 9.2, h: 0.08, fill: { color: C.primary } });
  s.addText("定義", {
    x: 0.5, y: 1.08, w: 1.2, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("神経系の「構造」には明らかな損傷がないにもかかわらず、\n神経学的な症状（麻痺・感覚障害・けいれんなど）が出現する疾患", {
    x: 0.6, y: 1.5, w: 8.8, h: 0.75,
    fontSize: 17, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // コンピュータ比喩カード
  addCard(s, 0.4, 2.5, 9.2, 1.5);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.5, w: 0.12, h: 1.5, fill: { color: C.orange } });
  s.addText("💡 コンピュータに例えると", {
    x: 0.7, y: 2.55, w: 8.7, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });
  s.addText("「ハードウェア（神経・脳の構造）」は正常", {
    x: 0.8, y: 3.0, w: 8.4, h: 0.38,
    fontSize: 17, fontFace: FONT_JP, color: C.green, bold: true, margin: 0,
  });
  s.addText("「ソフトウェア（脳の処理プログラム）」に問題が起きている", {
    x: 0.8, y: 3.42, w: 8.4, h: 0.38,
    fontSize: 17, fontFace: FONT_JP, color: C.red, bold: true, margin: 0,
  });

  // 旧名称
  addCard(s, 0.4, 4.15, 9.2, 0.75);
  s.addText("かつての呼び名：「転換性障害」「ヒステリー」「心身症」→ 現在は FND が正式名称", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.6,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 4: FNDの頻度
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "FNDはどれくらい多いのか");
  addPageNum(s, "4");

  // メイン統計
  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 1.1, w: 7.0, h: 1.6, fill: { color: C.primary }, shadow: makeShadow() });
  s.addText("脳神経内科外来で", {
    x: 1.5, y: 1.2, w: 7.0, h: 0.55,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("2番目に多い疾患", {
    x: 1.5, y: 1.72, w: 7.0, h: 0.8,
    fontSize: 32, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", margin: 0,
  });

  s.addText("（1位：頭痛）", {
    x: 1.5, y: 2.75, w: 7.0, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  // 比較カード
  var items = [
    { label: "多発性硬化症", flag: "より多い", color: C.accent },
    { label: "パーキンソン病", flag: "より多い", color: C.accent },
    { label: "年間新規発症", flag: "10万人あたり4〜12人", color: C.green },
  ];

  items.forEach(function(item, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 3.2, 3.0, 1.5);
    s.addText(item.label, {
      x: xPos + 0.1, y: 3.3, w: 2.8, h: 0.55,
      fontSize: 15, fontFace: FONT_JP, color: C.text, bold: true, align: "center", margin: 0,
    });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.6, y: 3.88, w: 1.8, h: 0.03, fill: { color: C.accent } });
    s.addText(item.flag, {
      x: xPos + 0.1, y: 3.95, w: 2.8, h: 0.55,
      fontSize: 13, fontFace: FONT_JP, color: item.color, bold: true, align: "center", margin: 0,
    });
  });

  s.addText("「まれな病気」ではなく、外来でよく出会う神経疾患", {
    x: 0.5, y: 4.85, w: 9.0, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 5: 脳の中で何が起きているか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳の中で何が起きているか ― 最新の神経科学");
  addPageNum(s, "5");

  // メインメッセージ
  addCard(s, 0.4, 1.0, 9.2, 0.85);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.0, w: 9.2, h: 0.08, fill: { color: C.accent } });
  s.addText("fMRI・PET研究で、脳の特定回路に機能的変化が確認されている", {
    x: 0.6, y: 1.1, w: 8.8, h: 0.65,
    fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });

  // 2カラム
  addCard(s, 0.4, 2.0, 4.4, 2.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.0, w: 4.4, h: 0.45, fill: { color: C.primary } });
  s.addText("関与する脳領域", {
    x: 0.5, y: 2.02, w: 4.2, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  var regions = ["前補足運動野（運動の準備）", "島皮質（身体感覚の統合）", "扁桃体（情動・ストレス応答）"];
  regions.forEach(function(r, i) {
    s.addText("• " + r, {
      x: 0.6, y: 2.55 + i * 0.5, w: 4.0, h: 0.45,
      fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  addCard(s, 5.2, 2.0, 4.4, 2.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.0, w: 4.4, h: 0.45, fill: { color: C.orange } });
  s.addText("引き金となる要因", {
    x: 5.3, y: 2.02, w: 4.2, h: 0.4,
    fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });
  var triggers = ["強いストレス・トラウマ", "身体的イベント（けが・感染症）", "不安・注意の過剰集中"];
  triggers.forEach(function(t, i) {
    s.addText("• " + t, {
      x: 5.4, y: 2.55 + i * 0.5, w: 4.0, h: 0.45,
      fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  // ポイント
  addCard(s, 0.4, 4.25, 9.2, 0.75);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.25, w: 0.1, h: 0.75, fill: { color: C.yellow } });
  s.addText("注意の向け方が症状を維持させる ― 治療では「注意の分散」が重要", {
    x: 0.65, y: 4.3, w: 8.8, h: 0.6,
    fontSize: 15, fontFace: FONT_JP, color: C.text, bold: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 6: 「心の病気」ではない
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "「心の病気」ではない、という意味");
  addPageNum(s, "6");

  var points = [
    {
      num: "1",
      title: "心理的背景が「必ず」あるわけではない",
      body: "すべてのFND患者に明確な心理的ストレスが特定できるわけではありません。\n大規模研究でも、トラウマや精神疾患が必須ではないことが示されています。",
      color: C.primary,
    },
    {
      num: "2",
      title: "「作っている」「弱い」ではない",
      body: "「心の病気」という言葉は患者さんに誤解を与えます。\nFNDは意図して症状を作り出しているわけではありません。",
      color: C.orange,
    },
  ];

  points.forEach(function(p, i) {
    var yPos = 1.1 + i * 2.0;
    addCard(s, 0.4, yPos, 9.2, 1.8);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 0.65, h: 1.8, fill: { color: p.color } });
    s.addText(p.num, {
      x: 0.4, y: yPos + 0.55, w: 0.65, h: 0.65,
      fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
    });
    s.addText(p.title, {
      x: 1.2, y: yPos + 0.15, w: 8.2, h: 0.55,
      fontSize: 18, fontFace: FONT_JP, color: p.color, bold: true, margin: 0,
    });
    s.addText(p.body, {
      x: 1.2, y: yPos + 0.75, w: 8.2, h: 0.9,
      fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });

  // 結論バー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.1, w: 9.2, h: 0.45, fill: { color: C.dark } });
  s.addText("FNDは脳のソフトウェアに生じた、実在する神経疾患", {
    x: 0.5, y: 5.12, w: 9.0, h: 0.4,
    fontSize: 17, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", margin: 0,
  });
})();

// ============================================================
// SLIDE 7: 症状の種類
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "FNDの症状サブタイプ");
  addPageNum(s, "7");

  var types = [
    {
      icon: "🦵",
      title: "機能性運動障害",
      body: "脱力・麻痺・振戦（ふるえ）・異常姿勢\n注意を向けると悪化、分散すると軽減",
      color: C.primary,
    },
    {
      icon: "⚡",
      title: "解離性発作（PNES）",
      body: "てんかん様発作だが脳波は正常\nてんかん外来受診患者の15〜30%",
      color: C.red,
    },
    {
      icon: "✋",
      title: "機能性感覚障害",
      body: "感覚消失・しびれ・過敏\n正中線で左右に分かれるパターンが特徴",
      color: C.orange,
    },
    {
      icon: "🗣️",
      title: "機能性言語障害",
      body: "機能性失声・機能性発語障害\n声が出ない・ことばが出にくい",
      color: C.green,
    },
  ];

  types.forEach(function(t, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var xPos = 0.3 + col * 4.85;
    var yPos = 1.1 + row * 2.05;

    addCard(s, xPos, yPos, 4.55, 1.85);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 4.55, h: 0.5, fill: { color: t.color } });
    s.addText(t.icon + " " + t.title, {
      x: xPos + 0.1, y: yPos + 0.05, w: 4.35, h: 0.4,
      fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    });
    s.addText(t.body, {
      x: xPos + 0.15, y: yPos + 0.6, w: 4.25, h: 1.1,
      fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });
})();

// ============================================================
// SLIDE 8: FND・脳卒中・心因反応の違い
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "FND・脳卒中・心因反応の違い");
  addPageNum(s, "8");

  // 比較テーブル
  var headers = ["比較項目", "FND", "脳卒中・器質性", "一般的な心因反応"];
  var colWidths = [2.2, 2.6, 2.6, 2.6];
  var colXs = [0.0, 2.2, 4.8, 7.4];

  // ヘッダー行
  headers.forEach(function(h, i) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: colXs[i], y: 1.0, w: colWidths[i], h: 0.5, fill: { color: C.primary }
    });
    s.addText(h, {
      x: colXs[i] + 0.05, y: 1.0, w: colWidths[i] - 0.05, h: 0.5,
      fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  var rows = [
    ["神経症状", "麻痺・感覚障害・発作\n（明確あり）", "麻痺・言語障害\n（明確あり）", "神経症状は通常なし"],
    ["画像・検査", "正常または症状と\n不一致", "明確な\n異常所見あり", "正常"],
    ["診察所見", "Hoover徴候など\nFND特有サイン", "上位・下位MN\n障害パターン", "神経学的異常なし"],
    ["心理的背景", "必ずしも\n必要でない", "通常は\n直接関係しない", "ストレスが\n直接原因"],
  ];

  rows.forEach(function(row, ri) {
    var bgColor = ri % 2 === 0 ? C.warmBg : C.white;
    row.forEach(function(cell, ci) {
      var textColor = C.text;
      if (ci === 1 && ri === 0) textColor = C.primary;
      if (ci === 2 && ri === 1) textColor = C.red;
      s.addShape(pres.shapes.RECTANGLE, {
        x: colXs[ci], y: 1.5 + ri * 0.92, w: colWidths[ci], h: 0.92, fill: { color: bgColor }
      });
      s.addText(cell, {
        x: colXs[ci] + 0.05, y: 1.5 + ri * 0.92, w: colWidths[ci] - 0.05, h: 0.92,
        fontSize: 12, fontFace: FONT_JP, color: textColor, align: "center", valign: "middle", margin: 0,
      });
    });
  });

  // Hoover解説
  addCard(s, 0.3, 5.2, 9.4, 0.38);
  s.addText("Hoover徴候：健側を屈曲すると麻痺側に力が入る ― FNDのポジティブサインで診断に使用", {
    x: 0.4, y: 5.22, w: 9.2, h: 0.35,
    fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 9: 治療と回復
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "治療と回復");
  addPageNum(s, "9");

  // 最重要ポイント
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.0, w: 9.2, h: 0.75, fill: { color: C.light } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.0, w: 0.1, h: 0.75, fill: { color: C.primary } });
  s.addText("最初かつ最重要：正確な診断と説明", {
    x: 0.65, y: 1.08, w: 8.8, h: 0.58,
    fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  var treatments = [
    {
      num: "1",
      title: "神経リハビリテーション",
      body: "FND特化の理学療法・作業療法\n注意の分散を使った動作課題が特に有効",
      color: C.primary,
    },
    {
      num: "2",
      title: "認知行動療法（CBT）",
      body: "症状を維持させる思考パターン・行動を変える\n無作為化比較試験（RCT）で有効性確認",
      color: C.accent,
    },
    {
      num: "3",
      title: "患者教育・自己管理",
      body: "neurosymptoms.org など信頼できる情報源\n「脳のソフトウェアの問題」と理解することが第一歩",
      color: C.green,
    },
  ];

  treatments.forEach(function(t, i) {
    var yPos = 1.95 + i * 1.05;
    addCard(s, 0.4, yPos, 9.2, 0.9);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 0.7, h: 0.9, fill: { color: t.color } });
    s.addText(t.num, {
      x: 0.4, y: yPos + 0.12, w: 0.7, h: 0.65,
      fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
    });
    s.addText(t.title, {
      x: 1.25, y: yPos + 0.07, w: 3.5, h: 0.4,
      fontSize: 16, fontFace: FONT_JP, color: t.color, bold: true, margin: 0,
    });
    s.addText(t.body, {
      x: 1.25, y: yPos + 0.48, w: 8.1, h: 0.38,
      fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });
})();

// ============================================================
// SLIDE 10: 予後と早期診断の重要性
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "予後と早期診断の重要性");
  addPageNum(s, "10");

  // 予後カード
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.05, w: 9.2, h: 1.3, fill: { color: C.primary }, shadow: makeShadow() });
  s.addText("適切な診断・治療を受ければ", {
    x: 0.5, y: 1.12, w: 9.0, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("症状が改善・安定する患者さんは相当数", {
    x: 0.5, y: 1.6, w: 9.0, h: 0.6,
    fontSize: 24, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", margin: 0,
  });

  // 警告カード
  addCard(s, 0.4, 2.55, 9.2, 0.85);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.55, w: 0.1, h: 0.85, fill: { color: C.orange } });
  s.addText("⚠ 診断が遅れると予後が悪化。PNESでは診断まで平均7年以上かかることも", {
    x: 0.65, y: 2.62, w: 8.8, h: 0.65,
    fontSize: 15, fontFace: FONT_JP, color: C.orange, bold: true, margin: 0,
  });

  // 緊急受診
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.55, w: 9.2, h: 0.42, fill: { color: C.red } });
  s.addText("急いで受診が必要なケース", {
    x: 0.5, y: 3.58, w: 9.0, h: 0.36,
    fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
  });

  var emergencies = [
    "突然の片側の脱力・感覚障害（脳卒中との鑑別が必要）",
    "意識を失うエピソードが繰り返す（てんかん・PNESの鑑別が必要）",
    "症状が急速に進行している",
  ];
  emergencies.forEach(function(e, i) {
    s.addText("• " + e, {
      x: 0.7, y: 4.05 + i * 0.38, w: 9.0, h: 0.35,
      fontSize: 14, fontFace: FONT_JP, color: C.text, margin: 0,
    });
  });
})();

// ============================================================
// SLIDE 11: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("まとめ　Take Home Message", {
    x: 0.5, y: 0.2, w: 9, h: 0.55,
    fontSize: 22, fontFace: FONT_JP, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { text: "FNDは「仮病」でも「心の病気」でもなく、\n脳のソフトウェアに生じた実在する神経疾患", color: C.yellow },
    { text: "脳神経内科外来で2番目に多い ― 決して「まれ」ではない", color: C.accent },
    { text: "Hoover徴候などのポジティブサインで診断する\n（消去法ではない）", color: C.green },
    { text: "早期の正確な診断と治療（神経リハビリ・CBT）で\n改善の可能性がある", color: C.orange },
  ];

  messages.forEach(function(m, i) {
    var yPos = 0.95 + i * 1.05;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 0.5, h: 0.85, fill: { color: m.color } });
    s.addText(String(i + 1), {
      x: 0.5, y: yPos + 0.1, w: 0.5, h: 0.65,
      fontSize: 22, fontFace: FONT_EN, color: C.dark, bold: true, align: "center", margin: 0,
    });
    s.addText(m.text, {
      x: 1.15, y: yPos + 0.05, w: 8.5, h: 0.78,
      fontSize: 16, fontFace: FONT_JP, color: C.white, margin: 0, lineSpacingMultiple: 1.2,
    });
  });

  s.addText("医知創造ラボ / チャンネル登録・コメントお待ちしています", {
    x: 0.5, y: 5.15, w: 9, h: 0.38,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });
})();

// ============================================================
// 出力
// ============================================================
var outPath = __dirname + "/機能性神経障害_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
