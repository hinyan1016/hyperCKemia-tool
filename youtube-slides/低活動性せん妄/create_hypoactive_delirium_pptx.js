var pptxgen = require("pptxgenjs");
var pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "低活動性せん妄 ― 「静かなせん妄」を見逃さない";

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
  purple: "7B2D8E",
};

var FJ = "BIZ UDPGothic";
var FE = "Segoe UI";

function shd() {
  return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 };
}

function hdr(s, title, bgColor) {
  bgColor = bgColor || C.primary;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: bgColor } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, fontFace: FJ, color: C.white, bold: true, margin: 0 });
}

function pageNum(s, num) {
  s.addText(String(num), {
    x: 9.2, y: 5.15, w: 0.6, h: 0.35,
    fontSize: 11, fontFace: FE, color: C.sub, align: "right", valign: "bottom"
  });
}

// ─────────────────────────────────────────────
// Slide 1: タイトルスライド
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.18, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.45, w: 10, h: 0.18, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 2.85, w: 7, h: 0.04, fill: { color: C.accent } });

  s.addText("低活動性せん妄", {
    x: 0.5, y: 0.7, w: 9, h: 1.2,
    fontSize: 44, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });
  s.addText("「静かなせん妄」を見逃さない", {
    x: 0.5, y: 1.85, w: 9, h: 0.75,
    fontSize: 26, fontFace: FJ, color: C.accent, bold: false,
    align: "center", valign: "middle"
  });

  // HYPO-D cards
  var letters = ["H", "Y", "P", "O", "D"];
  var labels = ["過眠", "注意力", "受動性", "摂取低下", "昼夜逆転"];
  var cardColors = [C.primary, C.purple, C.red, C.accent, C.green];
  for (var i = 0; i < letters.length; i++) {
    var kx = 1.0 + i * 1.6;
    s.addShape(pres.shapes.RECTANGLE, {
      x: kx, y: 3.0, w: 1.3, h: 1.3,
      fill: { color: cardColors[i] }, shadow: shd(),
      line: { color: C.white, width: 1.5 }
    });
    s.addText(letters[i], {
      x: kx, y: 2.95, w: 1.3, h: 0.9,
      fontSize: 38, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(labels[i], {
      x: kx, y: 3.7, w: 1.3, h: 0.5,
      fontSize: 13, fontFace: FJ, color: C.white, bold: false,
      align: "center", valign: "middle"
    });
  }

  s.addText("今村久司", {
    x: 0.5, y: 4.8, w: 9, h: 0.4,
    fontSize: 18, fontFace: FJ, color: "AABBD0",
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 2: この動画で学べること
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "この動画で学べること");

  var bullets = [
    { icon: "1", text: "低活動性せん妄がなぜ見逃されるのか ― その構造的理由" },
    { icon: "2", text: "「HYPO-D」で覚える臨床的特徴" },
    { icon: "3", text: "うつ病・認知症・アパシーとの鑑別ポイント" },
    { icon: "4", text: "「DELIRIUM」で原因を系統的にチェック" },
    { icon: "5", text: "非薬物療法を中心とした実践的治療" },
  ];

  for (var i = 0; i < bullets.length; i++) {
    var by = 1.1 + i * 0.82;
    s.addShape(pres.shapes.OVAL, {
      x: 0.45, y: by + 0.08, w: 0.52, h: 0.52,
      fill: { color: C.primary }, shadow: shd()
    });
    s.addText(bullets[i].icon, {
      x: 0.45, y: by + 0.08, w: 0.52, h: 0.52,
      fontSize: 20, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.1, y: by, w: 8.5, h: 0.65,
      fill: { color: C.white }, shadow: shd()
    });
    s.addText(bullets[i].text, {
      x: 1.25, y: by, w: 8.2, h: 0.65,
      fontSize: 19, fontFace: FJ, color: C.text, bold: false,
      valign: "middle"
    });
  }
  pageNum(s, 2);
})();

// ─────────────────────────────────────────────
// Slide 3: なぜ見逃されるのか
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "なぜ低活動性せん妄は見逃されるのか");

  // 3 reason cards
  var reasons = [
    { icon: "1", title: "「困る」行動がない", desc: "興奮もナースコールもない\n→ 「おとなしい患者」と評価" },
    { icon: "2", title: "他の病態と誤認", desc: "うつ病・認知症の進行\n疲労・鎮静薬の残効と解釈" },
    { icon: "3", title: "スクリーニング不足", desc: "一般病棟でCAMの定期使用\nが普及していない" },
  ];

  for (var i = 0; i < reasons.length; i++) {
    var cx = 0.35 + i * 3.2;
    // Card background
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.1, w: 2.95, h: 3.0,
      fill: { color: C.white }, shadow: shd()
    });
    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: cx + 1.05, y: 1.3, w: 0.7, h: 0.7,
      fill: { color: C.red }
    });
    s.addText(reasons[i].icon, {
      x: cx + 1.05, y: 1.3, w: 0.7, h: 0.7,
      fontSize: 24, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    // Title
    s.addText(reasons[i].title, {
      x: cx + 0.15, y: 2.15, w: 2.65, h: 0.55,
      fontSize: 18, fontFace: FJ, color: C.primary, bold: true,
      align: "center", valign: "middle"
    });
    // Description
    s.addText(reasons[i].desc, {
      x: cx + 0.15, y: 2.7, w: 2.65, h: 1.2,
      fontSize: 15, fontFace: FJ, color: C.text,
      align: "center", valign: "top", lineSpacingMultiple: 1.3
    });
  }

  // Bottom warning
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.3, w: 9.3, h: 0.75,
    fill: { color: C.lightRed }
  });
  s.addText("見逃しの代償: 入院期間延長 / 機能予後悪化 / 死亡率上昇", {
    x: 0.5, y: 4.3, w: 9.0, h: 0.75,
    fontSize: 18, fontFace: FJ, color: C.red, bold: true,
    align: "center", valign: "middle"
  });
  pageNum(s, 3);
})();

// ─────────────────────────────────────────────
// Slide 4: せん妄のサブタイプ（疫学）
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "せん妄のサブタイプ ― 低活動型が最多");

  var types = [
    { name: "過活動型", pct: "15-25%", desc: "興奮・攻撃性\n幻覚・不穏", color: C.red, w: 1.8 },
    { name: "低活動型", pct: "40-50%", desc: "無関心・傾眠\n動作緩慢・食欲低下", color: C.primary, w: 3.2 },
    { name: "混合型", pct: "30-40%", desc: "過活動と低活動が\n交互に出現", color: C.purple, w: 2.5 },
  ];

  var xpos = 0.5;
  for (var i = 0; i < types.length; i++) {
    var t = types[i];
    // Bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: xpos, y: 1.2, w: t.w, h: 1.8,
      fill: { color: t.color }, shadow: shd()
    });
    s.addText(t.name, {
      x: xpos, y: 1.25, w: t.w, h: 0.5,
      fontSize: 18, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(t.pct, {
      x: xpos, y: 1.7, w: t.w, h: 0.6,
      fontSize: 28, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(t.desc, {
      x: xpos, y: 2.3, w: t.w, h: 0.6,
      fontSize: 13, fontFace: FJ, color: C.white,
      align: "center", valign: "middle", lineSpacingMultiple: 1.2
    });
    xpos += t.w + 0.25;
  }

  // Highlight box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.3, w: 9.0, h: 1.8,
    fill: { color: C.white }, shadow: shd()
  });
  s.addText("低活動型は最も頻度が高いのに、最も認識率が低い", {
    x: 0.7, y: 3.4, w: 8.6, h: 0.55,
    fontSize: 20, fontFace: FJ, color: C.red, bold: true,
    align: "center", valign: "middle"
  });
  s.addText("高齢者・終末期・術後で特に多い\n過活動型の認識率 >> 低活動型の認識率\nDSM-5 (2013) でサブタイプ分類が正式導入", {
    x: 0.7, y: 3.95, w: 8.6, h: 1.0,
    fontSize: 16, fontFace: FJ, color: C.text,
    align: "center", valign: "middle", lineSpacingMultiple: 1.4
  });
  pageNum(s, 4);
})();

// ─────────────────────────────────────────────
// Slide 5: 臨床像 HYPO-D
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "臨床像 ―「HYPO-D」で覚える");

  var items = [
    { letter: "H", word: "Hypersomnia", ja: "過眠・傾眠", ex: "日中うとうと、覚醒が鈍い", col: C.primary },
    { letter: "Y", word: "Yielding attention", ja: "注意力低下", ex: "会話が続かない、質問に答えられない", col: C.purple },
    { letter: "P", word: "Passivity", ja: "受動性・無関心", ex: "自発的な動きが減少", col: C.red },
    { letter: "O", word: "Oral intake decline", ja: "経口摂取低下", ex: "食事量の急激な低下", col: C.accent },
    { letter: "D", word: "Day-night reversal", ja: "昼夜逆転", ex: "夜間不眠、日中傾眠", col: C.green },
  ];

  for (var i = 0; i < items.length; i++) {
    var iy = 1.05 + i * 0.85;
    // Letter circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.35, y: iy + 0.05, w: 0.65, h: 0.65,
      fill: { color: items[i].col }, shadow: shd()
    });
    s.addText(items[i].letter, {
      x: 0.35, y: iy + 0.05, w: 0.65, h: 0.65,
      fontSize: 26, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.15, y: iy, w: 8.5, h: 0.72,
      fill: { color: C.white }, shadow: shd()
    });
    s.addText(items[i].word, {
      x: 1.3, y: iy, w: 2.8, h: 0.36,
      fontSize: 14, fontFace: FE, color: items[i].col, bold: true,
      valign: "bottom"
    });
    s.addText(items[i].ja, {
      x: 1.3, y: iy + 0.34, w: 2.8, h: 0.34,
      fontSize: 13, fontFace: FJ, color: C.sub,
      valign: "top"
    });
    s.addText(items[i].ex, {
      x: 4.2, y: iy, w: 5.3, h: 0.72,
      fontSize: 16, fontFace: FJ, color: C.text,
      valign: "middle"
    });
  }

  // Key point
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 5.0, w: 9.3, h: 0.45,
    fill: { color: C.light }
  });
  s.addText("最重要: これらが「急性に」出現し「日内変動」があること → せん妄を疑う", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.45,
    fontSize: 15, fontFace: FJ, color: C.primary, bold: true,
    align: "center", valign: "middle"
  });
  pageNum(s, 5);
})();

// ─────────────────────────────────────────────
// Slide 6: 鑑別診断
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "鑑別診断 ― うつ病・認知症・アパシーとの違い");

  // Table header
  var cols = ["", "低活動性\nせん妄", "うつ病", "認知症", "アパシー"];
  var colW = [1.5, 2.0, 2.0, 2.2, 2.0];
  var xStart = 0.15;

  // Header row
  var hx = xStart;
  for (var c = 0; c < cols.length; c++) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: hx, y: 1.0, w: colW[c], h: 0.65,
      fill: { color: c === 0 ? C.dark : (c === 1 ? C.primary : "4A6FA5") }
    });
    s.addText(cols[c], {
      x: hx, y: 1.0, w: colW[c], h: 0.65,
      fontSize: 13, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle", lineSpacingMultiple: 1.1
    });
    hx += colW[c] + 0.03;
  }

  // Data rows
  var rows = [
    ["発症", "急性\n(時間〜日)", "亜急性\n(数週間)", "緩徐\n(月〜年)", "緩徐〜\n亜急性"],
    ["日内変動", "著明\n(夕〜夜増悪)", "朝方増悪", "軽度", "一定"],
    ["注意障害", "高度\n(中核症状)", "軽度", "初期は保持", "軽度"],
    ["意識", "変動あり", "清明", "清明", "清明"],
    ["可逆性", "原因除去で\n改善", "治療で\n改善", "不可逆が\n多い", "原因に\nよる"],
  ];

  for (var r = 0; r < rows.length; r++) {
    var ry = 1.68 + r * 0.72;
    var rx = xStart;
    for (var c2 = 0; c2 < rows[r].length; c2++) {
      var bgCol = c2 === 0 ? "E8EDF2" : (r % 2 === 0 ? C.white : C.warmBg);
      if (c2 === 1) bgCol = r % 2 === 0 ? C.light : "D4E6F9";
      s.addShape(pres.shapes.RECTANGLE, {
        x: rx, y: ry, w: colW[c2], h: 0.69,
        fill: { color: bgCol }, line: { color: "DEE2E6", width: 0.5 }
      });
      s.addText(rows[r][c2], {
        x: rx, y: ry, w: colW[c2], h: 0.69,
        fontSize: 12, fontFace: FJ, color: c2 === 0 ? C.dark : C.text,
        bold: c2 === 0, align: "center", valign: "middle", lineSpacingMultiple: 1.1
      });
      rx += colW[c2] + 0.03;
    }
  }

  // Pitfall box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 4.55, w: 9.3, h: 0.75,
    fill: { color: C.lightYellow }
  });
  s.addText("「入院してから急にぼけが進んだ」= 低活動性せん妄を強く疑う", {
    x: 0.5, y: 4.55, w: 9.0, h: 0.75,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true,
    align: "center", valign: "middle"
  });
  pageNum(s, 6);
})();

// ─────────────────────────────────────────────
// Slide 7: 原因検索 DELIRIUM
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "原因検索 ―「DELIRIUM」で系統チェック");

  var items = [
    { l: "D", en: "Drug", ja: "薬剤", ex: "抗コリン薬、ベンゾ、オピオイド、H2B", col: C.red },
    { l: "E", en: "Electrolyte", ja: "電解質", ex: "Na異常、Ca異常、低血糖、高血糖", col: C.accent },
    { l: "L", en: "Lack of sleep", ja: "睡眠障害", ex: "不眠、睡眠覚醒リズム障害", col: C.purple },
    { l: "I", en: "Infection", ja: "感染症", ex: "尿路感染、肺炎、胆管炎、敗血症", col: C.green },
    { l: "R", en: "Retention", ja: "排泄障害", ex: "尿閉、便秘", col: "5B8C5A" },
    { l: "I", en: "Intracranial", ja: "頭蓋内", ex: "脳卒中、硬膜下血腫、てんかん", col: C.primary },
    { l: "U", en: "Undernutrition", ja: "低栄養", ex: "脱水、VitB1欠乏、腎不全、肝不全", col: "8B6914" },
    { l: "M", en: "Myocardial", ja: "心肺", ex: "心不全、低酸素血症、呼吸不全", col: C.red },
  ];

  for (var i = 0; i < items.length; i++) {
    var iy = 1.0 + i * 0.55;
    // Letter box
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: iy, w: 0.5, h: 0.48,
      fill: { color: items[i].col }
    });
    s.addText(items[i].l, {
      x: 0.3, y: iy, w: 0.5, h: 0.48,
      fontSize: 22, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    // English + Japanese
    s.addText(items[i].en, {
      x: 0.9, y: iy, w: 1.7, h: 0.28,
      fontSize: 13, fontFace: FE, color: items[i].col, bold: true,
      valign: "bottom"
    });
    s.addText(items[i].ja, {
      x: 0.9, y: iy + 0.25, w: 1.7, h: 0.23,
      fontSize: 11, fontFace: FJ, color: C.sub,
      valign: "top"
    });
    // Examples
    s.addText(items[i].ex, {
      x: 2.7, y: iy, w: 7.0, h: 0.48,
      fontSize: 15, fontFace: FJ, color: C.text,
      valign: "middle"
    });
  }

  // Bottom emphasis
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 5.0, w: 9.4, h: 0.45,
    fill: { color: C.lightRed }
  });
  s.addText("最優先: 原因薬剤（抗コリン薬・ベンゾジアゼピン）の中止・変更", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.45,
    fontSize: 15, fontFace: FJ, color: C.red, bold: true,
    align: "center", valign: "middle"
  });
  pageNum(s, 7);
})();

// ─────────────────────────────────────────────
// Slide 8: 治療 ― 非薬物療法
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "治療 ― 非薬物療法が柱");

  var interventions = [
    { icon: "1", title: "不快刺激の除去", desc: "身体抑制の回避、不要なカテーテル抜去", col: C.red },
    { icon: "2", title: "見当識の維持", desc: "カレンダー・時計、入院目的の繰り返し説明", col: C.primary },
    { icon: "3", title: "感覚入力の最適化", desc: "眼鏡・補聴器・義歯の使用、疼痛管理", col: C.purple },
    { icon: "4", title: "早期離床", desc: "座位→立位→歩行の段階的リハビリ", col: C.green },
    { icon: "5", title: "睡眠衛生", desc: "夜間処置の最小化、日中覚醒の促進", col: C.accent },
    { icon: "6", title: "栄養・水分", desc: "脱水補正、栄養管理、口腔ケア", col: "5B8C5A" },
  ];

  for (var i = 0; i < interventions.length; i++) {
    var col = i < 3 ? 0 : 1;
    var row = i < 3 ? i : i - 3;
    var cx = 0.3 + col * 4.9;
    var cy = 1.1 + row * 1.35;

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: 4.7, h: 1.15,
      fill: { color: C.white }, shadow: shd()
    });
    // Number
    s.addShape(pres.shapes.OVAL, {
      x: cx + 0.15, y: cy + 0.22, w: 0.6, h: 0.6,
      fill: { color: interventions[i].col }
    });
    s.addText(interventions[i].icon, {
      x: cx + 0.15, y: cy + 0.22, w: 0.6, h: 0.6,
      fontSize: 22, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    // Title
    s.addText(interventions[i].title, {
      x: cx + 0.9, y: cy + 0.08, w: 3.6, h: 0.5,
      fontSize: 17, fontFace: FJ, color: interventions[i].col, bold: true,
      valign: "middle"
    });
    // Description
    s.addText(interventions[i].desc, {
      x: cx + 0.9, y: cy + 0.55, w: 3.6, h: 0.5,
      fontSize: 14, fontFace: FJ, color: C.text,
      valign: "middle"
    });
  }

  // Bottom note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 5.05, w: 9.4, h: 0.4,
    fill: { color: C.lightGreen }
  });
  s.addText("エビデンス: 多職種包括的介入でせん妄発症率を30〜40%低減（HELP試験）", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FJ, color: C.green, bold: true,
    align: "center", valign: "middle"
  });
  pageNum(s, 8);
})();

// ─────────────────────────────────────────────
// Slide 9: 薬物療法
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "薬物療法 ― トラゾドンが第一候補");

  // Trazodone highlight card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.1, w: 9.4, h: 1.8,
    fill: { color: C.white }, shadow: shd()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.1, w: 0.15, h: 1.8,
    fill: { color: C.primary }
  });
  s.addText("トラゾドン 25mg（眠前または夕食後）", {
    x: 0.7, y: 1.15, w: 8.8, h: 0.5,
    fontSize: 22, fontFace: FJ, color: C.primary, bold: true,
    valign: "middle"
  });

  var trBullets = [
    "5-HT2A受容体拮抗 → 睡眠覚醒リズム改善",
    "抗コリン作用が少ない → せん妄を悪化させにくい",
    "ベンゾジアゼピン系からの切り替え先として有用",
    "「不眠時指示をトラゾドンに変更するだけでも予防と治療に」",
  ];
  for (var i = 0; i < trBullets.length; i++) {
    s.addText("  " + trBullets[i], {
      x: 0.7, y: 1.65 + i * 0.3, w: 8.8, h: 0.3,
      fontSize: 14, fontFace: FJ, color: C.text,
      valign: "middle", bullet: { code: "2022" }
    });
  }

  // Other drugs
  var drugs = [
    { name: "ラメルテオン", pos: "予防エビデンスあり", note: "メラトニン受容体作動薬", col: C.green },
    { name: "スボレキサント", pos: "不眠対策", note: "オレキシン受容体拮抗薬", col: C.accent },
    { name: "抗精神病薬", pos: "低活動型への使用は限定的", note: "2024 APA GL: 効果限定的〜なし\n高齢者の死亡率上昇リスク", col: C.red },
  ];

  for (var j = 0; j < drugs.length; j++) {
    var dx = 0.3 + j * 3.2;
    s.addShape(pres.shapes.RECTANGLE, {
      x: dx, y: 3.15, w: 3.0, h: 1.65,
      fill: { color: C.white }, shadow: shd()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: dx, y: 3.15, w: 3.0, h: 0.08,
      fill: { color: drugs[j].col }
    });
    s.addText(drugs[j].name, {
      x: dx + 0.1, y: 3.28, w: 2.8, h: 0.4,
      fontSize: 16, fontFace: FJ, color: drugs[j].col, bold: true,
      valign: "middle"
    });
    s.addText(drugs[j].pos, {
      x: dx + 0.1, y: 3.65, w: 2.8, h: 0.35,
      fontSize: 13, fontFace: FJ, color: C.text,
      valign: "middle"
    });
    s.addText(drugs[j].note, {
      x: dx + 0.1, y: 4.0, w: 2.8, h: 0.7,
      fontSize: 11, fontFace: FJ, color: C.sub,
      valign: "top", lineSpacingMultiple: 1.2
    });
  }

  // Warning
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 5.0, w: 9.4, h: 0.45,
    fill: { color: C.lightYellow }
  });
  s.addText("原則: 低活動型に抗精神病薬は使わない。薬物より原因除去と非薬物療法を優先", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.45,
    fontSize: 14, fontFace: FJ, color: C.accent, bold: true,
    align: "center", valign: "middle"
  });
  pageNum(s, 9);
})();

// ─────────────────────────────────────────────
// Slide 10: まとめ
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.12, fill: { color: C.accent } });

  s.addText("まとめ ― 「静かだから大丈夫」ではない", {
    x: 0.5, y: 0.25, w: 9, h: 0.7,
    fontSize: 26, fontFace: FJ, color: C.accent, bold: true,
    align: "center", valign: "middle"
  });

  var summaryItems = [
    { num: "1", text: "低活動性せん妄は全体の約50% ― 最多かつ最も見逃されやすい" },
    { num: "2", text: "「HYPO-D」の急性変化を見たらせん妄を疑い、CAMでスクリーニング" },
    { num: "3", text: "鑑別のカギ: 発症速度・日内変動・注意障害の程度" },
    { num: "4", text: "「DELIRIUM」で原因を系統的に検索し、除去する" },
    { num: "5", text: "治療の柱は非薬物療法。薬物はトラゾドン 25mgが第一候補" },
    { num: "6", text: "多職種包括的介入で発症率を30〜40%低減できる" },
  ];

  for (var i = 0; i < summaryItems.length; i++) {
    var sy = 1.1 + i * 0.7;
    s.addShape(pres.shapes.OVAL, {
      x: 0.5, y: sy + 0.05, w: 0.5, h: 0.5,
      fill: { color: C.accent }
    });
    s.addText(summaryItems[i].num, {
      x: 0.5, y: sy + 0.05, w: 0.5, h: 0.5,
      fontSize: 18, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(summaryItems[i].text, {
      x: 1.2, y: sy, w: 8.3, h: 0.6,
      fontSize: 17, fontFace: FJ, color: C.white,
      valign: "middle"
    });
  }

  // Bottom
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.3, w: 10, h: 0.12, fill: { color: C.accent } });
  s.addText("ご視聴ありがとうございました", {
    x: 0.5, y: 5.05, w: 9, h: 0.35,
    fontSize: 14, fontFace: FJ, color: "AABBD0",
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 11: 参考文献
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "参考文献");

  var refs = [
    "1. APA. DSM-5. 2013.",
    "2. 日本サイコオンコロジー学会. せん妄ガイドライン 2022.",
    "3. Kiely DK, et al. J Am Geriatr Soc. 2009;57(1):55-61.",
    "4. Meagher DJ, et al. J Neuropsychiatry Clin Neurosci. 2008;20(2):185-93.",
    "5. Maldonado JR. Am J Geriatr Psychiatry. 2013;21(12):1190-222.",
    "6. 矢吹拓. 内科病棟診療の極意 第13回. 金芳堂.",
    "7. Inouye SK, et al. Ann Intern Med. 1990;113(12):941-8.",
    "8. Ely EW, et al. Crit Care Med. 2001;29(7):1370-9.",
    "9. APA. Practice Guideline for Delirium. Am J Psychiatry. 2024.",
    "10. NICE CG103. Delirium guideline. 2023 update.",
    "11. Inouye SK, et al. HELP. N Engl J Med. 1999;340(9):669-76.",
    "12. PADIS Guidelines 2018; 2025 focused update.",
  ];

  for (var i = 0; i < refs.length; i++) {
    s.addText(refs[i], {
      x: 0.5, y: 1.0 + i * 0.37, w: 9.0, h: 0.35,
      fontSize: 12, fontFace: FJ, color: C.text,
      valign: "middle"
    });
  }
  pageNum(s, 11);
})();

// ─────────────────────────────────────────────
// Export
// ─────────────────────────────────────────────
var outPath = __dirname + "/低活動性せん妄.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
