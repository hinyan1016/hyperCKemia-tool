var pptxgen = require("pptxgenjs");
var pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "rt-PA静注療法の適応と禁忌 ― 「時・血・画・出・偽」で覚える";

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

function shd() {
  return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 };
}

function hdr(s, title, bgColor) {
  bgColor = bgColor || C.primary;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: bgColor } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, fontFace: FJ, color: C.white, bold: true, margin: 0 });
}

// ─────────────────────────────────────────────
// Slide 1: Title slide
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  // Dark background
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark } });
  // Orange accent bar top
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.18, fill: { color: C.accent } });
  // Orange accent bar bottom
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.45, w: 10, h: 0.18, fill: { color: C.accent } });

  // Center decoration: horizontal line with glow
  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 2.85, w: 7, h: 0.04, fill: { color: C.accent }, line: { color: C.accent } });

  // Main title
  s.addText("脳梗塞のrt-PA静注療法", {
    x: 0.5, y: 0.8, w: 9, h: 1.1,
    fontSize: 42, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });
  // Subtitle
  s.addText("適応と禁忌を「時・血・画・出・偽」で覚える", {
    x: 0.5, y: 1.85, w: 9, h: 0.75,
    fontSize: 26, fontFace: FJ, color: C.accent, bold: false,
    align: "center", valign: "middle"
  });

  // 5 keywords strip
  var keywords = ["時", "血", "画", "出", "偽"];
  var kwColors = [C.primary, C.red, C.green, "E8850C", "7B2D8E"];
  for (var i = 0; i < keywords.length; i++) {
    var kx = 1.0 + i * 1.6;
    s.addShape(pres.shapes.RECTANGLE, {
      x: kx, y: 3.0, w: 1.3, h: 1.1,
      fill: { color: kwColors[i] }, shadow: shd(),
      line: { color: C.white, width: 1.5 }
    });
    s.addText(keywords[i], {
      x: kx, y: 3.0, w: 1.3, h: 1.1,
      fontSize: 40, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
  }

  // Author
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
    { icon: "1", text: "rt-PA（アルテプラーゼ）の適応の「本質」を1文で理解できる" },
    { icon: "2", text: "「時・血・画・出・偽」5つのキーワードで禁忌を体系的に覚えられる" },
    { icon: "3", text: "DWI/FLAIRミスマッチなど2025年最新の適応拡大ポイントを押さえられる" },
    { icon: "4", text: "当直で使える30秒チェックフローを身につけられる" },
    { icon: "5", text: "血栓回収療法との使い分け方針を理解できる" },
  ];

  for (var i = 0; i < bullets.length; i++) {
    var by = 1.1 + i * 0.82;
    // Numbered circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.45, y: by + 0.08, w: 0.52, h: 0.52,
      fill: { color: C.primary }, shadow: shd()
    });
    s.addText(bullets[i].icon, {
      x: 0.45, y: by + 0.08, w: 0.52, h: 0.52,
      fontSize: 20, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.1, y: by, w: 8.5, h: 0.65,
      fill: { color: C.white }, shadow: shd(),
      line: { color: "D8E6F3", width: 1 }
    });
    s.addText(bullets[i].text, {
      x: 1.25, y: by, w: 8.2, h: 0.65,
      fontSize: 20, fontFace: FJ, color: C.text, bold: false,
      valign: "middle"
    });
  }
})();

// ─────────────────────────────────────────────
// Slide 3: なぜ「丸暗記」ではうまくいかないのか
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "なぜ「丸暗記」ではうまくいかないのか", C.red);

  // Problem card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 9.2, h: 1.25,
    fill: { color: C.lightRed }, shadow: shd(),
    line: { color: C.red, width: 2 }
  });
  s.addText("問題", {
    x: 0.55, y: 1.12, w: 0.9, h: 0.4,
    fontSize: 15, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle",
    fill: { color: C.red }
  });
  s.addText("適応と禁忌を「同じ重み」で並列暗記しようとすると…項目が多すぎて破綻する", {
    x: 1.6, y: 1.1, w: 7.8, h: 1.1,
    fontSize: 21, fontFace: FJ, color: C.text,
    valign: "middle"
  });

  // Arrow
  s.addShape(pres.shapes.CHEVRON, {
    x: 4.6, y: 2.42, w: 0.8, h: 0.55,
    fill: { color: C.accent }, rotate: 90
  });

  // Solution card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.1, w: 9.2, h: 1.1,
    fill: { color: C.lightGreen }, shadow: shd(),
    line: { color: C.green, width: 2 }
  });
  s.addText("解決策", {
    x: 0.55, y: 3.18, w: 0.9, h: 0.38,
    fontSize: 15, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle",
    fill: { color: C.green }
  });
  s.addText("「臨床判断の順番」で覚える ― 禁忌を順にくぐり抜けた症例が適応になる", {
    x: 1.6, y: 3.1, w: 7.8, h: 1.1,
    fontSize: 21, fontFace: FJ, color: C.text,
    valign: "middle"
  });

  // Key phrase
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.5, y: 4.35, w: 7, h: 0.85,
    fill: { color: C.accent }, shadow: shd()
  });
  s.addText("「適応は、禁忌をくぐり抜けた結果」", {
    x: 1.5, y: 4.35, w: 7, h: 0.85,
    fontSize: 24, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 4: 1文で覚える適応の本質
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark } });
  // Top accent
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.18, fill: { color: C.accent } });

  s.addText("1文で覚える適応の本質", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 26, fontFace: FJ, color: "AABBD0", bold: false,
    align: "center", valign: "middle"
  });

  // Main message card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.1, w: 9, h: 2.8,
    fill: { color: C.primary }, shadow: shd(),
    line: { color: C.accent, width: 3 }
  });

  s.addText([
    { text: "4.5時間以内", options: { color: C.yellow, bold: true, fontSize: 36 } },
    { text: "で、", options: { color: C.white, bold: false, fontSize: 32 } },
    { text: "まだ助かる脳があり", options: { color: C.lightGreen || "D4EDDA", bold: true, fontSize: 36 } },
    { text: "、\n", options: { color: C.white, bold: false, fontSize: 32 } },
    { text: "出血させすぎない", options: { color: "F8A5B0", bold: true, fontSize: 36 } },
    { text: "症例", options: { color: C.white, bold: false, fontSize: 32 } },
  ], {
    x: 0.8, y: 1.2, w: 8.4, h: 2.6,
    fontFace: FJ,
    align: "center", valign: "middle",
    paraSpaceAfter: 8
  });

  // Three pillars below
  var pillars = [
    { label: "時間", color: C.yellow, desc: "4.5時間以内" },
    { label: "画像", color: C.green, desc: "助かる脳が残っている" },
    { label: "出血", color: C.red, desc: "止血能が保たれている" },
  ];
  for (var i = 0; i < pillars.length; i++) {
    var px = 0.6 + i * 3.0;
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: 4.1, w: 2.7, h: 1.1,
      fill: { color: "223A5E" }, shadow: shd(),
      line: { color: pillars[i].color, width: 2 }
    });
    s.addText(pillars[i].label, {
      x: px, y: 4.1, w: 2.7, h: 0.45,
      fontSize: 20, fontFace: FJ, color: pillars[i].color, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(pillars[i].desc, {
      x: px, y: 4.5, w: 2.7, h: 0.65,
      fontSize: 16, fontFace: FJ, color: C.white,
      align: "center", valign: "middle"
    });
  }
})();

// ─────────────────────────────────────────────
// Slide 5: 判断フレーム「時・血・画・出・偽」overview
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "判断フレーム「時・血・画・出・偽」");

  var cards = [
    { char: "時", color: C.primary,  meaning: "時間内か", detail: "4.5時間以内（LKW基準）" },
    { char: "血", color: C.red,      meaning: "数値異常", detail: "BP / 血糖 / 血小板 / INR" },
    { char: "画", color: C.green,    meaning: "画像禁忌", detail: "出血・広範梗塞・圧排" },
    { char: "出", color: "E8850C",   meaning: "出血歴", detail: "3か月 / 21日 / 14日" },
    { char: "偽", color: "7B2D8E",   meaning: "Mimic除外", detail: "痙攣・低血糖・SAH" },
  ];

  for (var i = 0; i < cards.length; i++) {
    var cx = 0.35 + i * 1.88;
    // Shadow base
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.05, w: 1.65, h: 3.9,
      fill: { color: cards[i].color }, shadow: shd()
    });
    // Big kanji
    s.addText(cards[i].char, {
      x: cx, y: 1.1, w: 1.65, h: 1.5,
      fontSize: 56, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    // Divider
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx + 0.15, y: 2.55, w: 1.35, h: 0.04,
      fill: { color: C.white }
    });
    // Meaning
    s.addText(cards[i].meaning, {
      x: cx, y: 2.65, w: 1.65, h: 0.75,
      fontSize: 18, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    // Detail box at bottom
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 3.5, w: 1.65, h: 1.3,
      fill: { color: "FFFFFF", transparency: 20 }
    });
    s.addText(cards[i].detail, {
      x: cx + 0.05, y: 3.52, w: 1.55, h: 1.26,
      fontSize: 15, fontFace: FJ, color: C.text, bold: false,
      align: "center", valign: "middle"
    });
  }

  // Flow arrow label
  s.addText("この順番で禁忌を確認し、すべてクリアした症例が適応", {
    x: 0.3, y: 5.0, w: 9.4, h: 0.45,
    fontSize: 18, fontFace: FJ, color: C.primary, bold: true,
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 6: 「時」― 時間の判断
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "「時」― 時間の判断", C.primary);

  // Big character
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.05, w: 1.5, h: 1.5,
    fill: { color: C.primary }, shadow: shd()
  });
  s.addText("時", {
    x: 0.3, y: 1.05, w: 1.5, h: 1.5,
    fontSize: 64, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });

  // Main rule
  s.addShape(pres.shapes.RECTANGLE, {
    x: 2.0, y: 1.05, w: 7.6, h: 1.5,
    fill: { color: C.light }, shadow: shd(),
    line: { color: C.primary, width: 2 }
  });
  s.addText("発症（Last Known Well）から", {
    x: 2.2, y: 1.1, w: 7.2, h: 0.55,
    fontSize: 20, fontFace: FJ, color: C.text, bold: false,
    valign: "middle"
  });
  s.addText("4.5時間以内", {
    x: 2.2, y: 1.55, w: 7.2, h: 0.75,
    fontSize: 34, fontFace: FJ, color: C.primary, bold: true,
    valign: "middle"
  });

  // Exception
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 2.75, w: 9.4, h: 1.3,
    fill: { color: C.lightYellow }, shadow: shd(),
    line: { color: C.yellow, width: 2 }
  });
  s.addText("例外：発症時刻不明 / 起床時発症", {
    x: 0.5, y: 2.8, w: 9.0, h: 0.45,
    fontSize: 20, fontFace: FJ, color: C.text, bold: true,
    valign: "middle"
  });
  s.addText("DWI陽性 ＋ FLAIR陰性（ミスマッチ）→ 4.5時間以内の可能性あり → 投与を検討", {
    x: 0.5, y: 3.15, w: 9.0, h: 0.75,
    fontSize: 19, fontFace: FJ, color: C.text,
    valign: "middle"
  });

  // Mnemonic
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.5, y: 4.2, w: 7.0, h: 0.95,
    fill: { color: C.primary }, shadow: shd()
  });
  s.addText("覚え方：「時計は1本、例外はMRI」", {
    x: 1.5, y: 4.2, w: 7.0, h: 0.95,
    fontSize: 22, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 7: DWI/FLAIRミスマッチとは
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "DWI / FLAIRミスマッチとは", C.primary);

  // Two MRI boxes side by side
  var panels = [
    { label: "DWI（拡散強調）", state: "高信号（＋）", color: C.red, icon: "発症後すぐ異常" },
    { label: "FLAIR", state: "低信号 / 正常（－）", color: C.green, icon: "発症4.5h以内の特徴" },
  ];
  for (var i = 0; i < panels.length; i++) {
    var px = 0.4 + i * 4.8;
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: 1.05, w: 4.3, h: 2.2,
      fill: { color: C.white }, shadow: shd(),
      line: { color: panels[i].color, width: 2.5 }
    });
    s.addText(panels[i].label, {
      x: px, y: 1.1, w: 4.3, h: 0.55,
      fontSize: 20, fontFace: FJ, color: panels[i].color, bold: true,
      align: "center", valign: "middle"
    });
    // Simulated MRI box
    s.addShape(pres.shapes.RECTANGLE, {
      x: px + 0.3, y: 1.65, w: 3.7, h: 1.0,
      fill: { color: "2D3436" }
    });
    s.addText(panels[i].state, {
      x: px + 0.3, y: 1.65, w: 3.7, h: 1.0,
      fontSize: 22, fontFace: FJ, color: i === 0 ? C.yellow : C.green, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(panels[i].icon, {
      x: px, y: 2.7, w: 4.3, h: 0.45,
      fontSize: 17, fontFace: FJ, color: C.text,
      align: "center", valign: "middle"
    });
  }

  // Plus sign between panels
  s.addText("+", {
    x: 4.55, y: 1.6, w: 0.9, h: 0.9,
    fontSize: 48, fontFace: FE, color: C.sub, bold: true,
    align: "center", valign: "middle"
  });

  // Arrow down
  s.addText("↓", {
    x: 4.4, y: 3.2, w: 1.2, h: 0.5,
    fontSize: 36, fontFace: FE, color: C.accent, bold: true,
    align: "center", valign: "middle"
  });

  // Result box (green)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 3.75, w: 9.2, h: 1.25,
    fill: { color: C.lightGreen }, shadow: shd(),
    line: { color: C.green, width: 2 }
  });
  s.addText("まだ発症4.5時間以内の可能性が高い → rt-PA投与を検討", {
    x: 0.6, y: 3.8, w: 8.8, h: 0.55,
    fontSize: 22, fontFace: FJ, color: C.text, bold: true,
    valign: "middle"
  });
  s.addText("根拠：WAKE-UP trial（NEJM 2018）― DWI/FLAIRミスマッチ例への投与で転帰改善", {
    x: 0.6, y: 4.3, w: 8.8, h: 0.55,
    fontSize: 17, fontFace: FJ, color: C.text,
    valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 8: 「血」― 数字で覚える
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "「血」― 数字で覚える禁忌閾値", C.red);

  // Big character
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.05, w: 0.95, h: 4.2,
    fill: { color: C.red }, shadow: shd()
  });
  s.addText("血", {
    x: 0.3, y: 1.05, w: 0.95, h: 4.2,
    fontSize: 44, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });

  var rows = [
    { label: "血圧",    threshold: "185/110 mmHg",  unit: "以上",  note: "投与前にコントロール" },
    { label: "血糖",    threshold: "＜50 / ＞400",  unit: "mg/dL", note: "代謝異常・mimicのリスク" },
    { label: "血小板",  threshold: "10万 /μL",      unit: "未満",  note: "止血能不十分" },
    { label: "PT-INR",  threshold: "1.7",           unit: "超",    note: "抗凝固薬使用・肝不全" },
    { label: "aPTT",    threshold: "延長",          unit: "（基準上限超）", note: "ヘパリン使用中" },
  ];

  for (var i = 0; i < rows.length; i++) {
    var ry = 1.1 + i * 0.85;
    // Row background
    var rowBg = i % 2 === 0 ? C.white : C.light;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.4, y: ry, w: 8.25, h: 0.78,
      fill: { color: rowBg }, shadow: shd(),
      line: { color: "D8E6F3", width: 1 }
    });
    // Label
    s.addText(rows[i].label, {
      x: 1.5, y: ry, w: 1.4, h: 0.78,
      fontSize: 20, fontFace: FJ, color: C.text, bold: true,
      valign: "middle"
    });
    // Threshold (red)
    s.addText(rows[i].threshold, {
      x: 2.9, y: ry, w: 2.8, h: 0.78,
      fontSize: 24, fontFace: FJ, color: C.red, bold: true,
      valign: "middle"
    });
    // Unit
    s.addText(rows[i].unit, {
      x: 5.65, y: ry, w: 1.3, h: 0.78,
      fontSize: 18, fontFace: FJ, color: C.text,
      valign: "middle"
    });
    // Note
    s.addText(rows[i].note, {
      x: 6.95, y: ry, w: 2.6, h: 0.78,
      fontSize: 15, fontFace: FJ, color: C.sub,
      valign: "middle"
    });
  }
})();

// ─────────────────────────────────────────────
// Slide 9: 数字の意味づけ
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "「血」― 数字の意味を理解して覚える", C.red);

  var groups = [
    {
      nums: "185 / 110",
      label: "血圧の閾値",
      meaning: "高血圧 → 再灌流後の出血リスクが激増",
      icon: "血が怖い",
      color: C.red
    },
    {
      nums: "＜50 / ＞400",
      label: "血糖の閾値",
      meaning: "低血糖は神経症候を呈す（Stroke mimic）\n高血糖は梗塞後の転帰悪化",
      icon: "本当に脳梗塞？",
      color: "7B2D8E"
    },
    {
      nums: "10万 / INR1.7",
      label: "凝固系の閾値",
      meaning: "止血能が不十分 → 出血を止められない",
      icon: "血が怖い",
      color: C.red
    },
  ];

  for (var i = 0; i < groups.length; i++) {
    var gx = 0.25 + i * 3.22;
    s.addShape(pres.shapes.RECTANGLE, {
      x: gx, y: 1.05, w: 3.0, h: 4.1,
      fill: { color: C.white }, shadow: shd(),
      line: { color: groups[i].color, width: 2.5 }
    });
    // Number
    s.addShape(pres.shapes.RECTANGLE, {
      x: gx, y: 1.05, w: 3.0, h: 0.9,
      fill: { color: groups[i].color }
    });
    s.addText(groups[i].nums, {
      x: gx, y: 1.05, w: 3.0, h: 0.9,
      fontSize: 24, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(groups[i].label, {
      x: gx + 0.1, y: 2.0, w: 2.8, h: 0.5,
      fontSize: 18, fontFace: FJ, color: C.text, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(groups[i].meaning, {
      x: gx + 0.1, y: 2.5, w: 2.8, h: 1.6,
      fontSize: 16, fontFace: FJ, color: C.text,
      align: "center", valign: "middle"
    });
    // Icon badge
    s.addShape(pres.shapes.RECTANGLE, {
      x: gx + 0.2, y: 4.3, w: 2.6, h: 0.6,
      fill: { color: groups[i].color }
    });
    s.addText(groups[i].icon, {
      x: gx + 0.2, y: 4.3, w: 2.6, h: 0.6,
      fontSize: 17, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
  }
})();

// ─────────────────────────────────────────────
// Slide 10: 「画」― 画像禁忌
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "「画」― 画像禁忌 ＝「血・広い・ずれてる」", C.green);

  // Big character
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.05, w: 1.2, h: 1.2,
    fill: { color: C.green }, shadow: shd()
  });
  s.addText("画", {
    x: 0.3, y: 1.05, w: 1.2, h: 1.2,
    fontSize: 50, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });

  var cards = [
    { title: "血", detail: "出血がある", sub: "頭蓋内出血・くも膜下出血", color: C.red },
    { title: "広い", detail: "広範な早期虚血性変化", sub: "ASPECTS低値・1/3以上の梗塞", color: C.accent },
    { title: "ずれてる", detail: "圧排所見・正中線偏位", sub: "脳浮腫による mass effect", color: "7B2D8E" },
  ];

  for (var i = 0; i < cards.length; i++) {
    var cx = 1.7 + i * 2.7;
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.05, w: 2.45, h: 2.4,
      fill: { color: C.white }, shadow: shd(),
      line: { color: cards[i].color, width: 2.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.05, w: 2.45, h: 0.7,
      fill: { color: cards[i].color }
    });
    s.addText(cards[i].title, {
      x: cx, y: 1.05, w: 2.45, h: 0.7,
      fontSize: 26, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(cards[i].detail, {
      x: cx + 0.1, y: 1.8, w: 2.25, h: 0.65,
      fontSize: 19, fontFace: FJ, color: C.text, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(cards[i].sub, {
      x: cx + 0.1, y: 2.45, w: 2.25, h: 0.85,
      fontSize: 15, fontFace: FJ, color: C.text,
      align: "center", valign: "middle"
    });
  }

  // SAH note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 3.65, w: 9.4, h: 0.8,
    fill: { color: C.lightRed }, shadow: shd(),
    line: { color: C.red, width: 1.5 }
  });
  s.addText("注意：くも膜下出血が疑われる場合も禁忌 — 特に突然発症の頭痛 + 神経症状", {
    x: 0.5, y: 3.65, w: 9.1, h: 0.8,
    fontSize: 18, fontFace: FJ, color: C.text,
    valign: "middle"
  });

  // Key message
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 4.6, w: 9.4, h: 0.75,
    fill: { color: C.green }, shadow: shd()
  });
  s.addText("すでに壊れすぎた脳への投与 → 出血転化リスクが転帰改善を上回る", {
    x: 0.5, y: 4.6, w: 9.0, h: 0.75,
    fontSize: 19, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 11: 「出」― 既往歴は時期で覚える
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "「出」― 既往歴は「時期」で覚える", C.accent);

  // Big character
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.05, w: 1.0, h: 3.95,
    fill: { color: C.accent }, shadow: shd()
  });
  s.addText("出", {
    x: 0.3, y: 1.05, w: 1.0, h: 3.95,
    fontSize: 44, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });

  var timeCards = [
    {
      period: "3か月以内",
      items: ["・脳梗塞既往", "・頭蓋内 / 脊髄手術", "・頭部外傷（重症）"],
      color: C.red
    },
    {
      period: "21日以内",
      items: ["・消化管出血", "・尿路出血（肉眼的血尿など）"],
      color: "E8850C"
    },
    {
      period: "14日以内",
      items: ["・大手術", "・重篤な外傷"],
      color: C.yellow
    },
  ];

  for (var i = 0; i < timeCards.length; i++) {
    var ty = 1.1 + i * 1.25;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.45, y: ty, w: 8.2, h: 1.1,
      fill: { color: C.white }, shadow: shd(),
      line: { color: timeCards[i].color, width: 2 }
    });
    // Period badge
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.45, y: ty, w: 1.9, h: 1.1,
      fill: { color: timeCards[i].color }
    });
    s.addText(timeCards[i].period, {
      x: 1.45, y: ty, w: 1.9, h: 1.1,
      fontSize: 22, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(timeCards[i].items.join("  "), {
      x: 3.45, y: ty, w: 6.1, h: 1.1,
      fontSize: 18, fontFace: FJ, color: C.text,
      valign: "middle"
    });
  }

  // Absolute contraindications box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.45, y: 4.6, w: 8.2, h: 0.8,
    fill: { color: C.lightRed }, shadow: shd(),
    line: { color: C.red, width: 2 }
  });
  s.addText("時期に関わらず禁忌：頭蓋内出血既往 / AVM / 動脈瘤（破裂リスク）/ 頭蓋内腫瘍", {
    x: 1.6, y: 4.6, w: 7.9, h: 0.8,
    fontSize: 17, fontFace: FJ, color: C.text,
    valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 12: 「出」の覚え方 ― タイムライン
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.18, fill: { color: C.accent } });

  s.addText("「出」の覚え方 ― タイムラインで整理", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.7,
    fontSize: 26, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });

  // Mnemonic banner
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.05, w: 8.4, h: 1.1,
    fill: { color: C.accent }, shadow: shd()
  });
  s.addText("「脳は3か月、出血は21日、手術は14日」", {
    x: 0.8, y: 1.05, w: 8.4, h: 1.1,
    fontSize: 30, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });

  // Timeline bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 2.5, w: 9.0, h: 0.12,
    fill: { color: C.white }
  });
  // Markers
  var markers = [
    { x: 0.5, label: "現在", color: C.green },
    { x: 2.8, label: "14日前\n大手術", color: C.yellow },
    { x: 5.3, label: "21日前\n出血", color: C.accent },
    { x: 9.5, label: "3か月前\n脳梗塞\n頭蓋内手術", color: C.red },
  ];
  for (var i = 0; i < markers.length; i++) {
    var mx = markers[i].x;
    s.addShape(pres.shapes.OVAL, {
      x: mx - 0.18, y: 2.35, w: 0.36, h: 0.36,
      fill: { color: markers[i].color }
    });
    s.addText(markers[i].label, {
      x: mx - 0.7, y: 2.8, w: 1.4, h: 1.2,
      fontSize: 17, fontFace: FJ, color: markers[i].color, bold: true,
      align: "center", valign: "top"
    });
  }

  // Arrow (past direction)
  s.addText("← 過去", {
    x: 0.5, y: 2.55, w: 2.0, h: 0.4,
    fontSize: 15, fontFace: FJ, color: "AABBD0",
    valign: "middle"
  });

  // Key insight
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 4.55, w: 8.4, h: 0.8,
    fill: { color: "223A5E" }, shadow: shd(),
    line: { color: C.accent, width: 1.5 }
  });
  s.addText("この期間内のイベントがあれば→投与禁忌（脳・出血・手術の順で期間が長い）", {
    x: 1.0, y: 4.55, w: 8.0, h: 0.8,
    fontSize: 18, fontFace: FJ, color: C.white,
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 13: 「偽」― それ、本当に脳梗塞？
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "「偽」― それ、本当に脳梗塞ですか？", "7B2D8E");

  // Big character
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.05, w: 1.2, h: 1.2,
    fill: { color: "7B2D8E" }, shadow: shd()
  });
  s.addText("偽", {
    x: 0.3, y: 1.05, w: 1.2, h: 1.2,
    fontSize: 50, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });

  var mimics = [
    {
      icon: "痙攣発作",
      risk: "Todd麻痺",
      detail: "発作後の一過性麻痺 → 脳梗塞と見分けがつきにくい",
      action: "既往・目撃者情報を必ず確認",
      color: "7B2D8E"
    },
    {
      icon: "低血糖",
      risk: "血糖＜50 mg/dL",
      detail: "神経症候を呈することがある\n補正後に症状消失するかを確認",
      action: "血糖補正後に再評価",
      color: C.accent
    },
    {
      icon: "くも膜下出血",
      risk: "SAH疑い",
      detail: "突然発症の頭痛 ＋ 神経症状\n→ 病態が全く異なる",
      action: "rt-PAは絶対禁忌",
      color: C.red
    },
  ];

  for (var i = 0; i < mimics.length; i++) {
    var my = 2.45 + i * 1.02;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: my, w: 9.4, h: 0.9,
      fill: { color: C.white }, shadow: shd(),
      line: { color: mimics[i].color, width: 2 }
    });
    // Badge
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: my, w: 2.0, h: 0.9,
      fill: { color: mimics[i].color }
    });
    s.addText(mimics[i].icon + "\n" + mimics[i].risk, {
      x: 0.3, y: my, w: 2.0, h: 0.9,
      fontSize: 15, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(mimics[i].detail, {
      x: 2.4, y: my, w: 4.7, h: 0.9,
      fontSize: 17, fontFace: FJ, color: C.text,
      valign: "middle"
    });
    // Action
    s.addShape(pres.shapes.RECTANGLE, {
      x: 7.1, y: my + 0.1, w: 2.45, h: 0.7,
      fill: { color: mimics[i].color }
    });
    s.addText(mimics[i].action, {
      x: 7.1, y: my + 0.1, w: 2.45, h: 0.7,
      fontSize: 14, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
  }

  // Warning
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 5.15, w: 9.4, h: 0.3,
    fill: { color: "7B2D8E" }
  });
  s.addText("Stroke mimicへのrt-PA投与 → 不必要な出血リスク", {
    x: 0.5, y: 5.15, w: 9.0, h: 0.3,
    fontSize: 15, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 14: 当直で使える30秒チェックフロー
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "当直で使える 30秒チェックフロー", C.dark);

  var steps = [
    { n: "1", text: "LKW（最終確認良好時刻）は？　→ 4.5時間以内か", cat: "時" },
    { n: "2", text: "CT/MRIで出血なし？　広範な梗塞・圧排なし？", cat: "画" },
    { n: "3", text: "くも膜下出血が疑われない？", cat: "偽" },
    { n: "4", text: "BP 185/110 mmHg 未満にコントロールできるか？", cat: "血" },
    { n: "5", text: "血糖 50〜400 mg/dL の範囲内？", cat: "血" },
    { n: "6", text: "血小板 ＞10万/μL　かつ　PT-INR ≤ 1.7？", cat: "血" },
    { n: "7", text: "3か月以内に脳梗塞・頭蓋内手術・重症頭部外傷なし？", cat: "出" },
    { n: "8", text: "21日以内のGI/尿路出血なし？　14日以内の大手術なし？", cat: "出" },
    { n: "9", text: "痙攣・低血糖・SAHなどのStroke mimic除外済み？", cat: "偽" },
    { n: "→", text: "すべてクリア → 開始！インフォームドコンセントと体重確認", cat: "GO" },
  ];

  var catColors = { "時": C.primary, "画": C.green, "血": C.red, "出": C.accent, "偽": "7B2D8E", "GO": C.green };

  for (var i = 0; i < steps.length; i++) {
    var col = i < 5 ? 0 : 1;
    var row = i < 5 ? i : i - 5;
    var sx = 0.2 + col * 5.0;
    var sy = 1.05 + row * 0.9;
    var cc = catColors[steps[i].cat];

    s.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: sy, w: 4.7, h: 0.78,
      fill: { color: steps[i].n === "→" ? C.lightGreen : C.white }, shadow: shd(),
      line: { color: cc, width: steps[i].n === "→" ? 3 : 1.5 }
    });
    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: sx + 0.08, y: sy + 0.13, w: 0.48, h: 0.48,
      fill: { color: cc }
    });
    s.addText(steps[i].n, {
      x: sx + 0.08, y: sy + 0.13, w: 0.48, h: 0.48,
      fontSize: 17, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(steps[i].text, {
      x: sx + 0.65, y: sy + 0.02, w: 3.95, h: 0.74,
      fontSize: 14, fontFace: FJ, color: C.text,
      valign: "middle"
    });
  }
})();

// ─────────────────────────────────────────────
// Slide 15: 血栓回収療法との関係
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "血栓回収療法との関係（2025年改訂ポイント）", C.dark);

  // Two-column comparison
  var panels = [
    {
      title: "rt-PA + 血栓回収",
      sub: "ブリッジング療法",
      color: C.primary,
      points: [
        "大血管閉塞に対して",
        "rt-PAを先に投与しながら",
        "血栓回収（IVT→EVT）へ",
        "発症〜2時間20分以内が目安",
      ]
    },
    {
      title: "直接血栓回収",
      sub: "EVT単独",
      color: C.accent,
      points: [
        "大血管閉塞 ＋ rt-PA禁忌の場合",
        "発症2時間20分以降の大血管閉塞",
        "IVT追加のベネフィットが限られる",
        "施設の体制・症例に応じて判断",
      ]
    },
  ];

  for (var i = 0; i < panels.length; i++) {
    var px = 0.3 + i * 5.0;
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: 1.05, w: 4.55, h: 3.8,
      fill: { color: C.white }, shadow: shd(),
      line: { color: panels[i].color, width: 2.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: px, y: 1.05, w: 4.55, h: 0.95,
      fill: { color: panels[i].color }
    });
    s.addText(panels[i].title, {
      x: px, y: 1.1, w: 4.55, h: 0.5,
      fontSize: 22, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(panels[i].sub, {
      x: px, y: 1.6, w: 4.55, h: 0.35,
      fontSize: 16, fontFace: FJ, color: C.white,
      align: "center", valign: "middle"
    });
    for (var j = 0; j < panels[i].points.length; j++) {
      s.addText("・" + panels[i].points[j], {
        x: px + 0.25, y: 2.1 + j * 0.6, w: 4.1, h: 0.55,
        fontSize: 17, fontFace: FJ, color: C.text,
        valign: "middle"
      });
    }
  }

  // 2025 update note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 5.0, w: 9.4, h: 0.45,
    fill: { color: C.light },
    line: { color: C.primary, width: 1 }
  });
  s.addText("2025改訂：「2時間20分」の目安が明記 ― より早い再灌流戦略の選択が推奨", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.45,
    fontSize: 16, fontFace: FJ, color: C.primary, bold: false,
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 16: まとめ
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.18, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.45, w: 10, h: 0.18, fill: { color: C.accent } });

  s.addText("まとめ", {
    x: 0.5, y: 0.2, w: 9, h: 0.7,
    fontSize: 30, fontFace: FJ, color: C.white, bold: true,
    align: "center", valign: "middle"
  });

  // 1文
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.05, w: 9.2, h: 1.35,
    fill: { color: C.primary }, shadow: shd(),
    line: { color: C.accent, width: 2 }
  });
  s.addText("1文", {
    x: 0.55, y: 1.1, w: 0.7, h: 0.4,
    fontSize: 14, fontFace: FJ, color: C.accent, bold: true,
    align: "center", valign: "middle"
  });
  s.addText("「4.5時間以内で、まだ助かる脳があり、出血させすぎない症例」", {
    x: 1.3, y: 1.1, w: 8.1, h: 1.2,
    fontSize: 22, fontFace: FJ, color: C.white, bold: true,
    valign: "middle"
  });

  // 5 keywords
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 2.55, w: 9.2, h: 0.6,
    fill: { color: "223A5E" }
  });
  s.addText("5つのキーワード", {
    x: 0.6, y: 2.55, w: 9.0, h: 0.6,
    fontSize: 18, fontFace: FJ, color: C.accent, bold: true,
    align: "center", valign: "middle"
  });

  var kwords = [
    { char: "時", color: C.primary },
    { char: "血", color: C.red },
    { char: "画", color: C.green },
    { char: "出", color: C.accent },
    { char: "偽", color: "7B2D8E" },
  ];
  for (var i = 0; i < kwords.length; i++) {
    var kx = 0.6 + i * 1.75;
    s.addShape(pres.shapes.RECTANGLE, {
      x: kx, y: 3.25, w: 1.5, h: 1.1,
      fill: { color: kwords[i].color }, shadow: shd()
    });
    s.addText(kwords[i].char, {
      x: kx, y: 3.25, w: 1.5, h: 1.1,
      fontSize: 44, fontFace: FJ, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
  }

  // Key numbers
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 4.5, w: 9.2, h: 0.75,
    fill: { color: "223A5E" }, shadow: shd()
  });
  s.addText("数字：4.5時間 ｜ 185/110 ｜ 50〜400 ｜ 10万 ｜ INR1.7 ｜ 3か月 / 21日 / 14日", {
    x: 0.6, y: 4.5, w: 8.8, h: 0.75,
    fontSize: 17, fontFace: FJ, color: C.accent, bold: true,
    align: "center", valign: "middle"
  });
})();

// ─────────────────────────────────────────────
// Slide 17: 参考文献
// ─────────────────────────────────────────────
(function() {
  var s = pres.addSlide();
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.warmBg } });
  hdr(s, "参考文献", C.dark);

  var refs = [
    { num: "1", text: "日本脳卒中学会 脳卒中治療ガイドライン2025（改訂版）" },
    { num: "2", text: "Yamaguchi T, et al. Randomized double-blind study of tissue-type plasminogen activator (t-PA). Stroke 2006; 37:2873–2879. [J-ACT]" },
    { num: "3", text: "Hacke W, et al. Thrombolysis with alteplase 3 to 4.5 hours after acute ischemic stroke. N Engl J Med 2008; 359:1317–1329. [ECASS III]" },
    { num: "4", text: "Thomalla G, et al. MRI-guided thrombolysis for stroke with unknown time of onset. N Engl J Med 2018; 379:611–622. [WAKE-UP]" },
    { num: "5", text: "アルテプラーゼ静注用0.6mg/kg 添付文書（最新版）" },
    { num: "6", text: "Powers WJ, et al. 2019 AHA/ASA Guideline for the Early Management of Patients With Acute Ischemic Stroke. Stroke 2019; 50:e344–e418." },
  ];

  for (var i = 0; i < refs.length; i++) {
    var ry = 1.1 + i * 0.73;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: ry, w: 9.4, h: 0.65,
      fill: { color: C.white }, shadow: shd(),
      line: { color: "D8E6F3", width: 1 }
    });
    s.addShape(pres.shapes.OVAL, {
      x: 0.38, y: ry + 0.09, w: 0.44, h: 0.44,
      fill: { color: C.primary }
    });
    s.addText(refs[i].num, {
      x: 0.38, y: ry + 0.09, w: 0.44, h: 0.44,
      fontSize: 15, fontFace: FE, color: C.white, bold: true,
      align: "center", valign: "middle"
    });
    s.addText(refs[i].text, {
      x: 0.92, y: ry, w: 8.65, h: 0.65,
      fontSize: 14, fontFace: FJ, color: C.text,
      valign: "middle"
    });
  }
})();

// ─────────────────────────────────────────────
// Output
// ─────────────────────────────────────────────
var outputPath = "youtube-slides/rt-PA血栓溶解療法/rt-PA血栓溶解療法.pptx";
pres.writeFile({ fileName: outputPath }).then(function() {
  console.log("PPTX created: " + outputPath);
}).catch(function(err) {
  console.error("Error:", err);
});
