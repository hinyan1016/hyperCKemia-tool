// ============================================================
// 食事指導シリーズ — 共通テンプレート
// 各回の create_slides.js から require して使用する
// ============================================================
//
// 使い方:
//   var T = require("./../_shared/series_template.js");
//   var pres = T.createPres("第1回 タイトル");
//   var C = T.C;        // カラーパレット
//   var FJ = T.FJ;      // 日本語フォント
//   var FE = T.FE;      // 英語フォント
//   T.addTitleSlide(pres, 1, "メインタイトル", "サブタイトル");
//   // ... 本編スライド ...
//   T.addSummarySlide(pres, ["ポイント1", "ポイント2", "ポイント3"]);
//   T.addNextEpisodeSlide(pres, 2, "次回タイトル");
//   T.addEndSlide(pres);
//   T.save(pres, "slides.pptx");

var pptxgen = require("pptxgenjs");

// --- シリーズ共通カラーパレット ---
var C = {
  dark: "1B5E20",       // ダーク緑
  primary: "2E7D32",    // メイン緑
  accent: "FF8F00",     // オレンジ（強調）
  light: "E8F5E9",      // ライト緑
  warmBg: "FAFAFA",     // 背景
  white: "FFFFFF",
  text: "2D3436",       // 本文
  sub: "636E72",        // サブテキスト
  red: "DC3545",
  green: "28A745",
  yellow: "F5C518",
  lightRed: "F8D7DA",
  lightYellow: "FFF3CD",
  lightGreen: "D4EDDA",
};

// --- フォント ---
var FJ = "BIZ UDPGothic";
var FE = "Segoe UI";

// --- シリーズ情報 ---
var SERIES_NAME = "外来でそのまま使える食事指導シリーズ";
var EPISODES = [
  "総論 ― 食事指導の基本フレーム",
  "高血圧の食事指導",
  "脂質異常症の食事指導",
  "糖尿病の食事指導",
  "肥満の食事指導",
  "CKDの食事指導",
  "肝疾患の食事指導",
  "嚥下障害の食事指導",
  "消化管術後の食事指導",
  "フレイル・多疾患併存の食事指導",
];

// --- ヘルパー関数 ---
function shd() {
  return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 };
}

function hdr(s, pres, title, bgColor) {
  bgColor = bgColor || C.primary;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: bgColor } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, fontFace: FJ, color: C.white, bold: true, margin: 0 });
}

function footerBar(s, pres, episodeNum) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.35, w: 10, h: 0.28, fill: { color: C.dark } });
  s.addText(SERIES_NAME + "  第" + episodeNum + "回", {
    x: 0.3, y: 5.35, w: 9.4, h: 0.28,
    fontSize: 10, fontFace: FJ, color: "A5D6A7", align: "right", valign: "middle", margin: 0,
  });
}

// --- プレゼンテーション作成 ---
function createPres(title) {
  var pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "今村久司";
  pres.title = title;
  return pres;
}

// --- タイトルスライド（各回共通） ---
function addTitleSlide(pres, episodeNum, mainTitle, subtitle) {
  var s = pres.addSlide();
  s.background = { color: C.dark };

  // 上下のアクセントライン
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  // シリーズ名
  s.addText(SERIES_NAME, {
    x: 0.6, y: 0.5, w: 8.8, h: 0.5,
    fontSize: 16, fontFace: FJ, color: "A5D6A7", align: "center", margin: 0,
  });

  // 回数バッジ
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 3.8, y: 1.2, w: 2.4, h: 0.5, rectRadius: 0.25, fill: { color: C.accent },
  });
  s.addText("第" + episodeNum + "回 / 全10回", {
    x: 3.8, y: 1.2, w: 2.4, h: 0.5,
    fontSize: 14, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // メインタイトル
  s.addText(mainTitle, {
    x: 0.6, y: 2.0, w: 8.8, h: 1.4,
    fontSize: 34, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 区切り線
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 3.6, w: 5, h: 0, line: { color: C.accent, width: 2 } });

  // サブタイトル
  if (subtitle) {
    s.addText(subtitle, {
      x: 0.6, y: 3.9, w: 8.8, h: 0.6,
      fontSize: 18, fontFace: FJ, color: "B0BEC5", align: "center", italic: true, margin: 0,
    });
  }

  return s;
}

// --- まとめスライド ---
function addSummarySlide(pres, episodeNum, points) {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, pres, "まとめ — Key Takeaways", C.primary);

  for (var i = 0; i < points.length; i++) {
    var yPos = 1.2 + i * 0.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.8, y: yPos, w: 8.4, h: 0.6, rectRadius: 0.1, fill: { color: C.white }, shadow: shd(),
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.8, y: yPos, w: 0.12, h: 0.6, rectRadius: 0, fill: { color: C.accent },
    });
    s.addText("✓  " + points[i], {
      x: 1.2, y: yPos, w: 7.8, h: 0.6,
      fontSize: 18, fontFace: FJ, color: C.text, valign: "middle", margin: 0,
    });
  }

  footerBar(s, pres, episodeNum);
  return s;
}

// --- 次回予告スライド ---
function addNextEpisodeSlide(pres, nextEpisodeNum, nextTitle) {
  var s = pres.addSlide();
  s.background = { color: C.dark };

  s.addText("次回予告", {
    x: 0.6, y: 1.2, w: 8.8, h: 0.6,
    fontSize: 16, fontFace: FJ, color: "A5D6A7", align: "center", margin: 0,
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.5, y: 2.0, w: 7, h: 2.0, rectRadius: 0.15, fill: { color: C.dark, transparency: 75 },
  });

  s.addText("第" + nextEpisodeNum + "回", {
    x: 1.5, y: 2.2, w: 7, h: 0.6,
    fontSize: 20, fontFace: FJ, color: C.accent, bold: true, align: "center", margin: 0,
  });
  s.addText(nextTitle || EPISODES[nextEpisodeNum - 1], {
    x: 1.5, y: 2.9, w: 7, h: 0.8,
    fontSize: 26, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });

  s.addText("チャンネル登録で見逃さないようにしましょう", {
    x: 0.6, y: 4.6, w: 8.8, h: 0.5,
    fontSize: 14, fontFace: FJ, color: "B0BEC5", align: "center", margin: 0,
  });

  return s;
}

// --- エンドスライド ---
function addEndSlide(pres) {
  var s = pres.addSlide();
  s.background = { color: C.dark };

  s.addText("ご視聴ありがとうございました", {
    x: 0.6, y: 1.2, w: 8.8, h: 0.8,
    fontSize: 30, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 2.2, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  var items = [
    "👍 高評価・チャンネル登録よろしくお願いします",
    "📋 再生リストから他の回もご覧ください",
    "💬 コメント欄で質問・リクエストお待ちしています",
  ];
  for (var i = 0; i < items.length; i++) {
    s.addText(items[i], {
      x: 1.5, y: 2.6 + i * 0.7, w: 7, h: 0.5,
      fontSize: 18, fontFace: FJ, color: "E8F5E9", align: "center", margin: 0,
    });
  }

  s.addText(SERIES_NAME, {
    x: 0.6, y: 4.8, w: 8.8, h: 0.4,
    fontSize: 12, fontFace: FJ, color: "A5D6A7", align: "center", margin: 0,
  });

  return s;
}

// --- ファイル保存 ---
function save(pres, filename) {
  pres.writeFile({ fileName: filename }).then(function() {
    console.log("✅ " + filename + " を作成しました");
  });
}

// --- エクスポート ---
module.exports = {
  C: C,
  FJ: FJ,
  FE: FE,
  SERIES_NAME: SERIES_NAME,
  EPISODES: EPISODES,
  shd: shd,
  hdr: hdr,
  footerBar: footerBar,
  createPres: createPres,
  addTitleSlide: addTitleSlide,
  addSummarySlide: addSummarySlide,
  addNextEpisodeSlide: addNextEpisodeSlide,
  addEndSlide: addEndSlide,
  save: save,
};
