var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "「最近つまずきやすい」は脳の問題？ ― 歩行とつまずきの原因を解剖学レベルで解説";

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

  s.addText("脳神経内科医が答える からだの不思議 #15", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("「最近つまずきやすい」\nは脳の問題？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 歩行とつまずきの原因を解剖学レベルで解説 ―", {
    x: 0.8, y: 2.9, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.7, w: 4, h: 0, line: { color: C.accent, width: 2 } });

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
// SLIDE 2: こんな経験ありませんか？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんな経験ありませんか？");

  var quotes = [
    "「段差でもないのに、なぜかつまずいた」",
    "「足が地面に引っかかってよろめく」",
    "「歩き始めの一歩がうまく出ない」",
  ];
  quotes.forEach(function(q, i) {
    var yPos = 1.2 + i * 1.35;
    addCard(s, 1.0, yPos, 8.0, 1.15);
    s.addText(q, {
      x: 1.3, y: yPos + 0.1, w: 7.4, h: 0.95,
      fontSize: 20, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
      lineSpacingMultiple: 1.2,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 5.0, w: 8.0, h: 0.5, fill: { color: C.light } });
  s.addText("「老化だから仕方ない」と放置するのは危険です", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.red, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: 歩行を支える5つのシステム
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "歩行を支える5つのシステム ― 脳から筋肉まで");

  var systems = [
    { name: "大脳運動野\n・補足運動野", role: "随意的な歩行開始\n歩行プログラムの指令", color: C.primary },
    { name: "基底核", role: "歩行リズム調整\n動作の自動化", color: C.accent },
    { name: "小脳", role: "バランス\n四肢の協調運動", color: C.green },
    { name: "脊髄・末梢神経", role: "運動指令を足へ伝達\n足底感覚を脳へ伝達", color: C.orange },
    { name: "筋骨格系", role: "地面を蹴る\n体を支える", color: C.red },
  ];

  systems.forEach(function(sys, i) {
    var xPos = 0.15 + i * 1.95;
    addCard(s, xPos, 1.1, 1.75, 3.8);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 1.75, h: 0.6, fill: { color: sys.color } });
    s.addText(sys.name, {
      x: xPos, y: 1.1, w: 1.75, h: 0.6,
      fontSize: 12, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 2,
    });
    s.addText(sys.role, {
      x: xPos + 0.1, y: 1.8, w: 1.55, h: 2.9,
      fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0,
      lineSpacingMultiple: 1.4,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: 5.0, w: 9.65, h: 0.5, fill: { color: C.light } });
  s.addText("どこかに障害が生じると、特有の「歩き方の乱れ」やつまずきが現れる", {
    x: 0.3, y: 5.0, w: 9.4, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: 解剖学レベル別の原因（比較表）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "解剖学レベル別 ― つまずきの原因");

  // ヘッダ行
  s.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: 1.0, w: 2.2, h: 0.5, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.35, y: 1.0, w: 2.6, h: 0.5, fill: { color: C.primary } });
  s.addShape(pres.shapes.RECTANGLE, { x: 4.95, y: 1.0, w: 4.9, h: 0.5, fill: { color: C.accent } });
  s.addText("障害レベル", { x: 0.15, y: 1.0, w: 2.2, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("代表的な疾患", { x: 2.35, y: 1.0, w: 2.6, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("特徴的な歩行・症状", { x: 4.95, y: 1.0, w: 4.9, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var rows = [
    { level: "中枢神経\n（脳・脊髄）", disease: "正常圧水頭症（NPH）", feature: "小刻み・すり足・極端に狭い歩幅。認知低下・尿失禁を伴う", bg: C.warmBg },
    { level: "", disease: "パーキンソン病", feature: "前傾姿勢・すくみ足（前に踏み出せない）。方向転換で転倒", bg: C.white },
    { level: "", disease: "頸椎症性脊髄症", feature: "手指の巧緻障害＋両足のぎこちなさ。階段の下りが特に不安定", bg: C.warmBg },
    { level: "", disease: "小脳疾患・脳幹梗塞", feature: "千鳥足のふらつき・足を広げて歩く。目を閉じると悪化", bg: C.white },
    { level: "末梢神経", disease: "多発神経障害（PN）", feature: "足底の感覚鈍麻・しびれ・垂れ足。暗所や凹凸で転倒しやすい", bg: C.warmBg },
    { level: "", disease: "腓骨神経麻痺", feature: "片側のつま先が上がらない「ぱたぱた」歩き。膝下外側のしびれ", bg: C.white },
    { level: "筋骨格系", disease: "サルコペニア・フレイル", feature: "全体的な筋力低下。椅子から立つのが辛い。段差でつまずく", bg: C.warmBg },
  ];

  rows.forEach(function(r, i) {
    var yPos = 1.5 + i * 0.57;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: yPos, w: 9.65, h: 0.57, fill: { color: r.bg } });
    if (r.level !== "") {
      s.addText(r.level, { x: 0.15, y: yPos, w: 2.2, h: 0.57, fontSize: 11, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    }
    s.addText(r.disease, { x: 2.35, y: yPos, w: 2.6, h: 0.57, fontSize: 11, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: [0, 4, 0, 4] });
    s.addText(r.feature, { x: 4.95, y: yPos, w: 4.9, h: 0.57, fontSize: 11, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 4, 0, 4] });
  });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: 正常圧水頭症・パーキンソン病・頸椎症
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "中枢神経の原因 ― 早期対処で改善できる疾患");

  var diseases = [
    {
      name: "正常圧水頭症\n（NPH）",
      color: C.red,
      points: ["小刻み・すり足歩行", "認知機能低下を伴う", "尿失禁の3徴", "シャント術で改善する\n数少ない「治せる」歩行障害"],
    },
    {
      name: "パーキンソン病",
      color: C.orange,
      points: ["すくみ足・前傾姿勢", "歩き始めの一歩が出ない", "方向転換で転倒しやすい", "振戦・動作緩慢を伴う\n薬物療法が有効"],
    },
    {
      name: "頸椎症性脊髄症",
      color: C.accent,
      points: ["両足のぎこちない歩き", "手指の巧緻障害を伴う", "ボタン・箸が難しくなる", "中高年に多い\n整形外科・脳神経外科で評価"],
    },
  ];

  diseases.forEach(function(d, i) {
    var xPos = 0.2 + i * 3.25;
    addCard(s, xPos, 1.1, 3.05, 4.35);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 3.05, h: 0.65, fill: { color: d.color } });
    s.addText(d.name, {
      x: xPos, y: 1.1, w: 3.05, h: 0.65,
      fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 2,
    });
    d.points.forEach(function(pt, j) {
      var yPt = 1.9 + j * 0.85;
      s.addShape(pres.shapes.OVAL, { x: xPos + 0.15, y: yPt + 0.05, w: 0.35, h: 0.35, fill: { color: d.color } });
      s.addText(pt, {
        x: xPos + 0.6, y: yPt, w: 2.35, h: 0.72,
        fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
      });
    });
  });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: 末梢神経障害・腓骨神経麻痺
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "末梢神経の原因 ― 感覚障害・垂れ足");

  // 左カード：多発神経障害
  addCard(s, 0.2, 1.1, 4.5, 4.35);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 1.1, w: 4.5, h: 0.6, fill: { color: C.primary } });
  s.addText("多発神経障害（PN）", { x: 0.2, y: 1.1, w: 4.5, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var pnItems = [
    "足底のしびれ・感覚鈍麻",
    "足先が上がりにくい（垂れ足）",
    "暗所・凹凸での転倒",
    "糖尿病・飲酒・\nビタミンB12欠乏が主な原因",
    "目を閉じると不安定さ増強",
  ];
  pnItems.forEach(function(item, i) {
    var yPos = 1.85 + i * 0.72;
    s.addShape(pres.shapes.OVAL, { x: 0.4, y: yPos + 0.08, w: 0.32, h: 0.32, fill: { color: C.primary } });
    s.addText(item, { x: 0.85, y: yPos, w: 3.7, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // 右カード：腓骨神経麻痺
  addCard(s, 5.1, 1.1, 4.7, 4.35);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.7, h: 0.6, fill: { color: C.green } });
  s.addText("腓骨神経麻痺", { x: 5.1, y: 1.1, w: 4.7, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var peItems = [
    "片側の足背屈障害\n（つま先が上がらない）",
    "「ぱたぱた」した歩き方",
    "膝下外側〜足背のしびれ",
    "長時間の足組み・\n急激な体重減少が誘因",
    "多くは保存的治療で回復",
  ];
  peItems.forEach(function(item, i) {
    var yPos = 1.85 + i * 0.72;
    s.addShape(pres.shapes.OVAL, { x: 5.3, y: yPos + 0.08, w: 0.32, h: 0.32, fill: { color: C.green } });
    s.addText(item, { x: 5.75, y: yPos, w: 3.85, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: サルコペニア・フレイル
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "筋骨格系の原因 ― サルコペニア・フレイル");

  // 定義カード
  addCard(s, 0.5, 1.1, 9.0, 1.1);
  s.addText("サルコペニア", {
    x: 0.8, y: 1.15, w: 3.5, h: 0.45,
    fontSize: 18, fontFace: FONT_EN, color: C.red, bold: true, italic: true, margin: 0,
  });
  s.addText("加齢による筋肉量・筋力・身体機能の低下", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 特徴カード 3列
  var feats = [
    { title: "主な症状", items: ["椅子から立つのが辛い", "段差でよくつまずく", "歩くのが遅くなった", "握力の低下"], color: C.red },
    { title: "チェック方法", items: ["握力: 男性<28kg\n女性<18kg", "歩行速度: <1.0m/s", "椅子立ち上がりテスト\n5回で>12秒"], color: C.orange },
    { title: "対策・予防", items: ["タンパク質を1日\n体重×1.2g摂取", "レジスタンス運動\n週2〜3回", "日光浴・ビタミンD補充", "有酸素運動との組み合わせ"], color: C.green },
  ];

  feats.forEach(function(f, i) {
    var xPos = 0.2 + i * 3.25;
    addCard(s, xPos, 2.5, 3.05, 3.0);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.5, w: 3.05, h: 0.5, fill: { color: f.color } });
    s.addText(f.title, { x: xPos, y: 2.5, w: 3.05, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    f.items.forEach(function(it, j) {
      var yIt = 3.15 + j * 0.57;
      s.addText("• " + it, { x: xPos + 0.1, y: yIt, w: 2.85, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });
    });
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: 受診判断フロー ― こんなときは要注意
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "受診判断フロー ― こんなときは要注意");

  // 左: 急いで受診
  addCard(s, 0.2, 1.1, 4.5, 4.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 1.1, w: 4.5, h: 0.55, fill: { color: C.red } });
  s.addText("早めに受診が必要", { x: 0.2, y: 1.1, w: 4.5, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  var urgents = [
    "数週〜数か月で急に歩行悪化",
    "ふるえ・動作の遅さが出てきた\n→ パーキンソン病を疑う",
    "認知機能低下＋尿失禁を伴う\n→ 正常圧水頭症を疑う",
    "首・手のしびれが同時に悪化\n→ 頸椎症を疑う",
    "片側の足先だけ上がらない",
    "糖尿病治療中の足のしびれ",
  ];
  urgents.forEach(function(u, i) {
    var yPos = 1.8 + i * 0.55;
    s.addShape(pres.shapes.OVAL, { x: 0.35, y: yPos + 0.1, w: 0.3, h: 0.3, fill: { color: C.red } });
    s.addText(u, { x: 0.75, y: yPos, w: 3.8, h: 0.5, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
  });

  // 右: 救急受診
  addCard(s, 5.1, 1.1, 4.7, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.7, h: 0.55, fill: { color: C.dark } });
  s.addText("すぐ救急受診", { x: 5.1, y: 1.1, w: 4.7, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.yellow, bold: true, align: "center", valign: "middle", margin: 0 });

  var emergencies = [
    "突然の強いふらつき・片側脱力・言語障害\n→ 脳卒中の疑い",
    "転倒後の激しい頭痛・意識がぼんやり\n→ 硬膜下血腫の疑い",
    "両足が急に動かない・排尿ができない\n→ 脊髄圧迫の疑い",
  ];
  emergencies.forEach(function(em, i) {
    var yPos = 1.8 + i * 0.42;
    s.addText("! " + em, { x: 5.2, y: yPos, w: 4.5, h: 0.38, fontSize: 11, fontFace: FONT_JP, color: C.red, valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
  });

  // 右: 相談先
  addCard(s, 5.1, 3.3, 4.7, 1.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 3.3, w: 4.7, h: 0.5, fill: { color: C.green } });
  s.addText("相談先の目安", { x: 5.1, y: 3.3, w: 4.7, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText([
    { text: "神経症状を伴う → 脳神経内科\n", options: { color: C.primary, bold: true } },
    { text: "痛みが主体 → 整形外科\n", options: { color: C.text } },
    { text: "迷ったら → かかりつけ医", options: { color: C.sub } },
  ], { x: 5.2, y: 3.9, w: 4.5, h: 1.1, fontSize: 13, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.5 });

  addPageNum(s, "8/9");
})();

// ============================================================
// SLIDE 9: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.2, w: 9, h: 0.6,
    fontSize: 22, fontFace: FONT_EN, color: C.accent, bold: true, align: "center", margin: 0,
  });

  var msgs = [
    { num: "1", text: "つまずきの原因は脳・脊髄・末梢神経・筋骨格系の\nどのレベルかによって全く異なる", color: C.accent },
    { num: "2", text: "正常圧水頭症・パーキンソン病・頸椎症は\n早期対処で改善できる可能性がある", color: C.green },
    { num: "3", text: "「どんな状況で」「どちらの足が」「他にどんな症状が」\nを観察することが診断の鍵", color: C.yellow },
    { num: "4", text: "「年だから仕方ない」と放置せず、\n気になったら専門家に相談を", color: C.red },
  ];

  msgs.forEach(function(m, i) {
    var yPos = 1.05 + i * 1.1;
    s.addShape(pres.shapes.OVAL, { x: 0.3, y: yPos + 0.1, w: 0.7, h: 0.7, fill: { color: m.color } });
    s.addText(m.num, { x: 0.3, y: yPos + 0.1, w: 0.7, h: 0.7, fontSize: 22, fontFace: FONT_EN, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, {
      x: 1.2, y: yPos, w: 8.5, h: 0.9,
      fontSize: 17, fontFace: FONT_JP, color: C.white, valign: "middle", margin: 0, lineSpacingMultiple: 1.3,
    });
  });

  s.addText("医知創造ラボ", {
    x: 0.5, y: 5.2, w: 9, h: 0.35,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "9/9");
})();

// ============================================================
// 出力
// ============================================================
var outPath = __dirname + "/つまずき_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
