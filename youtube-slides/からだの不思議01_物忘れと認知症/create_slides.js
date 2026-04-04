var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "「あの人の名前が出てこない」は認知症のサイン？ ― 脳神経内科医が解説";

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

  s.addText("脳神経内科医が答える からだの不思議 #01", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("「あの人の名前が出てこない」\nは認知症のサイン？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― ふつうの物忘れとの違いを解説 ―", {
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

  // 吹き出し風カード
  var quotes = [
    "「あの俳優の名前、何だっけ……\n  顔は浮かんでるのに」",
    "「昨日会った人の名前が\n  どうしても出てこない」",
    "「最近、人の名前が出てこなくて……\n  認知症でしょうか？」",
  ];
  quotes.forEach(function(q, i) {
    var yPos = 1.2 + i * 1.35;
    addCard(s, 1.0, yPos, 8.0, 1.15);
    s.addText(q, {
      x: 1.3, y: yPos + 0.1, w: 7.4, h: 0.95,
      fontSize: 18, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
      lineSpacingMultiple: 1.2,
    });
  });

  // 底部メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 5.0, w: 8.0, h: 0.5, fill: { color: C.light } });
  s.addText("外来で最も多い相談のひとつ", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2/11");
})();

// ============================================================
// SLIDE 3: 舌先現象（TOT現象）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "舌先現象（TOT現象）とは");

  // メインカード
  addCard(s, 0.5, 1.2, 9.0, 1.6);
  s.addText("Tip-of-the-Tongue (TOT) 現象", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.5,
    fontSize: 22, fontFace: FONT_EN, color: C.primary, bold: true, margin: 0,
  });
  s.addText("「喉まで出かかっているのに、あと一歩が出てこない」現象", {
    x: 0.8, y: 1.8, w: 8.4, h: 0.8,
    fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 頻度カード
  var freqs = [
    { label: "若い人", val: "週1回程度", color: C.green },
    { label: "中高年", val: "週数回", color: C.orange },
    { label: "高齢者", val: "ほぼ毎日", color: C.red },
  ];
  freqs.forEach(function(f, i) {
    var xPos = 0.5 + i * 3.1;
    addCard(s, xPos, 3.1, 2.9, 1.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.1, w: 2.9, h: 0.5, fill: { color: f.color } });
    s.addText(f.label, { x: xPos, y: 3.1, w: 2.9, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.val, { x: xPos, y: 3.7, w: 2.9, h: 0.8, fontSize: 20, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // ポイント
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.85, w: 9.0, h: 0.6, fill: { color: C.light } });
  s.addText("「意味」は検索できている → 足りないのは「音（発音パターン）」だけ", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.6,
    fontSize: 16, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "3/11");
})();

// ============================================================
// SLIDE 4: 脳の中で何が起きているか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "脳の中で何が起きているか");

  // 3つの脳領域カード
  var regions = [
    { name: "左下前頭回", role: "ことばの検索", icon: "🔍" },
    { name: "白質神経線維", role: "つなぎ役\n（加齢で衰える）", icon: "🔗" },
    { name: "側頭極", role: "固有名詞の\n貯蔵庫", icon: "📚" },
  ];
  regions.forEach(function(r, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.2, 3.0, 2.3);
    s.addText(r.icon, { x: xPos, y: 1.3, w: 3.0, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(r.name, { x: xPos, y: 1.9, w: 3.0, h: 0.6, fontSize: 18, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", margin: 0 });
    s.addText(r.role, { x: xPos, y: 2.5, w: 3.0, h: 0.8, fontSize: 15, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0 });
  });

  // 矢印（領域間の接続を示す）
  s.addText("→", { x: 3.0, y: 1.8, w: 0.6, h: 0.6, fontSize: 28, color: C.accent, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 6.2, y: 1.8, w: 0.6, h: 0.6, fontSize: 28, color: C.accent, align: "center", valign: "middle", margin: 0 });

  // 結論ボックス
  addCard(s, 0.5, 3.8, 9.0, 1.6);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.8, w: 9.0, h: 0.5, fill: { color: C.primary } });
  s.addText("核心メッセージ", { x: 0.8, y: 3.8, w: 8.4, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "名前が出てこない ≠ 記憶が消えた", options: { fontSize: 18, bold: true, color: C.red, breakLine: true } },
    { text: "「検索ルートが一時的に混雑している」状態", options: { fontSize: 16, color: C.text, breakLine: true } },
    { text: "→ あとから思い出せるのは、別ルートで音の情報に到達したから", options: { fontSize: 14, color: C.sub } },
  ], { x: 0.8, y: 4.3, w: 8.4, h: 1.0, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "4/11");
})();

// ============================================================
// SLIDE 5: ふつうの物忘れと認知症の比較
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ふつうの物忘れ vs 認知症の物忘れ");

  var rows = [
    { scene: "食事", normal: "メニューを忘れる", warning: "食べたこと自体を忘れる" },
    { scene: "約束", normal: "時間を忘れるが\n言われれば思い出す", warning: "約束そのものを忘れ\nヒントでも思い出せない" },
    { scene: "物の置き場所", normal: "鍵の場所を忘れるが\n探して見つけられる", warning: "不自然な場所に置き\n「盗まれた」と言う" },
    { scene: "道順", normal: "初めての場所で迷う", warning: "慣れた近所で迷う" },
  ];

  // テーブルヘッダ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.15, w: 1.8, h: 0.55, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.1, y: 1.15, w: 3.8, h: 0.55, fill: { color: C.green } });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: 1.15, w: 3.8, h: 0.55, fill: { color: C.red } });
  s.addText("場面", { x: 0.3, y: 1.15, w: 1.8, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("ふつうの物忘れ", { x: 2.1, y: 1.15, w: 3.8, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("認知症が疑われる物忘れ", { x: 5.9, y: 1.15, w: 3.8, h: 0.55, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.7 + i * 0.95;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.95, fill: { color: bgColor } });
    s.addText(r.scene, { x: 0.3, y: yPos, w: 1.8, h: 0.95, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.normal, { x: 2.1, y: yPos, w: 3.8, h: 0.95, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 8, 0, 8] });
    s.addText(r.warning, { x: 5.9, y: yPos, w: 3.8, h: 0.95, fontSize: 13, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: [0, 8, 0, 8] });
  });

  addPageNum(s, "5/11");
})();

// ============================================================
// SLIDE 6: 見分けの核心 ― 3つのポイント
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "見分けの核心 ― 3つのポイント");

  var points = [
    { num: "1", title: "ヒントで思い出せるか？", desc: "思い出せる → 正常な物忘れの可能性が高い", color: C.green },
    { num: "2", title: "一部を忘れたか、全体を忘れたか？", desc: "メニューを忘れる vs 食事したこと自体を忘れる\n→ 質がまったく違う", color: C.orange },
    { num: "3", title: "本人に自覚があるか？", desc: "本人が自覚 → 正常な物忘れ\n家族だけが心配 → 要注意", color: C.red },
  ];

  points.forEach(function(p, i) {
    var yPos = 1.2 + i * 1.35;
    addCard(s, 0.5, yPos, 9.0, 1.15);
    // 番号サークル
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.15, w: 0.85, h: 0.85, fill: { color: p.color } });
    s.addText(p.num, { x: 0.7, y: yPos + 0.15, w: 0.85, h: 0.85, fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // タイトル + 説明
    s.addText(p.title, { x: 1.8, y: yPos + 0.05, w: 7.4, h: 0.45, fontSize: 18, fontFace: FONT_JP, color: C.text, bold: true, margin: 0 });
    s.addText(p.desc, { x: 1.8, y: yPos + 0.5, w: 7.4, h: 0.6, fontSize: 14, fontFace: FONT_JP, color: C.sub, margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "6/11");
})();

// ============================================================
// SLIDE 7: MCI（軽度認知障害）とは
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "MCI（軽度認知障害）とは");

  // 定義カード
  addCard(s, 0.5, 1.2, 9.0, 1.2);
  s.addText("Mild Cognitive Impairment", {
    x: 0.8, y: 1.25, w: 8.4, h: 0.4,
    fontSize: 18, fontFace: FONT_EN, color: C.primary, bold: true, italic: true, margin: 0,
  });
  s.addText("年齢のわりに認知機能が低下しているが、日常生活は自立してできている状態", {
    x: 0.8, y: 1.7, w: 8.4, h: 0.6,
    fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 統計カード（厚労省2022）
  s.addText("65歳以上の高齢者（厚生労働省 2022年調査）", {
    x: 0.5, y: 2.6, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  var stats = [
    { label: "認知症", val: "約12.3%", sub: "443万人", color: C.red },
    { label: "MCI", val: "約15.5%", sub: "559万人", color: C.orange },
    { label: "合計", val: "約4人に1人", sub: "1,002万人", color: C.primary },
  ];
  stats.forEach(function(st, i) {
    var xPos = 0.5 + i * 3.1;
    addCard(s, xPos, 3.1, 2.9, 2.2);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.1, w: 2.9, h: 0.5, fill: { color: st.color } });
    s.addText(st.label, { x: xPos, y: 3.1, w: 2.9, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.val, { x: xPos, y: 3.7, w: 2.9, h: 0.8, fontSize: 24, fontFace: FONT_JP, color: st.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.sub, { x: xPos, y: 4.5, w: 2.9, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", valign: "middle", margin: 0 });
  });

  addPageNum(s, "7/11");
})();

// ============================================================
// SLIDE 8: MCIの行方
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "MCIの行方 ― すべてが認知症になるわけではない");

  // 3つの転帰カード
  var outcomes = [
    { label: "認知症に進展", val: "年間\n10〜15%", color: C.red, icon: "↗" },
    { label: "正常に回復", val: "14〜40%", color: C.green, icon: "↙" },
    { label: "MCIのまま安定", val: "残りの\n多くの方", color: C.accent, icon: "→" },
  ];
  outcomes.forEach(function(o, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.2, 3.0, 2.5);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.8, y: 1.4, w: 1.4, h: 1.4, fill: { color: o.color } });
    s.addText(o.val, { x: xPos + 0.8, y: 1.4, w: 1.4, h: 1.4, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(o.label, { x: xPos, y: 3.0, w: 3.0, h: 0.6, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // メッセージ
  addCard(s, 0.5, 4.1, 9.0, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.1, w: 0.15, h: 1.3, fill: { color: C.green } });
  s.addText([
    { text: "MCIの段階で気づければ", options: { bold: true, color: C.primary, breakLine: true } },
    { text: "生活習慣の改善や適切な介入で、進行を遅らせたり回復したりする可能性がある", options: { color: C.text, breakLine: true } },
    { text: "→ これが「早めの受診」が勧められる最大の理由", options: { color: C.green, bold: true } },
  ], { x: 0.9, y: 4.2, w: 8.4, h: 1.1, fontSize: 15, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "8/11");
})();

// ============================================================
// SLIDE 9: セルフチェック ― こんなときは要注意
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "セルフチェック ― こんなときは要注意");

  var checks = [
    "数分前の会話内容を覚えていない",
    "同じ話や質問を何度も繰り返す",
    "大切な約束をメモがあっても忘れる",
    "日付や曜日が頻繁にわからなくなる",
    "慣れた場所（近所）で迷う",
    "長年の料理の手順がわからなくなった",
    "お金の管理・計算でミスが著しく増えた",
    "趣味や社会活動への興味を急に失った",
    "性格が変わった（怒りっぽい・無気力）",
    "身だしなみに無頓着になった",
  ];

  // 2列レイアウト
  checks.forEach(function(c, i) {
    var col = (i < 5) ? 0 : 1;
    var row = (i < 5) ? i : i - 5;
    var xPos = 0.3 + col * 4.8;
    var yPos = 1.15 + row * 0.78;

    addCard(s, xPos, yPos, 4.6, 0.65);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.1, y: yPos + 0.15, w: 0.35, h: 0.35, fill: { color: C.red } });
    s.addText("" + (i + 1), { x: xPos + 0.1, y: yPos + 0.15, w: 0.35, h: 0.35, fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c, { x: xPos + 0.55, y: yPos, w: 3.9, h: 0.65, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  // 底部注記
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.1, w: 9.4, h: 0.45, fill: { color: C.light } });
  s.addText("複数該当 + 以前と比べて変化 → 一度相談する価値あり", {
    x: 0.5, y: 5.1, w: 9.0, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "9/11");
})();

// ============================================================
// SLIDE 10: 受診のタイミングと相談先
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "受診のタイミングと相談先");

  // 「もの忘れ外来」推奨カード
  addCard(s, 0.5, 1.2, 4.3, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.45, fill: { color: C.primary } });
  s.addText("迷ったら「もの忘れ外来」へ", { x: 0.7, y: 1.2, w: 3.9, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText("脳神経内科が得意とする領域\n認知機能検査～画像検査まで\nワンストップで評価", {
    x: 0.7, y: 1.75, w: 3.9, h: 0.6,
    fontSize: 13, fontFace: FONT_JP, color: C.text, margin: 0, lineSpacingMultiple: 1.3,
  });

  // 早めに受診すべき場合
  addCard(s, 5.2, 1.2, 4.5, 3.4);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.5, h: 0.45, fill: { color: C.red } });
  s.addText("早めに受診したほうがいい場合", { x: 5.4, y: 1.2, w: 4.1, h: 0.45, fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var urgents = [
    "家族や周囲の人が心配している",
    "数か月の間に目に見えて悪化",
    "慣れた道で迷うようになった",
    "お金・服薬の管理ができなくなった",
    "性格や行動に明らかな変化がある",
  ];
  urgents.forEach(function(u, i) {
    var yPos = 1.8 + i * 0.52;
    s.addText("●  " + u, {
      x: 5.5, y: yPos, w: 4.0, h: 0.48,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  // 受診をためらう方へ
  addCard(s, 0.5, 2.6, 4.3, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.6, w: 0.12, h: 2.0, fill: { color: C.green } });
  s.addText("受診をためらう方へ", { x: 0.8, y: 2.7, w: 3.8, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.green, bold: true, margin: 0 });
  s.addText([
    { text: "「何も問題なかった」も大きな価値", options: { bold: true, color: C.text, breakLine: true } },
    { text: "→ ベースラインの記録として将来の比較データになる", options: { color: C.sub, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "本人が嫌がる場合:", options: { bold: true, color: C.text, breakLine: true } },
    { text: "「健康診断のついで」「頭痛の相談」などきっかけ作りを", options: { color: C.sub } },
  ], { x: 0.8, y: 3.1, w: 3.8, h: 1.4, fontSize: 13, fontFace: FONT_JP, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

  addPageNum(s, "10/11");
})();

// ============================================================
// SLIDE 11: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.3, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "「舌先現象」は正常な加齢変化。若い人にも起こる" },
    { num: "2", text: "見分けのポイントは\n「ヒントで思い出せるか」「全体を忘れるか」「本人の自覚」" },
    { num: "3", text: "MCIの段階で気づけば、回復の可能性もある" },
    { num: "4", text: "気になったら「もの忘れ外来」へ\n何もなくても将来の比較データに" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.1 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.2, h: 0.9, fontSize: 16, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  // 底部メッセージ
  s.addShape(pres.shapes.LINE, { x: 2, y: 5.1, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.2, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "11/11");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/物忘れと認知症_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
