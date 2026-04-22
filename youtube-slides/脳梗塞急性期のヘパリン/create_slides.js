var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "脳梗塞急性期のヘパリン 結局どう使う？";

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
var TOTAL = 11;

function makeShadow() { return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 }; }

function addHeader(s, title) {
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 22, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
}

function addCard(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: C.white }, shadow: makeShadow() });
}

function addPageNum(s, num) {
  s.addText(num + "/" + TOTAL, { x: 9.2, y: 5.2, w: 0.6, h: 0.4, fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "right", margin: 0 });
}

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("脳神経内科・研修医・専攻医のための", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("脳梗塞急性期のヘパリン\n結局どう使う？", {
    x: 0.8, y: 1.1, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0, lineSpacingMultiple: 1.15,
  });

  s.addText("1万単位 vs APTT管理 / ブリッジング不要論 / HIT対応", {
    x: 0.8, y: 2.85, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.5, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  s.addText("脳卒中治療ガイドライン2021〔改訂2025〕対応", {
    x: 0.8, y: 3.7, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.light, align: "center", italic: true, margin: 0,
  });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.3, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });

  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.8, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 3つの論点
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "若手医師が悩む3つの論点");

  var qs = [
    { n: "1", t: "用量・モニタリング", d: "1万単位/日固定か？\nAPTTで調整すべきか？", color: C.primary },
    { n: "2", t: "ブリッジング", d: "心原性脳塞栓症で\nDOACへの橋渡しは必要？", color: C.orange },
    { n: "3", t: "HIT", d: "低用量でも発症する？\n気づけるか？", color: C.red },
  ];
  qs.forEach(function(p, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.3, 3.1, 3.6);
    s.addShape(pres.shapes.OVAL, { x: xPos + 1.15, y: 1.5, w: 0.8, h: 0.8, fill: { color: p.color } });
    s.addText(p.n, { x: xPos + 1.15, y: 1.5, w: 0.8, h: 0.8, fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.t, { x: xPos + 0.15, y: 2.4, w: 2.8, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.d, { x: xPos + 0.15, y: 3.0, w: 2.8, h: 1.8, fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.1, w: 9.4, h: 0.4, fill: { color: C.light } });
  s.addText([
    { text: "本日のゴール: ", options: { bold: true, color: C.primary, fontSize: 13 } },
    { text: "最新GL (2025) + ELAN・OPTIMAS に基づき、明日から使える実践方針を整理", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.6, y: 5.1, w: 8.8, h: 0.4, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, 2);
})();

// ============================================================
// SLIDE 3: ヘパリンのジレンマ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ヘパリンのジレンマ ― 再発は防ぐが転帰は改善せず");

  addCard(s, 0.3, 1.1, 9.4, 1.2);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.4, h: 0.4, fill: { color: C.primary } });
  s.addText("『脳卒中治療ガイドライン2021〔改訂2025〕』", { x: 0.6, y: 1.1, w: 9, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "脳梗塞急性期の未分画ヘパリン使用は ", options: { color: C.text, fontSize: 14 } },
    { text: "「考慮しても良い（推奨度C、エビデンスレベル中）」", options: { color: C.red, fontSize: 15, bold: true } },
    { text: "\n欧米GL: 再発予防・症候改善目的のルーチン抗凝固は推奨されない", options: { color: C.sub, fontSize: 13 } },
  ], { x: 0.6, y: 1.55, w: 9, h: 0.7, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });

  addCard(s, 0.3, 2.5, 9.4, 2.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 2.5, w: 9.4, h: 0.4, fill: { color: C.accent } });
  s.addText("International Stroke Trial (IST, n=19,435, Lancet 1997)", { x: 0.6, y: 2.5, w: 9, h: 0.4, fontSize: 13, fontFace: FONT_EN, color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "14日以内の脳梗塞再発  ", options: { color: C.text, fontSize: 14 } },
    { text: "2.9% ", options: { color: C.green, fontSize: 16, bold: true } },
    { text: "vs アスピリン ", options: { color: C.text, fontSize: 14 } },
    { text: "3.8%", options: { color: C.sub, fontSize: 14 } },
  ], { x: 0.6, y: 3.0, w: 9, h: 0.4, fontFace: FONT_JP, valign: "middle", margin: 0 });

  s.addText([
    { text: "しかし ", options: { color: C.text, fontSize: 14 } },
    { text: "出血性脳卒中が有意に増加", options: { color: C.red, fontSize: 15, bold: true } },
    { text: " → ベネフィットを相殺", options: { color: C.text, fontSize: 14 } },
  ], { x: 0.6, y: 3.45, w: 9, h: 0.4, fontFace: FONT_JP, valign: "middle", margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.95, w: 8.8, h: 1.2, fill: { color: C.light } });
  s.addText([
    { text: "結論: ", options: { bold: true, color: C.primary, fontSize: 14 } },
    { text: "虚血再発を抑える力はあるが、出血増加で全体転帰は改善せず\n", options: { color: C.text, fontSize: 14 } },
    { text: "→ 漫然投与ではなく ", options: { color: C.text, fontSize: 14 } },
    { text: "「病態を選んで使う個別化治療」", options: { color: C.primary, fontSize: 14, bold: true } },
    { text: " が求められる", options: { color: C.text, fontSize: 14 } },
  ], { x: 0.8, y: 3.95, w: 8.4, h: 1.2, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.4 });

  addPageNum(s, 3);
})();

// ============================================================
// SLIDE 4: 1万単位 vs APTT管理
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "1万単位固定 vs APTT管理 ― どちらが正解？");

  var rows = [
    { a: "体重ベース治療量", c: C.orange, d: "欧米標準: 80 U/kg ボーラス + 18 U/kg/hr", e: "60kg例で>25,000 U/日 → 脳梗塞では出血リスク大" },
    { a: "固定用量1万 U/日", c: C.accent, d: "日本の経験的開始用量（添付文書ベース）", e: "簡便・出血リスク低いが、効果不足例あり" },
    { a: "APTT調整（前値の2〜2.5倍）", c: C.green, d: "ガイドラインが言及する目標値", e: "固定で開始 → APTTで微調整するのが現実的" },
  ];

  rows.forEach(function(r, i) {
    var yy = 1.2 + i * 1.25;
    addCard(s, 0.3, yy, 9.4, 1.1);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yy, w: 3.2, h: 1.1, fill: { color: r.c } });
    s.addText(r.a, { x: 0.4, y: yy, w: 3.0, h: 1.1, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", align: "center", margin: 0 });
    s.addText(r.d, { x: 3.6, y: yy + 0.1, w: 6.0, h: 0.45, fontSize: 13, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(r.e, { x: 3.6, y: yy + 0.55, w: 6.0, h: 0.5, fontSize: 12, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.05, w: 9.4, h: 0.45, fill: { color: C.light } });
  s.addText([
    { text: "現実解: ", options: { bold: true, color: C.primary, fontSize: 13 } },
    { text: "固定用量で開始 → 早期APTT測定 → 目標値へ微調整 (= 3方式の折衷)", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.6, y: 5.05, w: 8.8, h: 0.45, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, 4);
})();

// ============================================================
// SLIDE 5: 実践タイムライン
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "実践タイムライン ― 開始から微調整まで");

  var steps = [
    { t: "0h", title: "開始", d: "未分画ヘパリン 10,000〜15,000 U/日 持続静注\n（体重・腎機能・出血リスク考慮）", color: C.primary },
    { t: "6〜12h", title: "初回APTT測定", d: "目標: 投与前値の 2〜2.5倍\n（施設により1.5〜2.5倍）", color: C.accent },
    { t: "調整", title: "用量増減", d: "目標未達 → 増量\n過延長 → 減量または一時中止", color: C.green },
    { t: "毎日〜隔日", title: "継続モニタ", d: "APTT + 血小板数\nHIT発症時期 (5〜14日) に注意", color: C.orange },
  ];

  steps.forEach(function(p, i) {
    var xPos = 0.3 + i * 2.4;
    addCard(s, xPos, 1.3, 2.3, 3.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.3, w: 2.3, h: 0.55, fill: { color: p.color } });
    s.addText(p.t, { x: xPos, y: 1.3, w: 2.3, h: 0.55, fontSize: 14, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.title, { x: xPos + 0.1, y: 1.95, w: 2.1, h: 0.55, fontSize: 15, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.d, { x: xPos + 0.1, y: 2.55, w: 2.1, h: 2.1, fontSize: 11, fontFace: FONT_JP, color: C.sub, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.RIGHT_TRIANGLE, { x: xPos + 2.32, y: 2.85, w: 0.15, h: 0.3, fill: { color: C.sub }, rotate: 90 });
    }
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.5, fill: { color: C.light } });
  s.addText([
    { text: "KEY: ", options: { bold: true, color: C.primary, fontSize: 13 } },
    { text: "固定用量は「開始用量」に過ぎない。APTTでの微調整が不可欠", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.6, y: 5.0, w: 8.8, h: 0.5, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, 5);
})();

// ============================================================
// SLIDE 6: ブリッジング不要論
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "心原性脳塞栓症 ― ブリッジングは過去のもの");

  addCard(s, 0.3, 1.1, 9.4, 0.9);
  s.addText([
    { text: "かつて: ", options: { bold: true, color: C.sub, fontSize: 14 } },
    { text: "ワルファリン治療域到達までヘパリン持続静注で「橋渡し」\n", options: { color: C.sub, fontSize: 13 } },
    { text: "現在: ", options: { bold: true, color: C.green, fontSize: 14 } },
    { text: "DOAC早期開始が新しい標準 (GL2025 推奨度B) → ルーチンのヘパリンブリッジング不要", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.5, y: 1.15, w: 9.0, h: 0.85, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });

  var trials = [
    { n: "TIMING", j: "Circulation 2022", r: "発症4日以内開始は\n5〜10日後開始に非劣性" },
    { n: "ELAN", j: "NEJM 2023", r: "30日複合エンドポイント\n早期 2.9% vs 遅延 4.1%" },
    { n: "OPTIMAS", j: "Lancet 2024", r: "≤4日 vs 7〜14日で非劣性\n症候性頭蓋内出血 0.6%" },
  ];
  trials.forEach(function(p, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 2.15, 3.1, 2.5);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.15, w: 3.1, h: 0.5, fill: { color: C.primary } });
    s.addText(p.n, { x: xPos, y: 2.15, w: 3.1, h: 0.5, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.j, { x: xPos + 0.15, y: 2.75, w: 2.8, h: 0.4, fontSize: 11, fontFace: FONT_EN, color: C.sub, italic: true, align: "center", margin: 0 });
    s.addText(p.r, { x: xPos + 0.15, y: 3.2, w: 2.8, h: 1.4, fontSize: 13, fontFace: FONT_JP, color: C.text, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.85, w: 9.4, h: 0.7, fill: { color: C.light } });
  s.addText([
    { text: "結論: ", options: { bold: true, color: C.primary, fontSize: 14 } },
    { text: "NVAF合併急性期脳梗塞 → ", options: { color: C.text, fontSize: 13 } },
    { text: "血圧管理の下、早期DOAC開始が標準", options: { color: C.primary, fontSize: 14, bold: true } },
    { text: " (ヘパリンブリッジング不要)", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.6, y: 4.85, w: 8.8, h: 0.7, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, 6);
})();

// ============================================================
// SLIDE 7: 1-2-3-4日ルール
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "1-2-3-4日ルール ― 梗塞サイズで開始日を決める");

  s.addText("国立循環器病研究センターの提唱 (旧: 1-3-6-12日ルール)", {
    x: 0.5, y: 0.98, w: 9, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  var rules = [
    { day: "1日", sev: "TIA", color: C.green, d: "症状は24時間以内に消失\n画像上の急性期梗塞なし" },
    { day: "2日", sev: "軽症", color: C.accent, d: "小梗塞 (≤1.5cm)\nNIHSS低値" },
    { day: "3日", sev: "中等症", color: C.orange, d: "中梗塞\n出血リスク中等度" },
    { day: "4日", sev: "重症", color: C.red, d: "大梗塞 (皮質含む)\n出血性変化リスク高" },
  ];
  rules.forEach(function(r, i) {
    var xPos = 0.3 + i * 2.4;
    addCard(s, xPos, 1.5, 2.3, 3.3);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.5, w: 2.3, h: 0.7, fill: { color: r.color } });
    s.addText(r.day, { x: xPos, y: 1.5, w: 2.3, h: 0.7, fontSize: 22, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(r.sev, { x: xPos + 0.1, y: 2.3, w: 2.1, h: 0.5, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(r.d, { x: xPos + 0.1, y: 2.85, w: 2.1, h: 1.9, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.0, w: 9.4, h: 0.5, fill: { color: C.light } });
  s.addText([
    { text: "応用: ", options: { bold: true, color: C.primary, fontSize: 13 } },
    { text: "梗塞サイズ・出血リスクを評価 → 該当日数後にDOAC開始 (血圧140/90以下も確認)", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.6, y: 5.0, w: 8.8, h: 0.5, fontFace: FONT_JP, valign: "middle", margin: 0 });

  addPageNum(s, 7);
})();

// ============================================================
// SLIDE 8: ヘパリンが主役の病態
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "ヘパリンが主役となる4つの病態");

  var dxs = [
    { n: "脳静脈・静脈洞血栓症 (CSVT)", d: "頭蓋内出血合併例でも\n急性期の第一選択は未分画ヘパリン持続静注", g: "推奨度B", color: C.red },
    { n: "癌関連脳梗塞 (Trousseau症候群)", d: "多発性・進行性の急性期は\nヘパリンによる強力抗凝固が主体", g: "DOACも選択肢", color: C.orange },
    { n: "抗リン脂質抗体症候群 (APS)", d: "急性期はヘパリン\n維持期はワルファリン (PT-INR 2〜3)", g: "DOAC非推奨", color: C.primary },
    { n: "頭蓋外動脈解離", d: "虚血発症急性期の選択肢\n抗血小板薬と並んで考慮", g: "推奨度B", color: C.accent },
  ];

  dxs.forEach(function(p, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var xPos = 0.3 + col * 4.8;
    var yPos = 1.15 + row * 2.05;
    addCard(s, xPos, yPos, 4.6, 1.9);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 4.6, h: 0.5, fill: { color: p.color } });
    s.addText(p.n, { x: xPos + 0.15, y: yPos, w: 4.3, h: 0.5, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
    s.addText(p.d, { x: xPos + 0.15, y: yPos + 0.55, w: 4.3, h: 0.9, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.15, y: yPos + 1.45, w: 4.3, h: 0.35, fill: { color: C.light } });
    s.addText(p.g, { x: xPos + 0.25, y: yPos + 1.45, w: 4.1, h: 0.35, fontSize: 11, fontFace: FONT_JP, color: p.color, bold: true, valign: "middle", margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.25, w: 9.4, h: 0.3, fill: { color: C.warmBg } });
  s.addText("ルーチンは不要だが、これらの病態では依然として「強力な武器」", {
    x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 11, fontFace: FONT_JP, color: C.sub, italic: true, valign: "middle", margin: 0,
  });

  addPageNum(s, 8);
})();

// ============================================================
// SLIDE 9: HIT 頻度と免疫原性
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "HIT ― 低用量でも発症する免疫原性疾患");

  addCard(s, 0.3, 1.1, 4.7, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 4.7, h: 0.45, fill: { color: C.red } });
  s.addText("頻度 (神経内科領域)", { x: 0.5, y: 1.1, w: 4.4, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "ヘパリン投与患者: ", options: { color: C.text, fontSize: 13 } },
    { text: "0.1〜7%\n", options: { color: C.red, fontSize: 15, bold: true } },
    { text: "急性期脳梗塞: ", options: { color: C.text, fontSize: 13 } },
    { text: "definite HIT 1.7%\n", options: { color: C.red, fontSize: 14, bold: true } },
    { text: "神経内科全体: ", options: { color: C.text, fontSize: 13 } },
    { text: "約 2.5%", options: { color: C.red, fontSize: 14, bold: true } },
  ], { x: 0.5, y: 1.6, w: 4.4, h: 1.4, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.4 });

  addCard(s, 5.1, 1.1, 4.6, 2.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.6, h: 0.45, fill: { color: C.orange } });
  s.addText("発症時期", { x: 5.3, y: 1.1, w: 4.3, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "通常型: ", options: { color: C.text, fontSize: 13 } },
    { text: "投与開始後 5〜14日\n", options: { color: C.orange, fontSize: 14, bold: true } },
    { text: "迅速発症型: ", options: { color: C.text, fontSize: 13 } },
    { text: "数時間〜24時間\n", options: { color: C.orange, fontSize: 14, bold: true } },
    { text: "(過去30日以内のヘパリン曝露歴あり)", options: { color: C.sub, fontSize: 11 } },
  ], { x: 5.3, y: 1.6, w: 4.3, h: 1.4, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.4 });

  addCard(s, 0.3, 3.2, 9.4, 1.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.2, w: 9.4, h: 0.45, fill: { color: C.primary } });
  s.addText("病態 ― 用量依存ではなく免疫原性", { x: 0.5, y: 3.2, w: 9.2, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });
  s.addText([
    { text: "ヘパリン ＋ PF4 (血小板第4因子) 複合体に対するIgG抗体 → 血小板強力活性化\n", options: { color: C.text, fontSize: 13 } },
    { text: "大量トロンビン産生 → ", options: { color: C.text, fontSize: 13 } },
    { text: "血小板減少を伴いながら「強力な血栓傾向 (白色血栓)」", options: { color: C.red, fontSize: 13, bold: true } },
  ], { x: 0.5, y: 3.7, w: 9.2, h: 0.75, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.65, w: 9.4, h: 0.9, fill: { color: "FFF3CD" } });
  s.addText([
    { text: "⚠ 重要: ", options: { bold: true, color: C.red, fontSize: 14 } },
    { text: "HITは用量依存ではない → ", options: { color: C.text, fontSize: 13 } },
    { text: "DVT予防の低用量 (5,000 U皮下注) でも発症しうる\n", options: { color: C.red, fontSize: 13, bold: true } },
    { text: "→ 「低用量だから安全」は成り立たない。必ず血小板数をモニタリング", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.5, y: 4.65, w: 9.2, h: 0.9, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, 9);
})();

// ============================================================
// SLIDE 10: 4Tsスコア
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "4Tsスコア ― 検査に進む前のスクリーニング");

  var headers = ["項目", "2点", "1点", "0点"];
  var widths = [2.6, 2.4, 2.4, 2.0];

  // ヘッダー
  var xOff = 0.3;
  headers.forEach(function(h, i) {
    s.addShape(pres.shapes.RECTANGLE, { x: xOff, y: 1.15, w: widths[i], h: 0.45, fill: { color: C.primary } });
    s.addText(h, { x: xOff, y: 1.15, w: widths[i], h: 0.45, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    xOff += widths[i];
  });

  var rows = [
    ["Thrombocytopenia\n(血小板減少)", ">50%低下\n最低値≥20,000/μL", "30〜50%低下\nまたは10,000〜19,000", "<30%低下\nまたは<10,000"],
    ["Timing\n(発症時期)", "5〜10日、または\n過去30日以内曝露", ">10日、または\n過去30〜100日曝露", "≤4日\n(曝露歴なし)"],
    ["Thrombosis\n(血栓症)", "新規血栓/皮膚壊死\n急性全身反応", "血栓進行/皮膚紅斑\n血栓疑い", "なし"],
    ["oTher causes\n(他原因)", "他原因なし", "他原因の可能性", "他原因が明らか"],
  ];

  rows.forEach(function(row, i) {
    var yy = 1.6 + i * 0.68;
    var xOff2 = 0.3;
    row.forEach(function(cell, j) {
      s.addShape(pres.shapes.RECTANGLE, { x: xOff2, y: yy, w: widths[j], h: 0.68, fill: { color: j === 0 ? C.light : C.white }, line: { color: C.sub, width: 0.5 } });
      s.addText(cell, { x: xOff2 + 0.05, y: yy, w: widths[j] - 0.1, h: 0.68, fontSize: 10, fontFace: FONT_JP, color: j === 0 ? C.primary : C.text, bold: j === 0, align: j === 0 ? "left" : "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
      xOff2 += widths[j];
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.4, w: 9.4, h: 1.15, fill: { color: C.warmBg } });
  s.addText([
    { text: "0〜3点: ", options: { bold: true, color: C.green, fontSize: 14 } },
    { text: "低確率 (陰性的中率 99%)\n", options: { color: C.text, fontSize: 13 } },
    { text: "4〜5点: ", options: { bold: true, color: C.orange, fontSize: 14 } },
    { text: "中等度 → 抗PF4抗体検査へ\n", options: { color: C.text, fontSize: 13 } },
    { text: "6〜8点: ", options: { bold: true, color: C.red, fontSize: 14 } },
    { text: "高確率 → 検査と並行してヘパリン中止＋代替抗凝固", options: { color: C.text, fontSize: 13 } },
  ], { x: 0.6, y: 4.4, w: 8.8, h: 1.15, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.4 });

  addPageNum(s, 10);
})();

// ============================================================
// SLIDE 11: HIT対応とまとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "HIT疑い時の4ステップ ＆ 本日のまとめ");

  // HIT対応4ステップ
  addCard(s, 0.3, 1.1, 9.4, 1.95);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.4, h: 0.45, fill: { color: C.red } });
  s.addText("4Tsスコア≥4点で即実施 (検査結果を待たない)", { x: 0.5, y: 1.1, w: 9.2, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var actions = [
    "① すべてのヘパリン即時中止 (持続静注・皮下注・ルートフラッシュ・コートカテーテル)",
    "② アルガトロバン開始 0.7 μg/kg/min → APTT前値の1.5〜3倍に調整",
    "③ ワルファリン禁忌 (血小板≥150,000/μL回復まで) ― 皮膚壊死・四肢壊疽回避",
    "④ 血小板輸血は原則禁忌 (重篤な出血時のみ) ― 血栓助長を防ぐ",
  ];
  actions.forEach(function(a, i) {
    s.addText(a, { x: 0.5, y: 1.6 + i * 0.36, w: 9.2, h: 0.32, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  // まとめ
  addCard(s, 0.3, 3.2, 9.4, 2.3);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.2, w: 9.4, h: 0.45, fill: { color: C.primary } });
  s.addText("本日の Take Home Message", { x: 0.5, y: 3.2, w: 9.2, h: 0.45, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  var takeaway = [
    { t: "ルーチンのヘパリンブリッジングは不要", c: C.green },
    { t: "使うなら 1万〜1万5千 U/日 開始 → APTT 2〜2.5倍に微調整", c: C.accent },
    { t: "投与後 5〜14日 は血小板数を毎日確認", c: C.orange },
    { t: "4Tsスコア≥4点 → ヘパリン即中止＋アルガトロバン", c: C.red },
  ];
  takeaway.forEach(function(p, i) {
    var yy = 3.75 + i * 0.42;
    s.addShape(pres.shapes.OVAL, { x: 0.55, y: yy + 0.05, w: 0.25, h: 0.25, fill: { color: p.c } });
    s.addText(p.t, { x: 0.95, y: yy, w: 8.8, h: 0.4, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  addPageNum(s, 11);
})();

pres.writeFile({ fileName: "脳梗塞急性期のヘパリン_スライド.pptx" })
  .then(function(fn) { console.log("Saved: " + fn); })
  .catch(function(e) { console.error(e); });
