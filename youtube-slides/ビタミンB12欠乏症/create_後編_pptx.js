var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "ビタミンB12欠乏症【後編】治療・フォローアップ・特殊集団";

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

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("ビタミンB12欠乏症【後編】", {
    x: 0.6, y: 1.0, w: 8.8, h: 1.2,
    fontSize: 38, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.5, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("治療・フォローアップ・特殊集団のマネジメント", {
    x: 0.6, y: 2.8, w: 8.8, h: 0.6,
    fontSize: 22, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
  s.addText("早期治療が予後を決める ― 原因別の長期戦略", {
    x: 0.6, y: 3.8, w: 8.8, h: 0.6,
    fontSize: 18, fontFace: FJ, color: "B0BEC5", align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: 投与経路の比較（メタアナリシス）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "経口・舌下・筋注に有意差なし（2025メタアナリシス）");

  // エビデンスカード
  card(s, 0.5, 1.2, 9, 1.6, C.light);
  s.addText([
    { text: "16研究・6,098名の解析結果\n", options: { fontSize: 18, fontFace: FJ, color: C.text, bold: true } },
    { text: "血清B12上昇: ", options: { fontSize: 16, fontFace: FJ, color: C.sub } },
    { text: "+402.6 pg/mL", options: { fontSize: 16, fontFace: FE, color: C.primary, bold: true } },
    { text: "　ホモシステイン低下: ", options: { fontSize: 16, fontFace: FJ, color: C.sub } },
    { text: "-4.83 μmol/L", options: { fontSize: 16, fontFace: FE, color: C.primary, bold: true } },
  ], { x: 0.8, y: 1.3, w: 8.4, h: 1.4, valign: "middle", margin: [8, 12, 8, 12] });

  // 3経路カード
  var routes = [
    { title: "筋注", desc: "重症・急性神経症状\n吸収障害", color: C.red },
    { title: "高用量経口", desc: "1,000〜2,000μg/日\n軽〜中等症・維持", color: C.green },
    { title: "舌下", desc: "経口と同等\n嚥下困難時", color: C.primary },
  ];
  routes.forEach(function(r, i) {
    var xPos = 0.5 + i * 3.1;
    card(s, xPos, 3.1, 2.9, 2.2, C.white, r.color);
    s.addText(r.title, { x: xPos + 0.2, y: 3.2, w: 2.5, h: 0.5, fontSize: 20, fontFace: FJ, color: r.color, bold: true, margin: 0 });
    s.addText(r.desc, { x: xPos + 0.2, y: 3.8, w: 2.5, h: 1.2, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });

  // 注記
  s.addText("⚠ 異質性が高い（I²>80%）ため全例で完全同等とは言い切れない", {
    x: 0.5, y: 5.0, w: 9, h: 0.4, fontSize: 12, fontFace: FJ, color: C.sub, italic: true, margin: 0
  });
})();

// ============================================================
// SLIDE 3: 投与スケジュール
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "初期補充と維持療法の投与スケジュール");

  var tOpts = { x: 0.5, y: 1.2, w: 9, h: 0.01, colW: [2.5, 3.25, 3.25], rowH: [0.5, 0.7, 0.7], margin: [6,8,6,8], autoPage: false };
  var rows = [
    [
      { text: "フェーズ", options: { fontSize: 16, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary } } },
      { text: "筋注レジメン", options: { fontSize: 16, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary } } },
      { text: "経口レジメン", options: { fontSize: 16, fontFace: FJ, color: C.white, bold: true, fill: { color: C.primary } } },
    ],
    [
      { text: "初期補充\n（ローディング）", options: { fontSize: 14, fontFace: FJ, bold: true, fill: { color: C.light } } },
      { text: "1,000μg 筋注\n連日〜隔日 × 1〜2週間", options: { fontSize: 14, fontFace: FJ, fill: { color: C.white } } },
      { text: "1,000〜2,000μg/日\n経口", options: { fontSize: 14, fontFace: FJ, fill: { color: C.white } } },
    ],
    [
      { text: "維持療法", options: { fontSize: 14, fontFace: FJ, bold: true, fill: { color: C.light } } },
      { text: "1,000μg 筋注\n1〜3ヶ月ごと", options: { fontSize: 14, fontFace: FJ, fill: { color: C.white } } },
      { text: "1,000〜2,000μg/日\n経口を継続", options: { fontSize: 14, fontFace: FJ, fill: { color: C.white } } },
    ],
  ];
  s.addTable(rows, tOpts);

  // 推奨サマリー
  card(s, 0.5, 3.2, 9, 1.0, C.lightYellow, C.yellow);
  s.addText([
    { text: "Delphi 2024: ", options: { fontSize: 15, fontFace: FE, color: C.text, bold: true } },
    { text: "急性・重篤な症状 → 筋注を第一選択\n", options: { fontSize: 15, fontFace: FJ, color: C.text } },
    { text: "Am Fam Physician 2025: ", options: { fontSize: 15, fontFace: FE, color: C.text, bold: true } },
    { text: "経口は筋注に非劣性。重症例以外では経口も推奨", options: { fontSize: 15, fontFace: FJ, color: C.text } },
  ], { x: 0.7, y: 3.3, w: 8.6, h: 0.8, valign: "middle", margin: 0 });

  // 原因別戦略
  card(s, 0.5, 4.5, 9, 1.0, C.lightRed, C.red);
  s.addText([
    { text: "NICE 2024: ", options: { fontSize: 15, fontFace: FE, color: C.red, bold: true } },
    { text: "悪性貧血・胃全摘後 → lifelong筋注を推奨", options: { fontSize: 15, fontFace: FJ, color: C.text } },
  ], { x: 0.7, y: 4.6, w: 8.6, h: 0.7, valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 4: 治療開始時の注意点
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "治療開始時に押さえるべき4つの注意点");

  var items = [
    { icon: "🔍", title: "葉酸の同時評価", desc: "masking効果の回避\nB12欠乏が否定されるまで\n葉酸単独補充は避ける", color: C.primary },
    { icon: "⚡", title: "低カリウム血症", desc: "まれだが重症巨赤芽球性\n貧血の初期補充時に\n造血回復でK取り込み↑", color: C.yellow },
    { icon: "🩸", title: "鉄欠乏の合併", desc: "特にAIGでは鉄欠乏が\nB12欠乏に先行・合併\n鉄の同時評価が必要", color: C.accent },
    { icon: "⏰", title: "可逆性の限界", desc: "長期放置された神経障害\nは補充しても完全回復\nしない場合がある", color: C.red },
  ];
  items.forEach(function(item, i) {
    var xPos = 0.3 + i * 2.4;
    card(s, xPos, 1.2, 2.2, 3.5, C.white, item.color);
    s.addText(item.icon, { x: xPos, y: 1.3, w: 2.2, h: 0.6, fontSize: 28, align: "center", margin: 0 });
    s.addText(item.title, { x: xPos + 0.15, y: 2.0, w: 1.9, h: 0.5, fontSize: 15, fontFace: FJ, color: item.color, bold: true, align: "center", margin: 0 });
    s.addText(item.desc, { x: xPos + 0.15, y: 2.6, w: 1.9, h: 1.8, fontSize: 12, fontFace: FJ, color: C.text, align: "center", margin: 0 });
  });
})();

// ============================================================
// SLIDE 5: フォローアップ：改善の時間経過
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "治療反応のタイムライン：いつ何が改善するか");

  var timeline = [
    { time: "3〜5日", item: "網赤血球↑", note: "最も早い指標\nピーク7〜10日", color: C.green },
    { time: "6〜8週", item: "Hb改善", note: "正常化まで\n約8週間", color: C.primary },
    { time: "数週〜月", item: "MCV正常化\n末梢神経改善", note: "程度と期間に\n依存", color: C.accent },
    { time: "月〜年", item: "SCD改善", note: "進行例は\n不可逆的", color: C.red },
  ];
  timeline.forEach(function(t, i) {
    var xPos = 0.4 + i * 2.35;
    // Arrow connector
    if (i > 0) {
      s.addShape(pres.shapes.LINE, { x: xPos - 0.35, y: 2.8, w: 0.3, h: 0, line: { color: C.sub, width: 2, dashType: "dash" } });
    }
    card(s, xPos, 1.2, 2.15, 3.2, C.white, t.color);
    s.addText(t.time, { x: xPos, y: 1.3, w: 2.15, h: 0.5, fontSize: 18, fontFace: FE, color: t.color, bold: true, align: "center", margin: 0 });
    s.addText(t.item, { x: xPos + 0.1, y: 2.0, w: 1.95, h: 0.7, fontSize: 15, fontFace: FJ, color: C.text, bold: true, align: "center", margin: 0 });
    s.addText(t.note, { x: xPos + 0.1, y: 2.8, w: 1.95, h: 1.0, fontSize: 12, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  });

  // Bottom message
  card(s, 0.5, 4.7, 9, 0.7, C.lightGreen, C.green);
  s.addText("早期治療ほど可逆性が高い ─ 治療の遅れが予後を決める", {
    x: 0.7, y: 4.8, w: 8.6, h: 0.5, fontSize: 16, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0
  });
})();

// ============================================================
// SLIDE 6: 再検査と治療反応の評価
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "フォローアップ：再検査と長期管理");

  // 再検査
  card(s, 0.5, 1.2, 4.2, 2.5, C.white, C.primary);
  s.addText("再検査のポイント", { x: 0.7, y: 1.3, w: 3.8, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "• 治療開始後1〜2ヶ月でCBC再検\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "• バイオマーカーによる注射頻度の\n  「滴定」は推奨されない\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "• 臨床症状の改善を主体に評価", options: { fontSize: 13, fontFace: FJ, color: C.text, bold: true } },
  ], { x: 0.7, y: 1.8, w: 3.8, h: 1.8, valign: "top", margin: 0 });

  // 長期管理
  card(s, 5.3, 1.2, 4.2, 2.5, C.white, C.green);
  s.addText("長期モニタリング", { x: 5.5, y: 1.3, w: 3.8, h: 0.4, fontSize: 17, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText([
    { text: "• 原因持続 → 生涯の補充療法\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "• 薬剤性 → 年1回以上のB12評価\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "• AIG → 胃NETサーベイランス併施", options: { fontSize: 13, fontFace: FJ, color: C.text } },
  ], { x: 5.5, y: 1.8, w: 3.8, h: 1.8, valign: "top", margin: 0 });

  // 警告
  card(s, 0.5, 4.0, 9, 1.2, C.lightYellow, C.yellow);
  s.addText([
    { text: "⚠ 補充なしでB12 > 1,000 pg/mLが持続する場合\n", options: { fontSize: 15, fontFace: FJ, color: C.red, bold: true } },
    { text: "固形腫瘍や血液悪性腫瘍との関連が報告されている（因果関係は未確立）。精査を考慮。", options: { fontSize: 14, fontFace: FJ, color: C.text } },
  ], { x: 0.7, y: 4.1, w: 8.6, h: 1.0, valign: "middle", margin: 0 });
})();

// ============================================================
// SLIDE 7: 特殊集団：高齢者
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "特殊集団①：高齢者 ─ 最大のリスク集団");

  // 3つの柱
  var pillars = [
    { title: "胃酸分泌低下", desc: "食物結合B12\n吸収障害の主因", icon: "🔻" },
    { title: "食事摂取量減少", desc: "低栄養・独居\n食事の偏り", icon: "🍽" },
    { title: "ポリファーマシー", desc: "メトホルミン\n+PPI長期使用", icon: "💊" },
  ];
  pillars.forEach(function(p, i) {
    var xPos = 0.5 + i * 3.1;
    card(s, xPos, 1.2, 2.9, 2.2, C.white, C.primary);
    s.addText(p.icon + " " + p.title, { x: xPos + 0.15, y: 1.3, w: 2.6, h: 0.5, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(p.desc, { x: xPos + 0.15, y: 1.9, w: 2.6, h: 1.2, fontSize: 14, fontFace: FJ, color: C.text, margin: 0 });
  });

  // データ
  card(s, 0.5, 3.7, 9, 0.8, C.light);
  s.addText("高齢者のB12欠乏は報告により3〜43%と幅がある（厚労省eJIM）", {
    x: 0.7, y: 3.8, w: 8.6, h: 0.5, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0
  });

  // 推奨
  card(s, 0.5, 4.8, 9, 0.7, C.lightGreen, C.green);
  s.addText("75歳以上ではスクリーニングを検討すべき", {
    x: 0.7, y: 4.9, w: 8.6, h: 0.5, fontSize: 16, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0
  });
})();

// ============================================================
// SLIDE 8: 特殊集団：妊婦
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "特殊集団②：妊婦 ─ NTDリスクとの関連");

  // データカード3枚
  var data = [
    { label: "妊婦B12欠乏\n世界プール有病率", value: "17.1%", sub: "95%CI 8.8-30.6%", color: C.primary },
    { label: "母体B12低値と\nNTDリスク", value: "SMD -0.23", sub: "p<0.001\n葉酸に差がなくても有意", color: C.red },
    { label: "B12補充による\n欠乏率低下", value: "OR 0.43", sub: "95%CI 0.19-0.95\n産後投与で母乳B12も改善", color: C.green },
  ];
  data.forEach(function(d, i) {
    var xPos = 0.4 + i * 3.15;
    card(s, xPos, 1.2, 2.95, 2.8, C.white, d.color);
    s.addText(d.label, { x: xPos + 0.15, y: 1.3, w: 2.65, h: 0.7, fontSize: 13, fontFace: FJ, color: C.sub, margin: 0 });
    s.addText(d.value, { x: xPos + 0.15, y: 2.1, w: 2.65, h: 0.6, fontSize: 24, fontFace: FE, color: d.color, bold: true, align: "center", margin: 0 });
    s.addText(d.sub, { x: xPos + 0.15, y: 2.7, w: 2.65, h: 1.0, fontSize: 12, fontFace: FJ, color: C.sub, align: "center", margin: 0 });
  });

  card(s, 0.5, 4.3, 9, 0.8, C.lightYellow, C.yellow);
  s.addText("高リスク妊婦への周産期B12スクリーニング導入が提言されている", {
    x: 0.7, y: 4.4, w: 8.6, h: 0.6, fontSize: 15, fontFace: FJ, color: C.text, bold: true, valign: "middle", margin: 0
  });
})();

// ============================================================
// SLIDE 9: 特殊集団：胃切除後・IBD + 予防・スクリーニング
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "胃切除後・IBD ＋ 予防とスクリーニング");

  // 左：胃切除後・IBD
  card(s, 0.3, 1.2, 4.6, 2.2, C.white, C.primary);
  s.addText("胃全摘後・IBD", { x: 0.5, y: 1.3, w: 4.2, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "• 胃全摘後: 術後4〜5年で必発\n  → lifelong筋注（NICE 2024）\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "• IBD（回腸病変/切除後）:\n  → B12モニタリング推奨（AGA 2024）", options: { fontSize: 13, fontFace: FJ, color: C.text } },
  ], { x: 0.5, y: 1.8, w: 4.2, h: 1.4, valign: "top", margin: 0 });

  // 右：ハイリスク群
  card(s, 5.1, 1.2, 4.6, 2.2, C.white, C.green);
  s.addText("予防的補充を考慮すべき群", { x: 5.3, y: 1.3, w: 4.2, h: 0.4, fontSize: 17, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText([
    { text: "✅ ヴィーガン・厳格ベジタリアン\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "✅ 75歳以上の高齢者\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "✅ 胃切除後・肥満外科手術後\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "✅ メトホルミン 4ヶ月以上\n", options: { fontSize: 13, fontFace: FJ, color: C.text } },
    { text: "✅ PPI/H2RA 12ヶ月以上", options: { fontSize: 13, fontFace: FJ, color: C.text } },
  ], { x: 5.3, y: 1.8, w: 4.2, h: 1.4, valign: "top", margin: 0 });

  // メトホルミン/PPIモニタリング
  card(s, 0.3, 3.7, 9.4, 1.6, C.white, C.accent);
  s.addText("メトホルミン/PPI使用者のモニタリング（糖尿病標準診療マニュアル2025）", {
    x: 0.5, y: 3.8, w: 9, h: 0.4, fontSize: 15, fontFace: FJ, color: C.accent, bold: true, margin: 0
  });
  s.addText([
    { text: "① メトホルミン開始後、少なくとも年1回の血清B12測定\n", options: { fontSize: 14, fontFace: FJ, color: C.text } },
    { text: "② PPI併用例ではより頻回のモニタリングを考慮\n", options: { fontSize: 14, fontFace: FJ, color: C.text } },
    { text: "③ 原因不明の貧血・神経障害・認知機能低下 → 速やかにB12評価", options: { fontSize: 14, fontFace: FJ, color: C.text } },
  ], { x: 0.5, y: 4.3, w: 9, h: 0.9, valign: "top", margin: 0 });
})();

// ============================================================
// SLIDE 10: 後編まとめ
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("4つのTake-home Message", { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 26, fontFace: FJ, color: C.white, bold: true, margin: 0 });

  var msgs = [
    { num: "1", text: "治療は個別化\n重症→筋注 / 維持→経口も可", color: C.primary },
    { num: "2", text: "悪性貧血・胃全摘後は\nlifelong筋注（NICE 2024）", color: C.red },
    { num: "3", text: "早期治療が予後を決める\n進行した神経障害は不可逆的", color: C.accent },
    { num: "4", text: "ハイリスク群を把握し\n積極的にスクリーニング", color: C.green },
  ];
  msgs.forEach(function(m, i) {
    var xPos = 0.3 + i * 2.45;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 2.25, h: 2.8, fill: { color: "24466E" }, shadow: shd(), rectRadius: 0.1 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.2, w: 2.25, h: 0.08, fill: { color: m.color } });
    s.addText(m.num, { x: xPos, y: 1.4, w: 2.25, h: 0.6, fontSize: 28, fontFace: FE, color: m.color, bold: true, align: "center", margin: 0 });
    s.addText(m.text, { x: xPos + 0.1, y: 2.1, w: 2.05, h: 1.6, fontSize: 14, fontFace: FJ, color: C.white, align: "center", valign: "middle", margin: 0 });
  });

  // フッター
  s.addText("ご視聴ありがとうございました", {
    x: 0, y: 4.6, w: 10, h: 0.6,
    fontSize: 20, fontFace: FJ, color: C.accent, align: "center", margin: 0,
  });
})();

// ============================================================
// SAVE
// ============================================================
var outPath = "C:/Users/jsber/OneDrive/Documents/Claude_task/youtube-slides/ビタミンB12欠乏症/B12欠乏症_後編_slides.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX saved: " + outPath);
}).catch(function(e) {
  console.error("Error: " + e);
});
