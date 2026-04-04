var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "けいれんを目撃したときの正しい対応 ― 脳神経内科医が解説【からだの不思議 #16】";

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

function addPageNum(s, num, total) {
  s.addText(num + "/" + total, { x: 9.0, y: 5.2, w: 0.8, h: 0.4, fontSize: 11, fontFace: FONT_B, color: C.sub, align: "right", margin: 0 });
}

var TOTAL = 10;

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("脳神経内科医が答える からだの不思議 #16", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("けいれんを目撃したとき\nどうすればいい？", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 正しい対応と絶対にやってはいけないこと ―", {
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
// SLIDE 2: けいれんとは ― 脳の電気的な嵐
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "けいれんとは ― 脳の電気的な嵐");

  // メイン説明カード
  addCard(s, 0.5, 1.1, 9.0, 1.5);
  s.addText("けいれん発作（全般発作）のとき、脳では", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.45,
    fontSize: 16, fontFace: FONT_JP, color: C.text, margin: 0,
  });
  s.addText("神経細胞の電気的な嵐（異常放電）が広範囲に起きている", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.7,
    fontSize: 20, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });

  // 3つのフェーズカード
  var phases = [
    { name: "強直期", sub: "tonic phase", desc: "全身硬直\n呼吸停止のことも", duration: "30〜60秒", color: C.red },
    { name: "間代期", sub: "clonic phase", desc: "規則的な\nガクガク収縮", duration: "30〜60秒", color: C.orange },
    { name: "発作後期", sub: "postictal phase", desc: "深い眠り\n錯乱状態", duration: "数分〜30分", color: C.accent },
  ];
  phases.forEach(function(p, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 2.85, 3.0, 2.4);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.85, w: 3.0, h: 0.55, fill: { color: p.color } });
    s.addText(p.name, { x: xPos, y: 2.85, w: 3.0, h: 0.35, fontSize: 18, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.sub, { x: xPos, y: 3.2, w: 3.0, h: 0.2, fontSize: 10, fontFace: FONT_EN, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText(p.desc, { x: xPos, y: 3.4, w: 3.0, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.text, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.3, y: 4.4, w: 2.4, h: 0.35, fill: { color: C.light } });
    s.addText(p.duration, { x: xPos + 0.3, y: 4.4, w: 2.4, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  });

  // 矢印
  s.addText("→", { x: 3.1, y: 3.5, w: 0.4, h: 0.5, fontSize: 22, color: C.sub, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 6.3, y: 3.5, w: 0.4, h: 0.5, fontSize: 22, color: C.sub, align: "center", valign: "middle", margin: 0 });

  // 重要メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.05, w: 9.0, h: 0.45, fill: { color: C.light } });
  s.addText("ほとんどの発作は 2〜3分以内に自然に止まります", {
    x: 0.8, y: 5.05, w: 8.4, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2", TOTAL);
})();

// ============================================================
// SLIDE 3: 正しい対応 ― 7つのこと
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "発作中にやるべきこと");

  var actions = [
    { icon: "1", text: "時間を確認する", desc: "発作が始まった時刻を記録。5分超えたら救急車" },
    { icon: "2", text: "周囲の危険物を除く", desc: "テーブルの角・椅子の脚・硬い障害物を遠ざける" },
    { icon: "3", text: "頭を守る", desc: "折りたたんだ上着やカバンを頭の下に置く" },
    { icon: "4", text: "横向きに寝かせる（回復体位）", desc: "嘔吐時に気道がふさがらないよう顔を横に向ける" },
    { icon: "5", text: "体を締め付けるものを外す", desc: "眼鏡・ネクタイ・ベルトなど" },
    { icon: "6", text: "発作の様子を観察・記録する", desc: "どこが震えたか、眼の向き、意識の有無" },
    { icon: "7", text: "声をかけ続ける", desc: "発作後「大丈夫ですよ」と安心させる" },
  ];

  actions.forEach(function(a, i) {
    var col = i < 4 ? 0 : 1;
    var row = i < 4 ? i : i - 4;
    var xPos = 0.3 + col * 4.85;
    var yPos = 1.05 + row * 1.05;
    if (col === 1 && row === 3) return; // 7項目 → col1に3つのみ
    addCard(s, xPos, yPos, 4.55, 0.9);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.1, y: yPos + 0.1, w: 0.7, h: 0.7, fill: { color: C.green } });
    s.addText(a.icon, { x: xPos + 0.1, y: yPos + 0.1, w: 0.7, h: 0.7, fontSize: 16, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(a.text, { x: xPos + 0.9, y: yPos + 0.05, w: 3.5, h: 0.4, fontSize: 14, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(a.desc, { x: xPos + 0.9, y: yPos + 0.45, w: 3.5, h: 0.4, fontSize: 11, fontFace: FONT_JP, color: C.sub, valign: "top", margin: 0 });
  });

  // 7番目（col1, row3）は幅広で下に配置
  addCard(s, 0.3, 5.0, 9.4, 0.55);
  s.addShape(pres.shapes.OVAL, { x: 0.4, y: 5.05, w: 0.45, h: 0.45, fill: { color: C.green } });
  s.addText("7", { x: 0.4, y: 5.05, w: 0.45, h: 0.45, fontSize: 13, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("声をかけ続ける", { x: 0.95, y: 5.05, w: 2.8, h: 0.25, fontSize: 13, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
  s.addText("発作後「大丈夫ですよ、そこにいますよ」と安心させる", { x: 0.95, y: 5.3, w: 8.5, h: 0.2, fontSize: 11, fontFace: FONT_JP, color: C.sub, valign: "top", margin: 0 });

  addPageNum(s, "3", TOTAL);
})();

// ============================================================
// SLIDE 4: やってはいけないこと（赤カード）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "やってはいけないこと");

  var donts = [
    { text: "口の中に何かを入れない", desc: "タオルや指を入れるのは最も危険。窒息・骨折のリスク" },
    { text: "体を押さえつけない", desc: "無理に動きを止めようとすると骨折・脱臼を引き起こす" },
    { text: "水を飲ませない", desc: "意識混濁中の飲水は誤嚥を引き起こす" },
    { text: "一人にしない", desc: "回復するまでそばに寄り添う" },
    { text: "仰向けのままにしない", desc: "嘔吐物が気道に流れ込まないよう横向きに保持" },
    { text: "発作中に食べ物・薬を口に入れない", desc: "意識のない状態での経口投与は窒息死につながる" },
  ];

  donts.forEach(function(d, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var xPos = 0.3 + col * 4.85;
    var yPos = 1.05 + row * 1.4;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 4.55, h: 1.25, fill: { color: "FFF5F5" }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 4.55, h: 0.45, fill: { color: C.red } });
    s.addText("×  " + d.text, {
      x: xPos + 0.1, y: yPos, w: 4.35, h: 0.45,
      fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
    });
    s.addText(d.desc, {
      x: xPos + 0.1, y: yPos + 0.5, w: 4.35, h: 0.65,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.15,
    });
  });

  addPageNum(s, "4", TOTAL);
})();

// ============================================================
// SLIDE 5: 救急車を呼ぶ基準 ― 5分ルール
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "救急車を呼ぶべき状況 ― 5分ルール");

  // 5分ルール強調
  addCard(s, 0.5, 1.05, 9.0, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.05, w: 2.5, h: 1.1, fill: { color: C.red } });
  s.addText("5分ルール", { x: 0.5, y: 1.05, w: 2.5, h: 1.1, fontSize: 26, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("発作が 5分以上 続いたら即119番", {
    x: 3.1, y: 1.05, w: 6.2, h: 1.1,
    fontSize: 22, fontFace: FONT_JP, color: C.red, bold: true, valign: "middle", margin: 0,
  });

  // 救急を呼ぶべき9項目
  var criteria = [
    "発作が5分以上続いている（てんかん重積状態の可能性）",
    "発作が止まっても意識が戻らないまま次の発作が始まる",
    "発作後、数分経っても意識が戻らない・呼びかけに反応しない",
    "呼吸が回復しない・チアノーゼ（唇や爪が紫色）がある",
    "水の中（プール・浴槽）で発作が起きた",
    "頭部を打撲した後に発作が起きた",
    "妊娠中に発作が起きた（子癇の可能性）",
    "初めての発作と思われる（てんかんの診断がない人）",
    "発作後に麻痺・ろれつが回らないなど新たな神経症状がある",
  ];

  var yStart = 2.35;
  var itemH = 0.37;
  criteria.forEach(function(c, i) {
    var bgColor = (i % 2 === 0) ? "FFF5F5" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yStart + i * itemH, w: 9.0, h: itemH, fill: { color: bgColor } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yStart + i * itemH, w: 0.3, h: itemH, fill: { color: C.red } });
    s.addText("!", { x: 0.5, y: yStart + i * itemH, w: 0.3, h: itemH, fontSize: 12, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c, {
      x: 0.9, y: yStart + i * itemH, w: 8.5, h: itemH,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  addPageNum(s, "5", TOTAL);
})();

// ============================================================
// SLIDE 6: 救急隊員に伝える情報表
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "救急隊員・医師への情報提供");

  // リード文
  s.addText("目撃者の証言は診断に不可欠。次の情報をできる限り伝えてください。", {
    x: 0.5, y: 1.0, w: 9.0, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // テーブル
  var rows = [
    { item: "発作の始まり方", detail: "突然か、体の一部から始まったか。どの部位から始まったか" },
    { item: "発作の持続時間", detail: "発作が始まった時刻・止まった時刻" },
    { item: "体の動き", detail: "全身か、右半身か左半身か。硬直→ガクガクの順か" },
    { item: "眼の動き", detail: "眼が白目を向いたか、どちらを向いていたか" },
    { item: "意識の有無", detail: "声をかけて反応があったか、呼びかけに気づいていたか" },
    { item: "発作後の状態", detail: "すぐに意識が戻ったか、麻痺や言語障害はないか" },
    { item: "既往歴", detail: "てんかんや糖尿病の持病、内服薬（わかれば）" },
  ];

  // ヘッダー
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.55, w: 2.8, h: 0.45, fill: { color: C.primary } });
  s.addShape(pres.shapes.RECTANGLE, { x: 3.2, y: 1.55, w: 6.4, h: 0.45, fill: { color: C.primary } });
  s.addText("確認項目", { x: 0.4, y: 1.55, w: 2.8, h: 0.45, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("具体的に観察・記録するポイント", { x: 3.2, y: 1.55, w: 6.4, h: 0.45, fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 2.0 + i * 0.49;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yPos, w: 9.2, h: 0.49, fill: { color: bgColor } });
    s.addText(r.item, { x: 0.5, y: yPos, w: 2.6, h: 0.49, fontSize: 12, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: [0, 4, 0, 4] });
    s.addText(r.detail, { x: 3.2, y: yPos, w: 6.2, h: 0.49, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 4, 0, 4] });
  });

  addPageNum(s, "6", TOTAL);
})();

// ============================================================
// SLIDE 7: よくある誤解 3つ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "よくある誤解を解く");

  var myths = [
    {
      q: "「舌を飲み込む」のを防ぐために口に何かを入れるべき？",
      a: "誤りです",
      detail: "舌は口腔底に固定されており飲み込まれることはありません。口に物を入れると歯の骨折・異物誤嚥・咬傷のリスクがあります。",
      color: C.red,
    },
    {
      q: "体を押さえつけると発作が短くなる？",
      a: "誤りです",
      detail: "発作は脳内の電気的現象。体を固定しても止まりません。むしろ肩関節脱臼・肋骨骨折などの二次的損傷が生じます。",
      color: C.orange,
    },
    {
      q: "けいれん中に水を飲ませるとよい？",
      a: "誤りです",
      detail: "意識のない状態での飲水は誤嚥の危険があります。完全に意識が回復してから本人が希望する場合にのみ行ってください。",
      color: C.primary,
    },
  ];

  myths.forEach(function(m, i) {
    var yPos = 1.1 + i * 1.5;
    addCard(s, 0.5, yPos, 9.0, 1.35);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 9.0, h: 0.5, fill: { color: m.color } });
    s.addText("Q" + (i + 1) + ": " + m.q, {
      x: 0.7, y: yPos, w: 7.5, h: 0.5,
      fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0,
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 8.4, y: yPos + 0.05, w: 1.0, h: 0.4, fill: { color: C.yellow } });
    s.addText(m.a, {
      x: 8.4, y: yPos + 0.05, w: 1.0, h: 0.4,
      fontSize: 12, fontFace: FONT_JP, color: C.dark, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(m.detail, {
      x: 0.7, y: yPos + 0.55, w: 8.6, h: 0.75,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.25,
    });
  });

  addPageNum(s, "7", TOTAL);
})();

// ============================================================
// SLIDE 8: 発作後の対応と回復体位
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "発作後の対応 ― 発作後朦朧状態（Postictal state）");

  // 説明カード
  addCard(s, 0.5, 1.1, 9.0, 1.3);
  s.addText("発作が終わると脳は「発作後朦朧状態（postictal state）」に入ります", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.5,
    fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, margin: 0,
  });
  s.addText("本人はぼんやり・眠る・何が起きたか覚えていない ― これは正常な経過です", {
    x: 0.8, y: 1.7, w: 8.4, h: 0.6,
    fontSize: 15, fontFace: FONT_JP, color: C.text, margin: 0,
  });

  // 発作後の対応3ステップ
  var steps = [
    { num: "1", title: "回復体位を保つ", desc: "横向きのまま安静。嘔吐に備える", color: C.green },
    { num: "2", title: "意識の戻りを確認", desc: "声かけを続けながら反応を観察", color: C.accent },
    { num: "3", title: "慌てて揺り起こさない", desc: "意識回復を焦らず待つ", color: C.primary },
  ];
  steps.forEach(function(st, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 2.6, 3.0, 2.0);
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.85, y: 2.75, w: 1.3, h: 1.3, fill: { color: st.color } });
    s.addText(st.num, { x: xPos + 0.85, y: 2.75, w: 1.3, h: 1.3, fontSize: 36, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.title, { x: xPos, y: 4.15, w: 3.0, h: 0.4, fontSize: 15, fontFace: FONT_JP, color: C.text, bold: true, align: "center", margin: 0 });
    s.addText(st.desc, { x: xPos, y: 4.55, w: 3.0, h: 0.4, fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0 });
  });

  // 矢印
  s.addText("→", { x: 3.1, y: 3.2, w: 0.4, h: 0.5, fontSize: 22, color: C.sub, align: "center", valign: "middle", margin: 0 });
  s.addText("→", { x: 6.3, y: 3.2, w: 0.4, h: 0.5, fontSize: 22, color: C.sub, align: "center", valign: "middle", margin: 0 });

  // 下部メッセージ
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.12, w: 9.0, h: 0.45, fill: { color: C.light } });
  s.addText("意識が完全に戻るまでそばを離れない。一人にしない。", {
    x: 0.8, y: 5.12, w: 8.4, h: 0.45,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "8", TOTAL);
})();

// ============================================================
// SLIDE 9: てんかんと紛らわしい発作（鑑別）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "けいれんに見える、でもてんかんではないかもしれない");

  s.addText("目撃者にできることは変わりません。診断は医師が行います。", {
    x: 0.5, y: 1.0, w: 9.0, h: 0.45,
    fontSize: 15, fontFace: FONT_JP, color: C.sub, margin: 0,
  });

  var ddx = [
    { name: "失神（syncope）", desc: "脳への血流低下で起こる意識消失。四肢がぴくっと動くことがあり発作と間違われやすい", color: C.accent },
    { name: "低血糖発作", desc: "糖尿病患者で血糖が著しく低下したときに起こる。意識障害・発作様症状が現れる", color: C.orange },
    { name: "解離性（心因性）非てんかん発作（PNES）", desc: "脳波上の異常なく発作様症状が出る。ストレス関連のことが多い", color: C.primary },
    { name: "熱性けいれん", desc: "乳幼児に多い。高熱（38°C以上）に伴って起きるが、ほとんどは数分で自然に止まる", color: C.yellow },
  ];

  ddx.forEach(function(d, i) {
    var yPos = 1.6 + i * 0.96;
    addCard(s, 0.5, yPos, 9.0, 0.85);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yPos, w: 3.0, h: 0.85, fill: { color: d.color } });
    s.addText(d.name, {
      x: 0.6, y: yPos, w: 2.8, h: 0.85,
      fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: [0, 4, 0, 4], lineSpacingMultiple: 1.1,
    });
    s.addText(d.desc, {
      x: 3.65, y: yPos, w: 5.7, h: 0.85,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 4, 0, 4], lineSpacingMultiple: 1.1,
    });
  });

  // 底部ボックス
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.42, w: 9.0, h: 0.15, fill: { color: C.light } });

  addPageNum(s, "9", TOTAL);
})();

// ============================================================
// SLIDE 10: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "まとめ ― Take Home Message");

  var msgs = [
    { num: "1", text: "けいれん中、脳では神経細胞の電気的な嵐が起きている。発作は通常2〜3分で自然に止まる", color: C.primary },
    { num: "2", text: "やるべきことは「安全な環境・回復体位・時間計測・観察」の4つ", color: C.green },
    { num: "3", text: "口に何かを入れることは最も危険な誤った行動。絶対にしない", color: C.red },
    { num: "4", text: "5分以上続く・意識が戻らない・繰り返すなどの場合は即119番", color: C.red },
    { num: "5", text: "目撃者の「発作の様子」の情報は診断に不可欠。できる限りメモしておく", color: C.accent },
  ];

  msgs.forEach(function(m, i) {
    var yPos = 1.05 + i * 0.88;
    addCard(s, 0.5, yPos, 9.0, 0.78);
    s.addShape(pres.shapes.OVAL, { x: 0.65, y: yPos + 0.04, w: 0.7, h: 0.7, fill: { color: m.color } });
    s.addText(m.num, {
      x: 0.65, y: yPos + 0.04, w: 0.7, h: 0.7,
      fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(m.text, {
      x: 1.55, y: yPos, w: 7.8, h: 0.78,
      fontSize: 15, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 8, 0, 0], lineSpacingMultiple: 1.2,
    });
  });

  // チャンネル名
  s.addText("医知創造ラボ ― 脳神経内科医がAIで紡ぐ最新医療情報", {
    x: 0.5, y: 5.25, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });

  addPageNum(s, "10", TOTAL);
})();

// ============================================================
// OUTPUT
// ============================================================
var outPath = __dirname + "/けいれん対応_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("Created: " + outPath);
}).catch(function(err) {
  console.error(err);
});
