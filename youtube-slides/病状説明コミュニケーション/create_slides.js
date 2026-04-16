var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "理想的な病状説明とは何か ― 脳神経内科医が考える「伝わる」コミュニケーション";

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
  teal: "17A2B8",
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

  s.addText("脳神経内科医の視点", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("理想的な病状説明とは何か", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.2,
    fontSize: 36, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 「伝わる」コミュニケーションの技術 ―", {
    x: 0.8, y: 2.5, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.LINE, { x: 3, y: 3.3, w: 4, h: 0, line: { color: C.accent, width: 2 } });

  // 3つのキーワード
  var keys = [
    { text: "SPIKESプロトコル", color: C.primary },
    { text: "3つの条件", color: C.green },
    { text: "実践セリフ例", color: C.orange },
  ];
  keys.forEach(function(k, i) {
    var xPos = 0.5 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 3.5, w: 2.8, h: 0.55, fill: { color: k.color } });
    s.addText(k.text, {
      x: xPos, y: 3.5, w: 2.8, h: 0.55,
      fontSize: 16, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  s.addText("医知創造ラボ  今村久司", {
    x: 0.8, y: 4.3, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT_JP, color: C.light, align: "center", margin: 0,
  });
  s.addText("脳神経内科専門医", {
    x: 0.8, y: 4.85, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", italic: true, margin: 0,
  });
})();

// ============================================================
// SLIDE 2: なぜ病状説明は難しいのか？
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "なぜ病状説明は難しいのか？");

  var challenges = [
    { icon: "1", text: "正確な医学情報を伝えなければならない", detail: "ガイドライン・エビデンスに基づく内容" },
    { icon: "2", text: "患者さん・ご家族は極度の不安の中にいる", detail: "感情的な受容と心理的サポートが必要" },
    { icon: "3", text: "「今後どうなるのか」という見通しを求められる", detail: "不確実性の中でも誠実に答える必要がある" },
    { icon: "4", text: "限られた時間の中で行わなければならない", detail: "特に急性期は一刻を争う状況" },
  ];

  challenges.forEach(function(c, i) {
    var yPos = 1.1 + i * 1.05;
    addCard(s, 0.4, yPos, 9.2, 0.9);
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: yPos + 0.18, w: 0.54, h: 0.54, fill: { color: C.red } });
    s.addText(c.icon, { x: 0.6, y: yPos + 0.18, w: 0.54, h: 0.54, fontSize: 18, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.text, { x: 1.35, y: yPos + 0.05, w: 7.5, h: 0.45, fontSize: 16, fontFace: FONT_JP, color: C.text, bold: true, valign: "middle", margin: 0 });
    s.addText(c.detail, { x: 1.35, y: yPos + 0.48, w: 7.5, h: 0.35, fontSize: 13, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.15, w: 9.2, h: 0.4, fill: { color: C.light } });
  s.addText("→ エビデンスに基づいた「型」を学ぶことで、及第点の説明ができるようになる", {
    x: 0.6, y: 5.15, w: 8.8, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "2/9");
})();

// ============================================================
// SLIDE 3: SPIKESプロトコル概要
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "SPIKESプロトコル ― 6つのステップ");

  var steps = [
    { letter: "S", name: "Setting", ja: "状況の設定", desc: "静かな環境、家族同席、座って話す", color: C.primary },
    { letter: "P", name: "Perception", ja: "認識の確認", desc: "相手の理解度・認識のズレを把握", color: C.teal },
    { letter: "I", name: "Invitation", ja: "情報提供への招待", desc: "どこまで知りたいかを確認する", color: C.green },
    { letter: "K", name: "Knowledge", ja: "知識の提供", desc: "専門用語を避け、少しずつ伝える", color: C.orange },
    { letter: "E", name: "Emotions", ja: "感情への共感", desc: "感情を受け止め、言葉で寄り添う", color: C.red },
    { letter: "S", name: "Strategy", ja: "戦略と要約", desc: "治療計画を話し合い、要約する", color: C.dark },
  ];

  steps.forEach(function(st, i) {
    var yPos = 1.05 + i * 0.72;
    addCard(s, 0.3, yPos, 9.4, 0.62);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.65, h: 0.62, fill: { color: st.color } });
    s.addText(st.letter, { x: 0.3, y: yPos, w: 0.65, h: 0.62, fontSize: 24, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(st.name, { x: 1.1, y: yPos + 0.02, w: 1.7, h: 0.3, fontSize: 14, fontFace: FONT_EN, color: st.color, bold: true, valign: "middle", margin: 0 });
    s.addText(st.ja, { x: 1.1, y: yPos + 0.32, w: 1.7, h: 0.26, fontSize: 12, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 0 });
    s.addText(st.desc, { x: 3.0, y: yPos, w: 6.5, h: 0.62, fontSize: 14, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
  });

  s.addText("Baile WF, et al. Oncologist. 2000;5(4):302-311.", {
    x: 0.3, y: 5.35, w: 9.4, h: 0.3,
    fontSize: 10, fontFace: FONT_EN, color: C.sub, italic: true, margin: 0,
  });

  addPageNum(s, "3/9");
})();

// ============================================================
// SLIDE 4: SPIKES詳細 S・P・I
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "SPIKES 前半 ― 準備と確認のステップ");

  var items = [
    {
      letter: "S", name: "Setting（状況の設定）", color: C.primary,
      points: [
        "プライバシーが保たれる静かな環境を用意",
        "ご家族の同席を促し、全員が座れるようにする",
        "時間的な制約や中断を最小限にする",
      ],
    },
    {
      letter: "P", name: "Perception（認識の確認）", color: C.teal,
      points: [
        "「今のご本人の状態をどのようにお聞きですか？」",
        "相手の理解度と認識のズレを把握する",
        "医学知識のレベルに合わせた説明の出発点にする",
      ],
    },
    {
      letter: "I", name: "Invitation（情報提供への招待）", color: C.green,
      points: [
        "「詳しい検査結果をお伝えしてよろしいですか？」",
        "詳細な予後を聞きたくない人もいることに配慮",
        "情報の受け入れ態勢があるかを確認する",
      ],
    },
  ];

  items.forEach(function(it, i) {
    var yPos = 1.05 + i * 1.5;
    addCard(s, 0.3, yPos, 9.4, 1.35);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.6, h: 1.35, fill: { color: it.color } });
    s.addText(it.letter, { x: 0.3, y: yPos, w: 0.6, h: 0.6, fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(it.name, { x: 1.1, y: yPos + 0.05, w: 8.4, h: 0.35, fontSize: 16, fontFace: FONT_JP, color: it.color, bold: true, valign: "middle", margin: 0 });
    it.points.forEach(function(p, j) {
      s.addText("・" + p, { x: 1.1, y: yPos + 0.42 + j * 0.3, w: 8.4, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
    });
  });

  addPageNum(s, "4/9");
})();

// ============================================================
// SLIDE 5: SPIKES詳細 K・E・S
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "SPIKES 後半 ― 伝達・共感・計画のステップ");

  var items = [
    {
      letter: "K", name: "Knowledge（知識の提供）", color: C.orange,
      points: [
        "専門用語を避け、わかりやすい言葉で伝える",
        "一度に大量の情報を与えず、理解度を確認しながら",
        "個別の状況に合わせてデータを「翻訳」する",
      ],
    },
    {
      letter: "E", name: "Emotions（感情への共感）", color: C.red,
      points: [
        "感情的反応（悲しみ・怒り・沈黙）を受け止める",
        "「お辛いですね」「突然のことで驚かれますよね」",
        "適度な間（ポーズ）を取り、感情の揺れに寄り添う",
      ],
    },
    {
      letter: "S", name: "Strategy & Summary（戦略と要約）", color: C.dark,
      points: [
        "今後の治療方針・計画を一緒に話し合う",
        "説明内容を3点程度に要約して確認する",
        "質問の機会を保証し、次回面談の予定を立てる",
      ],
    },
  ];

  items.forEach(function(it, i) {
    var yPos = 1.05 + i * 1.5;
    addCard(s, 0.3, yPos, 9.4, 1.35);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.6, h: 1.35, fill: { color: it.color } });
    s.addText(it.letter, { x: 0.3, y: yPos, w: 0.6, h: 0.6, fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(it.name, { x: 1.1, y: yPos + 0.05, w: 8.4, h: 0.35, fontSize: 16, fontFace: FONT_JP, color: it.color, bold: true, valign: "middle", margin: 0 });
    it.points.forEach(function(p, j) {
      s.addText("・" + p, { x: 1.1, y: yPos + 0.42 + j * 0.3, w: 8.4, h: 0.3, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0 });
    });
  });

  addPageNum(s, "5/9");
})();

// ============================================================
// SLIDE 6: 理想的な病状説明の3つの条件
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "理想的な病状説明の3つの条件");

  var conditions = [
    {
      num: "1", title: "情報の透明性と正確性",
      sub: "エビデンスに基づく説明",
      points: "最新のガイドラインに基づいた正確な内容\n個別の状況に合わせてデータを「翻訳」\n不確実性も誠実に伝え、信頼関係を構築",
      color: C.primary,
    },
    {
      num: "2", title: "心理的安全性と共感的な態度",
      sub: "感情に寄り添う姿勢",
      points: "相手の目を見て話し、適度な間を取る\n「質問が思い浮かばなくても大丈夫」\n後で聞ける逃げ道を用意する",
      color: C.green,
    },
    {
      num: "3", title: "生活を見据えた見通しの提示",
      sub: "退院後の生活まで見せる",
      points: "後遺症との付き合い・再発予防を説明\n退院先（自宅 / リハビリ病院 / 施設）\n早期から多職種（MSW・リハビリ）と連携",
      color: C.orange,
    },
  ];

  conditions.forEach(function(c, i) {
    var xPos = 0.2 + i * 3.27;
    addCard(s, xPos, 1.05, 3.1, 4.4);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.05, w: 3.1, h: 0.7, fill: { color: c.color } });
    s.addShape(pres.shapes.OVAL, { x: xPos + 0.1, y: 1.12, w: 0.55, h: 0.55, fill: { color: C.white } });
    s.addText(c.num, { x: xPos + 0.1, y: 1.12, w: 0.55, h: 0.55, fontSize: 22, fontFace: FONT_EN, color: c.color, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(c.title, { x: xPos + 0.7, y: 1.05, w: 2.3, h: 0.7, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
    s.addText(c.sub, { x: xPos + 0.15, y: 1.85, w: 2.8, h: 0.35, fontSize: 12, fontFace: FONT_JP, color: c.color, bold: true, valign: "middle", margin: 0 });
    s.addText(c.points, { x: xPos + 0.15, y: 2.25, w: 2.8, h: 3.0, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "top", margin: 0, lineSpacingMultiple: 1.4 });
  });

  addPageNum(s, "6/9");
})();

// ============================================================
// SLIDE 7: 実践サンプル 前半（導入～診断伝達～治療方針）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "【実践】急性期脳梗塞の病状説明（前半）");

  s.addText("状況：70代男性、右半身麻痺＋失語症で搬送 → アテローム血栓性脳梗塞と診断。妻と長男に初回説明。", {
    x: 0.3, y: 0.95, w: 9.4, h: 0.4,
    fontSize: 12, fontFace: FONT_JP, color: C.sub, italic: true, margin: 0,
  });

  var sections = [
    {
      label: "導入と認識の確認", tag: "S・P", tagColor: C.primary,
      script: "「本日は急なことで大変驚かれたことと思います。私は脳神経内科の〇〇です。まずは、今ご本人の状態について、どのようにお聞きになっていますか？」",
    },
    {
      label: "病状と診断の伝達", tag: "K", tagColor: C.orange,
      script: "「MRI検査の結果、脳の左側の血管が詰まっており『脳梗塞』を起こしている状態です。アテローム血栓性脳梗塞といって、動脈硬化で血管が狭くなり血栓が詰まったタイプです。」",
    },
    {
      label: "治療方針の提示", tag: "K・S", tagColor: C.green,
      script: "「現在は『急性期』で非常に不安定な時期です。点滴で悪化を防ぐ治療を開始しています。最初の数日間は集中治療室（SCU）で厳重に管理します。」",
    },
  ];

  sections.forEach(function(sec, i) {
    var yPos = 1.45 + i * 1.3;
    addCard(s, 0.3, yPos, 9.4, 1.15);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.1, h: 1.15, fill: { color: sec.tagColor } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.55, y: yPos + 0.08, w: 0.55, h: 0.3, fill: { color: sec.tagColor } });
    s.addText(sec.tag, { x: 0.55, y: yPos + 0.08, w: 0.55, h: 0.3, fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(sec.label, { x: 1.2, y: yPos + 0.05, w: 3.5, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: sec.tagColor, bold: true, valign: "middle", margin: 0 });
    s.addText(sec.script, { x: 0.55, y: yPos + 0.42, w: 9.0, h: 0.68, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });
  });

  addPageNum(s, "7/9");
})();

// ============================================================
// SLIDE 8: 実践サンプル 後半（共感～見通し～要約）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "【実践】急性期脳梗塞の病状説明（後半）");

  var sections = [
    {
      label: "共感と不確実性の伝達", tag: "E・K", tagColor: C.red,
      script: "「本当に突然のことで、ご不安ですよね。正直に申し上げると、現時点では麻痺や言葉の障害がどこまで回復するか予測は困難です。しかし早期リハビリは回復に非常に重要です。」",
    },
    {
      label: "再発予防と今後の見通し", tag: "S", tagColor: C.dark,
      script: "「脳梗塞は再発しやすい病気です。原因を調べ再発予防の治療も行います。約1〜2週間後には回復期リハビリ病院への転院が必要になる可能性があります。詳細はMSWからもご案内します。」",
    },
    {
      label: "要約と質問の確認", tag: "S", tagColor: C.primary,
      script: "「要約すると①左脳の脳梗塞で右半身麻痺と言葉の障害、②点滴で悪化を防ぐ治療中、③明日からリハビリ開始、の3点です。質問が思い浮かばなくても、いつでもメモして聞いてくださいね。」",
    },
  ];

  sections.forEach(function(sec, i) {
    var yPos = 1.1 + i * 1.3;
    addCard(s, 0.3, yPos, 9.4, 1.15);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 0.1, h: 1.15, fill: { color: sec.tagColor } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.55, y: yPos + 0.08, w: 0.55, h: 0.3, fill: { color: sec.tagColor } });
    s.addText(sec.tag, { x: 0.55, y: yPos + 0.08, w: 0.55, h: 0.3, fontSize: 11, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(sec.label, { x: 1.2, y: yPos + 0.05, w: 3.5, h: 0.35, fontSize: 14, fontFace: FONT_JP, color: sec.tagColor, bold: true, valign: "middle", margin: 0 });
    s.addText(sec.script, { x: 0.55, y: yPos + 0.42, w: 9.0, h: 0.68, fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.05, w: 9.4, h: 0.5, fill: { color: C.light } });
  s.addText("ポイント：「型」を意識することで、パターナリズム（一方的説明）を防ぎ、協働する姿勢を築ける", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.5,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  addPageNum(s, "8/9");
})();

// ============================================================
// SLIDE 9: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.2, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "病状説明は「情報伝達」ではなく「対話のプロセス」\nSPIKESプロトコルの「型」を活用しよう" },
    { num: "2", text: "まず相手の理解度を確認（Perception）してから説明を始める\n一方通行の説明にならないための最初の一歩" },
    { num: "3", text: "感情に寄り添う（Emotions）＋不確実性を正直に伝える\n「分からないこと」を伝えることも信頼構築の一部" },
    { num: "4", text: "病名・治療だけでなく「生活の見通し」まで伝える\n退院後の生活・リハビリ・再発予防 → 多職種連携" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.05 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.2, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.2, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  ｜  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.3, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "9/9");
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/病状説明コミュニケーション_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
