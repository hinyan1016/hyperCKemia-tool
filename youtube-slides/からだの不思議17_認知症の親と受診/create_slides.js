var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "認知症の親を病院に連れて行くには ― 受診を嫌がる場合の対処法";

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
  s.addText(num + "/" + total, { x: 9.0, y: 5.2, w: 0.8, h: 0.4, fontSize: 11, fontFace: FONT_EN, color: C.sub, align: "right", margin: 0 });
}

var TOTAL = "9";

// ============================================================
// SLIDE 1: タイトル
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("脳神経内科医が答える からだの不思議 #17", {
    x: 0.8, y: 0.5, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_JP, color: C.accent, align: "center", margin: 0,
  });

  s.addText("認知症の親を\n病院に連れて行くには", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.8,
    fontSize: 38, fontFace: FONT_JP, color: C.white, bold: true, align: "center", margin: 0,
    lineSpacingMultiple: 1.2,
  });

  s.addText("― 受診を嫌がるときの実践的アプローチ ―", {
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
// SLIDE 2: こんな経験、ありませんか？（家族の葛藤）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "こんな経験、ありませんか？");

  var quotes = [
    "「父の物忘れがひどい。もの忘れ外来に連れて行こうとしたら\n　『私はボケていない！』と頑なに拒否された」",
    "「かかりつけ医に相談してみたら、本人がいる前では\n　何も言えなくて……結局また次回に持ち越し」",
    "「どう説得すればいいのか、もう何が正解かわからない。\n　家族みんなが疲弊してしまっている」",
  ];

  quotes.forEach(function(q, i) {
    var yPos = 1.15 + i * 1.35;
    addCard(s, 0.8, yPos, 8.4, 1.15);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yPos, w: 0.12, h: 1.15, fill: { color: C.accent } });
    s.addText(q, {
      x: 1.1, y: yPos + 0.1, w: 7.9, h: 0.95,
      fontSize: 16, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
      lineSpacingMultiple: 1.25,
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 5.0, w: 8.4, h: 0.5, fill: { color: C.light } });
  s.addText("この状況は医学的な背景がある ─ 「頑固」ではなく「脳の変化」かもしれません", {
    x: 1.0, y: 5.0, w: 8.0, h: 0.5,
    fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  addPageNum(s, "2", TOTAL);
})();

// ============================================================
// SLIDE 3: なぜ受診を嫌がるのか
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "なぜ受診を嫌がるのか ― 3つの医学的背景");

  var reasons = [
    {
      icon: "🧠",
      title: "病識低下（アノソグノシア）",
      desc: "脳の萎縮が「自分の状態を客観視する能力」に影響。\n「私はおかしくない」は強がりではなく本当にそう感じている。\nアルツハイマー型の 40〜81% にみられる",
      color: C.red,
    },
    {
      icon: "😨",
      title: "スティグマへの恐れ",
      desc: "「診断されたら施設に入れられる」「家族に迷惑をかける」\n長年、一家の柱として生きてきた世代ほど\nプライドや将来への不安が強い",
      color: C.orange,
    },
    {
      icon: "❓",
      title: "医療不信・先送り",
      desc: "「病院に行ったら何をされるかわからない」\n「年をとれば誰でも忘れる」\n「忙しい、今度でいい」という否認・先送り",
      color: C.accent,
    },
  ];

  reasons.forEach(function(r, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.1, 3.0, 4.2);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 3.0, h: 0.5, fill: { color: r.color } });
    s.addText(r.icon + " " + r.title, {
      x: xPos, y: 1.1, w: 3.0, h: 0.5,
      fontSize: 13, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 2,
    });
    s.addText(r.desc, {
      x: xPos + 0.1, y: 1.65, w: 2.8, h: 3.5,
      fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "top", margin: 4,
      lineSpacingMultiple: 1.3,
    });
  });

  addPageNum(s, "3", TOTAL);
})();

// ============================================================
// SLIDE 4: 3つの実践的アプローチ
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "受診につなげる 3つの実践的アプローチ");

  var approaches = [
    {
      num: "1",
      title: "「健康診断」「脳ドック」という名目を使う",
      desc: "「認知症の検査をしよう」と正面から言わない。\n「血圧薬のついでに脳も」「頭のMRIを一度撮っておこう」など\n別の目的を前面に出すのがポイント",
      color: C.green,
    },
    {
      num: "2",
      title: "かかりつけ医に相談し、紹介してもらう",
      desc: "長年の付き合いがある医師の言葉は家族より響く。\nまず家族だけでかかりつけ医に相談し、\n「先生から本人に受診を勧めてもらえますか」とお願いする",
      color: C.orange,
    },
    {
      num: "3",
      title: "まず「家族だけ」で相談に行く",
      desc: "もの忘れ外来・地域包括支援センターは\n患者本人がいなくても相談できる。\n「どうすれば連れてこられるか」のアドバイスを事前に得られる",
      color: C.accent,
    },
  ];

  approaches.forEach(function(a, i) {
    var yPos = 1.1 + i * 1.45;
    addCard(s, 0.5, yPos, 9.0, 1.25);
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: yPos + 0.2, w: 0.85, h: 0.85, fill: { color: a.color } });
    s.addText(a.num, { x: 0.7, y: yPos + 0.2, w: 0.85, h: 0.85, fontSize: 28, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(a.title, { x: 1.75, y: yPos + 0.05, w: 7.5, h: 0.45, fontSize: 17, fontFace: FONT_JP, color: C.text, bold: true, margin: 0 });
    s.addText(a.desc, { x: 1.75, y: yPos + 0.5, w: 7.5, h: 0.7, fontSize: 13, fontFace: FONT_JP, color: C.sub, margin: 0, lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "4", TOTAL);
})();

// ============================================================
// SLIDE 5: 受診前準備チェックリスト
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "受診前に家族が準備すること ― チェックリスト");

  var items = [
    { cat: "持参物", list: ["健康保険証・診察券（他院のものも）", "服用中の薬の一覧（お薬手帳）", "直近の血液・画像検査結果"] },
    { cat: "症状のメモ", list: ["症状が始まった時期・最初に気づいたきっかけ", "具体的なエピソード（「1日3回同じ話をする」など）", "日常生活への支障（料理・家計・外出・服薬）"] },
    { cat: "生活歴", list: ["最終学歴・職歴・趣味", "既往歴（高血圧・糖尿病・脳卒中・うつなど）", "家族歴（認知症の親族はいるか）"] },
  ];

  items.forEach(function(item, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.1, 3.0, 4.2);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 3.0, h: 0.48, fill: { color: C.primary } });
    s.addText(item.cat, {
      x: xPos, y: 1.1, w: 3.0, h: 0.48,
      fontSize: 15, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    item.list.forEach(function(txt, j) {
      var yPos = 1.7 + j * 1.1;
      s.addShape(pres.shapes.OVAL, { x: xPos + 0.15, y: yPos + 0.05, w: 0.28, h: 0.28, fill: { color: C.red } });
      s.addText(txt, {
        x: xPos + 0.5, y: yPos, w: 2.4, h: 0.95,
        fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
      });
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.1, w: 9.4, h: 0.45, fill: { color: C.light } });
  s.addText("診察前に「家族だけ少し話せますか」とスタッフに申し出るのも有効", {
    x: 0.5, y: 5.1, w: 9.0, h: 0.45,
    fontSize: 13, fontFace: FONT_JP, color: C.primary, bold: true, valign: "middle", align: "center", margin: 0,
  });

  addPageNum(s, "5", TOTAL);
})();

// ============================================================
// SLIDE 6: もの忘れ外来の流れ（表）
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "もの忘れ外来では何をするのか");

  var rows = [
    { step: "問診", detail: "いつから・どんな症状か・日常生活への影響を確認。家族からの情報が診断の質を左右する" },
    { step: "神経心理検査", detail: "MMSE・MoCAなどの筆記・口頭テスト（30分程度）。「健康チェックです」と伝えると安心しやすい" },
    { step: "血液検査", detail: "甲状腺機能低下症・ビタミンB12欠乏など「認知症もどき」の除外。治療可能な原因がないか確認" },
    { step: "頭部MRI", detail: "萎縮の部位・脳血管病変の確認。初診は予約が必要なことも多い" },
    { step: "結果と説明", detail: "「今の段階では白黒つかない」ことも多い。ベースラインの記録が将来の比較データになる" },
  ];

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.05, w: 2.2, h: 0.5, fill: { color: C.dark } });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 1.05, w: 7.2, h: 0.5, fill: { color: C.primary } });
  s.addText("内容", { x: 0.3, y: 1.05, w: 2.2, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  s.addText("詳細", { x: 2.5, y: 1.05, w: 7.2, h: 0.5, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

  rows.forEach(function(r, i) {
    var yPos = 1.55 + i * 0.8;
    var bgColor = (i % 2 === 0) ? C.warmBg : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: yPos, w: 9.4, h: 0.8, fill: { color: bgColor } });
    s.addText(r.step, { x: 0.4, y: yPos, w: 2.0, h: 0.8, fontSize: 14, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2 });
    s.addText(r.detail, { x: 2.5, y: yPos, w: 7.0, h: 0.8, fontSize: 13, fontFace: FONT_JP, color: C.text, valign: "middle", margin: [0, 6, 0, 6], lineSpacingMultiple: 1.2 });
  });

  addPageNum(s, "6", TOTAL);
})();

// ============================================================
// SLIDE 7: 相談窓口
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "受診前後に使える相談窓口");

  var windows = [
    {
      name: "地域包括支援センター",
      tag: "まず相談",
      tagColor: C.green,
      desc: "各市区町村に設置。無料で相談できる。\n医療機関への橋渡し・介護保険の案内など\n「どこに行けばいいかわからない」段階から頼れる",
      how: "市区町村の窓口 or「地域包括支援センター＋市区町村名」で検索",
    },
    {
      name: "認知症疾患医療センター",
      tag: "専門診断",
      tagColor: C.orange,
      desc: "都道府県が指定する専門拠点。\nBPSDが強い・かかりつけ医での診断が困難な\n複雑なケースで紹介を受けることが多い",
      how: "かかりつけ医または地域包括支援センターから紹介",
    },
    {
      name: "認知症の人と家族の会",
      tag: "家族支援",
      tagColor: C.accent,
      desc: "同じ状況の家族同士が支え合う場。\n受診前後を問わず、経験談を聞ける。\n電話相談窓口: 0120-294-456",
      how: "alzheimer.or.jp からアクセス可",
    },
  ];

  windows.forEach(function(w, i) {
    var xPos = 0.3 + i * 3.2;
    addCard(s, xPos, 1.1, 3.0, 4.3);
    s.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.1, w: 3.0, h: 0.48, fill: { color: w.tagColor } });
    s.addText(w.tag, { x: xPos, y: 1.1, w: 3.0, h: 0.48, fontSize: 14, fontFace: FONT_JP, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(w.name, {
      x: xPos + 0.1, y: 1.65, w: 2.8, h: 0.65,
      fontSize: 15, fontFace: FONT_JP, color: C.primary, bold: true, align: "center", valign: "middle", margin: 2,
    });
    s.addShape(pres.shapes.LINE, { x: xPos + 0.3, y: 2.35, w: 2.4, h: 0, line: { color: C.light, width: 1 } });
    s.addText(w.desc, {
      x: xPos + 0.1, y: 2.45, w: 2.8, h: 1.6,
      fontSize: 12, fontFace: FONT_JP, color: C.text, valign: "top", margin: 4, lineSpacingMultiple: 1.3,
    });
    s.addShape(pres.shapes.RECTANGLE, { x: xPos + 0.1, y: 4.12, w: 2.8, h: 0.6, fill: { color: C.warmBg } });
    s.addText(w.how, {
      x: xPos + 0.1, y: 4.12, w: 2.8, h: 0.6,
      fontSize: 11, fontFace: FONT_JP, color: C.sub, valign: "middle", margin: 4, lineSpacingMultiple: 1.1,
    });
  });

  addPageNum(s, "7", TOTAL);
})();

// ============================================================
// SLIDE 8: 緊急受診が必要なとき
// ============================================================
(function() {
  var s = pres.addSlide();
  addHeader(s, "緊急受診が必要なとき ― もの忘れ外来ではなく救急へ");

  // 緊急サイン
  var signs = [
    "急に（数日〜数週間で）認知機能が落ちた",
    "手足の麻痺・言語障害・激しい頭痛を伴う",
    "高熱・意識の変容がある",
    "強い幻視・幻聴・妄想で本人や家族が危険な状況",
  ];

  addCard(s, 0.5, 1.1, 9.0, 3.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 9.0, h: 0.5, fill: { color: C.red } });
  s.addText("こんな症状は急いで受診を", { x: 0.7, y: 1.1, w: 8.6, h: 0.5, fontSize: 17, fontFace: FONT_JP, color: C.white, bold: true, valign: "middle", margin: 0 });

  signs.forEach(function(sign, i) {
    var yPos = 1.75 + i * 0.55;
    s.addShape(pres.shapes.OVAL, { x: 0.75, y: yPos + 0.1, w: 0.3, h: 0.3, fill: { color: C.red } });
    s.addText(sign, {
      x: 1.2, y: yPos, w: 7.9, h: 0.5,
      fontSize: 16, fontFace: FONT_JP, color: C.text, valign: "middle", margin: 0,
    });
  });

  // 疾患リスト
  addCard(s, 0.5, 4.25, 9.0, 1.1);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.25, w: 0.15, h: 1.1, fill: { color: C.red } });
  s.addText([
    { text: "これらは脳卒中・脳炎・せん妄など", options: { bold: true, color: C.red, breakLine: true } },
    { text: "緊急疾患の可能性があります。もの忘れ外来でなく救急外来へ。", options: { color: C.text } },
  ], { x: 0.8, y: 4.3, w: 8.5, h: 1.0, fontSize: 15, fontFace: FONT_JP, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });

  addPageNum(s, "8", TOTAL);
})();

// ============================================================
// SLIDE 9: まとめ（Take Home Message）
// ============================================================
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("Take Home Message", {
    x: 0.5, y: 0.25, w: 9.0, h: 0.6,
    fontSize: 26, fontFace: FONT_EN, color: C.white, bold: true, align: "center", margin: 0,
  });

  var messages = [
    { num: "1", text: "受診拒否の背景には「病識低下」という医学的原因がある。\n頑固なだけではない" },
    { num: "2", text: "「認知症の検査」とは言わず、健康診断・脳ドック・\nかかりつけ医など別の入口を使う" },
    { num: "3", text: "まず家族だけで相談できる。地域包括支援センター・\nもの忘れ外来で受け付けている" },
    { num: "4", text: "1回で白黒つかなくてよい。ベースラインの記録を残すことが\n将来の比較データになる" },
  ];

  messages.forEach(function(m, i) {
    var yPos = 1.05 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos, w: 8.4, h: 0.9, fill: { color: "243B5C" }, rectRadius: 0.1 });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fill: { color: C.accent } });
    s.addText(m.num, { x: 1.0, y: yPos + 0.15, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT_EN, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(m.text, { x: 1.8, y: yPos, w: 7.2, h: 0.9, fontSize: 15, fontFace: FONT_JP, color: C.light, valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
  });

  s.addShape(pres.shapes.LINE, { x: 2, y: 5.15, w: 6, h: 0, line: { color: C.accent, width: 1 } });
  s.addText("医知創造ラボ  今村久司  |  チャンネル登録よろしくお願いします", {
    x: 0.5, y: 5.22, w: 9.0, h: 0.38,
    fontSize: 13, fontFace: FONT_JP, color: C.sub, align: "center", margin: 0,
  });

  addPageNum(s, "9", TOTAL);
})();

// ============================================================
// ファイル出力
// ============================================================
var outPath = __dirname + "/認知症の親と受診_スライド.pptx";
pres.writeFile({ fileName: outPath }).then(function() {
  console.log("PPTX created: " + outPath);
}).catch(function(err) {
  console.error("Error:", err);
});
