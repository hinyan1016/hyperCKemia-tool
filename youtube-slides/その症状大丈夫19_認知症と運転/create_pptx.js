var pptxgen = require("pptxgenjs");
var pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "【その症状、大丈夫？#19】認知症と運転免許 ― 家族が知っておくべき制度と説得のコツ";

var C = {
  dark: "1B3A5C", primary: "2C5AA0", accent: "E8850C", light: "E8F4FD",
  warmBg: "F8F9FA", white: "FFFFFF", text: "2D3436", sub: "636E72",
  red: "DC3545", green: "28A745", yellow: "F5C518",
  lightRed: "F8D7DA", lightYellow: "FFF3CD", lightGreen: "D4EDDA",
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
  if (borderColor) { s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.12, h: h, fill: { color: borderColor }, rectRadius: 0.05 }); }
}
function ftr(s) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.25, w: 10, h: 0.38, fill: { color: C.dark } });
  s.addText("医知創造ラボ ｜ その症状、大丈夫？ #19", { x: 0.4, y: 5.27, w: 5, h: 0.3, fontSize: 10, fontFace: FJ, color: C.white, margin: 0 });
}

// SLIDE 1: タイトル
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("🚗", { x: 0.6, y: 0.4, w: 8.8, h: 0.9, fontSize: 52, align: "center", margin: 0 });
  s.addText("認知症と運転免許", {
    x: 0.6, y: 1.35, w: 8.8, h: 0.85,
    fontSize: 40, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 2.5, y: 2.3, w: 5, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("家族が知っておくべき\n制度と説得のコツ", {
    x: 0.6, y: 2.45, w: 8.8, h: 0.85,
    fontSize: 26, fontFace: FJ, color: C.accent, align: "center", lineSpacingMultiple: 1.2, margin: 0,
  });
  s.addText("その症状、大丈夫？ #19", {
    x: 0.6, y: 3.95, w: 8.8, h: 0.4,
    fontSize: 16, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.45, w: 8.8, h: 0.4,
    fontSize: 12, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
})();

// SLIDE 2: こんな経験ありませんか？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "こんな経験、ありませんか？");
  var items = [
    { emoji: "👴", text: "親の運転が心配だが、どう伝えればよいかわからない" },
    { emoji: "🚗", text: "駐車が下手になった、道に迷うことが増えた気がする" },
    { emoji: "😟", text: "「まだ運転できる」と言い張って聞いてくれない" },
    { emoji: "📋", text: "免許返納を勧めたいが、制度がよくわからない" },
    { emoji: "🏡", text: "免許を返したら生活が不便になるのでは、と心配" },
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.78;
    card(s, 0.6, y, 8.8, 0.65, C.white, C.primary);
    s.addText(item.emoji, { x: 0.85, y: y + 0.05, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(item.text, { x: 1.6, y: y + 0.05, w: 7.5, h: 0.55, fontSize: 22, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 3: 認知症と運転 ― なぜ危険？
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "認知症と運転 ― なぜ危険？");
  card(s, 0.5, 1.1, 9.0, 0.55, C.light, C.primary);
  s.addText("安全な運転には、複数の脳機能が同時に働く必要があります", { x: 0.8, y: 1.18, w: 8.5, h: 0.42, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var risks = [
    { icon: "🧠", title: "記憶障害", body: "道順を忘れる・駐車場所を忘れる" },
    { icon: "⚡", title: "判断力の低下", body: "飛び出しへの反応が遅れる・優先判断を誤る" },
    { icon: "🗺️", title: "空間認識の障害", body: "車線変更を誤る・縁石に乗り上げる" },
    { icon: "🎯", title: "注意力の分散", body: "複数のことに同時対応できなくなる" },
  ];
  risks.forEach(function(r, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var x = col === 0 ? 0.5 : 5.3;
    var y = 1.9 + row * 1.45;
    card(s, x, y, 4.6, 1.2, C.white, C.red);
    s.addText(r.icon, { x: x + 0.2, y: y + 0.1, w: 0.7, h: 0.9, fontSize: 28, align: "center", margin: 0 });
    s.addText(r.title, { x: x + 0.95, y: y + 0.1, w: 3.5, h: 0.38, fontSize: 16, fontFace: FJ, color: C.red, bold: true, margin: 0 });
    s.addText(r.body, { x: x + 0.95, y: y + 0.52, w: 3.5, h: 0.55, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  card(s, 0.5, 4.55, 9.0, 0.6, C.lightRed, C.red);
  s.addText("⚠️ 認知症と診断されていなくても、MCIの段階から運転リスクは上昇します", { x: 0.8, y: 4.6, w: 8.5, h: 0.5, fontSize: 15, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0 });
  ftr(s);
})();

// SLIDE 4: 運転に必要な脳の機能
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "運転に必要な脳の機能");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("運転は「脳の総合テスト」― 多くの認知機能を同時に使います", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var funcs = [
    { icon: "👁️", name: "注意・集中力", desc: "標識・歩行者・対向車を同時に把握する", area: "前頭葉・頭頂葉" },
    { icon: "⚖️", name: "判断力", desc: "「曲がれるか」「止まるべきか」を瞬時に決める", area: "前頭前野" },
    { icon: "🗺️", name: "空間認識", desc: "車の位置・距離・車線の幅を感じる", area: "頭頂葉" },
    { icon: "⚡", name: "反応速度", desc: "危険を察知してブレーキを踏む", area: "小脳・運動野" },
  ];
  funcs.forEach(function(f, i) {
    var y = 1.8 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(f.icon, { x: 0.7, y: y + 0.07, w: 0.7, h: 0.55, fontSize: 26, align: "center", margin: 0 });
    s.addText(f.name, { x: 1.5, y: y + 0.07, w: 2.2, h: 0.35, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(f.desc, { x: 1.5, y: y + 0.42, w: 5.5, h: 0.28, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
    s.addText("脳領域: " + f.area, { x: 7.1, y: y + 0.2, w: 2.2, h: 0.28, fontSize: 12, fontFace: FJ, color: C.sub, align: "right", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 5: 認知症の初期に現れる運転の変化
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "認知症の初期に現れる運転の変化");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("「まだ大丈夫」と本人が言っていても、こんな変化が出ていたら要注意", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var signs = [
    { emoji: "🅿️", text: "駐車がうまくできない・縁石に乗り上げる" },
    { emoji: "🗺️", text: "いつも走る道なのに迷ってしまう" },
    { emoji: "🚦", text: "信号を見落とす・一時停止を無視する" },
    { emoji: "💥", text: "小さなこすり傷や接触が増えた" },
    { emoji: "😤", text: "あおり運転や急ブレーキが増えた" },
    { emoji: "🔑", text: "駐車した場所を忘れる・エンジンを切り忘れる" },
  ];
  signs.forEach(function(sign, i) {
    var col = i % 2;
    var row = Math.floor(i / 2);
    var x = col === 0 ? 0.5 : 5.3;
    var y = 1.85 + row * 1.05;
    card(s, x, y, 4.6, 0.88, C.white, C.accent);
    s.addText(sign.emoji, { x: x + 0.15, y: y + 0.08, w: 0.55, h: 0.7, fontSize: 22, align: "center", margin: 0 });
    s.addText(sign.text, { x: x + 0.82, y: y + 0.08, w: 3.6, h: 0.7, fontSize: 14, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 6: 日本の制度 ― 75歳以上の免許更新
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "日本の制度 ― 75歳以上の免許更新");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("道路交通法（2022年改正）で75歳以上の運転者への規制が強化されました", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var steps = [
    { num: "①", title: "認知機能検査（更新時・3年ごと）", body: "時計の描写・記憶力テストなど。「第1分類」と判定されると医師の診断が必要になります。" },
    { num: "②", title: "高齢者講習（75歳以上全員）", body: "実車評価を含む講習。運転技能を客観的に確認します。" },
    { num: "③", title: "運転技能検査（2022年改正で新設）", body: "過去3年間に一定の交通違反がある75歳以上は、合格しないと免許更新ができません。" },
  ];
  steps.forEach(function(step, i) {
    var y = 1.82 + i * 1.1;
    card(s, 0.5, y, 9.0, 0.95, C.white, C.primary);
    s.addText(step.num, { x: 0.72, y: y + 0.08, w: 0.6, h: 0.75, fontSize: 22, fontFace: FJ, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(step.title, { x: 1.45, y: y + 0.08, w: 7.8, h: 0.32, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(step.body, { x: 1.45, y: y + 0.44, w: 7.8, h: 0.42, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
  });
  card(s, 0.5, 5.12, 9.0, 0.05, C.accent);
  ftr(s);
})();

// SLIDE 7: 臨時適性検査とは
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "臨時適性検査とは ― 診断後の手続き");
  card(s, 0.5, 1.1, 9.0, 0.55, C.lightRed, C.red);
  s.addText("⚠️ 認知症と診断された場合、医師には公安委員会への通知義務があります（2017年改正）", { x: 0.8, y: 1.15, w: 8.5, h: 0.45, fontSize: 15, fontFace: FJ, color: C.red, bold: true, valign: "middle", margin: 0 });

  s.addText("免許取り消しまでの流れ", { x: 0.5, y: 1.85, w: 9.0, h: 0.35, fontSize: 18, fontFace: FJ, color: C.dark, bold: true, margin: 0 });
  var flow = [
    { step: "診断", body: "専門医が認知症（アルツハイマー型・レビー小体型など）と診断" },
    { step: "通知", body: "医師が都道府県公安委員会に診断内容を通知" },
    { step: "検査", body: "公安委員会が臨時適性検査を実施（または医師に診断書を提出）" },
    { step: "処分", body: "認知症と確認されると免許取り消し・停止の行政処分" },
  ];
  flow.forEach(function(f, i) {
    var y = 2.3 + i * 0.72;
    card(s, 0.5, y, 9.0, 0.6, C.white, C.red);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: y + 0.1, w: 0.9, h: 0.38, fill: { color: C.red }, rectRadius: 0.05 });
    s.addText(f.step, { x: 0.7, y: y + 0.1, w: 0.9, h: 0.38, fontSize: 13, fontFace: FJ, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(f.body, { x: 1.75, y: y + 0.1, w: 7.5, h: 0.38, fontSize: 15, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
    if (i < 3) { s.addText("↓", { x: 0.7, y: y + 0.68, w: 0.9, h: 0.25, fontSize: 14, fontFace: FJ, color: C.red, align: "center", margin: 0 }); }
  });
  ftr(s);
})();

// SLIDE 8: 運転をやめてもらうための家族の対応
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "家族の対応 ― 説得のコツ");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("対立ではなく「一緒に考える」姿勢が鍵です", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 18, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var tips = [
    { ok: "✅ やること", ng: "❌ やらないこと" },
  ];
  var doList = [
    "医師から直接伝えてもらう（医師の言葉は権威がある）",
    "具体的なエピソードで話す（「先日○○で…」）",
    "「家族が心配している」という気持ちを伝える",
    "免許返納後の生活サポートを先に提示する",
    "地域の返納支援制度を一緒に調べる",
  ];
  var dontList = [
    "「あなたは認知症だから」と病名で攻める",
    "感情的に怒鳴ったり責める",
    "頭ごなしに「絶対ダメ」と禁止する",
    "一度の会話で決めようとする",
    "本人の尊厳を傷つける言い方をする",
  ];
  s.addText("やること（推奨）", { x: 0.5, y: 1.8, w: 4.3, h: 0.35, fontSize: 15, fontFace: FJ, color: C.green, bold: true, margin: 0 });
  s.addText("やらないこと（NG）", { x: 5.2, y: 1.8, w: 4.3, h: 0.35, fontSize: 15, fontFace: FJ, color: C.red, bold: true, margin: 0 });
  doList.forEach(function(item, i) {
    var y = 2.22 + i * 0.57;
    card(s, 0.5, y, 4.5, 0.48, C.lightGreen, C.green);
    s.addText("• " + item, { x: 0.78, y: y + 0.04, w: 4.1, h: 0.42, fontSize: 12, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  dontList.forEach(function(item, i) {
    var y = 2.22 + i * 0.57;
    card(s, 5.2, y, 4.5, 0.48, C.lightRed, C.red);
    s.addText("• " + item, { x: 5.48, y: y + 0.04, w: 4.1, h: 0.42, fontSize: 12, fontFace: FJ, color: C.text, valign: "middle", margin: 0 });
  });
  ftr(s);
})();

// SLIDE 9: 運転をやめた後の生活支援
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "運転をやめた後の生活支援");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("「車がないと生活できない」は今や過去の話 ― 代替手段を一緒に整えましょう", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 16, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var supports = [
    { emoji: "🚕", title: "タクシー助成券・乗車クーポン", body: "多くの市区町村が高齢者向けに補助。免許返納者は特典を受けられることも。" },
    { emoji: "🚌", title: "コミュニティバス・デマンド交通", body: "地域の乗合サービス。予約制で自宅付近まで来てくれる地域も増加。" },
    { emoji: "🛒", title: "買い物・食事の宅配サービス", body: "スーパーの宅配・食材キット・弁当配送。日常の買い物困難を解決。" },
    { emoji: "👨‍👩‍👧", title: "家族・近隣とのサポート体制", body: "定期的な送迎の役割分担。「○曜日は私が送る」の明示が重要。" },
    { emoji: "📱", title: "移動支援アプリ・シニア向け交通", body: "Uber・乗合タクシーアプリ。自治体のシニア向けライドシェアも各地で普及中。" },
  ];
  supports.forEach(function(sup, i) {
    var y = 1.82 + i * 0.67;
    card(s, 0.5, y, 9.0, 0.57, C.white, C.green);
    s.addText(sup.emoji, { x: 0.7, y: y + 0.05, w: 0.6, h: 0.47, fontSize: 22, align: "center", margin: 0 });
    s.addText(sup.title, { x: 1.42, y: y + 0.04, w: 2.8, h: 0.28, fontSize: 14, fontFace: FJ, color: C.green, bold: true, margin: 0 });
    s.addText(sup.body, { x: 1.42, y: y + 0.32, w: 7.8, h: 0.22, fontSize: 12, fontFace: FJ, color: C.text, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 10: 受診の目安 ― 緊急度3段階
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "受診の目安 ― 緊急度3段階");

  var levels = [
    {
      color: C.red, bgColor: C.lightRed, label: "🔴 すぐに受診",
      items: [
        "逆走・信号無視など重大な交通違反があった",
        "事故を起こしたが本人が覚えていない",
        "帰宅できずに警察に保護された",
      ]
    },
    {
      color: C.yellow, bgColor: C.lightYellow, label: "🟡 早めに受診",
      items: [
        "小さな接触事故・こすり傷が増えてきた",
        "いつもの道で迷うことが複数回あった",
        "家族が運転に同乗して不安を感じた",
      ]
    },
    {
      color: C.green, bgColor: C.lightGreen, label: "🟢 かかりつけ医に相談",
      items: [
        "駐車が少し下手になった気がする",
        "標識への反応が遅くなった（同乗者の感覚）",
        "本人が「運転が怖い」と言い始めた",
      ]
    },
  ];
  levels.forEach(function(lv, i) {
    var y = 1.05 + i * 1.42;
    card(s, 0.5, y, 9.0, 1.28, lv.bgColor, lv.color);
    s.addText(lv.label, { x: 0.78, y: y + 0.08, w: 8.5, h: 0.34, fontSize: 17, fontFace: FJ, color: lv.color, bold: true, margin: 0 });
    lv.items.forEach(function(item, j) {
      s.addText("• " + item, { x: 0.78, y: y + 0.46 + j * 0.28, w: 8.5, h: 0.26, fontSize: 13, fontFace: FJ, color: C.text, margin: 0 });
    });
  });
  ftr(s);
})();

// SLIDE 11: 早期発見のためにできること
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "早期発見のためにできること");
  card(s, 0.5, 1.1, 9.0, 0.5, C.light, C.primary);
  s.addText("認知症は早期発見・早期対応で進行を遅らせることができます", { x: 0.8, y: 1.15, w: 8.5, h: 0.4, fontSize: 17, fontFace: FJ, color: C.primary, bold: true, valign: "middle", margin: 0 });

  var actions = [
    { emoji: "📋", title: "かかりつけ医への相談", body: "「物忘れが気になる」と伝えるだけでOK。血液検査・簡単な認知機能テスト（MMSE・HDS-R）を受けられます。" },
    { emoji: "🏥", title: "もの忘れ外来・神経内科の受診", body: "専門医によるMRI・神経心理検査で認知症の種類と程度を詳しく評価します。" },
    { emoji: "🧩", title: "認知機能のセルフチェック", body: "「認知症チェックリスト」は市区町村や厚生労働省のサイトで無料公開。まず自分でチェック。" },
    { emoji: "👨‍👩‍👧", title: "家族が気になることをメモする", body: "「いつから、どんなこと」を具体的に記録して受診時に持参すると診断がスムーズです。" },
  ];
  actions.forEach(function(a, i) {
    var y = 1.82 + i * 0.82;
    card(s, 0.5, y, 9.0, 0.7, C.white, C.primary);
    s.addText(a.emoji, { x: 0.7, y: y + 0.07, w: 0.65, h: 0.55, fontSize: 24, align: "center", margin: 0 });
    s.addText(a.title, { x: 1.48, y: y + 0.07, w: 7.8, h: 0.28, fontSize: 15, fontFace: FJ, color: C.primary, bold: true, margin: 0 });
    s.addText(a.body, { x: 1.48, y: y + 0.38, w: 7.8, h: 0.3, fontSize: 12, fontFace: FJ, color: C.text, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 12: まとめ
(function() {
  var s = pres.addSlide();
  s.background = { color: C.warmBg };
  hdr(s, "まとめ");

  var points = [
    {
      num: "1",
      title: "認知症は運転能力に直接影響する",
      body: "記憶・判断・空間認識・反応速度が低下し、事故リスクが高まります。MCI段階からリスクは上昇します。",
      color: C.primary,
    },
    {
      num: "2",
      title: "制度を知ることが家族の力になる",
      body: "75歳以上の認知機能検査・臨時適性検査・免許取り消し制度を把握し、「本人だけの問題」にしない。",
      color: C.accent,
    },
    {
      num: "3",
      title: "対立せず、代替案を示しながら説得を",
      body: "医師の力を借り、生活支援を先に提示する。免許返納後も「生活できる」と見せることが鍵です。",
      color: C.green,
    },
  ];
  points.forEach(function(p, i) {
    var y = 1.08 + i * 1.35;
    card(s, 0.5, y, 9.0, 1.18, C.white, p.color);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: y + 0.14, w: 0.55, h: 0.55, fill: { color: p.color }, rectRadius: 0.08 });
    s.addText(p.num, { x: 0.7, y: y + 0.14, w: 0.55, h: 0.55, fontSize: 20, fontFace: FE, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(p.title, { x: 1.42, y: y + 0.1, w: 7.8, h: 0.35, fontSize: 17, fontFace: FJ, color: p.color, bold: true, margin: 0 });
    s.addText(p.body, { x: 1.42, y: y + 0.5, w: 7.8, h: 0.6, fontSize: 13, fontFace: FJ, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
  });
  ftr(s);
})();

// SLIDE 13: エンディング
(function() {
  var s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("ご視聴ありがとうございました", {
    x: 0.6, y: 0.5, w: 8.8, h: 0.6,
    fontSize: 28, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 1.5, y: 1.22, w: 7, h: 0, line: { color: C.accent, width: 1.5 } });

  card(s, 1.0, 1.42, 8.0, 1.35, "1E2D40", null);
  s.addText("次回 #20", { x: 1.25, y: 1.55, w: 7.5, h: 0.3, fontSize: 14, fontFace: FJ, color: C.accent, align: "center", margin: 0 });
  s.addText("脳ドックは受けるべき？", { x: 1.25, y: 1.9, w: 7.5, h: 0.38, fontSize: 22, fontFace: FJ, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText("― 検査の意味と結果の活かし方", { x: 1.25, y: 2.3, w: 7.5, h: 0.3, fontSize: 16, fontFace: FJ, color: "90A4AE", align: "center", margin: 0 });

  s.addText("チャンネル登録・高評価をよろしくお願いします！", {
    x: 0.6, y: 3.15, w: 8.8, h: 0.45,
    fontSize: 18, fontFace: FJ, color: C.accent, bold: true, align: "center", margin: 0,
  });
  s.addText("概要欄にブログ記事・参考文献リンクを掲載しています", {
    x: 0.6, y: 3.7, w: 8.8, h: 0.35,
    fontSize: 14, fontFace: FJ, color: "90A4AE", align: "center", margin: 0,
  });
  s.addText("その症状、大丈夫？ #19", {
    x: 0.6, y: 4.25, w: 8.8, h: 0.3,
    fontSize: 14, fontFace: FJ, color: "78909C", align: "center", margin: 0,
  });
  s.addText("医知創造ラボ ～脳神経内科医がAIで紡ぐ最新医療情報～", {
    x: 0.6, y: 4.62, w: 8.8, h: 0.3,
    fontSize: 11, fontFace: FJ, color: "607D8B", align: "center", margin: 0,
  });
})();

pres.writeFile({ fileName: __dirname + "/dementia_driving_slides.pptx" })
  .then(function() { console.log("PPTX saved."); })
  .catch(function(err) { console.error(err); });
