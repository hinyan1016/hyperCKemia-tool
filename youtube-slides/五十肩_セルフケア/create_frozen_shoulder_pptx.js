const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "今村久司";
pres.title = "五十肩のセルフケア ～自宅でできる運動療法～";

// --- Color Palette (Health / Warm / Patient-friendly) ---
const C = {
  dark: "1B4332",      // deep green (title bg)
  primary: "2D6A4F",   // forest green
  accent: "52B788",    // fresh green accent
  light: "D8F3DC",     // pale mint
  warmBg: "F7F9F7",    // off-white green tint
  white: "FFFFFF",
  text: "2D3436",      // charcoal
  sub: "636E72",       // gray
  coral: "E17055",     // warm coral for emphasis
  orange: "F39C12",    // amber for highlights
};

const FONT_H = "Arial Black";
const FONT_B = "Arial";

// Helper: create shadow (fresh object each time)
const makeShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });

// Image URLs (from the ChatGPT page)
const IMGS = {
  hotpack1: "https://m.media-amazon.com/images/I/312Us8D8U-L._AC_UF350%2C350_QL80_.jpg",
  hotpack2: "https://m.media-amazon.com/images/I/71a6c8bjXhL.jpg",
  hotpack3: "https://cdn.shopify.com/s/files/1/0764/0209/8488/files/1_0542416c-b88b-4093-b013-71d56ce8208c.webp?v=1684432914",
  pendulum1: "https://i.pinimg.com/474x/ed/15/06/ed1506ddd3477804adbf6ac674c9e520.jpg",
  pendulum2: "https://content.healthwise.net/resources/14.8/en-us/media/medical/hw/aco5777_460x300.jpg",
  towel1: "https://static1.squarespace.com/static/5891f1bb86e6c0e1ef425901/5891f3f517bffc16ca16f8c0/5e32ce53dccb3e3e70baeac2/1673252812057/Frozen-Shoulder-Stretch-3.jpg",
  towel2: "https://www.ipcphysicaltherapy.com/pic/Exercise-Images/473.jpg",
  towel3: "https://ipcphysicaltherapy.com/pic/Exercise-Images/473b.jpg",
  wall1: "https://cdn.shopify.com/s/files/1/0601/5460/9845/files/003.jpg",
  wall2: "https://media.univcomm.cornell.edu/photos/1280x720/600EBB67-C2C6-47F3-CFE391A0456833EE.jpeg",
  wall3: "https://ipcphysicaltherapy.com/pic/Exercise-Images/352b.jpg",
};

// ============================================================
// SLIDE 1: Title
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  // Decorative accent bar top
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  // Decorative accent bar bottom
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("五十肩のセルフケア", {
    x: 0.8, y: 1.2, w: 8.4, h: 1.2,
    fontSize: 44, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("～ 自宅でできる運動療法 ～", {
    x: 0.8, y: 2.4, w: 8.4, h: 0.8,
    fontSize: 28, fontFace: FONT_B, color: C.accent, align: "center", margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 3.4, w: 4, h: 0, line: { color: C.accent, width: 2 } });
  s.addText("痛くない範囲で、毎日少しずつ動かしましょう", {
    x: 0.8, y: 3.7, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_B, color: C.light, align: "center", italic: true, margin: 0,
  });
}

// ============================================================
// SLIDE 2: 五十肩とは？
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  // Header bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("五十肩とは？", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Main message in a card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.3, w: 8.8, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
  s.addText([
    { text: "正式名称：肩関節周囲炎", options: { fontSize: 24, bold: true, color: C.primary, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "肩の周りが炎症を起こし、\n動かしにくくなる病気です。", options: { fontSize: 22, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "自然に治ることが多いですが、\n放っておくと固まってしまいます。", options: { fontSize: 22, color: C.coral, bold: true } },
  ], { x: 1.2, y: 1.6, w: 7.6, h: 3.2, fontFace: FONT_B, align: "center", valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 3: なぜ運動が大切？
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("なぜ運動が大切？", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Left card - Bad
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.6, fill: { color: C.coral } });
  s.addText("動かさないと...", { x: 0.5, y: 1.3, w: 4.3, h: 0.6, fontSize: 20, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "痛い → 動かさない", options: { fontSize: 20, breakLine: true } },
    { text: "→ 肩が固まる", options: { fontSize: 20, breakLine: true } },
    { text: "→ もっと動かない", options: { fontSize: 20, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "悪循環に陥ります", options: { fontSize: 22, bold: true, color: C.coral } },
  ], { x: 0.8, y: 2.2, w: 3.7, h: 2.6, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });

  // Right card - Good
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.6, fill: { color: C.accent } });
  s.addText("少しずつ動かすと...", { x: 5.2, y: 1.3, w: 4.3, h: 0.6, fontSize: 20, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
  s.addText([
    { text: "少し動かす → 柔らかくなる", options: { fontSize: 20, breakLine: true } },
    { text: "→ もっと動く", options: { fontSize: 20, breakLine: true } },
    { text: "→ 回復が早まる", options: { fontSize: 20, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "好循環が生まれます", options: { fontSize: 22, bold: true, color: C.primary } },
  ], { x: 5.5, y: 2.2, w: 3.7, h: 2.6, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 4: 治療の2つの目標
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("治療の2つの目標", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Goal 1
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.4, w: 4.3, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("1", { x: 1.5, y: 1.6, w: 1, h: 1, fontSize: 48, fontFace: FONT_H, color: C.accent, bold: true, align: "center", margin: 0 });
  s.addText("痛みを\n悪化させない", {
    x: 0.8, y: 2.7, w: 3.7, h: 1.8, fontSize: 26, fontFace: FONT_H, color: C.text, align: "center", valign: "middle", margin: 0,
  });

  // Goal 2
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.4, w: 4.3, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
  s.addText("2", { x: 6.2, y: 1.6, w: 1, h: 1, fontSize: 48, fontFace: FONT_H, color: C.accent, bold: true, align: "center", margin: 0 });
  s.addText("可動域を\n維持・回復させる", {
    x: 5.5, y: 2.7, w: 3.7, h: 1.8, fontSize: 26, fontFace: FONT_H, color: C.text, align: "center", valign: "middle", margin: 0,
  });
}

// ============================================================
// SLIDE 5: 運動の前に ― まず温めよう
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("運動の前に ― まず温めよう", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Message
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.2, fill: { color: C.light } });
  s.addText("温めてから運動すると、肩が動かしやすくなります", {
    x: 0.8, y: 1.3, w: 8.4, h: 1, fontSize: 24, fontFace: FONT_B, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0,
  });

  // 3 images side by side
  s.addImage({ path: IMGS.hotpack1, x: 0.6, y: 2.7, w: 2.6, h: 2.4, sizing: { type: "contain", w: 2.6, h: 2.4 } });
  s.addImage({ path: IMGS.hotpack2, x: 3.7, y: 2.7, w: 2.6, h: 2.4, sizing: { type: "contain", w: 2.6, h: 2.4 } });
  s.addImage({ path: IMGS.hotpack3, x: 6.8, y: 2.7, w: 2.6, h: 2.4, sizing: { type: "contain", w: 2.6, h: 2.4 } });
}

// ============================================================
// SLIDE 6: 温め方のポイント
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("温め方のポイント", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // 3 methods as cards
  const methods = [
    { title: "蒸しタオル", desc: "電子レンジで\n温めたタオルを\n肩にあてる" },
    { title: "入浴", desc: "お風呂に\nゆっくり浸かって\n肩を温める" },
    { title: "ホットパック", desc: "市販の\nホットパックを\n肩にあてる" },
  ];
  methods.forEach((m, i) => {
    const x = 0.5 + i * 3.15;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.85, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.85, h: 0.6, fill: { color: C.accent } });
    s.addText(m.title, { x, y: 1.3, w: 2.85, h: 0.6, fontSize: 20, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(m.desc, { x: x + 0.15, y: 2.2, w: 2.55, h: 2.6, fontSize: 20, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
  });

  // Warning
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.15, w: 9, h: 0.4, fill: { color: "FFF3CD" } });
  s.addText("注意：炎症が強い急性期（夜間痛が強い時期）は温めすぎないこと", {
    x: 0.7, y: 5.15, w: 8.6, h: 0.4, fontSize: 14, fontFace: FONT_B, color: C.coral, align: "center", margin: 0,
  });
}

// ============================================================
// SLIDE 7: 運動1 振り子運動とは
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("運動 1", {
    x: 0.5, y: 1.5, w: 9, h: 0.8,
    fontSize: 24, fontFace: FONT_B, color: C.accent, align: "center", margin: 0,
  });
  s.addText("振り子運動（Codman運動）", {
    x: 0.5, y: 2.3, w: 9, h: 1.2,
    fontSize: 40, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("一番簡単で、痛みが少ない運動です", {
    x: 0.5, y: 3.6, w: 9, h: 0.6,
    fontSize: 22, fontFace: FONT_B, color: C.light, italic: true, align: "center", margin: 0,
  });
}

// ============================================================
// SLIDE 8: 振り子運動のやり方
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("振り子運動のやり方", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Left: images
  s.addImage({ path: IMGS.pendulum1, x: 0.5, y: 1.2, w: 4.2, h: 4.0, sizing: { type: "contain", w: 4.2, h: 4.0 } });

  // Right: steps
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 4.0, fill: { color: C.white }, shadow: makeShadow() });
  s.addText([
    { text: "1. 前かがみになる", options: { fontSize: 22, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "2. 腕をだらんと垂らす", options: { fontSize: 22, bold: true, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "3. 体を揺らして\n    腕を振り子のように動かす", options: { fontSize: 22, bold: true } },
  ], { x: 5.5, y: 1.5, w: 3.7, h: 3.4, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 9: 振り子運動のポイント
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("振り子運動のポイント", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Left: image
  s.addImage({ path: IMGS.pendulum2, x: 0.5, y: 1.2, w: 4.2, h: 3.2, sizing: { type: "contain", w: 4.2, h: 3.2 } });

  // Right: key points
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  s.addText([
    { text: "力を入れない！", options: { fontSize: 26, bold: true, color: C.coral, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "自分の筋力で上げない", options: { fontSize: 20, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "重力と反動で動かす", options: { fontSize: 20 } },
  ], { x: 5.5, y: 1.5, w: 3.7, h: 2.6, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });

  // Bottom tip box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 9, h: 0.8, fill: { color: C.light } });
  s.addText("「力を抜いて、腕をぶらぶらさせるだけでOK」", {
    x: 0.8, y: 4.65, w: 8.4, h: 0.7, fontSize: 22, fontFace: FONT_B, color: C.primary, bold: true, align: "center", margin: 0,
  });
}

// ============================================================
// SLIDE 10: 運動2 タオル体操 セクションタイトル
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("運動 2", {
    x: 0.5, y: 1.5, w: 9, h: 0.8,
    fontSize: 24, fontFace: FONT_B, color: C.accent, align: "center", margin: 0,
  });
  s.addText("タオル体操（後方ストレッチ）", {
    x: 0.5, y: 2.3, w: 9, h: 1.2,
    fontSize: 40, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("家にあるタオル1本でできる運動です", {
    x: 0.5, y: 3.6, w: 9, h: 0.6,
    fontSize: 22, fontFace: FONT_B, color: C.light, italic: true, align: "center", margin: 0,
  });
}

// ============================================================
// SLIDE 11: タオル体操のやり方
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("タオル体操のやり方", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // 3 images
  s.addImage({ path: IMGS.towel1, x: 0.3, y: 1.1, w: 3.0, h: 3.0, sizing: { type: "contain", w: 3.0, h: 3.0 } });
  s.addImage({ path: IMGS.towel2, x: 3.5, y: 1.1, w: 3.0, h: 3.0, sizing: { type: "contain", w: 3.0, h: 3.0 } });
  s.addImage({ path: IMGS.towel3, x: 6.7, y: 1.1, w: 3.0, h: 3.0, sizing: { type: "contain", w: 3.0, h: 3.0 } });

  // Steps below
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.2, w: 9, h: 1.2, fill: { color: C.white }, shadow: makeShadow() });
  s.addText([
    { text: "1. 背中でタオルを持つ　　", options: { fontSize: 20, bold: true } },
    { text: "2. 健側で上に引く　　", options: { fontSize: 20, bold: true } },
    { text: "3. 患側の肩を後ろに伸ばす", options: { fontSize: 20, bold: true } },
  ], { x: 0.8, y: 4.3, w: 8.4, h: 1.0, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 12: タオル体操の効果
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("タオル体操の効果", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Two effect cards
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.5, w: 3.8, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.5, w: 0.08, h: 3.2, fill: { color: C.accent } });
  s.addText("内旋可動域の\n改善", {
    x: 1.2, y: 1.8, w: 3.1, h: 2.6, fontSize: 28, fontFace: FONT_H, color: C.text, align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.5, w: 3.8, h: 3.2, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.5, w: 0.08, h: 3.2, fill: { color: C.accent } });
  s.addText("肩関節拘縮の\n改善", {
    x: 5.8, y: 1.8, w: 3.1, h: 2.6, fontSize: 28, fontFace: FONT_H, color: C.text, align: "center", valign: "middle", margin: 0,
  });

  // Tip
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.9, w: 8.4, h: 0.5, fill: { color: C.light } });
  s.addText("背中に手が回しにくい方に特に効果的です", {
    x: 0.8, y: 4.9, w: 8.4, h: 0.5, fontSize: 18, fontFace: FONT_B, color: C.primary, align: "center", margin: 0,
  });
}

// ============================================================
// SLIDE 13: 運動3 壁のぼり運動 セクションタイトル
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addText("運動 3", {
    x: 0.5, y: 1.5, w: 9, h: 0.8,
    fontSize: 24, fontFace: FONT_B, color: C.accent, align: "center", margin: 0,
  });
  s.addText("壁のぼり運動", {
    x: 0.5, y: 2.3, w: 9, h: 1.2,
    fontSize: 44, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0,
  });
  s.addText("壁さえあればどこでもできます", {
    x: 0.5, y: 3.6, w: 9, h: 0.6,
    fontSize: 22, fontFace: FONT_B, color: C.light, italic: true, align: "center", margin: 0,
  });
}

// ============================================================
// SLIDE 14: 壁のぼり運動のやり方
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("壁のぼり運動のやり方", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // 3 images
  s.addImage({ path: IMGS.wall1, x: 0.3, y: 1.1, w: 3.0, h: 3.0, sizing: { type: "contain", w: 3.0, h: 3.0 } });
  s.addImage({ path: IMGS.wall2, x: 3.5, y: 1.1, w: 3.0, h: 3.0, sizing: { type: "contain", w: 3.0, h: 3.0 } });
  s.addImage({ path: IMGS.wall3, x: 6.7, y: 1.1, w: 3.0, h: 3.0, sizing: { type: "contain", w: 3.0, h: 3.0 } });

  // Steps
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.2, w: 9, h: 1.2, fill: { color: C.white }, shadow: makeShadow() });
  s.addText([
    { text: "1. 壁に指をつける　　", options: { fontSize: 20, bold: true } },
    { text: "2. 指を少しずつ上へ　　", options: { fontSize: 20, bold: true } },
    { text: "3. 痛くないところまで", options: { fontSize: 20, bold: true } },
  ], { x: 0.8, y: 4.3, w: 8.4, h: 1.0, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 15: 壁のぼり運動のポイント
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("壁のぼり運動のポイント", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Big message card
  s.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1.5, w: 8, h: 3.5, fill: { color: C.white }, shadow: makeShadow() });
  s.addText([
    { text: "痛みが出たら止める", options: { fontSize: 32, bold: true, color: C.coral, breakLine: true } },
    { text: "", options: { fontSize: 16, breakLine: true } },
    { text: "少しずつ高さを上げていく", options: { fontSize: 26, color: C.text, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "毎日の記録をつけると\nモチベーションが上がります", options: { fontSize: 20, color: C.sub } },
  ], { x: 1.5, y: 1.8, w: 7, h: 2.9, fontFace: FONT_B, align: "center", valign: "middle", margin: 0 });
}

// ============================================================
// SLIDE 16: 姿勢改善も大切
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("姿勢改善も大切", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1, fill: { color: C.light } });
  s.addText("肩甲骨の可動性低下が五十肩を悪化させます", {
    x: 0.8, y: 1.3, w: 8.4, h: 0.8, fontSize: 22, fontFace: FONT_B, color: C.coral, bold: true, align: "center", margin: 0,
  });

  // 3 cards
  const postures = [
    { title: "猫背を改善", desc: "背中を丸めない\n姿勢を意識する" },
    { title: "胸を張る", desc: "胸を開いて\n肩を後ろに引く" },
    { title: "肩甲骨を寄せる", desc: "肩甲骨を中央に\n寄せる運動をする" },
  ];
  postures.forEach((p, i) => {
    const x = 0.5 + i * 3.15;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.5, w: 2.85, h: 2.8, fill: { color: C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.5, w: 2.85, h: 0.55, fill: { color: C.accent } });
    s.addText(p.title, { x, y: 2.5, w: 2.85, h: 0.55, fontSize: 20, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(p.desc, { x: x + 0.1, y: 3.3, w: 2.65, h: 1.7, fontSize: 20, fontFace: FONT_B, color: C.text, align: "center", valign: "middle", margin: 0 });
  });
}

// ============================================================
// SLIDE 17: 五十肩の3つの時期
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("五十肩の3つの時期", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Timeline-style 3 cards
  const phases = [
    { title: "炎症期", period: "3〜6ヶ月", feature: "夜間痛が強い", color: C.coral },
    { title: "拘縮期", period: "6〜12ヶ月", feature: "肩が動かない", color: C.orange },
    { title: "回復期", period: "12ヶ月〜", feature: "徐々に改善", color: C.accent },
  ];
  phases.forEach((p, i) => {
    const x = 0.5 + i * 3.15;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.85, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.85, h: 0.7, fill: { color: p.color } });
    s.addText(p.title, { x, y: 1.3, w: 2.85, h: 0.7, fontSize: 24, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0 });
    s.addText(p.period, { x, y: 2.2, w: 2.85, h: 0.5, fontSize: 18, fontFace: FONT_B, color: C.sub, align: "center", margin: 0 });
    s.addText(p.feature, { x, y: 2.9, w: 2.85, h: 0.8, fontSize: 22, fontFace: FONT_H, color: C.text, align: "center", valign: "middle", margin: 0 });
  });

  // Arrow connectors
  s.addShape(pres.shapes.LINE, { x: 3.4, y: 1.65, w: 0.2, h: 0, line: { color: C.sub, width: 2 } });
  s.addShape(pres.shapes.LINE, { x: 6.55, y: 1.65, w: 0.2, h: 0, line: { color: C.sub, width: 2 } });
}

// ============================================================
// SLIDE 18: 時期別の運動
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("時期に合わせた運動を", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  // Table
  const tableData = [
    [
      { text: "時期", options: { fill: { color: C.primary }, color: "FFFFFF", bold: true, fontSize: 20, fontFace: FONT_H } },
      { text: "おすすめ運動", options: { fill: { color: C.primary }, color: "FFFFFF", bold: true, fontSize: 20, fontFace: FONT_H } },
      { text: "注意点", options: { fill: { color: C.primary }, color: "FFFFFF", bold: true, fontSize: 20, fontFace: FONT_H } },
    ],
    [
      { text: "炎症期", options: { bold: true, fontSize: 20, fontFace: FONT_B, color: C.coral } },
      { text: "振り子運動のみ", options: { fontSize: 20, fontFace: FONT_B } },
      { text: "無理な運動は避ける", options: { fontSize: 20, fontFace: FONT_B, color: C.coral } },
    ],
    [
      { text: "拘縮期", options: { bold: true, fontSize: 20, fontFace: FONT_B, color: C.orange } },
      { text: "ストレッチ全般\nタオル体操\n壁のぼり", options: { fontSize: 20, fontFace: FONT_B } },
      { text: "痛くない範囲で\n積極的に", options: { fontSize: 20, fontFace: FONT_B } },
    ],
    [
      { text: "回復期", options: { bold: true, fontSize: 20, fontFace: FONT_B, color: C.primary } },
      { text: "筋力訓練を追加", options: { fontSize: 20, fontFace: FONT_B } },
      { text: "可動域を広げていく", options: { fontSize: 20, fontFace: FONT_B } },
    ],
  ];
  s.addTable(tableData, {
    x: 0.5, y: 1.2, w: 9, h: 4,
    border: { pt: 1, color: "DEE2E6" },
    colW: [2.5, 3.5, 3],
    rowH: [0.7, 0.8, 1.2, 0.8],
  });
}

// ============================================================
// SLIDE 19: 毎日のおすすめ運動 トップ3
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.warmBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("毎日のおすすめ運動 トップ3", {
    x: 0.5, y: 0.1, w: 9, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
  });

  const exercises = [
    { num: "1", title: "振り子運動", desc: "力を抜いて\n腕をぶらぶら", color: C.accent },
    { num: "2", title: "壁のぼり", desc: "痛くないところまで\n指を上へ", color: C.primary },
    { num: "3", title: "入浴後ストレッチ", desc: "温まった状態で\n肩を動かす", color: "5D9B9B" },
  ];
  exercises.forEach((e, i) => {
    const x = 0.5 + i * 3.15;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.85, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
    // Number circle
    s.addShape(pres.shapes.OVAL, { x: x + 0.9, y: 1.5, w: 1.05, h: 1.05, fill: { color: e.color } });
    s.addText(e.num, { x: x + 0.9, y: 1.5, w: 1.05, h: 1.05, fontSize: 40, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(e.title, { x, y: 2.8, w: 2.85, h: 0.6, fontSize: 22, fontFace: FONT_H, color: C.text, bold: true, align: "center", margin: 0 });
    s.addText(e.desc, { x: x + 0.1, y: 3.5, w: 2.65, h: 1.3, fontSize: 18, fontFace: FONT_B, color: C.sub, align: "center", valign: "middle", margin: 0 });
  });
}

// ============================================================
// SLIDE 20: まとめ ― 今日からできること
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.565, w: 10, h: 0.06, fill: { color: C.accent } });

  s.addText("今日からできること", {
    x: 0.5, y: 0.6, w: 9, h: 0.8,
    fontSize: 36, fontFace: FONT_H, color: C.white, bold: true, align: "center", margin: 0,
  });

  // Central message card
  s.addShape(pres.shapes.RECTANGLE, { x: 1.2, y: 1.6, w: 7.6, h: 2.8, fill: { color: C.accent }, transparency: 15, shadow: makeShadow() });
  s.addText([
    { text: "五十肩は放っておくと固まる病気です。", options: { fontSize: 26, bold: true, color: C.white, breakLine: true } },
    { text: "", options: { fontSize: 12, breakLine: true } },
    { text: "痛くない範囲で\n毎日少し動かすことが一番大事です。", options: { fontSize: 28, bold: true, color: C.white } },
  ], { x: 1.5, y: 1.8, w: 7, h: 2.4, fontFace: FONT_B, align: "center", valign: "middle", margin: 0 });

  s.addText("お風呂上がりの5分で、今日から始めましょう！", {
    x: 0.5, y: 4.7, w: 9, h: 0.6,
    fontSize: 22, fontFace: FONT_B, color: C.light, italic: true, align: "center", margin: 0,
  });
}

// ============================================================
// OUTPUT
// ============================================================
pres.writeFile({ fileName: "C:/Users/jsber/OneDrive/Documents/Claude_task/五十肩_セルフケア_YouTube用.pptx" })
  .then(() => console.log("PPTX created successfully!"))
  .catch(err => console.error("Error:", err));
