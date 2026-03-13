const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaHeart, FaHandHoldingHeart, FaUsers, FaChild, FaSmile, FaBalanceScale,
  FaBrain, FaFlask, FaDna, FaStar, FaHeartbeat, FaSeedling, FaQuoteLeft,
  FaLightbulb, FaInfinity, FaBook
} = require("react-icons/fa");

function renderIconSvg(Ic, color, size = 256) {
  return ReactDOMServer.renderToStaticMarkup(React.createElement(Ic, { color, size: String(size) }));
}
async function icon(Ic, color, size = 256) {
  const svg = renderIconSvg(Ic, color.startsWith("#") ? color : "#" + color, size);
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

async function main() {
  // ── Palette ──
  const DARK = "2D1B2E", BERRY = "8B1A4A", ROSE = "C74B7A", DUSTY = "A26769";
  const LPINK = "F8E8EE", CREAM = "FFF8F5", WHITE = "FFFFFF", GOLD = "D4A574";
  const TXTD = "2D1B2E", TXTM = "5A3D4A";

  // ── Fonts ──
  const JP = "Yu Gothic", EN = "Segoe UI";

  const shd = () => ({ type: "outer", blur: 4, offset: 1.5, color: "000000", opacity: 0.08, angle: 135 });

  // ── Icons ──
  const I = {
    heart:     await icon(FaHeart,            BERRY),
    handHeart: await icon(FaHandHoldingHeart, BERRY),
    users:     await icon(FaUsers,            DUSTY),
    child:     await icon(FaChild,            GOLD),
    smile:     await icon(FaSmile,            ROSE),
    balance:   await icon(FaBalanceScale,     DUSTY),
    brain:     await icon(FaBrain,            BERRY),
    flask:     await icon(FaFlask,            ROSE),
    dna:       await icon(FaDna,              DUSTY),
    star:      await icon(FaStar,             GOLD),
    heartbeat: await icon(FaHeartbeat,        ROSE),
    seedling:  await icon(FaSeedling,         DUSTY),
    quoteG:    await icon(FaQuoteLeft,        GOLD),
    lightbulb: await icon(FaLightbulb,        GOLD),
    infinity:  await icon(FaInfinity,         BERRY),
    book:      await icon(FaBook,             BERRY),
    heartW:    await icon(FaHeart,            WHITE),
  };

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "愛とは何か？";

  // ═══ Helper: Section divider ═══
  function section(num, title, sub, notes) {
    const s = pres.addSlide();
    s.background = { color: BERRY };
    s.addImage({ data: I.heartW, x: 0.6, y: 0.4, w: 0.5, h: 0.5, transparency: 75 });
    s.addImage({ data: I.heartW, x: 8.8, y: 4.5, w: 0.4, h: 0.4, transparency: 80 });
    s.addText(num, {
      x: 0.5, y: 1.2, w: 9, h: 0.7,
      fontSize: 24, fontFace: EN, color: GOLD,
      align: "center", valign: "middle", margin: 0, charSpacing: 6
    });
    s.addText(title, {
      x: 0.5, y: 2.0, w: 9, h: 1.0,
      fontSize: 40, fontFace: JP, color: WHITE,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.LINE, { x: 3.5, y: 3.15, w: 3, h: 0, line: { color: GOLD, width: 2 } });
    s.addText(sub, {
      x: 0.5, y: 3.4, w: 9, h: 0.6,
      fontSize: 22, fontFace: JP, color: LPINK,
      align: "center", valign: "middle", margin: 0
    });
    s.addNotes(notes);
  }

  // ═══ Helper: Row item (icon + title + desc) ═══
  function addRow(s, y, icData, title, desc) {
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: y + 0.05, w: 0.7, h: 0.7, fill: { color: LPINK } });
    s.addImage({ data: icData, x: 0.9, y: y + 0.15, w: 0.5, h: 0.5 });
    s.addText(title, {
      x: 1.7, y, w: 2.5, h: 0.8,
      fontSize: 22, fontFace: JP, color: BERRY, bold: true, valign: "middle", margin: 0
    });
    s.addText(desc, {
      x: 4.3, y, w: 5.2, h: 0.8,
      fontSize: 20, fontFace: JP, color: TXTM, valign: "middle", margin: 0
    });
  }

  // ══════════════════════════════════════
  //  1. タイトル
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: DARK };
    s.addImage({ data: I.heartW, x: 4.25, y: 0.5, w: 1.0, h: 1.0, transparency: 30 });
    s.addImage({ data: I.heartW, x: 0.8, y: 0.6, w: 0.45, h: 0.45, transparency: 80 });
    s.addImage({ data: I.heartW, x: 8.5, y: 0.9, w: 0.6, h: 0.6, transparency: 75 });
    s.addText("愛とは何か？", {
      x: 0.5, y: 1.5, w: 9, h: 1.2,
      fontSize: 52, fontFace: JP, color: WHITE,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.LINE, { x: 3.0, y: 2.9, w: 4.0, h: 0, line: { color: GOLD, width: 2 } });
    s.addText("哲学・心理学・科学が解き明かす\n「愛」の正体", {
      x: 0.5, y: 3.1, w: 9, h: 1.0,
      fontSize: 24, fontFace: JP, color: LPINK,
      align: "center", valign: "middle", margin: 0
    });
    s.addText("YouTube Education Series", {
      x: 0.5, y: 4.6, w: 9, h: 0.5,
      fontSize: 20, fontFace: EN, color: GOLD,
      align: "center", valign: "middle", margin: 0, charSpacing: 4
    });
    s.addNotes(
      "皆さん、こんにちは。今日は「愛とは何か？」という人類最大のテーマの一つに挑みます。\n\n" +
      "古代ギリシャの哲学者たちから現代の神経科学まで、人類は何千年もの間、愛の正体を解き明かそうとしてきました。\n\n" +
      "この動画では、哲学・心理学・科学という3つのレンズを通して「愛」を多角的に理解していきます。"
    );
  }

  // ══════════════════════════════════════
  //  2. この動画について
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("この動画について", {
      x: 0.5, y: 0.3, w: 9, h: 0.7,
      fontSize: 30, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    const items = [
      "「愛」は誰もが経験するが、\n正確に定義するのは難しい",
      "哲学・心理学・科学の3つの視点から\n愛を多角的に分析する",
      "古代ギリシャの6つの愛の分類から\n最新の脳科学まで網羅する",
      "愛を「感じるもの」から\n「理解するもの」へ深める",
    ];
    items.forEach((t, i) => {
      const y = 1.2 + i * 1.0;
      s.addShape(pres.shapes.OVAL, { x: 1.2, y: y + 0.15, w: 0.5, h: 0.5, fill: { color: BERRY } });
      s.addText(String(i + 1), {
        x: 1.2, y: y + 0.15, w: 0.5, h: 0.5,
        fontSize: 20, fontFace: EN, color: WHITE,
        bold: true, align: "center", valign: "middle", margin: 0
      });
      s.addText(t, {
        x: 2.0, y, w: 7.2, h: 0.85,
        fontSize: 22, fontFace: JP, color: TXTD, valign: "middle", margin: 0
      });
    });
    s.addNotes(
      "なぜこの動画を作ったのか、その背景を説明します。\n\n" +
      "「愛」という言葉は日常的に使われますが、その定義を問われると答えに詰まる人が多いのではないでしょうか。\n\n" +
      "この動画では、哲学では2500年以上の歴史がある愛の議論、心理学ではスターンバーグの三角理論や愛着理論、科学ではドーパミンやオキシトシンなどの脳内物質の研究から、愛を体系的に理解していきます。\n\n" +
      "動画を見終わる頃には、「愛」に対する見方が一段と深まっていることでしょう。"
    );
  }

  // ══════════════════════════════════════
  //  3. 目次
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("目 次", {
      x: 0.5, y: 0.3, w: 9, h: 0.7,
      fontSize: 36, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    const toc = [
      { n: "01", t: "哲学の視点 — 愛の分類と名言", ic: I.heart },
      { n: "02", t: "心理学の視点 — 三角理論と愛着理論", ic: I.brain },
      { n: "03", t: "科学の視点 — 脳科学と進化", ic: I.flask },
      { n: "04", t: "まとめ — 愛の本質に迫る", ic: I.infinity },
    ];
    toc.forEach((it, i) => {
      const y = 1.4 + i * 0.9;
      s.addShape(pres.shapes.OVAL, { x: 1.5, y: y + 0.05, w: 0.55, h: 0.55, fill: { color: BERRY } });
      s.addText(it.n, {
        x: 1.5, y: y + 0.05, w: 0.55, h: 0.55,
        fontSize: 20, fontFace: EN, color: WHITE,
        bold: true, align: "center", valign: "middle", margin: 0
      });
      s.addImage({ data: it.ic, x: 2.3, y: y + 0.08, w: 0.5, h: 0.5 });
      s.addText(it.t, {
        x: 3.0, y, w: 5.5, h: 0.65,
        fontSize: 24, fontFace: JP, color: TXTD,
        bold: true, valign: "middle", margin: 0
      });
    });
    s.addNotes(
      "この動画は大きく4つのセクションで構成されています。\n\n" +
      "まず哲学の視点から、古代ギリシャの愛の分類や偉大な哲学者たちの言葉を紹介します。\n" +
      "次に心理学の視点から、スターンバーグの愛の三角理論や愛着理論を解説します。\n" +
      "そして科学の視点から、脳内物質や進化の観点で愛を分析します。\n" +
      "最後にすべてを統合してまとめます。"
    );
  }

  // ══════════════════════════════════════
  //  4. SECTION 01: 哲学の視点
  // ══════════════════════════════════════
  section("SECTION 01", "哲学の視点", "古代ギリシャから現代まで",
    "ここからは哲学の視点で「愛」を見ていきます。古代ギリシャ人は愛を6つの類型に分類しました。現代の恋愛観にも通じる深い洞察が含まれています。"
  );

  // ══════════════════════════════════════
  //  5. 古代ギリシャ人が考えた6つの愛
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("古代ギリシャ人が考えた6つの愛", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText("ギリシャ語には「愛」を表す6つの言葉があり、\nそれぞれ異なる愛の側面を指す", {
      x: 0.5, y: 0.85, w: 9, h: 0.6,
      fontSize: 20, fontFace: JP, color: TXTM,
      align: "center", valign: "middle", margin: 0
    });

    const types = [
      { name: "エロス（Eros）",       desc: "情熱的・性愛的な愛", ic: I.heart },
      { name: "フィリア（Philia）",   desc: "友情・仲間への深い愛", ic: I.users },
      { name: "アガペー（Agape）",    desc: "無条件の愛・神聖な愛", ic: I.handHeart },
      { name: "ストルゲー（Storge）", desc: "家族間の自然な愛情", ic: I.child },
      { name: "ルドゥス（Ludus）",    desc: "遊び心ある軽やかな愛", ic: I.smile },
      { name: "プラグマ（Pragma）",   desc: "長年の信頼に基づく実用的な愛", ic: I.balance },
    ];
    types.forEach((t, i) => {
      const y = 1.6 + i * 0.62;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y, w: 8.6, h: 0.55, fill: { color: i % 2 === 0 ? WHITE : LPINK }
      });
      s.addImage({ data: t.ic, x: 0.9, y: y + 0.07, w: 0.4, h: 0.4 });
      s.addText(t.name, {
        x: 1.5, y, w: 3.0, h: 0.55,
        fontSize: 20, fontFace: JP, color: TXTD, bold: true, valign: "middle", margin: 0
      });
      s.addText(t.desc, {
        x: 4.6, y, w: 4.5, h: 0.55,
        fontSize: 20, fontFace: JP, color: TXTM, valign: "middle", margin: 0
      });
    });
    s.addNotes(
      "古代ギリシャ語には、日本語や英語の「愛」とは異なり、愛を表す6つの異なる言葉がありました。\n\n" +
      "エロスは情熱的・性的な愛で、恋に落ちたときの激しい感情を指します。ギリシャ神話の愛の神エロスに由来します。\n\n" +
      "フィリアは友人や仲間との間に生まれる深い信頼と友情の愛です。アリストテレスはこれを最も価値ある愛と考えました。\n\n" +
      "アガペーは無条件の愛で、見返りを求めない利他的な愛です。後にキリスト教で「神の愛」として中心的な概念になりました。\n\n" +
      "ストルゲーは親子間に自然に生まれる家族の愛情です。ルドゥスは恋愛の初期段階にある遊び心のある軽やかな愛。プラグマは長年寄り添った夫婦に見られる実用的で成熟した愛です。\n\n" +
      "次の2枚のスライドでは、主要な3類型をさらに詳しく見ていきましょう。"
    );
  }

  // ══════════════════════════════════════
  //  6. エロス・フィリア・アガペー
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("3つの主要な愛の類型", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    const main3 = [
      { t: "エロス", en: "Eros", d: "「獲得する愛」と呼ばれ、\n美への憧れと情熱が原動力。\n自己中心的な側面を持つ。", col: BERRY, ic: I.heart },
      { t: "フィリア", en: "Philia", d: "対等な関係に基づく友愛。\n信頼と善意が基盤であり、\nアリストテレスは最も徳ある愛とした。", col: DUSTY, ic: I.users },
      { t: "アガペー", en: "Agape", d: "「与える愛」とも呼ばれ、\n見返りを求めない無条件の愛。\n神の愛として最高位に置かれる。", col: ROSE, ic: I.handHeart },
    ];
    main3.forEach((m, i) => {
      const y = 1.15 + i * 1.45;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y, w: 8.6, h: 1.25, fill: { color: WHITE }, shadow: shd()
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 0.08, h: 1.25, fill: { color: m.col } });
      s.addShape(pres.shapes.OVAL, { x: 1.0, y: y + 0.2, w: 0.75, h: 0.75, fill: { color: LPINK } });
      s.addImage({ data: m.ic, x: 1.12, y: y + 0.32, w: 0.5, h: 0.5 });
      s.addText(m.t, {
        x: 2.0, y, w: 2.2, h: 0.55,
        fontSize: 24, fontFace: JP, color: m.col, bold: true, valign: "middle", margin: 0
      });
      s.addText(m.en, {
        x: 2.0, y: y + 0.5, w: 2.2, h: 0.4,
        fontSize: 20, fontFace: EN, color: TXTM, valign: "top", margin: 0
      });
      s.addText(m.d, {
        x: 4.2, y: y + 0.1, w: 4.9, h: 1.05,
        fontSize: 20, fontFace: JP, color: TXTD, valign: "middle", margin: 0
      });
    });
    s.addNotes(
      "ここでは特に重要な3つの類型を掘り下げます。\n\n" +
      "【エロス】プラトンの『饗宴』で詳しく論じられた概念です。単なる肉体的欲求ではなく、美しいものを求め、最終的には「善のイデア」にまで上昇していく力とされました。しかし「獲得する愛」であるため、自己中心的な性質を持ちます。\n\n" +
      "【フィリア】アリストテレスが『ニコマコス倫理学』で3種類に分けました。快楽に基づく友情、利益に基づく友情、そして最高の形である徳に基づく友情です。お互いの善を心から願い合う対等な関係です。\n\n" +
      "【アガペー】新約聖書でパウロが「愛は忍耐強い、愛は情け深い」と述べた愛がまさにアガペーです。相手が誰であろうと、どんな状況であろうと変わらない無条件の愛を指します。"
    );
  }

  // ══════════════════════════════════════
  //  7. プラトンの愛の哲学
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: DARK };
    s.addImage({ data: I.quoteG, x: 1.0, y: 0.6, w: 0.7, h: 0.7 });
    s.addText("プラトンの愛の哲学", {
      x: 0.5, y: 0.3, w: 9, h: 0.7,
      fontSize: 30, fontFace: JP, color: WHITE,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    // Quote
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.0, y: 1.3, w: 8.0, h: 1.3, fill: { color: DARK, transparency: 30 }
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 1.3, w: 0.08, h: 1.3, fill: { color: GOLD } });
    s.addText("「愛とは善きものの\n永久の所有へ向けられたもの」", {
      x: 1.4, y: 1.3, w: 7.3, h: 0.8,
      fontSize: 24, fontFace: JP, color: WHITE, valign: "middle", margin: 0
    });
    s.addText("― プラトン『饗宴』", {
      x: 1.4, y: 2.1, w: 7.3, h: 0.45,
      fontSize: 20, fontFace: JP, color: GOLD, valign: "middle", margin: 0
    });

    // Key points
    s.addText("愛は肉体の美から始まり、やがて\n精神の美、学問の美へと上昇していく\n「エロスの梯子」として描かれた。\n愛とは不完全な存在が完全性を求める力である。", {
      x: 1.0, y: 2.9, w: 8.0, h: 1.6,
      fontSize: 22, fontFace: JP, color: LPINK, valign: "top", margin: 0
    });

    s.addNotes(
      "プラトンの『饗宴』は、西洋哲学における愛の議論の出発点です。\n\n" +
      "この対話篇では、宴会の席で参加者がそれぞれ愛（エロス）についてスピーチを行います。中でもソクラテスが紹介するディオティマの教えが最も重要です。\n\n" +
      "ディオティマによれば、愛とは「美しいものの中での出産」を求める力です。まず一つの美しい肉体に惹かれ、次にあらゆる肉体の美を認め、さらに精神の美、制度の美、学問の美へと上昇し、最終的に「美のイデア」そのものに到達する。これが有名な「エロスの梯子」です。\n\n" +
      "プラトンにとって愛は、不完全な存在である人間が完全なるものを求める根源的な力だったのです。"
    );
  }

  // ══════════════════════════════════════
  //  8. アリストテレスからフロムへ
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: DARK };
    s.addText("哲学者たちの「愛」の定義", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 30, fontFace: JP, color: WHITE,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    const quotes = [
      { p: "アリストテレス", q: "愛とは、相手のために\n善を願うことである" },
      { p: "ライプニッツ", q: "愛とは、他者の幸福を\n喜ぶことである" },
      { p: "エーリッヒ・フロム", q: "愛は技術であり、\n学び実践するものである" },
    ];
    quotes.forEach((item, i) => {
      const y = 1.2 + i * 1.3;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 1.0, y, w: 8.0, h: 1.1, fill: { color: DARK, transparency: 30 }
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y, w: 0.08, h: 1.1, fill: { color: GOLD } });
      s.addImage({ data: I.quoteG, x: 1.3, y: y + 0.15, w: 0.45, h: 0.45 });
      s.addText("「" + item.q + "」", {
        x: 1.95, y, w: 6.8, h: 0.65,
        fontSize: 22, fontFace: JP, color: WHITE, valign: "middle", margin: 0
      });
      s.addText("― " + item.p, {
        x: 1.95, y: y + 0.6, w: 6.8, h: 0.45,
        fontSize: 20, fontFace: JP, color: GOLD, valign: "middle", margin: 0
      });
    });

    s.addNotes(
      "哲学史を通じて、多くの哲学者が愛を定義しようと試みてきました。\n\n" +
      "【アリストテレス】愛の本質を「相手のために善を願うこと」としました。これは自分のためではなく、相手そのもののために善を願う姿勢です。利己的な愛ではなく、利他的な愛こそが真の愛だとしました。\n\n" +
      "【ライプニッツ】17世紀ドイツの哲学者ライプニッツは「他者の幸福を喜ぶこと」が愛だと定義しました。相手が幸せであること自体に喜びを見出す感情です。\n\n" +
      "【エーリッヒ・フロム】20世紀の社会心理学者フロムは1956年の著書『愛するということ』で、愛は自然に生まれる感情ではなく学習と実践によって身につける「技術」だと主張しました。音楽や医術と同じように、理論を学び、実践し、修練を積む必要があるのです。次のスライドでフロムの思想を詳しく見ていきます。"
    );
  }

  // ══════════════════════════════════════
  //  9. フロム「愛するということ」
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addImage({ data: I.book, x: 0.7, y: 0.3, w: 0.5, h: 0.5 });
    s.addText("フロム『愛するということ』の教え", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText("フロムが示す「愛の技術」に必要な4つの要素", {
      x: 0.5, y: 0.9, w: 9, h: 0.5,
      fontSize: 20, fontFace: JP, color: TXTM,
      align: "center", valign: "middle", margin: 0
    });

    const elems = [
      { t: "配慮（Care）", d: "愛する者の成長と幸福に\n積極的に関わること" },
      { t: "責任（Responsibility）", d: "相手のニーズに自発的に\n応答する姿勢" },
      { t: "尊重（Respect）", d: "相手をありのままに認め、\n支配しないこと" },
      { t: "知（Knowledge）", d: "相手を深く知ろうとする\n不断の努力" },
    ];
    elems.forEach((e, i) => {
      const y = 1.6 + i * 0.9;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y, w: 8.6, h: 0.75, fill: { color: WHITE }, shadow: shd()
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 0.08, h: 0.75, fill: { color: BERRY } });
      s.addText(e.t, {
        x: 1.0, y, w: 3.0, h: 0.75,
        fontSize: 22, fontFace: JP, color: BERRY, bold: true, valign: "middle", margin: 0
      });
      s.addText(e.d, {
        x: 4.1, y, w: 5.0, h: 0.75,
        fontSize: 20, fontFace: JP, color: TXTD, valign: "middle", margin: 0
      });
    });

    s.addNotes(
      "エーリッヒ・フロムは1956年に出版した『愛するということ（The Art of Loving）』で、愛の4つの基本要素を示しました。\n\n" +
      "【配慮】愛する人の生命と成長に積極的に関わること。植物を育てるように、相手の成長を見守り手助けすることです。\n\n" +
      "【責任】相手の求めに応じる能力と意思のこと。義務感からではなく、自発的に相手のニーズに応えようとする姿勢です。\n\n" +
      "【尊重】相手をありのままに受け止め、その人らしさを認めること。相手を自分の都合のいいように変えようとしない態度です。\n\n" +
      "【知】表面的な理解ではなく、相手の本質を深く知ろうとする不断の努力です。\n\n" +
      "フロムは、現代人の多くが『愛される方法』ばかり考えて、『愛する方法』を学ぼうとしないと批判しました。愛は受動的な感情ではなく、能動的な技術なのです。"
    );
  }

  // ══════════════════════════════════════
  //  10. SECTION 02: 心理学の視点
  // ══════════════════════════════════════
  section("SECTION 02", "心理学の視点", "愛の三角理論と愛着理論",
    "ここからは心理学の視点に移ります。心理学は愛を科学的に測定し、分類することに挑みました。特にスターンバーグの愛の三角理論と、ボウルビィの愛着理論は愛の理解に大きく貢献しています。"
  );

  // ══════════════════════════════════════
  //  11. 愛の三角理論 — 3つの要素
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("スターンバーグの愛の三角理論", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText("ロバート・スターンバーグ（1986）は、\n愛を3つの成分で説明した", {
      x: 0.5, y: 0.85, w: 9, h: 0.6,
      fontSize: 20, fontFace: JP, color: TXTM,
      align: "center", valign: "middle", margin: 0
    });

    const tri = [
      { t: "親密さ（Intimacy）", d: "相手との結びつき、\n温かさ、信頼感。\n友情の基盤にもなる感情的なつながり。", col: BERRY },
      { t: "情熱（Passion）", d: "身体的魅力やロマンスによる\n強い感情的・性的な引力。\n恋愛初期に特に強い。", col: ROSE },
      { t: "コミットメント\n（Commitment）", d: "関係を維持する意志的な決断。\n短期は「愛する決断」、\n長期は「関係を守る決断」。", col: GOLD },
    ];
    tri.forEach((t, i) => {
      const y = 1.6 + i * 1.2;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y, w: 8.6, h: 1.0, fill: { color: WHITE }, shadow: shd()
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 0.08, h: 1.0, fill: { color: t.col } });
      s.addText(t.t, {
        x: 1.0, y, w: 3.5, h: 1.0,
        fontSize: 22, fontFace: JP, color: t.col, bold: true, valign: "middle", margin: 0
      });
      s.addText(t.d, {
        x: 4.5, y, w: 4.6, h: 1.0,
        fontSize: 20, fontFace: JP, color: TXTD, valign: "middle", margin: 0
      });
    });

    s.addNotes(
      "1986年にイェール大学の心理学者ロバート・スターンバーグが提唱した「愛の三角理論」は、愛の構造を理解するための最も有力な枠組みの一つです。\n\n" +
      "【親密さ】温かさ、信頼感、お互いを深く理解しているという感覚です。友情にも恋愛にも共通する基盤的な要素です。\n\n" +
      "【情熱】ロマンティックな引力、身体的な魅力への反応です。恋愛の「燃え上がる」感覚がこれにあたります。恋愛初期に最も強く、時間とともに変化していきます。\n\n" +
      "【コミットメント】愛の認知的・意志的な側面です。「この人を愛している」という短期的な認識と、「この関係を守り続ける」という長期的な決断の両方を含みます。\n\n" +
      "次のスライドでは、この3要素の組み合わせから生まれる8つの愛のかたちを見ていきます。"
    );
  }

  // ══════════════════════════════════════
  //  12. 8つの愛のかたち
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("三角理論が示す8つの愛のかたち", {
      x: 0.5, y: 0.15, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    const kinds = [
      { name: "非愛",             combo: "なし",               ex: "知人・他人" },
      { name: "好意",             combo: "親密さのみ",          ex: "親しい友人" },
      { name: "心酔",             combo: "情熱のみ",            ex: "一目惚れ" },
      { name: "空虚な愛",         combo: "コミットメントのみ",    ex: "義務的な結婚" },
      { name: "ロマンティックな愛", combo: "親密さ＋情熱",        ex: "恋人関係の初期" },
      { name: "友愛的な愛",       combo: "親密さ＋コミットメント", ex: "長年の伴侶" },
      { name: "愚かな愛",         combo: "情熱＋コミットメント",   ex: "電撃婚" },
      { name: "完全な愛",         combo: "3つすべて",            ex: "理想の関係" },
    ];

    // Header
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 0.9, w: 8.6, h: 0.5, fill: { color: BERRY } });
    s.addText("愛のかたち", { x: 0.9, y: 0.9, w: 2.5, h: 0.5, fontSize: 20, fontFace: JP, color: WHITE, bold: true, valign: "middle", margin: 0 });
    s.addText("構成要素", { x: 3.5, y: 0.9, w: 3.0, h: 0.5, fontSize: 20, fontFace: JP, color: WHITE, bold: true, valign: "middle", margin: 0 });
    s.addText("例", { x: 6.8, y: 0.9, w: 2.3, h: 0.5, fontSize: 20, fontFace: JP, color: WHITE, bold: true, valign: "middle", margin: 0 });

    kinds.forEach((k, i) => {
      const y = 1.4 + i * 0.5;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 8.6, h: 0.48, fill: { color: i % 2 === 0 ? WHITE : LPINK } });
      s.addText(k.name, { x: 0.9, y, w: 2.5, h: 0.48, fontSize: 20, fontFace: JP, color: TXTD, bold: true, valign: "middle", margin: 0 });
      s.addText(k.combo, { x: 3.5, y, w: 3.0, h: 0.48, fontSize: 20, fontFace: JP, color: TXTM, valign: "middle", margin: 0 });
      s.addText(k.ex, { x: 6.8, y, w: 2.3, h: 0.48, fontSize: 20, fontFace: JP, color: TXTM, valign: "middle", margin: 0 });
    });

    s.addNotes(
      "3つの要素の有無により、理論上8つの異なる愛のかたちが生まれます。\n\n" +
      "【非愛】3要素のいずれもない状態。単なる知り合いとの関係です。\n" +
      "【好意】親密さだけがある状態で、親しい友情に相当します。\n" +
      "【心酔】情熱だけの状態。一目惚れや芸能人への熱狂がこれにあたります。\n" +
      "【空虚な愛】コミットメントだけの愛。情熱も親密さも失われた結婚生活に見られます。\n\n" +
      "【ロマンティックな愛】親密さと情熱の組み合わせ。恋人関係の初期に多い形です。\n" +
      "【友愛的な愛】親密さとコミットメントの組み合わせ。情熱は薄れたが深い信頼で結ばれた長年の夫婦に見られます。\n" +
      "【愚かな愛】情熱とコミットメントの組み合わせ。十分に知り合う前に結婚を決めるケースです。\n" +
      "【完全な愛】3要素すべてが備わった理想的な関係。達成は可能だが維持には努力が必要です。"
    );
  }

  // ══════════════════════════════════════
  //  13. 愛着理論
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("愛着理論（アタッチメント理論）", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText("幼少期の親との関係パターンが、\n大人の恋愛スタイルを形作る", {
      x: 0.5, y: 0.85, w: 9, h: 0.6,
      fontSize: 20, fontFace: JP, color: TXTM,
      align: "center", valign: "middle", margin: 0
    });

    const att = [
      { t: "安定型（Secure）", d: "信頼関係を築きやすく、\n親密さと自律のバランスが取れる。\n人口の約50%がこのタイプとされる。", col: DUSTY },
      { t: "回避型（Avoidant）", d: "親密さを避け独立性を重視する。\n感情を抑制しがちで、\n距離を置くことで安全を確保する。", col: BERRY },
      { t: "不安型（Anxious）", d: "見捨てられることへの恐怖が強い。\n関係に過度に執着し、\n相手の反応に敏感に反応する。", col: ROSE },
    ];
    att.forEach((a, i) => {
      const y = 1.6 + i * 1.2;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y, w: 8.6, h: 1.0, fill: { color: WHITE }, shadow: shd()
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 0.08, h: 1.0, fill: { color: a.col } });
      s.addText(a.t, {
        x: 1.0, y, w: 3.0, h: 1.0,
        fontSize: 22, fontFace: JP, color: a.col, bold: true, valign: "middle", margin: 0
      });
      s.addText(a.d, {
        x: 4.1, y, w: 5.0, h: 1.0,
        fontSize: 20, fontFace: JP, color: TXTD, valign: "middle", margin: 0
      });
    });

    s.addNotes(
      "愛着理論はイギリスの精神科医ジョン・ボウルビィが提唱し、後にメアリー・エインズワースが実験で実証しました。1980年代にはシェイバーとハザンが大人の恋愛関係にも同じ3パターンが見られることを発見しました。\n\n" +
      "【安定型】幼少期に養育者から一貫した応答を受けた人に多い。親密さを心地よく感じ、信頼に基づく関係を築ける。\n\n" +
      "【回避型】養育者が感情的に距離を置いていた環境で育った人に多い。深い関係を「窮屈」と感じやすい。\n\n" +
      "【不安型】養育者の応答が不安定だった環境で育った人に多い。相手に見捨てられないか常に不安を感じる。\n\n" +
      "重要なのは、愛着スタイルは固定ではなく、自己理解と意識的な努力で変化しうるということです。"
    );
  }

  // ══════════════════════════════════════
  //  14. SECTION 03: 科学の視点
  // ══════════════════════════════════════
  section("SECTION 03", "科学の視点", "脳科学と進化生物学",
    "ここからは科学の視点で愛を見ていきます。現代の神経科学は、愛という複雑な感情の背後にある脳内メカニズムを解明しつつあります。"
  );

  // ══════════════════════════════════════
  //  15. ドーパミン
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("ドーパミン — 恋の高揚感の正体", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: 1.1, w: 8.6, h: 4.0, fill: { color: WHITE }, shadow: shd()
    });
    s.addShape(pres.shapes.OVAL, { x: 1.2, y: 1.4, w: 1.2, h: 1.2, fill: { color: LPINK } });
    s.addImage({ data: I.star, x: 1.4, y: 1.6, w: 0.8, h: 0.8 });

    const dopa = [
      "恋愛初期に大量に分泌され、\n強い幸福感と高揚感をもたらす",
      "脳の報酬系を活性化し、\n相手に「依存」するような状態を作る",
      "コカインなどの薬物と\n同じ脳領域を刺激することが判明",
      "「もっと一緒にいたい」という\n衝動の神経科学的な基盤",
    ];
    dopa.forEach((d, i) => {
      const y = 1.5 + i * 0.85;
      s.addShape(pres.shapes.OVAL, { x: 2.8, y: y + 0.12, w: 0.2, h: 0.2, fill: { color: ROSE } });
      s.addText(d, {
        x: 3.2, y, w: 5.8, h: 0.75,
        fontSize: 20, fontFace: JP, color: TXTD, valign: "middle", margin: 0
      });
    });

    s.addNotes(
      "恋に落ちたとき、私たちの脳ではドーパミンが大量に放出されます。\n\n" +
      "ドーパミンは「報酬系」と呼ばれる脳の回路を活性化し、強い快感と幸福感をもたらします。恋愛初期の「相手のことが頭から離れない」「会えないと苦しい」という感覚は、まさにドーパミンの作用です。\n\n" +
      "興味深いことに、脳画像研究では、恋愛中の人の脳活動パターンはコカインなどの依存性薬物を使用した時と非常に似ていることが分かっています。つまり恋は一種の「依存状態」なのです。\n\n" +
      "このドーパミンによる高揚感は通常12〜18ヶ月で徐々に落ち着いていきます。しかし、これは愛が終わったのではなく、次に紹介するオキシトシンによるより深い絆の段階へと移行しているのです。"
    );
  }

  // ══════════════════════════════════════
  //  16. オキシトシン
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("オキシトシン — 絆と信頼のホルモン", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: 1.1, w: 8.6, h: 4.0, fill: { color: WHITE }, shadow: shd()
    });
    s.addShape(pres.shapes.OVAL, { x: 1.2, y: 1.4, w: 1.2, h: 1.2, fill: { color: LPINK } });
    s.addImage({ data: I.heartbeat, x: 1.4, y: 1.6, w: 0.8, h: 0.8 });

    const oxy = [
      "「愛情ホルモン」「絆ホルモン」\nとして広く知られる",
      "スキンシップ、授乳、\n信頼関係の構築時に分泌される",
      "パートナーへの愛着を深め、\n長期的な絆を維持する役割",
      "オキシトシンが分泌されるほど\n愛ある行動が増える好循環",
    ];
    oxy.forEach((o, i) => {
      const y = 1.5 + i * 0.85;
      s.addShape(pres.shapes.OVAL, { x: 2.8, y: y + 0.12, w: 0.2, h: 0.2, fill: { color: ROSE } });
      s.addText(o, {
        x: 3.2, y, w: 5.8, h: 0.75,
        fontSize: 20, fontFace: JP, color: TXTD, valign: "middle", margin: 0
      });
    });

    s.addNotes(
      "オキシトシンはドーパミンとは異なる役割を果たす「愛情ホルモン」です。\n\n" +
      "ドーパミンが恋愛の「燃え上がる」フェーズを担当するのに対し、オキシトシンは「穏やかで深い絆」のフェーズを支えます。抱擁、キス、スキンシップなどの身体接触で分泌が促進されます。\n\n" +
      "母親が授乳するときにもオキシトシンが大量に放出され、母子の絆を強化します。つまりロマンティックな愛と親子の愛は、神経化学的に共通の基盤を持っているのです。\n\n" +
      "さらに興味深いのは、オキシトシンが「好循環」を生むことです。愛情ある行動をするとオキシトシンが分泌され、オキシトシンが増えるとさらに愛情ある行動をしたくなる。このポジティブフィードバックが長期的な絆を支えています。"
    );
  }

  // ══════════════════════════════════════
  //  17. 脳と進化
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("脳と進化から見る愛", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 30, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    // VTA
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: 1.15, w: 8.6, h: 1.7, fill: { color: WHITE }, shadow: shd()
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.15, w: 0.08, h: 1.7, fill: { color: BERRY } });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: 1.4, w: 0.9, h: 0.9, fill: { color: LPINK } });
    s.addImage({ data: I.brain, x: 1.15, y: 1.55, w: 0.6, h: 0.6 });
    s.addText("腹側被蓋野（VTA）", {
      x: 2.2, y: 1.15, w: 3.5, h: 0.75,
      fontSize: 22, fontFace: JP, color: BERRY, bold: true, valign: "middle", margin: 0
    });
    s.addText("脳の深部にある報酬系の中枢。\n恋愛感情の源であり、母性愛にも関与する。\n「愛は本能の一部」を示す重要な証拠。", {
      x: 2.2, y: 1.85, w: 6.9, h: 0.95,
      fontSize: 20, fontFace: JP, color: TXTD, valign: "top", margin: 0
    });

    // Evolution
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: 3.15, w: 8.6, h: 1.7, fill: { color: WHITE }, shadow: shd()
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 3.15, w: 0.08, h: 1.7, fill: { color: DUSTY } });
    s.addShape(pres.shapes.OVAL, { x: 1.0, y: 3.4, w: 0.9, h: 0.9, fill: { color: LPINK } });
    s.addImage({ data: I.dna, x: 1.15, y: 3.55, w: 0.6, h: 0.6 });
    s.addText("進化的適応としての愛", {
      x: 2.2, y: 3.15, w: 3.5, h: 0.75,
      fontSize: 22, fontFace: JP, color: DUSTY, bold: true, valign: "middle", margin: 0
    });
    s.addText("愛は種の存続に不可欠な生物学的メカニズム。\n協力、子育て、集団の結束を促し、\n人類の生存を支えてきた適応戦略。", {
      x: 2.2, y: 3.85, w: 6.9, h: 0.95,
      fontSize: 20, fontFace: JP, color: TXTD, valign: "top", margin: 0
    });

    s.addNotes(
      "【腹側被蓋野（VTA）】\n脳の中脳にある小さな領域ですが、愛の神経科学において極めて重要な部位です。人類学者ヘレン・フィッシャーのfMRI研究により、恋愛中の人がパートナーの写真を見たときにVTAが強く活性化することが確認されました。\n\n" +
      "注目すべきは、VTAの活性化パターンが母親が子供を見たときにも似ていることです。ロマンティックな愛と母性愛は、脳の同じ報酬系を共有しているのです。\n\n" +
      "【進化的適応】\n進化生物学の視点では、愛は「遺伝子を次世代に伝える」ための適応メカニズムです。人間の赤ちゃんは非常に長い期間、親の世話を必要とします。パートナー間の愛の絆は、子育ての協力体制を維持するために進化したと考えられています。\n\n" +
      "また、集団への帰属意識（友愛）も、厳しい自然環境で協力して生き延びるための適応です。愛は「あれば素敵なもの」ではなく、人類の生存に不可欠な生物学的戦略なのです。"
    );
  }

  // ══════════════════════════════════════
  //  18. SECTION 04: まとめ
  // ══════════════════════════════════════
  section("SECTION 04", "現代の理解とまとめ", "愛の本質に迫る",
    "最後のセクションでは、現代哲学の新しい理論、愛がもたらす具体的な効果、そして全体のまとめをお届けします。"
  );

  // ══════════════════════════════════════
  //  19. 現代哲学の新理論
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("現代哲学の3つの愛の理論", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    const th = [
      { t: "合一論（Union View）", d: "愛する2人が「私たち」という\n新たな主体を形成する。\n個の境界を超えた一体化を愛の本質と捉える。", col: BERRY },
      { t: "配慮の理論\n（Robust Concern）", d: "相手の幸福を深く気にかけ行動すること。\nただし相手の自律性を尊重し、\n支配とは異なる。", col: ROSE },
      { t: "意志の理論\n（Volitional View）", d: "愛は感情ではなく意志に基づく選択。\n日々「愛する」と決断し続ける\n行為そのものが愛である。", col: GOLD },
    ];
    th.forEach((t, i) => {
      const y = 1.15 + i * 1.4;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y, w: 8.6, h: 1.2, fill: { color: WHITE }, shadow: shd()
      });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 0.08, h: 1.2, fill: { color: t.col } });
      s.addText(t.t, {
        x: 1.0, y, w: 3.2, h: 1.2,
        fontSize: 22, fontFace: JP, color: t.col, bold: true, valign: "middle", margin: 0
      });
      s.addText(t.d, {
        x: 4.3, y, w: 4.8, h: 1.2,
        fontSize: 20, fontFace: JP, color: TXTD, valign: "middle", margin: 0
      });
    });

    s.addNotes(
      "20世紀以降、哲学者たちは古代ギリシャの枠組みを超えた新しい愛の理論を構築してきました。\n\n" +
      "【合一論】ロバート・ノージックらが支持する理論です。愛する2人は「私」と「あなた」ではなく「私たち」という新しい主体を形成する。財産、喜び、悲しみ、人生の目標をすべて共有する状態を愛と捉えます。批判としては、個人のアイデンティティが失われるリスクがあるという指摘があります。\n\n" +
      "【配慮の理論】ハリー・フランクファートらが主張する理論です。愛の核心は相手の幸福を自分のことのように気にかける態度にあるとします。ただし、合一論と異なり、相手の自律性を尊重することが重要だとします。\n\n" +
      "【意志の理論】愛は移ろいやすい感情ではなく、意志に基づく決断だとする立場です。フロムの「愛は技術である」という主張にも通じます。毎日「この人を愛する」と選択し続ける行為そのものが愛なのです。"
    );
  }

  // ══════════════════════════════════════
  //  20. 愛がもたらす効果
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: CREAM };
    s.addText("愛がもたらす心身への効果", {
      x: 0.5, y: 0.25, w: 9, h: 0.7,
      fontSize: 28, fontFace: JP, color: BERRY,
      bold: true, align: "center", valign: "middle", margin: 0
    });

    addRow(s, 1.15, I.star,      "幸福感の向上",     "自己肯定感と生きがいが高まり、\n人生の満足度が上がる");
    addRow(s, 2.1,  I.heartbeat, "ストレス軽減",     "コルチゾールの低下、血圧低下、\nリラックス効果をもたらす");
    addRow(s, 3.05, I.seedling,  "健康増進と長寿",   "免疫力が向上し、心疾患リスクが低下。\n寿命延長にも寄与する");
    addRow(s, 4.0,  I.lightbulb, "人格の成長",       "共感力が深まり、思いやりの心と\n自己成長が促進される");

    s.addNotes(
      "愛は単なる感情ではなく、具体的な健康効果をもたらすことが多くの研究で示されています。\n\n" +
      "【幸福感】安定したパートナー関係にある人は、そうでない人に比べて主観的幸福度が有意に高いことが縦断研究で示されています。\n\n" +
      "【ストレス軽減】パートナーと手を繋ぐだけでストレスホルモンであるコルチゾールの値が低下することがバージニア大学の研究で確認されています。\n\n" +
      "【健康と長寿】ハーバード大学の75年にわたる追跡調査「Grant Study」では、人生の幸福と健康を最も強く予測する因子は、富でも名声でもなく「温かい人間関係」であることが示されました。\n\n" +
      "【人格の成長】愛する経験は、共感能力を高め、自己中心的な視点を超えて他者の立場に立つ力を育てます。"
    );
  }

  // ══════════════════════════════════════
  //  21. まとめ — 愛の本質とは
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: DARK };
    s.addImage({ data: I.heartW, x: 0.5, y: 0.3, w: 0.4, h: 0.4, transparency: 80 });
    s.addImage({ data: I.heartW, x: 9.0, y: 0.5, w: 0.5, h: 0.5, transparency: 75 });
    s.addText("まとめ — 愛の本質とは", {
      x: 0.5, y: 0.2, w: 9, h: 0.7,
      fontSize: 30, fontFace: JP, color: WHITE,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.LINE, { x: 3.5, y: 1.0, w: 3, h: 0, line: { color: GOLD, width: 2 } });

    const pts = [
      "愛は単一の感情ではなく、\n多面的で複層的な現象である",
      "哲学・心理学・科学は\n異なる角度から同じ真実を照射する",
      "求める愛と与える愛は、\n同一現象の異なる側面にすぎない",
      "愛を理解し、意識的に実践することが\n豊かな人生をつくる",
    ];
    pts.forEach((p, i) => {
      const y = 1.3 + i * 1.0;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 1.0, y, w: 8.0, h: 0.8, fill: { color: DARK, transparency: 30 }
      });
      s.addShape(pres.shapes.OVAL, { x: 1.2, y: y + 0.15, w: 0.5, h: 0.5, fill: { color: GOLD } });
      s.addText(String(i + 1), {
        x: 1.2, y: y + 0.15, w: 0.5, h: 0.5,
        fontSize: 20, fontFace: EN, color: DARK,
        bold: true, align: "center", valign: "middle", margin: 0
      });
      s.addText(p, {
        x: 2.0, y, w: 6.8, h: 0.8,
        fontSize: 22, fontFace: JP, color: WHITE, valign: "middle", margin: 0
      });
    });

    s.addNotes(
      "ここまで哲学・心理学・科学の3つの視点から愛を分析してきました。\n\n" +
      "哲学はエロス、フィリア、アガペーという愛の多様な姿を明らかにしました。心理学は親密さ、情熱、コミットメントという構成要素を特定しました。科学はドーパミンやオキシトシンという脳内メカニズムを解明しました。\n\n" +
      "一見バラバラに見えるこれらの知見は、実は同じ現象を異なるレンズで見ているだけです。プラトンの「善きものへの希求」も、フロムの「愛する技術」も、ドーパミンの報酬系も、すべて「愛」という一つの現象の異なる側面を捉えています。\n\n" +
      "最も大切なメッセージは、愛は受動的に「降ってくる」ものではなく、理解し、学び、日々実践するものだということです。フロムが教えてくれたように、愛は技術なのです。"
    );
  }

  // ══════════════════════════════════════
  //  22. おわりに
  // ══════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: DARK };
    s.addImage({ data: I.heartW, x: 4.25, y: 0.6, w: 1.2, h: 1.2, transparency: 25 });
    s.addText("ご視聴ありがとうございました", {
      x: 0.5, y: 2.0, w: 9, h: 1.0,
      fontSize: 36, fontFace: JP, color: WHITE,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.LINE, { x: 3.0, y: 3.2, w: 4, h: 0, line: { color: GOLD, width: 2 } });
    s.addText("チャンネル登録・高評価お願いします", {
      x: 0.5, y: 3.5, w: 9, h: 0.6,
      fontSize: 22, fontFace: JP, color: LPINK,
      align: "center", valign: "middle", margin: 0
    });
    s.addText("コメント欄であなたの「愛の定義」を\n教えてください", {
      x: 0.5, y: 4.2, w: 9, h: 0.7,
      fontSize: 22, fontFace: JP, color: GOLD,
      align: "center", valign: "middle", margin: 0
    });

    s.addNotes(
      "最後までご視聴いただきありがとうございました。\n\n" +
      "今日の動画では「愛とは何か」を哲学・心理学・科学の3つの視点から探求しました。\n\n" +
      "もしこの動画が参考になりましたら、チャンネル登録と高評価をお願いします。\n\n" +
      "そして、ぜひコメント欄であなたなりの「愛の定義」を教えてください。2500年の歴史を持つこの問いに、一人ひとりの答えがあるはずです。\n\n" +
      "それでは、また次の動画でお会いしましょう。"
    );
  }

  // ── Write ──
  const out = process.cwd() + "/愛とは何か_YouTube用資料.pptx";
  await pres.writeFile({ fileName: out });
  console.log("Created: " + out);
}

main().catch(err => { console.error(err); process.exit(1); });
