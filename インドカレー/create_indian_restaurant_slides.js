const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaPassport, FaUsers, FaStore, FaFlag,
  FaUtensils, FaBreadSlice,
  FaEye, FaSmile, FaCoins, FaBrain,
  FaGlobeAsia, FaExchangeAlt, FaLeaf, FaArrowDown
} = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// --- Color palette: Warm Terracotta (spice/curry theme) ---
const C = {
  primary:    "B35C1E",
  dark:       "3D1F0A",
  darkBg:     "2A1608",
  light:      "FDF6F0",
  sand:       "E8D5C4",
  accent:     "D4A017",
  white:      "FFFFFF",
  textDark:   "2A1608",
  textLight:  "F5EDE5",
  textMuted:  "8B7355",
  green:      "5A8A3C",
};

const JP = "Yu Gothic";
const EN = "Segoe UI";
const mkShadow = () => ({ type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.12 });

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Indian Restaurant Blog";
  pres.title = "なぜ日本のインド料理店はネパール人経営が多いのか";

  const icons = {
    passport:  await iconToBase64Png(FaPassport, "#" + C.accent, 256),
    users:     await iconToBase64Png(FaUsers, "#" + C.accent, 256),
    store:     await iconToBase64Png(FaStore, "#" + C.accent, 256),
    flag:      await iconToBase64Png(FaFlag, "#" + C.accent, 256),
    utensils:  await iconToBase64Png(FaUtensils, "#" + C.white, 256),
    bread:     await iconToBase64Png(FaBreadSlice, "#" + C.white, 256),
    eye:       await iconToBase64Png(FaEye, "#" + C.primary, 256),
    smile:     await iconToBase64Png(FaSmile, "#" + C.primary, 256),
    coins:     await iconToBase64Png(FaCoins, "#" + C.primary, 256),
    brain:     await iconToBase64Png(FaBrain, "#" + C.primary, 256),
    globe:     await iconToBase64Png(FaGlobeAsia, "#" + C.white, 256),
    exchange:  await iconToBase64Png(FaExchangeAlt, "#" + C.white, 256),
    leaf:      await iconToBase64Png(FaLeaf, "#" + C.green, 256),
    arrowDown: await iconToBase64Png(FaArrowDown, "#" + C.sand, 256),
  };

  // ============================================================
  // SLIDE 1: Title
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    s.addImage({ data: icons.utensils, x: 3.8, y: 0.6, w: 0.7, h: 0.7 });
    s.addImage({ data: icons.bread, x: 4.65, y: 0.6, w: 0.7, h: 0.7 });
    s.addImage({ data: icons.globe, x: 5.5, y: 0.6, w: 0.7, h: 0.7 });

    s.addText("なぜ日本の「インド料理店」は\nネパール人経営が多いのか", {
      x: 0.8, y: 1.6, w: 8.4, h: 1.6,
      fontSize: 34, fontFace: JP, color: C.white, bold: true,
      align: "center", valign: "middle", lineSpacingMultiple: 1.3
    });

    s.addText("そして、ナンの意外な真実", {
      x: 0.8, y: 3.3, w: 8.4, h: 0.6,
      fontSize: 24, fontFace: JP, color: C.accent,
      align: "center", valign: "middle"
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 3.5, y: 4.1, w: 3, h: 0.04, fill: { color: C.sand }
    });

    s.addText("出典：東京大学 カレル研究 / ブリタニカ百科事典 ほか", {
      x: 0.8, y: 4.4, w: 8.4, h: 0.6,
      fontSize: 18, fontFace: JP, color: C.textMuted, align: "center", valign: "middle"
    });
  }

  // ============================================================
  // SLIDE 2: Problem Statement
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.light };

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 0.4, w: 2.4, h: 0.45,
      fill: { color: C.primary }
    });
    s.addText("QUESTION", {
      x: 0.6, y: 0.4, w: 2.4, h: 0.45,
      fontSize: 18, fontFace: EN, color: C.white, bold: true,
      align: "center", valign: "middle"
    });

    s.addText("日本のインド料理店の多くは\nインド人ではなくネパール人が経営", {
      x: 0.6, y: 1.2, w: 8.8, h: 1.5,
      fontSize: 30, fontFace: JP, color: C.textDark, bold: true,
      align: "left", valign: "middle", lineSpacingMultiple: 1.4
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 2.9, w: 2.5, h: 0.04, fill: { color: C.accent }
    });

    s.addText("しかも「インドの主食」と信じている\nナンは、実際の日常食と異なる", {
      x: 0.6, y: 3.2, w: 8.8, h: 1.4,
      fontSize: 24, fontFace: JP, color: C.textMuted,
      align: "left", valign: "top", lineSpacingMultiple: 1.5
    });
  }

  // ============================================================
  // SLIDE 3: Reason 1 & 2
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.light };

    s.addText("ネパール人経営が増えた構造的理由 ①②", {
      x: 0.6, y: 0.3, w: 8.8, h: 0.7,
      fontSize: 26, fontFace: JP, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.0, w: 1.5, h: 0.04, fill: { color: C.accent }
    });

    // Card 1
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.3, w: 8.8, h: 1.8,
      fill: { color: C.white }, shadow: mkShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.3, w: 0.06, h: 1.8, fill: { color: C.primary }
    });
    s.addText("01", {
      x: 0.85, y: 1.4, w: 0.6, h: 0.5,
      fontSize: 24, fontFace: EN, color: C.primary, bold: true, margin: 0
    });
    s.addImage({ data: icons.passport, x: 1.5, y: 1.4, w: 0.4, h: 0.4 });
    s.addText("在留資格の仕組み", {
      x: 2.1, y: 1.35, w: 7.0, h: 0.55,
      fontSize: 22, fontFace: JP, color: C.textDark, bold: true, margin: 0, valign: "middle"
    });
    s.addText("「技能（コック）」枠で合法的に就労できる道が確立", {
      x: 0.85, y: 2.0, w: 8.2, h: 0.9,
      fontSize: 20, fontFace: JP, color: C.textMuted, margin: 0, valign: "top"
    });

    // Card 2
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 3.4, w: 8.8, h: 1.8,
      fill: { color: C.white }, shadow: mkShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 3.4, w: 0.06, h: 1.8, fill: { color: C.primary }
    });
    s.addText("02", {
      x: 0.85, y: 3.5, w: 0.6, h: 0.5,
      fontSize: 24, fontFace: EN, color: C.primary, bold: true, margin: 0
    });
    s.addImage({ data: icons.users, x: 1.5, y: 3.5, w: 0.4, h: 0.4 });
    s.addText("同郷ネットワーク（連鎖移住）", {
      x: 2.1, y: 3.45, w: 7.0, h: 0.55,
      fontSize: 22, fontFace: JP, color: C.textDark, bold: true, margin: 0, valign: "middle"
    });
    s.addText("ネパール西部の村から同じ業種へ連鎖的に移住", {
      x: 0.85, y: 4.1, w: 8.2, h: 0.9,
      fontSize: 20, fontFace: JP, color: C.textMuted, margin: 0, valign: "top"
    });
  }

  // ============================================================
  // SLIDE 4: Reason 3 & 4
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.light };

    s.addText("ネパール人経営が増えた構造的理由 ③④", {
      x: 0.6, y: 0.3, w: 8.8, h: 0.7,
      fontSize: 26, fontFace: JP, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.0, w: 1.5, h: 0.04, fill: { color: C.accent }
    });

    // Card 3
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.3, w: 8.8, h: 1.8,
      fill: { color: C.white }, shadow: mkShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.3, w: 0.06, h: 1.8, fill: { color: C.primary }
    });
    s.addText("03", {
      x: 0.85, y: 1.4, w: 0.6, h: 0.5,
      fontSize: 24, fontFace: EN, color: C.primary, bold: true, margin: 0
    });
    s.addImage({ data: icons.store, x: 1.5, y: 1.4, w: 0.4, h: 0.4 });
    s.addText("独立サイクルの確立", {
      x: 2.1, y: 1.35, w: 7.0, h: 0.55,
      fontSize: 22, fontFace: JP, color: C.textDark, bold: true, margin: 0, valign: "middle"
    });
    s.addText("コック → 経験 → 経営者へ独立 → 新コック呼び込み", {
      x: 0.85, y: 2.0, w: 8.2, h: 0.9,
      fontSize: 20, fontFace: JP, color: C.textMuted, margin: 0, valign: "top"
    });

    // Card 4
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 3.4, w: 8.8, h: 1.8,
      fill: { color: C.white }, shadow: mkShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 3.4, w: 0.06, h: 1.8, fill: { color: C.primary }
    });
    s.addText("04", {
      x: 0.85, y: 3.5, w: 0.6, h: 0.5,
      fontSize: 24, fontFace: EN, color: C.primary, bold: true, margin: 0
    });
    s.addImage({ data: icons.flag, x: 1.5, y: 3.5, w: 0.4, h: 0.4 });
    s.addText("「インドカレー」看板の合理性", {
      x: 2.1, y: 3.45, w: 7.0, h: 0.55,
      fontSize: 22, fontFace: JP, color: C.textDark, bold: true, margin: 0, valign: "middle"
    });
    s.addText("「ネパール料理」より「インドカレー」が集客力大", {
      x: 0.85, y: 4.1, w: 8.2, h: 0.9,
      fontSize: 20, fontFace: JP, color: C.textMuted, margin: 0, valign: "top"
    });
  }

  // ============================================================
  // SLIDE 5: The cycle diagram
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.light };

    s.addText("コックから経営者への独立サイクル", {
      x: 0.6, y: 0.3, w: 8.8, h: 0.7,
      fontSize: 26, fontFace: JP, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.0, w: 1.5, h: 0.04, fill: { color: C.accent }
    });

    const steps = [
      { num: "1", text: "「技能（コック）」ビザで来日" },
      { num: "2", text: "既存店で修業・経験を積む" },
      { num: "3", text: "日本語と経営感覚を習得" },
      { num: "4", text: "「経営・管理」ビザで独立開業" },
    ];

    steps.forEach((st, i) => {
      const y = 1.3 + i * 1.05;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.8, y: y, w: 5.8, h: 0.75,
        fill: { color: C.white }, shadow: mkShadow()
      });
      s.addShape(pres.shapes.OVAL, {
        x: 1.0, y: y + 0.1, w: 0.55, h: 0.55,
        fill: { color: C.primary }
      });
      s.addText(st.num, {
        x: 1.0, y: y + 0.1, w: 0.55, h: 0.55,
        fontSize: 22, fontFace: EN, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0
      });
      s.addText(st.text, {
        x: 1.75, y: y + 0.05, w: 4.6, h: 0.65,
        fontSize: 20, fontFace: JP, color: C.textDark,
        valign: "middle", margin: 0
      });
      if (i < steps.length - 1) {
        s.addImage({ data: icons.arrowDown, x: 1.12, y: y + 0.78, w: 0.3, h: 0.2 });
      }
    });

    // Right callout
    s.addShape(pres.shapes.RECTANGLE, {
      x: 7.0, y: 1.3, w: 2.6, h: 3.6,
      fill: { color: C.darkBg }, shadow: mkShadow()
    });
    s.addText("この循環が\n自己増殖的に\n繰り返される", {
      x: 7.0, y: 1.5, w: 2.6, h: 1.5,
      fontSize: 20, fontFace: JP, color: C.accent, bold: true,
      align: "center", valign: "middle", lineSpacingMultiple: 1.3
    });
    s.addText("新しい店が\nコックを呼び込み\nさらに独立が続く", {
      x: 7.1, y: 3.1, w: 2.4, h: 1.5,
      fontSize: 18, fontFace: JP, color: C.textLight,
      align: "center", valign: "top", lineSpacingMultiple: 1.3
    });
  }

  // ============================================================
  // SLIDE 6: Why "Indian" not "Nepalese"
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 0.35, w: 2.8, h: 0.5,
      fill: { color: C.accent }
    });
    s.addText("KEY INSIGHT", {
      x: 0.6, y: 0.35, w: 2.8, h: 0.5,
      fontSize: 20, fontFace: EN, color: C.darkBg, bold: true,
      align: "center", valign: "middle"
    });

    s.addText("なぜ「ネパール料理」ではなく\n「インド料理」を名乗るのか", {
      x: 0.6, y: 1.1, w: 8.8, h: 1.3,
      fontSize: 30, fontFace: JP, color: C.white, bold: true,
      align: "left", lineSpacingMultiple: 1.4
    });

    // Quote box
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 2.7, w: 8.8, h: 1.5,
      fill: { color: "3D1F0A" }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 2.7, w: 0.06, h: 1.5, fill: { color: C.accent }
    });
    s.addText("「日本人のイメージは\n　カレーと言ったらインド、ナンと言ったらインド」", {
      x: 1.0, y: 2.8, w: 8.0, h: 1.0,
      fontSize: 22, fontFace: JP, color: C.accent, italic: true,
      valign: "middle", lineSpacingMultiple: 1.3
    });
    s.addText("── ネパール出身の店主（関西テレビ取材）", {
      x: 1.0, y: 3.8, w: 8.0, h: 0.35,
      fontSize: 18, fontFace: JP, color: C.textMuted, align: "right", valign: "middle"
    });

    s.addText("国籍の表示ではなく、日本人に伝わる言葉への翻訳", {
      x: 0.6, y: 4.5, w: 8.8, h: 0.6,
      fontSize: 22, fontFace: JP, color: C.textLight,
      align: "center", valign: "middle"
    });
  }

  // ============================================================
  // SLIDE 7: Naan myth - transition
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.primary };

    s.addImage({ data: icons.exchange, x: 4.5, y: 0.6, w: 1.0, h: 1.0 });

    s.addText("「インドの主食はナン」\nは正確ではない", {
      x: 0.6, y: 1.8, w: 8.8, h: 1.6,
      fontSize: 36, fontFace: JP, color: C.white, bold: true,
      align: "center", valign: "middle", lineSpacingMultiple: 1.4
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 3.5, y: 3.6, w: 3.0, h: 0.04, fill: { color: C.accent }
    });

    s.addText("ナンは食べられている\nただし「毎日の家庭の主食」ではない", {
      x: 1.0, y: 3.9, w: 8.0, h: 1.0,
      fontSize: 22, fontFace: JP, color: C.light,
      align: "center", lineSpacingMultiple: 1.4
    });
  }

  // ============================================================
  // SLIDE 8: Chapati vs Naan comparison
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.light };

    s.addText("チャパティ vs ナン ― 何が違うのか", {
      x: 0.6, y: 0.3, w: 8.8, h: 0.7,
      fontSize: 26, fontFace: JP, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.0, w: 1.5, h: 0.04, fill: { color: C.accent }
    });

    const headerOpts = {
      fill: { color: C.primary }, color: C.white, bold: true,
      fontSize: 18, fontFace: JP, align: "center", valign: "middle"
    };
    const cellOpts = {
      fontSize: 18, fontFace: JP, color: C.textDark, valign: "middle", align: "center"
    };
    const altBg = { fill: { color: "F5EDE5" } };

    const tableRows = [
      [
        { text: "", options: headerOpts },
        { text: "チャパティ", options: headerOpts },
        { text: "ナン", options: headerOpts }
      ],
      [
        { text: "材料", options: { ...cellOpts, bold: true } },
        { text: "全粒粉", options: cellOpts },
        { text: "精製小麦粉", options: cellOpts }
      ],
      [
        { text: "発酵", options: { ...cellOpts, bold: true, ...altBg } },
        { text: "なし", options: { ...cellOpts, ...altBg } },
        { text: "酵母を使う", options: { ...cellOpts, ...altBg } }
      ],
      [
        { text: "焼き方", options: { ...cellOpts, bold: true } },
        { text: "鉄板（タワ）", options: cellOpts },
        { text: "タンドール窯", options: cellOpts }
      ],
      [
        { text: "家庭", options: { ...cellOpts, bold: true, ...altBg } },
        { text: "簡単に作れる", options: { ...cellOpts, ...altBg } },
        { text: "窯がなく不可", options: { ...cellOpts, ...altBg } }
      ],
      [
        { text: "位置づけ", options: { ...cellOpts, bold: true } },
        { text: "日常の主食", options: { ...cellOpts, bold: true, color: C.green } },
        { text: "外食のパン", options: { ...cellOpts, bold: true, color: C.primary } }
      ],
    ];

    s.addTable(tableRows, {
      x: 0.6, y: 1.2, w: 8.8, h: 3.6,
      colW: [1.8, 3.5, 3.5],
      border: { pt: 0.5, color: C.sand },
      rowH: [0.5, 0.55, 0.55, 0.55, 0.55, 0.6],
    });

    // Callout at bottom
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 4.9, w: 8.8, h: 0.55,
      fill: { color: "FFF3CD" }
    });
    s.addText("ブリタニカ：チャパティは「何世紀にもわたる家庭の主食」", {
      x: 0.8, y: 4.9, w: 8.4, h: 0.55,
      fontSize: 18, fontFace: JP, color: C.textDark,
      align: "center", valign: "middle"
    });
  }

  // ============================================================
  // SLIDE 9: Why naan became standard in Japan
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.light };

    s.addText("日本で「カレー＋ナン」が定番になった理由", {
      x: 0.6, y: 0.3, w: 8.8, h: 0.7,
      fontSize: 26, fontFace: JP, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.0, w: 1.5, h: 0.04, fill: { color: C.accent }
    });

    const reasons = [
      { icon: icons.eye,   title: "視覚的インパクト",   desc: "大きくて見栄えがよく本格感" },
      { icon: icons.smile,  title: "セットの満足感",     desc: "ボリュームでランチ満足度UP" },
      { icon: icons.brain,  title: "既存イメージに合致", desc: "「インド料理＝ナン」の記号" },
      { icon: icons.coins,  title: "原価の合理性",       desc: "材料費が安く利益率が高い" },
    ];

    reasons.forEach((r, i) => {
      const y = 1.25 + i * 1.05;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: y, w: 8.8, h: 0.9,
        fill: { color: C.white }, shadow: mkShadow()
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: y, w: 0.06, h: 0.9, fill: { color: C.accent }
      });
      s.addImage({ data: r.icon, x: 0.9, y: y + 0.2, w: 0.5, h: 0.5 });
      s.addText(r.title, {
        x: 1.6, y: y + 0.05, w: 3.0, h: 0.8,
        fontSize: 20, fontFace: JP, color: C.textDark, bold: true,
        valign: "middle", margin: 0
      });
      s.addText(r.desc, {
        x: 4.8, y: y + 0.05, w: 4.3, h: 0.8,
        fontSize: 20, fontFace: JP, color: C.textMuted,
        valign: "middle", margin: 0
      });
    });
  }

  // ============================================================
  // SLIDE 10: Summary table
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.light };

    s.addText("まとめ：イメージと現実のギャップ", {
      x: 0.6, y: 0.3, w: 8.8, h: 0.7,
      fontSize: 26, fontFace: JP, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.0, w: 1.5, h: 0.04, fill: { color: C.accent }
    });

    const hOpts = {
      fill: { color: C.primary }, color: C.white, bold: true,
      fontSize: 18, fontFace: JP, align: "center", valign: "middle"
    };
    const cL = { fontSize: 18, fontFace: JP, color: C.textDark, valign: "middle", align: "center" };
    const cR = { fontSize: 18, fontFace: JP, color: C.primary, bold: true, valign: "middle", align: "center" };
    const alt = { fill: { color: "F5EDE5" } };

    const summaryRows = [
      [
        { text: "日本でのイメージ", options: hOpts },
        { text: "実際の姿", options: hOpts }
      ],
      [
        { text: "インド人が作っている", options: cL },
        { text: "ネパール人経営が大半", options: cR }
      ],
      [
        { text: "ナンがインドの主食", options: { ...cL, ...alt } },
        { text: "庶民の主食はチャパティ", options: { ...cR, ...alt } }
      ],
      [
        { text: "ナンは家庭料理", options: cL },
        { text: "タンドール窯が必要な外食メニュー", options: cR }
      ],
      [
        { text: "北インド風のみ", options: { ...cL, ...alt } },
        { text: "南インドは別の食文化", options: { ...cR, ...alt } }
      ],
    ];

    s.addTable(summaryRows, {
      x: 0.6, y: 1.2, w: 8.8, h: 3.0,
      colW: [4.4, 4.4],
      border: { pt: 0.5, color: C.sand },
      rowH: [0.5, 0.6, 0.6, 0.6, 0.6],
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 4.5, w: 8.8, h: 0.7,
      fill: { color: C.darkBg }
    });
    s.addText("移民の歴史と日本市場への適応が生んだ独自の食文化", {
      x: 0.8, y: 4.5, w: 8.4, h: 0.7,
      fontSize: 20, fontFace: JP, color: C.accent, bold: true,
      align: "center", valign: "middle"
    });
  }

  // ============================================================
  // SLIDE 11: Closing
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent }
    });

    s.addImage({ data: icons.leaf, x: 4.5, y: 0.6, w: 0.8, h: 0.8 });

    s.addText("チャパティを出している店を\n見つけたら、ぜひ試してみてほしい", {
      x: 0.8, y: 1.6, w: 8.4, h: 1.4,
      fontSize: 30, fontFace: JP, color: C.white, bold: true,
      align: "center", valign: "middle", lineSpacingMultiple: 1.4
    });

    s.addShape(pres.shapes.RECTANGLE, {
      x: 3.5, y: 3.3, w: 3.0, h: 0.04, fill: { color: C.sand }
    });

    s.addText("インドの家庭の味に、一歩近づけるはずだ", {
      x: 1.0, y: 3.6, w: 8.0, h: 0.6,
      fontSize: 24, fontFace: JP, color: C.accent,
      align: "center", valign: "middle"
    });

    // References in 18pt
    s.addText("参考文献", {
      x: 0.8, y: 4.4, w: 8.4, h: 0.4,
      fontSize: 18, fontFace: JP, color: C.textMuted, bold: true,
      align: "center", valign: "middle"
    });
    s.addText("Kharel D. 東京大学 / Britannica / PMC / 関西テレビ / 好書好日", {
      x: 0.8, y: 4.8, w: 8.4, h: 0.5,
      fontSize: 19, fontFace: JP, color: C.textMuted,
      align: "center", valign: "middle"
    });
  }

  await pres.writeFile({ fileName: "C:/Users/jsber/OneDrive/Documents/Claude_task/indian_restaurant_slides.pptx" });
  console.log("PPTX created successfully!");
}

main().catch(err => { console.error(err); process.exit(1); });
