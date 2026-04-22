// 2026年上半期 内科Practice-Changing論文14選 スライド生成
// 医療PPTX仕様: BIZ UDPゴシック / Segoe UI / ネイビー基調

const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();

// レイアウト設定: 16:9 (10.0 x 5.625 inch)
pptx.defineLayout({ name: 'WIDE16x9', width: 10, height: 5.625 });
pptx.layout = 'WIDE16x9';
pptx.title = '2026年上半期 内科Practice-Changing論文14選';

// カラーパレット
const C = {
  navyDark: '1B3A5C',
  navy: '2C5AA0',
  blueLight: 'E8F4FD',
  red: 'DC3545',
  yellow: 'F5C518',
  green: '28A745',
  orange: 'E8850C',
  bg: 'FFFFFF',
  bgGray: 'F8F9FA',
  text: '2C3E50',
  sub: '6C757D',
  light: 'EAF2FB',
  accent: 'F0F8FF',
};

// フォント
const F = { ja: 'BIZ UDPゴシック', en: 'Segoe UI' };

// ========== 共通要素 ==========
function addFooter(slide, pageNum, totalPage) {
  slide.addText('© 2026 Internal Medicine Practice-Changing Papers', { x: 0.3, y: 5.30, w: 6, h: 0.25, fontSize: 9, color: C.sub, fontFace: F.en });
  slide.addText(`${pageNum} / ${totalPage}`, { x: 8.7, y: 5.30, w: 1.0, h: 0.25, fontSize: 9, color: C.sub, fontFace: F.en, align: 'right' });
}

function addTitleBar(slide, title, subtitle) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: C.navyDark }, line: { color: C.navyDark } });
  slide.addText(title, { x: 0.3, y: 0.05, w: 7.5, h: 0.4, fontSize: 20, color: 'FFFFFF', fontFace: F.ja, bold: true, valign: 'middle' });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.3, y: 0.40, w: 7.5, h: 0.28, fontSize: 11, color: 'CFE1F5', fontFace: F.en, valign: 'middle' });
  }
}

function addTrialMeta(slide, journal, date, badge) {
  slide.addShape(pptx.ShapeType.rect, { x: 7.85, y: 0.18, w: 1.95, h: 0.42, fill: { color: C.yellow }, line: { color: C.yellow } });
  slide.addText(`${journal}\n${date}`, { x: 7.85, y: 0.18, w: 1.95, h: 0.42, fontSize: 10, color: C.navyDark, fontFace: F.en, bold: true, align: 'center', valign: 'middle' });
}

const TOTAL = 22;
let p = 0;

// ========== Slide 1: タイトル ==========
{
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.navyDark };
  // 背景装飾
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 4.6, w: 10, h: 1.025, fill: { color: C.navy }, line: { color: C.navy } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 4.55, w: 10, h: 0.05, fill: { color: C.yellow }, line: { color: C.yellow } });

  s.addText('2026年上半期', { x: 0.5, y: 0.8, w: 9, h: 0.6, fontSize: 28, color: 'CFE1F5', fontFace: F.ja, bold: false, align: 'center' });
  s.addText('内科 Practice-Changing 論文 14選', { x: 0.5, y: 1.4, w: 9, h: 1.0, fontSize: 40, color: 'FFFFFF', fontFace: F.ja, bold: true, align: 'center' });

  s.addShape(pptx.ShapeType.rect, { x: 3.0, y: 2.6, w: 4, h: 0.04, fill: { color: C.yellow }, line: { color: C.yellow } });

  s.addText('— Less is More / 上流介入 / 予防医療進化 —', { x: 0.5, y: 2.75, w: 9, h: 0.5, fontSize: 18, color: C.yellow, fontFace: F.ja, align: 'center' });
  s.addText('NEJM・Lancet・JAMA から厳選', { x: 0.5, y: 3.3, w: 9, h: 0.4, fontSize: 16, color: 'CFE1F5', fontFace: F.ja, align: 'center' });

  s.addText('Internal Medicine Highlights 2026 Q1', { x: 0.5, y: 4.75, w: 9, h: 0.35, fontSize: 12, color: 'FFFFFF', fontFace: F.en, align: 'center', italic: true });
  s.addText('医師・研修医向け', { x: 0.5, y: 5.10, w: 9, h: 0.3, fontSize: 11, color: 'CFE1F5', fontFace: F.ja, align: 'center' });
}

// ========== Slide 2: 3つの潮流 概観 ==========
{
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addTitleBar(s, '2026年上半期の3つの潮流', 'Three Major Trends in Q1 2026');

  // 3カード横並び
  const cards = [
    { x: 0.3, color: C.navy, title: '① Less is More', subtitle: '標準治療の引き算', items: ['SMART-DECISION (β遮断薬中止)', 'WISDOM (リスクベース乳がん検診)', 'CTT (スタチン副作用)'] },
    { x: 3.55, color: C.green, title: '② 上流介入', subtitle: '新規モダリティの躍進', items: ['Iptacopan (補体)', 'Olezarsen (APOC3)', 'Semaglutide (MASH)', 'Tezepelumab (TSLP)', 'Treprostinil 吸入', 'CHAMPION-AF (LAAC)'] },
    { x: 6.80, color: C.orange, title: '③ 予防医療', subtitle: '発症前・再発前介入', items: ['Lenacapavir (HIV)', 'Abatacept (RA前治療)', 'Hydration (結石)', 'FMT (CDI)'] },
  ];

  cards.forEach(c => {
    s.addShape(pptx.ShapeType.rect, { x: c.x, y: 1.0, w: 2.95, h: 4.0, fill: { color: C.bg }, line: { color: c.color, width: 1.5 } });
    s.addShape(pptx.ShapeType.rect, { x: c.x, y: 1.0, w: 2.95, h: 0.85, fill: { color: c.color }, line: { color: c.color } });
    s.addText(c.title, { x: c.x, y: 1.05, w: 2.95, h: 0.4, fontSize: 16, color: 'FFFFFF', fontFace: F.ja, bold: true, align: 'center' });
    s.addText(c.subtitle, { x: c.x, y: 1.45, w: 2.95, h: 0.35, fontSize: 11, color: 'FFFFFF', fontFace: F.ja, align: 'center' });

    const bulletText = c.items.map(i => `• ${i}`).join('\n');
    s.addText(bulletText, { x: c.x + 0.15, y: 2.0, w: 2.65, h: 2.9, fontSize: 11, color: C.text, fontFace: F.ja, valign: 'top', paraSpaceAfter: 4 });
  });

  addFooter(s, p, TOTAL);
}

// ========== 論文スライド共通関数 ==========
function addTrialSlide(opts) {
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addTitleBar(s, opts.section, opts.sectionEn);
  addTrialMeta(s, opts.journal, opts.date);

  // 試験名
  s.addShape(pptx.ShapeType.rect, { x: 0.3, y: 0.85, w: 9.4, h: 0.55, fill: { color: C.blueLight }, line: { color: C.navy } });
  s.addText(opts.trialName, { x: 0.4, y: 0.88, w: 9.2, h: 0.5, fontSize: 16, color: C.navyDark, fontFace: F.ja, bold: true, valign: 'middle' });

  // 4ボックスレイアウト (PICO + Result)
  const boxes = [
    { x: 0.3, y: 1.55, w: 4.65, h: 1.45, label: '対象 (P)', text: opts.population, color: C.navy },
    { x: 5.05, y: 1.55, w: 4.65, h: 1.45, label: '介入 (I/C)', text: opts.intervention, color: C.navy },
    { x: 0.3, y: 3.10, w: 9.4, h: 1.4, label: '主要結果 (Result)', text: opts.result, color: C.green },
  ];

  boxes.forEach(b => {
    s.addShape(pptx.ShapeType.rect, { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: C.bg }, line: { color: b.color, width: 1 } });
    s.addShape(pptx.ShapeType.rect, { x: b.x, y: b.y, w: b.w, h: 0.32, fill: { color: b.color }, line: { color: b.color } });
    s.addText(b.label, { x: b.x + 0.1, y: b.y, w: b.w - 0.2, h: 0.32, fontSize: 11, color: 'FFFFFF', fontFace: F.ja, bold: true, valign: 'middle' });
    s.addText(b.text, { x: b.x + 0.15, y: b.y + 0.36, w: b.w - 0.3, h: b.h - 0.4, fontSize: 13, color: C.text, fontFace: F.ja, valign: 'top' });
  });

  // Take-home
  s.addShape(pptx.ShapeType.rect, { x: 0.3, y: 4.6, w: 9.4, h: 0.6, fill: { color: 'FFF4D6' }, line: { color: C.yellow, width: 1 } });
  s.addText('💡 Take-home: ' + opts.takeHome, { x: 0.4, y: 4.65, w: 9.2, h: 0.5, fontSize: 12, color: C.navyDark, fontFace: F.ja, bold: true, valign: 'middle' });

  addFooter(s, p, TOTAL);
}

// ========== 第1章 仕切り ==========
{
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 2.4, w: 10, h: 0.04, fill: { color: C.yellow }, line: { color: C.yellow } });
  s.addText('第 1 章', { x: 0.5, y: 1.6, w: 9, h: 0.6, fontSize: 24, color: 'CFE1F5', fontFace: F.ja, align: 'center' });
  s.addText('Less is More', { x: 0.5, y: 2.55, w: 9, h: 0.8, fontSize: 44, color: 'FFFFFF', fontFace: F.en, bold: true, align: 'center' });
  s.addText('— 標準治療の引き算 —', { x: 0.5, y: 3.4, w: 9, h: 0.5, fontSize: 18, color: C.yellow, fontFace: F.ja, align: 'center' });
  s.addText('SMART-DECISION  ・  WISDOM  ・  CTT (Statin)', { x: 0.5, y: 4.0, w: 9, h: 0.4, fontSize: 14, color: 'CFE1F5', fontFace: F.en, align: 'center' });
  addFooter(s, p, TOTAL);
}

// ========== 1. SMART-DECISION ==========
addTrialSlide({
  section: '1. 心筋梗塞後β遮断薬は中止してよいか',
  sectionEn: 'SMART-DECISION Trial',
  journal: 'NEJM',
  date: '2026/3/30',
  trialName: 'SMART-DECISION (Choi KH ら / 韓国)',
  population: '心筋梗塞後 1年以上β遮断薬継続中、\nLVEF ≥40% の安定患者\n2,540例 (平均63歳, 女性12.8%)',
  intervention: 'β遮断薬「中止」 vs 「継続」\n中央値追跡 3.1年',
  result: '全死亡＋再梗塞＋心不全入院: 中止群 7.2% vs 継続群 9.0% (HR 0.80, 95%CI 0.57–1.13)\n→ 「中止」の「継続」に対する 非劣性 を達成',
  takeHome: 'LVEFが保たれ安定した心筋梗塞既往患者では、1年以降のβ遮断薬中止は安全な選択肢',
});

// ========== 2. WISDOM ==========
addTrialSlide({
  section: '2. リスクベース vs 毎年マンモグラフィ',
  sectionEn: 'WISDOM Trial',
  journal: 'JAMA',
  date: '2026/3/3',
  trialName: 'WISDOM (Esserman LJ ら / 米国)',
  population: '40〜74歳の女性 28,372例\n(リスクベース 14,212 vs 毎年 14,160)',
  intervention: '個人リスクに応じた検診間隔 vs 一律毎年マンモ\n中央値追跡 5.1年',
  result: 'Stage IIB以上のがん発見率: リスクベース 30/10万人年 vs 毎年 48/10万人年\n→ 非劣性。最高リスク層では Stage IIB+ ゼロ (=より早期検出)',
  takeHome: '個別化検診（precision prevention）が、過剰診断を減らしつつ進行がん検出を担保',
});

// ========== 3. CTT スタチン副作用 ==========
addTrialSlide({
  section: '3. スタチン副作用の大半はノセボ効果',
  sectionEn: 'CTT Collaboration Meta-analysis',
  journal: 'Lancet',
  date: '2026/2',
  trialName: 'CTT Collaboration (Reith C ら, Oxford)',
  population: 'スタチン vs プラセボ 二重盲検RCT 19試験\n123,940例 (中央値追跡 4.5年)',
  intervention: 'スタチン群 と プラセボ群 の有害事象を\n個別患者データレベルで比較',
  result: '記憶障害・抑うつ・睡眠障害・性機能障害 → スタチンに起因しない\n因果が確認されたのは 横紋筋融解・新規糖尿病・肝酵素上昇 など6項目のみ',
  takeHome: 'スタチン副作用の大半はノセボ効果。患者教育と再投与（rechallenge）戦略が重要',
});

// ========== 第2章 仕切り ==========
{
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.green };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 2.4, w: 10, h: 0.04, fill: { color: C.yellow }, line: { color: C.yellow } });
  s.addText('第 2 章', { x: 0.5, y: 1.6, w: 9, h: 0.6, fontSize: 24, color: 'D4EDDA', fontFace: F.ja, align: 'center' });
  s.addText('Upstream Intervention', { x: 0.5, y: 2.55, w: 9, h: 0.8, fontSize: 40, color: 'FFFFFF', fontFace: F.en, bold: true, align: 'center' });
  s.addText('— 病態上流への介入：新規モダリティ —', { x: 0.5, y: 3.4, w: 9, h: 0.5, fontSize: 18, color: C.yellow, fontFace: F.ja, align: 'center' });
  s.addText('Iptacopan ・ Olezarsen ・ Semaglutide ・ Tezepelumab ・ Treprostinil ・ CHAMPION-AF', { x: 0.5, y: 4.0, w: 9, h: 0.4, fontSize: 12, color: 'D4EDDA', fontFace: F.en, align: 'center' });
  addFooter(s, p, TOTAL);
}

// ========== 4. APPLAUSE-IgAN ==========
addTrialSlide({
  section: '4. IgA腎症に補体B因子阻害',
  sectionEn: 'APPLAUSE-IgAN 24-month',
  journal: 'NEJM',
  date: '2026/3/29',
  trialName: 'APPLAUSE-IgAN 24ヶ月最終データ (Barratt J ら)',
  population: '原発性IgA腎症\n(標準治療下でもタンパク尿持続)',
  intervention: '補体B因子阻害薬 Iptacopan vs プラセボ\n投与期間 24ヶ月',
  result: '年間eGFR低下: −3.10 vs −6.12 mL/min/1.73㎡ (差 +3.02, p<0.001)\n腎複合エンドポイント: 21.4% vs 33.5% (HR 0.57)。重篤感染 6.7% vs 2.1% に注意',
  takeHome: '補体副経路の直接阻害が、IgA腎症の進行をほぼ半減 (ただし感染症対策必須)',
});

// ========== 5. CORE-TIMI 72 / Olezarsen ==========
addTrialSlide({
  section: '5. APOC3標的で重症高TG血症と膵炎を抑制',
  sectionEn: 'CORE-TIMI 72 (Olezarsen)',
  journal: 'NEJM',
  date: '2026',
  trialName: 'CORE-TIMI 72a / 72b (Marston NA ら)',
  population: '重症高トリグリセリド血症患者\n(TG ≥500 mg/dL)',
  intervention: 'APOC3標的アンチセンスオリゴ Olezarsen 50/80mg\n月1回 皮下注 vs プラセボ (12ヶ月)',
  result: '6ヶ月でTG値が大幅低下\n急性膵炎 平均発症率比 0.15 (95%CI 0.05–0.40, p<0.001) → 約 85% 減',
  takeHome: 'mRNA標的の新世代モダリティが、難治性高TG血症の膵炎リスクをほぼゼロに',
});

// ========== 6. ESSENCE Semaglutide MASH ==========
addTrialSlide({
  section: '6. セマグルチドが MASH に有効',
  sectionEn: 'ESSENCE Trial',
  journal: 'NEJM',
  date: '2025/4',
  trialName: 'ESSENCE (Sanyal AJ ら)',
  population: 'MASH (代謝機能障害関連脂肪肝炎)\n線維化Stage 2–3 / 1,197例',
  intervention: 'セマグルチド 2.4mg 週1回 皮下注 vs プラセボ\n投与期間 72週',
  result: '線維化悪化なし MASH消失率: 62.9% vs 34.3% (差 +28.7%, p<0.001)\nMASH消失＋線維化改善の同時達成: 32.7% vs 16.1%',
  takeHome: 'GLP-1受容体作動薬がMASH治療薬としての確固たる地位を獲得',
});

// ========== 7. WAYFINDER Tezepelumab ==========
addTrialSlide({
  section: '7. 重症喘息の経口ステロイド離脱',
  sectionEn: 'WAYFINDER (Tezepelumab Phase 3b)',
  journal: 'Lancet RM',
  date: '2026/2',
  trialName: 'WAYFINDER (フェーズ3b・単群)',
  population: 'OCS依存性 重症喘息成人 298例\n(プレドニン換算 5–40 mg/日)',
  intervention: 'TSLP阻害薬 Tezepelumab 210mg / 4週ごと SC\n投与期間 52週',
  result: '52週時点で OCS ≤5 mg 達成: 89.9%\nOCS 完全離脱: 50.3% (喘息コントロール悪化なし)',
  takeHome: 'TSLP阻害薬がOCS依存重症喘息の半数を「ステロイドフリー」に到達させる',
});

// ========== 8. TETON-2 Treprostinil IPF ==========
addTrialSlide({
  section: '8. 吸入Treprostinilが IPF の進行を抑制',
  sectionEn: 'TETON-2 Trial',
  journal: 'NEJM',
  date: '2026/3/11',
  trialName: 'TETON-2 (Nathan SD ら)',
  population: '特発性肺線維症 (IPF) 患者',
  intervention: '吸入 Treprostinil 12吸入×4回/日 vs プラセボ\n投与期間 52週',
  result: 'FVC変化中央値: −49.9 mL vs −136.4 mL (差 +95.6 mL)\n臨床的悪化イベント: 27.2% vs 39.0%',
  takeHome: 'PH-ILDだけでなく IPF 本体の進行を抑制。pirfenidone・nintedanibに次ぐ第3の選択肢',
});

// ========== 9. CHAMPION-AF ==========
addTrialSlide({
  section: '9. 左心耳閉鎖術 vs DOAC',
  sectionEn: 'CHAMPION-AF Trial',
  journal: 'NEJM',
  date: '2026/3/28',
  trialName: 'CHAMPION-AF (Leon MB ら / ACC.26)',
  population: '長期DOAC可能な心房細動患者\n3,000例',
  intervention: 'WATCHMAN FLX による LAAC 1,499例\nvs DOAC 1,501例 (3年追跡)',
  result: '心血管死＋脳卒中＋全身塞栓 → 非劣性\n非手技関連出血: 10.9% vs 19.0% (45%減, p<0.001) 優越性',
  takeHome: 'LAACが「DOAC不耐者の代替」から「DOAC候補患者でも選択肢の1つ」へシフト',
});

// ========== 第3章 仕切り ==========
{
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.orange };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 2.4, w: 10, h: 0.04, fill: { color: C.navyDark }, line: { color: C.navyDark } });
  s.addText('第 3 章', { x: 0.5, y: 1.6, w: 9, h: 0.6, fontSize: 24, color: 'FFE5C2', fontFace: F.ja, align: 'center' });
  s.addText('Prevention Evolved', { x: 0.5, y: 2.55, w: 9, h: 0.8, fontSize: 44, color: 'FFFFFF', fontFace: F.en, bold: true, align: 'center' });
  s.addText('— 予防医療の進化：発症前・再発前介入 —', { x: 0.5, y: 3.4, w: 9, h: 0.5, fontSize: 18, color: C.navyDark, fontFace: F.ja, align: 'center' });
  s.addText('PURPOSE-2 (Lenacapavir) ・ ALTO (Abatacept) ・ Hydration ・ FMT', { x: 0.5, y: 4.0, w: 9, h: 0.4, fontSize: 12, color: 'FFE5C2', fontFace: F.en, align: 'center' });
  addFooter(s, p, TOTAL);
}

// ========== 10. PURPOSE-2 Lenacapavir ==========
addTrialSlide({
  section: '10. 年2回注射でHIV感染をほぼゼロに',
  sectionEn: 'PURPOSE-2 Trial',
  journal: 'NEJM',
  date: '2024/11',
  trialName: 'PURPOSE-2 (Lenacapavir 二重盲検RCT)',
  population: 'シスゲイ男性・トランスジェンダー等\n3,265例 (2:1割付)',
  intervention: '長作用型カプシド阻害薬 Lenacapavir\n26週ごと SC vs 毎日 FTC/TDF (PrEP)',
  result: 'HIV発症率: Lenacapavir群 0.10/100人年 (2例)\nvs F/TDF群 0.93/100人年 (9例) 背景発症率より有意に低下',
  takeHome: '毎日服薬のアドヒアランス問題を根本解決。HIV予防の歴史的ブレイクスルー',
});

// ========== 11. ALTO / APIPPRA ==========
addTrialSlide({
  section: '11. 関節リウマチを発症前に「治療」する',
  sectionEn: 'ALTO (APIPPRA Long-term)',
  journal: 'Lancet Rheum',
  date: '2026/1',
  trialName: 'ALTO (APIPPRA Long-term Outcome)',
  population: '抗CCP抗体陽性＋関節痛の RA発症ハイリスク群\n213例',
  intervention: 'T細胞共刺激モジュレーター Abatacept\n1年間 投与 vs プラセボ (4–8年追跡)',
  result: '1年治療で 治療終了後 最長4年にわたり RA発症が有意に遅延\n広範な自己抗体プロファイルを持つ患者で特に有効',
  takeHome: '「症状が出る前に治療」する Pre-emptive therapy の概念実証',
});

// ========== 12. Hydration trial ==========
addTrialSlide({
  section: '12. 水分摂取で結石は再発予防できるか',
  sectionEn: 'Hydration Adherence Trial',
  journal: 'Lancet',
  date: '2026/3/21',
  trialName: 'Hydration trial (Desai AC ら / USDN)',
  population: '尿路結石既往患者',
  intervention: 'スマート水筒・コーチング・インセンティブによる\n水分摂取増加プログラム vs 通常ケア (2年)',
  result: '介入群で 尿量は有意に増加\nしかし症候性結石再発率には 有意差なし',
  takeHome: '行動介入だけでは不十分。食事制限・薬物療法を含む包括的アプローチが必要',
});

// ========== 13. FMT for CDI ==========
addTrialSlide({
  section: '13. 再発性CDIにおけるFMT・微生物製剤',
  sectionEn: 'FMT for Recurrent C. difficile (Review 2026)',
  journal: 'Reviews',
  date: '2026',
  trialName: 'FMT vs 精製マイクロバイオーム製剤 (FDA承認製剤の登場)',
  population: '再発性Clostridioides difficile感染症\n小児・成人',
  intervention: '従来型FMT vs FDA承認の精製微生物製剤\n(Fecal microbiota live-jslm / SER-109 等)',
  result: '小児ではFMT治癒率 約86% (単回 81%)\n重篤有害事象 2.0% — 高い有効性と安全性が再確認',
  takeHome: 'CDIマイクロバイオーム治療の「使い分け」が整理される時代へ',
});

// ========== 14. SSC 2026 ==========
{
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addTitleBar(s, '14. Surviving Sepsis Campaign 2026', 'SSC 2026 Guideline');
  addTrialMeta(s, 'Crit Care Med', '2026/4/1');

  s.addShape(pptx.ShapeType.rect, { x: 0.3, y: 0.85, w: 9.4, h: 0.55, fill: { color: C.blueLight }, line: { color: C.navy } });
  s.addText('SSC 2026 (Prescott HC・Antonelli M 共同議長 / 23ヶ国69名パネル)', { x: 0.4, y: 0.88, w: 9.2, h: 0.5, fontSize: 16, color: C.navyDark, fontFace: F.ja, bold: true, valign: 'middle' });

  // 主要変更点 4つ
  const items = [
    { icon: '🕐', t: '抗菌薬投与のタイミング', d: '敗血症性ショック疑い → 1時間以内 を維持\n敗血症（ショックなし）はリスク階層化で柔軟化' },
    { icon: '💧', t: '輸液蘇生', d: '固定 30 mL/kg/h ではなく\n動的指標による個別化' },
    { icon: '🧬', t: '129ステートメント中 46が新規', d: 'Precision Medicine（免疫表現型ベース）への展望が初めて明文化' },
    { icon: '🌍', t: '低中所得国の文脈', d: 'パネリストの 38% が低中所得国出身\n資源制約下の推奨も提示' },
  ];

  items.forEach((it, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 4.75;
    const y = 1.55 + row * 1.6;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 4.65, h: 1.45, fill: { color: C.bg }, line: { color: C.navy, width: 1 } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.5, h: 1.45, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(it.icon, { x: x + 0.05, y: y + 0.3, w: 0.4, h: 0.6, fontSize: 22, color: 'FFFFFF', align: 'center', valign: 'middle' });
    s.addText(it.t, { x: x + 0.6, y: y + 0.1, w: 3.95, h: 0.4, fontSize: 13, color: C.navyDark, fontFace: F.ja, bold: true, valign: 'middle' });
    s.addText(it.d, { x: x + 0.6, y: y + 0.5, w: 3.95, h: 0.9, fontSize: 11, color: C.text, fontFace: F.ja, valign: 'top' });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.3, y: 4.78, w: 9.4, h: 0.42, fill: { color: 'FFF4D6' }, line: { color: C.yellow, width: 1 } });
  s.addText('💡 Take-home: 1時間ルールは堅持しつつ、輸液・Precision Medicineで「個別化」へ進化', { x: 0.4, y: 4.80, w: 9.2, h: 0.38, fontSize: 12, color: C.navyDark, fontFace: F.ja, bold: true, valign: 'middle' });

  addFooter(s, p, TOTAL);
}

// ========== 15. SWiFT 全血輸血 ==========
addTrialSlide({
  section: '15. 病院前 全血輸血 vs 標準ケア',
  sectionEn: 'SWiFT Trial',
  journal: 'NEJM',
  date: '2026/3/17',
  trialName: 'SWiFT (Study of Whole blood in Frontline Trauma)',
  population: '生命危険を伴う出血性外傷 616例\n(救急航空機10機 / 19病院)',
  intervention: '病院前 全血 2単位 輸血\nvs 標準血液製剤輸血',
  result: '24時間以内の死亡 or 大量輸血の複合エンドポイントで\n優越性なし',
  takeHome: 'プレホスピタル全血戦略のロジ・コストを見直す根拠 (PDF原文「RePHILL-2」は誤記)',
});

// ========== 16. まとめ表 ==========
{
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  addTitleBar(s, '14試験 × 3つの潮流 マッピング', 'Summary Table');

  // ヘッダー行
  const cols = [
    { x: 0.3, w: 1.8, label: '潮流' },
    { x: 2.1, w: 4.0, label: '代表試験' },
    { x: 6.1, w: 3.6, label: '臨床への含意' },
  ];
  cols.forEach(c => {
    s.addShape(pptx.ShapeType.rect, { x: c.x, y: 1.0, w: c.w, h: 0.5, fill: { color: C.navyDark }, line: { color: C.navyDark } });
    s.addText(c.label, { x: c.x, y: 1.0, w: c.w, h: 0.5, fontSize: 13, color: 'FFFFFF', fontFace: F.ja, bold: true, align: 'center', valign: 'middle' });
  });

  const rows = [
    { trend: 'Less is More', trial: 'SMART-DECISION / WISDOM / CTT', impact: '漫然継続を見直し、個別化と引き算が標準へ', color: C.navy },
    { trend: '上流介入', trial: 'Iptacopan / Olezarsen / Semaglutide / Tezepelumab / Treprostinil / LAAC', impact: '分子・デバイスで病態上流を直接攻撃', color: C.green },
    { trend: '予防医療進化', trial: 'Lenacapavir / Abatacept / Hydration', impact: '発症前・再発前介入と、行動介入の限界', color: C.orange },
    { trend: '重症ケア', trial: 'SSC 2026 / SWiFT', impact: '個別化 (Precision Medicine) への進化と再評価', color: C.red },
  ];

  let y = 1.5;
  rows.forEach((r, i) => {
    const h = 0.85;
    s.addShape(pptx.ShapeType.rect, { x: 0.3, y, w: 1.8, h, fill: { color: r.color }, line: { color: 'FFFFFF', width: 1 } });
    s.addText(r.trend, { x: 0.3, y, w: 1.8, h, fontSize: 13, color: 'FFFFFF', fontFace: F.ja, bold: true, align: 'center', valign: 'middle' });
    s.addShape(pptx.ShapeType.rect, { x: 2.1, y, w: 4.0, h, fill: { color: i % 2 === 0 ? C.bgGray : 'FFFFFF' }, line: { color: 'DDDDDD', width: 1 } });
    s.addText(r.trial, { x: 2.2, y: y + 0.05, w: 3.8, h: h - 0.1, fontSize: 11, color: C.text, fontFace: F.ja, valign: 'middle' });
    s.addShape(pptx.ShapeType.rect, { x: 6.1, y, w: 3.6, h, fill: { color: i % 2 === 0 ? C.bgGray : 'FFFFFF' }, line: { color: 'DDDDDD', width: 1 } });
    s.addText(r.impact, { x: 6.2, y: y + 0.05, w: 3.4, h: h - 0.1, fontSize: 11, color: C.text, fontFace: F.ja, valign: 'middle' });
    y += h;
  });

  addFooter(s, p, TOTAL);
}

// ========== 17. クロージング ==========
{
  p++;
  const s = pptx.addSlide();
  s.background = { color: C.navyDark };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 4.55, w: 10, h: 0.05, fill: { color: C.yellow }, line: { color: C.yellow } });

  s.addText('まとめ', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 32, color: C.yellow, fontFace: F.ja, bold: true, align: 'center' });

  s.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.5, w: 8.4, h: 2.8, fill: { color: 'FFFFFF', transparency: 90 }, line: { color: 'FFFFFF', width: 1 } });

  const points = [
    '✓ 「漫然と続ける標準治療」を引き算する研究が増加',
    '✓ 補体・mRNA・長作用型注射など 病態上流のモダリティ が躍進',
    '✓ 発症前 / 再発前 の能動的介入が新しい予防医療の柱に',
    '✓ 1試験の結果だけでなく、対象集団・サブ解析を読み込むことが重要',
  ];

  points.forEach((pt, i) => {
    s.addText(pt, { x: 1.2, y: 1.7 + i * 0.55, w: 7.6, h: 0.5, fontSize: 15, color: 'FFFFFF', fontFace: F.ja, valign: 'middle' });
  });

  s.addText('Internal Medicine 2026 Q1 — End of Slides', { x: 0.5, y: 4.75, w: 9, h: 0.4, fontSize: 14, color: 'CFE1F5', fontFace: F.en, italic: true, align: 'center' });
  s.addText('参考文献は記事末尾を参照', { x: 0.5, y: 5.10, w: 9, h: 0.3, fontSize: 11, color: 'CFE1F5', fontFace: F.ja, align: 'center' });
}

// ========== 保存 ==========
pptx.writeFile({ fileName: '2026_internal_medicine_top14.pptx' })
  .then(fn => console.log('Saved:', fn));
