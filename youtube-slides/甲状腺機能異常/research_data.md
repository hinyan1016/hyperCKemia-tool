# 甲状腺機能異常（亢進・低下）鑑別診断支援ツール向け 医学情報の体系化

## 対象と設計思想
総合内科医・研修医が外来〜救急〜病棟で遭遇する甲状腺機能亢進症／低下症（潜在性、薬剤性、重症疾患関連を含む）を、1つのウェブアプリで段階的に鑑別できるようにする。
「検査パターン（TSH/FT4/FT3）→身体所見→抗体→画像→背景」の順に構造化。

臨床で最重要:
- 破壊性（甲状腺炎） vs 産生過剰（バセドウ病・中毒性結節など）の切り分け（治療が逆方向）
- TSHが抑制されない甲状腺中毒症（SITSH）や、TSHが高値にならない低下症（中枢性）を見落とさない
- 甲状腺クリーゼ／粘液水腫性昏睡のRed Flagを早期に拾う

## 鑑別疾患リスト

### 甲状腺機能亢進症・甲状腺中毒症（主にTSH低値）

1. バセドウ病 / Graves' disease
   - TRAb陽性の自己免疫性甲状腺機能亢進症。甲状腺眼症の主要原因。
   - びまん性甲状腺腫、眼球突出、頻脈・体重減少・手指振戦
   - 検査: TSH+FT4/FT3、TRAb/TSAb、シンチでびまん性高取り込み、エコー血流増加
   - 甲状腺クリーゼに移行しうる

2. 機能性結節（Plummer病）/ Toxic adenoma
   - 自律性にホルモン産生する単発結節。TSH抑制型。
   - 結節性甲状腺。眼症状なし。
   - 検査: TSH+FT4/FT3、シンチでhot nodule

3. 多結節性中毒性甲状腺腫 / Toxic multinodular goiter (TMNG)
   - 複数の自律性結節。高齢者に多い。
   - 多結節性甲状腺腫。眼症状なし。
   - 検査: TSH+FT4/FT3、シンチで不均一な取り込み

4. 亜急性甲状腺炎 / Subacute thyroiditis (SAT)
   - ウイルス感染後の破壊性甲状腺炎。有痛性。
   - 頸部圧痛、発熱、CRP/ESR高値
   - 検査: TSH+FT4/FT3、CRP/ESR、エコーで疼痛部低エコー、シンチで取り込み低下
   - 通常自然軽快

5. 無痛性甲状腺炎 / Painless thyroiditis (PT)
   - 無痛性の破壊性甲状腺炎。橋本病背景が多い。
   - 甲状腺腫はあっても軽度。圧痛なし。
   - 検査: TSH+FT4/FT3、TRAb陰性、シンチで取り込み低下

6. 産後甲状腺炎 / Postpartum thyroiditis (PPT)
   - 産後1年以内に発症。破壊性→低下→回復の経過。
   - 検査: TSH+FT4/FT3、TPOAb

7. アミオダロン誘発性甲状腺中毒症 / Amiodarone-induced thyrotoxicosis (AIT)
   - Type 1: 基礎疾患あり、ヨウ素過剰による産生過剰
   - Type 2: 正常甲状腺での破壊性
   - 検査: カラードプラ血流、RAIU

8. 免疫チェックポイント阻害薬による甲状腺炎 / ICI-induced thyroid dysfunction (ICI-TD)
   - 破壊性甲状腺炎が主体。抗PD-1で多い。
   - 検査: TSH+FT4/FT3経時フォロー

9. ヨウ素誘発性甲状腺中毒症 / Iodine-induced thyrotoxicosis (JOD)
   - ヨウ素大量摂取（造影剤等）後の中毒症。
   - 検査: 尿中ヨウ素、シンチ

10. 外因性甲状腺中毒症 / Factitious thyrotoxicosis (FACT)
    - 甲状腺ホルモン剤の過剰摂取。
    - 甲状腺腫なし、Tg低値。
    - 検査: Tg（低値が手がかり）、シンチで取り込み消失

11. TSH産生下垂体腺腫 / TSH-producing pituitary adenoma (TSHoma)
    - FT4高値なのにTSHが抑制されない（SITSH）。
    - 検査: TSH+FT4/FT3、α-subunit、下垂体MRI

12. 妊娠性一過性甲状腺中毒症 / Gestational transient thyrotoxicosis (GTT)
    - hCGによる刺激。つわりと関連。
    - 検査: TSH+FT4/FT3（妊娠期基準）、TRAb

13. 卵巣甲状腺腫 / Struma ovarii (SO)
    - 卵巣奇形腫内の甲状腺組織がホルモン産生。
    - 検査: シンチで頸部集積乏しく骨盤集積

### 甲状腺機能低下症（TSH高値が典型だが例外あり）

14. 橋本病（慢性甲状腺炎）/ Hashimoto thyroiditis
    - 原発性低下症の最多原因。TPOAb/TgAb陽性。
    - びまん性甲状腺腫大（萎縮例も）
    - 検査: TSH+FT4、TPOAb/TgAb、エコー

15. 医原性（術後）/ Post-thyroidectomy hypothyroidism
16. 医原性（RAI後）/ Post-RAI hypothyroidism

17. 薬剤性（リチウム）/ Lithium-induced hypothyroidism
18. アミオダロン誘発性低下症 / Amiodarone-induced hypothyroidism (AIH)
19. ICI誘発性低下症
20. ヨウ素過剰/欠乏による低下症

21. 亜急性甲状腺炎（回復期）— 一過性低下

22. 中枢性甲状腺機能低下症 / Central hypothyroidism (CeH)
    - FT4低値なのにTSHが低値〜正常。
    - 検査: 下垂体ホルモン評価、MRI

### 潜在性（サブクリニカル）

23. 潜在性甲状腺機能亢進症 / Subclinical hyperthyroidism (SHyper)
    - TSH低値＋FT4/FT3正常。Grade 1 (0.1-0.39) / Grade 2 (<0.1)
24. 潜在性甲状腺機能低下症 / Subclinical hypothyroidism (SHypo)
    - TSH高値＋FT4正常。TSH 4-10 / >10

### その他
25. 非甲状腺疾患 / NTIS (Sick euthyroid)
26. 甲状腺ホルモン不応症 / Resistance to thyroid hormone (RTH)

## 鑑別ステップ

### Step 1: TSH値
- ⬇️ 低値（抑制）→ 亢進症/中毒症系
- ⬆️ 高値 → 低下症系
- ↔️ 正常 → NTIS/回復期/CeH/SITSH
- ⚠️ 信頼できない → 再検

### Step 2: TSH×FT4/FT3パターン
- 🔥 TSH低値＋FT4/FT3高値（顕性中毒症）
- 🟥 TSH低値＋FT3のみ高値（T3優位）
- 🌿 TSH低値＋FT4/FT3正常（潜在性亢進）
- 🧊 TSH高値＋FT4低値（顕性低下）
- 🌧️ TSH高値＋FT4正常（潜在性低下）
- 🧠 FT4低値＋TSH低〜正常（中枢性/NTIS）
- ⚡ FT4高値＋TSH正常〜高値（SITSH）

### Step 3: 症状パターン
- 🔥 甲状腺クリーゼの徴候
- 🧊 粘液水腫性昏睡の徴候
- ❤️ 動悸/振戦/体重減少
- 🐢 倦怠感/寒がり/便秘/浮腫
- 🏥 ICU/重症感染/術後
- 😶 無症候

### Step 4: 身体所見
- 👁️ 眼球突出
- 🟠 びまん性甲状腺腫
- 🧩 結節性甲状腺腫
- 🤕 圧痛あり
- 🫥 圧痛なし
- 🧱 急速増大＋圧迫症状

### Step 5: 自己抗体
- 🧿 TRAb/TSAb陽性
- 🧾 TPOAb/TgAb陽性
- 🚫 抗体陰性

### Step 6: 画像所見（エコー・シンチ）
### Step 7: 臨床背景（薬剤歴・妊娠等）
### Step 8: 追加精査

## 重要な基準値・分類

### TSH判定アルゴリズム
- 顕性亢進: TSH低値＋FT4/FT3高値
- 潜在性亢進: TSH低値＋FT4/FT3正常 (Grade 1: 0.1-0.39, Grade 2: <0.1)
- 顕性低下: TSH高値＋FT4低値
- 潜在性低下: TSH高値＋FT4正常 (TSH 4-10 / >10)
- 中枢性低下: FT4低値＋TSH低〜正常
- SITSH: FT4高値＋TSH正常〜高値

### バセドウ病 vs 破壊性甲状腺中毒症
| 項目 | バセドウ病 | 破壊性 |
|------|-----------|--------|
| TRAb | 陽性 | 原則陰性 |
| 甲状腺痛 | なし | 亜急性は有痛 |
| シンチ取り込み | びまん性高取り込み | 低下/消失 |
| エコー血流 | 増加 | 低下傾向 |

### AIT1 vs AIT2
| 所見 | AIT1 | AIT2 |
|------|------|------|
| 背景甲状腺病変 | あり | 通常なし |
| カラードプラ血流 | 増加 | 亢進なし |
| RAIU | 低〜正常〜高 | 抑制 |

### 甲状腺クリーゼ: Burch-Wartofsky スコア
### 粘液水腫性昏睡: 日本の診断基準

## Red Flags
- 甲状腺クリーゼ: 高熱・頻脈・心不全・意識障害
- 粘液水腫性昏睡: 低体温・意識低下・低換気
- 急速増大する甲状腺腫: 悪性リンパ腫・未分化癌
- ICI-irAE
- 中枢性低下症（下垂体卒中の可能性）

## 参考ガイドライン
- ATA 2016 Guidelines for Hyperthyroidism (Ross DS, et al.)
- ATA/AACE Guidelines for Hypothyroidism (2012)
- ETA 2015 Guidelines: Subclinical Hyperthyroidism
- ETA 2018 Guidelines: Central Hypothyroidism
- ETA 2018 Guidelines: Amiodarone-Associated Thyroid Dysfunction
- 日本甲状腺学会 甲状腺疾患診断ガイドライン2024
- 日本内分泌学会/日本甲状腺学会 甲状腺クリーゼ診療ガイドライン
- Endotext (NCBI Bookshelf)
