from __future__ import annotations

import math
import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


W, H = 1080, 1920
ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "images" / "in"
DB_PATH = ROOT / "data" / "state.sqlite"


@dataclass(frozen=True)
class Section:
    heading: str
    bullets: tuple[str, ...]
    stats: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class Entry:
    file: str
    title: str
    subtitle: str
    colors: tuple[str, str, str]
    source: str
    sections: tuple[Section, ...]


ENTRIES: tuple[Entry, ...] = (
    Entry(
        "6802888565233321439.png",
        "診療支援AI",
        "国際的な活用状況 2025",
        ("#00BCD4", "#673AB7", "#1A3A5C"),
        "2025-09-15 参考記事を確認",
        (
            Section("実用化の進展", ("FDA認可AI医療機器は2024年8月時点で約950件", "年間認可は2015年6件から2023年221件へ急増", "多くは医用画像解析だが文書作成・要約にも拡大"), (("2015", "6件"), ("2023", "221件"), ("累計", "約950件"))),
            Section("各国の導入と規制", ("米国FDAはSaMDとして510(k)・De Novo・PMAで審査", "欧州MDR、中国NMPA、日本PMDAで枠組みが異なる", "中国は2023年7月時点で59件以上のAI医療機器を承認")),
            Section("医療用LLMの性能", ("Med-PaLM 2、GPT-4、Med-Geminiなどが医学QAで高性能", "Med-GeminiはUSMLE形式で91.1%と報告", "現場導入は文書作成・検索・要約から段階的に進む"), (("Med-PaLM 2", "86.5%"), ("Med-Gemini", "91.1%"))),
            Section("運用上の課題", ("AIは医師を置換せずケアチームを補強する位置づけ", "バイアス、ハルシネーション、透明性を継続検証", "市販後監視と再学習管理を含むガバナンスが必要")),
        ),
    ),
    Entry(
        "6802888565235866864.png",
        "肥満症 × 脳卒中",
        "最新エビデンスと予防",
        ("#FF9800", "#E53935", "#1A3A5C"),
        "2025-09-16 参考記事を確認",
        (
            Section("BMIとリスク", ("BMI上昇は脳卒中リスクと量的に関連", "BMI +1でリスク約6%増、BMI +5で虚血性脳卒中約21%増", "日本人女性BMI 30以上でリスク2.16倍との報告"), (("BMI +1", "+6%"), ("BMI +5", "+21%"), ("BMI>=30", "2.16倍"))),
            Section("機序", ("インスリン抵抗性と慢性炎症が血管障害を促進", "血管内皮機能障害、脂質異常、血圧上昇が重なる", "肥満は単独リスクであり生活習慣病の集積点でもある")),
            Section("GLP-1受容体作動薬", ("SUSTAIN-6、LEADER、SELECTなどで心血管保護を検討", "全脳卒中リスク16%低下、非致死的脳卒中13%低下", "NNTは約250で、生活習慣改善と併用が前提"), (("全脳卒中", "-16%"), ("非致死的", "-13%"), ("NNT", "約250"))),
            Section("臨床への含意", ("一次予防・二次予防に体重管理を統合", "心血管・腎保護まで見据えた介入を選ぶ", "肥満パラドックスは解釈に注意し個別化する")),
        ),
    ),
    Entry(
        "6802888565236418465.png",
        "成人てんかん × 睡眠",
        "最新の知見",
        ("#1A5276", "#6A1B9A", "#1A3A5C"),
        "2025-09-16 参考記事を確認",
        (
            Section("発作と睡眠相", ("発作はNREM睡眠中に多くREMでは少ない", "NREMの同期的徐波が発作閾値を下げる可能性", "夜間発作はSUDEPリスク評価でも重要"), (("NREM", "95%"), ("REM", "<1%"))),
            Section("睡眠障害の併存", ("不眠は約50%、OSAは10-35%で併存", "睡眠障害により発作リスクが上がり悪循環を作る", "日中眠気、服薬アドヒアランス、生活リズムを同時に評価"), (("不眠", "50%"), ("OSA", "10-35%"))),
            Section("CPAPの意義", ("OSA合併例ではCPAP治療で発作改善率が上昇", "未治療14%に対しCPAP治療群63%改善の報告", "睡眠併存症治療は抗発作薬調整と並ぶ介入点"), (("未治療", "14%"), ("CPAP", "63%"))),
            Section("薬剤とモニタリング", ("スボレキサント、ラコサミド、ペランパネル等の睡眠影響を確認", "ウェアラブル脳波やAI解析で在宅観察が進む", "睡眠の質と発作記録を合わせて長期管理する")),
        ),
    ),
    Entry(
        "6802888565237146188.png",
        "POEMS症候群",
        "疫学・診断・治療・予後",
        ("#512DA8", "#E91E63", "#1A3A5C"),
        "2025-09-17 参考記事を確認",
        (
            Section("疫学と定義", ("Crow-Fukase症候群とも呼ばれる希少疾患", "形質細胞異常とVEGF高値を伴うパラネオプラスティック疾患", "50歳代好発、男性やや多い"), (("有病率", "0.3/10万"), ("男:女", "1.5-2:1"))),
            Section("診断と病態", ("λ型軽鎖の単クローン性増殖が重要", "VEGF高値は診断と活動性評価に有用", "2024年新基準も踏まえ主要・副次基準を確認"), (("VEGF", ">1000"), ("感度", "100%"), ("特異度", "93%"))),
            Section("症状頻度", ("多発神経炎はほぼ必発", "内分泌異常、皮膚変化、臓器腫大、浮腫を組み合わせて評価", "Castleman病合併やM蛋白も確認"), (("神経症状", "100%"), ("内分泌", "74%"), ("皮膚変化", "59%"), ("臓器腫大", "54%"))),
            Section("治療と予後", ("自家末梢血幹細胞移植、レナリドミド、ボルテゾミブ等を選択", "VEGFを経過観察に活用", "長期生存は良好化しており早期治療が鍵"), (("5年PFS", "約55%"), ("10年生存", ">90%"))),
        ),
    ),
    Entry(
        "6802888565241143774.png",
        "V4椎骨動脈解離",
        "フォローアップ戦略",
        ("#D32F2F", "#0D47A1", "#1A3A5C"),
        "2025-09-19 参考記事を確認",
        (
            Section("背景・リスク", ("日本では椎骨動脈解離が動脈解離の70-80%を占める報告", "V4は頭蓋内椎骨動脈で脳幹・小脳梗塞やSAHの原因", "40-50歳代、頭痛・頚部痛で発症しうる"), (("虚血発症", "約33%"), ("SAH", "約30%"))),
            Section("抗血栓療法", ("抗血小板薬と抗凝固薬の有効性に明確な差はない", "通常は3-6か月を目安に画像所見で終了判断", "DAPTは routine ではなく出血リスクに注意"), (("期間", "3-6か月"), ("推奨", "個別化"))),
            Section("画像フォロー", ("発症後3か月を目安にMRA/CTAで再評価", "6か月、必要時12か月も確認", "VW-MRIは治癒過程の補助評価に有用"), (("3か月", "再評価"), ("6か月", "大半再開通"), ("治癒過程", "約60%"))),
            Section("疼痛と生活指導", ("後頭部から項部の強い痛みはNSAIDs・アセトアミノフェン等で管理", "収縮期140mmHg未満を目安に血圧管理", "1-2週間の安静、頚部過伸展や重量物を避ける")),
        ),
    ),
    Entry(
        "6802888565252131513.png",
        "術後PLMD",
        "鉄代謝と治療フロー",
        ("#1E88E5", "#43A047", "#1A3A5C"),
        "2025-09-24 参考記事を確認",
        (
            Section("診断", ("周期性四肢運動と睡眠分断・日中傾眠を確認", "睡眠ポリグラフではPLMI 15/時以上が目安", "術後発症・増悪では薬剤、疼痛、鉄代謝を再評価"), (("PLMI", ">=15/時"),)),
            Section("鉄代謝評価", ("血清フェリチン75 ng/mL未満で介入を検討", "TSAT 45%未満も補充判断に使う", "目標は少なくとも50-100 ng/mL以上を意識"), (("Ferritin", "<75"), ("TSAT", "<45%"))),
            Section("鉄補充", ("経口は元素鉄100-200 mgとビタミンCを組み合わせる", "吸収不良や不耐では静脈鉄を検討", "術中・周術期はロチゴチンパッチ等の代替も考慮"), (("元素鉄", "100-200mg"),)),
            Section("長期管理", ("短期はドパミン作動薬、長期はガバペンチノイドを軸に検討", "プラミペキソール、ロピニロール、ガバペンチン、プレガバリンを個別化", "依存・増悪症状を定期的に監視")),
        ),
    ),
    Entry(
        "6802888565264238669.png",
        "脳梗塞 × 胃薬",
        "3つの落とし穴",
        ("#C62828", "#F57F17", "#1A3A5C"),
        "2025-09-30 参考記事を確認",
        (
            Section("落とし穴1 保険適用", ("アスピリンと潰瘍既往ではPPI適用を確認", "他の抗血栓薬では症状・既往の記載が重要", "処方理由を患者・薬剤師と共有する")),
            Section("落とし穴2 相互作用", ("クロピドグレルはCYP2C19経由でPPIとの相互作用に注意", "オメプラゾール、ランソプラゾールでは効力低下が問題になりうる", "PCABボノプラザンや相互作用の少ない選択肢を検討")),
            Section("落とし穴3 長期リスク", ("PPI長期使用では腸管感染症、骨折、B12・鉄・Mg不足を再評価", "高齢者ほど必要性を定期的に見直す", "3年以上の漫然投与を避ける"), (("長期", "3年以上"),)),
            Section("相談ポイント", ("症状改善後の減量・中止を検討", "主治医と薬剤師に抗血栓薬名を伝える", "ラベプラゾール、エソメプラゾール等も候補")),
        ),
    ),
    Entry(
        "6802888565265733310.png",
        "DIAM",
        "薬剤性無菌性髄膜炎の診断",
        ("#0D47A1", "#FF6F00", "#1A3A5C"),
        "2025-10-01 参考記事を確認",
        (
            Section("定義と疫学", ("医薬品使用に伴う無菌性髄膜炎", "髄液培養は陰性だが細菌性髄膜炎と似る", "初回と再曝露で発症速度が異なる")),
            Section("原因薬剤", ("NSAIDs、ST合剤・βラクタム、抗てんかん薬", "IVIG、免疫チェックポイント阻害薬も重要", "ラモトリギンは1994-2009年に40症例のFDA報告"), (("原因群", "5分類"), ("ICI頻度", "<1%"), ("致死率", "約5%"))),
            Section("髄液所見", ("細胞数は数百から数千/μL", "初期は好中球優位で細菌性との鑑別が難しい", "蛋白は軽度から中等度上昇、糖は正常範囲が多い"), (("細胞数", "数百-数千"), ("蛋白", "約100mg/dL"))),
            Section("経過と治療", ("初回は数時間から数日、再曝露は数分から数時間", "原因薬中止で24-48時間以内に改善しやすい", "再投与は原則禁忌、薬歴聴取が鍵"), (("改善", "24-48h"),)),
        ),
    ),
    Entry(
        "6802888565265735201.png",
        "DIAM",
        "薬で起こる髄膜炎 5つの真実",
        ("#E53935", "#66BB6A", "#1A3A5C"),
        "2025-10-01 参考記事を確認",
        (
            Section("身近な薬で起こる", ("イブプロフェンなどNSAIDsが代表", "抗菌薬、抗てんかん薬、IVIG、ICIでも報告", "市販薬・処方薬の両方が原因になりうる")),
            Section("2回目は速い", ("初回は数時間から数日で発症", "再投与では数分から数時間で重く再発しうる", "同じ薬で似た症状があれば必ず伝える"), (("初回", "数時間-数日"), ("再投与", "数分-数時間"))),
            Section("診断が難しい", ("発熱、激しい頭痛、項部硬直は感染性髄膜炎と似る", "髄液所見も好中球優位になり混同されやすい", "薬手帳・市販薬・サプリ情報が診断の鍵")),
            Section("治療と新しい課題", ("原因薬中止で24-48時間以内に急速改善することが多い", "ICI関連では重症化しステロイド治療が必要な場合あり", "副作用情報は新薬登場ごとに更新が必要"), (("改善", "24-48h"), ("ICI", "<1%"))),
        ),
    ),
    Entry(
        "6802888565271050771.png",
        "ジルチアゼム供給不足",
        "3つの驚きの事実",
        ("#D32F2F", "#FBC02D", "#1A3A5C"),
        "2025-10-03 参考記事を確認",
        (
            Section("先発品の退場", ("ヘルベッサー錠が2025年3月末で販売終了", "需要がジェネリックへ一斉に移り注文が集中", "原薬遅延やGMP問題も重なった複合要因"), (("販売終了", "2025/3末"),)),
            Section("沢井製薬の長期問題", ("2021年から徐放性カプセル工程の問題が続いた", "是正対応の長期化で増産余力が限られた", "2025年には通常錠剤にも逼迫が波及"), (("工程問題", "2021年-"),)),
            Section("他社品影響", ("東和薬品は自社問題なしでも注文殺到で限定出荷", "一社の不足が健全なメーカーにも波及", "供給網の相互依存と脆弱性が露呈")),
            Section("現場の代替工夫", ("200mg徐放カプセルから30mg錠 1日3回へ再設計", "100mgカプセル2個への代替案も提示", "血圧・心拍を確認し薬剤師と連携"), (("200mg", "徐放"), ("30mg", "1日3回"), ("100mg", "2個"))),
        ),
    ),
    Entry(
        "6802888565272314020.png",
        "てんかん発作 × 乳酸",
        "診断マーカーとしての価値",
        ("#6A4C93", "#FFB84D", "#1A3A5C"),
        "2025-10-04 参考記事を確認",
        (
            Section("代謝病態", ("全身強直間代発作で筋活動と脳過剰発火が起こる", "嫌気的解糖が急速に亢進", "ANLSと酸素需給不均衡が乳酸上昇に関与")),
            Section("乳酸値の動態", ("基準値2.4 mmol/Lを超え発作後30分以内に急上昇", "平均8.7倍へ上がり、その後速やかに正常化", "発作模倣症との鑑別補助に使える"), (("基準", "2.4"), ("上昇", "8.7倍"), ("採血", "30分以内"))),
            Section("診断精度", ("全身強直間代発作の判別AUROC 0.94-0.97", "重積状態予測AUC 0.908", "プロラクチン、アンモニア、CKより有用な場面がある"), (("AUROC", "0.94-0.97"), ("AUC", "0.908"))),
            Section("ミトコンドリア病", ("MELASではてんかん合併が70-96%", "基礎乳酸高値に発作が重なりさらに上昇", "治療反応性や病勢評価にも注意"), (("MELAS合併", "70-96%"),)),
        ),
    ),
    Entry(
        "6802888565272474855.png",
        "ミトコンドリア脳筋症",
        "疫学・分類・診断・治療",
        ("#7B3FF2", "#4A90E2", "#1A3A5C"),
        "2025-10-04 参考記事を確認",
        (
            Section("疫学と遺伝", ("欧米20/10万、日本2.9/10万の報告", "日本患者数は2018-2019年で約3,629人", "mtDNA母系遺伝と核DNAメンデル遺伝が混在"), (("欧米", "20/10万"), ("日本", "2.9/10万"), ("患者数", "約3,629"))),
            Section("主要症候群", ("MELASはA3243G変異が80%以上", "MERRFはA8344G変異、KSSは眼瞼下垂など", "Leigh症は乳児脳幹病変が中心"), (("MELAS", "A3243G"), ("MERRF", "A8344G"))),
            Section("診断マーカー", ("血漿・髄液乳酸高値、CK上昇を確認", "筋生検RRF、遺伝子検査を組み合わせる", "MRI所見は本文側で原典画像を扱い、ここでは文字化")),
            Section("治療と新展開", ("CoQ10、L-カルニチン、イデベノンなどを検討", "運動療法と対症療法が重要", "遺伝子治療・ミトコンドリア置換療法は研究段階"), (("MELAS生存", "約17年"), ("Leigh", "約7年"))),
        ),
    ),
    Entry(
        "6802888565290397618.png",
        "副腎白質ジストロフィー",
        "ALD総説",
        ("#8B4513", "#50C878", "#1A3A5C"),
        "2025-10-19 参考記事を確認",
        (
            Section("遺伝と病因", ("ABCD1遺伝子変異によるX連鎖性疾患", "ALDP欠損でVLCFAが蓄積", "日本の指定難病で2019年度受給者証保持者249人"), (("頻度", "1/14,000-1/17,000"), ("日本", "249人"))),
            Section("臨床病型", ("小児大脳型は5-10歳発症で最頻", "思春期型、成人大脳型、AMN、アジソン型に分類", "AMNは成人男性ALDの約半数"), (("小児", "5-10歳"), ("AMN", "約50%"), ("副腎不全", "約80%"))),
            Section("診断", ("血漿VLCFA、C26:0、C24:0/C22:0比を測定", "C26:0-LPCは新生児マススクリーニングに応用", "ABCD1遺伝子解析で確定")),
            Section("治療", ("早期の造血幹細胞移植が予後を左右", "副腎ホルモン補充と対症療法を併用", "レリグリタゾンや遺伝子治療など新規治療も注目"), (("病勢停止", "35%/96週"),)),
        ),
    ),
    Entry(
        "6802888565291699308.png",
        "京都市宿泊税",
        "2026年3月の引き上げ",
        ("#5A4A7F", "#DAA520", "#1A3A5C"),
        "2025-10-19 参考記事を確認",
        (
            Section("政策背景", ("オーバーツーリズムと市民生活の摩擦が深刻化", "文化財保護と観光インフラ整備の財源を確保", "ホテル、旅館、簡易宿所、民泊が課税対象")),
            Section("新税率5段階", ("6,000円未満: 200円", "6,000-20,000円: 400円 / 20,000-50,000円: 1,000円", "50,000-100,000円: 4,000円 / 100,000円以上: 10,000円"), (("<6千円", "200円"), ("2-5万円", "1,000円"), (">=10万円", "10,000円"))),
            Section("全国比較", ("東京都は100-200円の定額方式", "大阪府は2025年9月改定で最高500円", "京都市の最高10,000円は全国最高水準"), (("東京", "100-200円"), ("大阪", "最高500円"), ("京都", "最高1万円"))),
            Section("税収用途", ("年間約126億円で現行52億円の倍以上を見込む", "交通混雑緩和、文化財保護、観光インフラに充当", "修学旅行は免除対象"), (("税収", "約126億円"), ("現行", "52億円"))),
        ),
    ),
    Entry(
        "6802888565297668024.png",
        "ECT(mECT)",
        "適応・手順・安全性・最新動向",
        ("#001F3F", "#2ECC40", "#1A3A5C"),
        "2025-10-25 参考記事を確認",
        (
            Section("適応と有効率", ("難治性大うつ病、躁病、統合失調症急性増悪、カタトニア等", "大うつ病の反応率50-80%、寛解30-60%", "躁病は著明改善約80%"), (("反応", "50-80%"), ("寛解", "30-60%"), ("躁病", "約80%"))),
            Section("実施手順", ("修正型ECTは全身麻酔と筋弛緩薬を併用", "両側性または片側性電極を選択", "週2-3回、急性期合計6-12回が目安"), (("頻度", "週2-3回"), ("急性期", "6-12回"), ("禁食", "6時間"))),
            Section("安全性", ("死亡リスクは約5-8万回に1回", "24時間以内死亡は10万回中2.4回", "錯乱、健忘、頭痛などを説明し観察"), (("死亡", "1/5-8万回"), ("24h", "2.4/10万"))),
            Section("ガイドラインと動向", ("日本精神神経学会2013、APA 2025改訂を踏まえる", "維持ECTは数週から数か月間隔", "片側性電極は記憶障害軽減が期待される")),
        ),
    ),
    Entry(
        "6802888565300846715.png",
        "V180I gCJD",
        "日本特有の緩徐型プリオン病",
        ("#4A148C", "#7B1818", "#1A3A5C"),
        "2025-10-27 参考記事を確認",
        (
            Section("疫学", ("gCJDはプリオン病全体の約10-15%", "日本のgCJDでV180Iは41.2%と最多", "平均発症は70歳前後で高齢発症"), (("gCJD", "10-15%"), ("V180I", "41.2%"), ("発症", "70歳前後"))),
            Section("臨床像", ("孤発性CJDより緩徐で1-2年の病程", "認知機能低下、物忘れ、見当識障害で始まりやすい", "ミオクローヌスや小脳症状は比較的少ない"), (("病程", "1-2年"),)),
            Section("診断マーカー", ("MRI DWIの皮質リボン状高信号が重要", "PSDは低頻度で脳波は決め手になりにくい", "RT-QuICのPrPSc検出率は約39%と低め"), (("PrPSc", "39%"), ("PSD", "低頻度"))),
            Section("病理と治療", ("大脳皮質中心の海綿状変性とグリオーシス", "確立した根治療法はない", "対症療法、家族支援、遺伝カウンセリングが柱")),
        ),
    ),
    Entry(
        "17179246901307699152.png",
        "運動分解 vs 測定障害",
        "小脳症候の鑑別",
        ("#1976D2", "#FB8C00", "#1A3A5C"),
        "2025-10-30 参考記事を確認",
        (
            Section("運動分解", ("複合関節運動が滑らかに統合されない", "肩、肘、手首が連動せず段階的に動く", "小脳脚障害で目立つことがある")),
            Section("測定障害", ("目標到達の距離と力加減の制御障害", "オーバーシュートまたはアンダーシュートが起こる", "終末点の精度低下として観察する")),
            Section("臨床評価", ("指鼻試験、踵膝試験で観察", "速度変化、方向転換、反復で悪化を確認", "両者は同時に存在することが多い")),
            Section("使い分け", ("運動分解は動作過程の質的評価", "測定障害は終末点の量的評価", "動画記録と反復観察で表現を統一する")),
        ),
    ),
    Entry(
        "17179246901310366056.png",
        "脊髄梗塞",
        "疫学・血管解剖・危険因子・予防",
        ("#B71C1C", "#D4A017", "#1A3A5C"),
        "2025-11-02 参考記事を確認",
        (
            Section("疫学", ("全脊髄血管障害の約1-2%", "脳卒中全体に占める割合は低いが重大", "危険因子合併率は75%以上"), (("頻度", "1-2%"), ("危険因子", ">75%"))),
            Section("血管解剖", ("ASAは脊髄前方2/3を灌流", "PSAは後方1/3を灌流", "Adamkiewicz動脈は胸腰髄供給で重要"), (("ASA", "前方2/3"), ("PSA", "後方1/3"), ("AKA", "T7-9"))),
            Section("危険因子", ("高血圧、脂質異常症、喫煙、糖尿病を確認", "大動脈解離・瘤・手術、心房細動も重要", "全身性血栓性疾患を見逃さない"), (("高血圧", "約40%"), ("脂質", "約29%"), ("喫煙", "約30%"), ("糖尿病", "約16%"))),
            Section("予防", ("血管危険因子を厳格管理", "禁煙と大動脈病変のサーベイランス", "抗血栓療法は病態に応じて個別判断")),
        ),
    ),
    Entry(
        "17179246901310649451.png",
        "沼・池・湖・潟",
        "自然地理学から見た違い",
        ("#0D47A1", "#2E7D8C", "#1A3A5C"),
        "2025-11-02 参考記事を確認",
        (
            Section("湖 Lake", ("水深5-10m以上の大規模水域", "岸辺植生は周縁部に限られる", "断層湖、カルデラ湖、堰止湖など成因が多様"), (("水深", "5-10m以上"),)),
            Section("沼 Marsh/Swamp", ("水深5m以内の浅い水域", "湖底全面に沈水植物が繁茂しやすい", "湖との境界は連続的で線引きが曖昧"), (("水深", "5m以内"),)),
            Section("池 Pond", ("小規模水域で人造の貯水池・ため池が多い", "農業用、防災用、景観用など目的が明確", "自然池との区別は成因と規模で考える")),
            Section("潟 Lagoon", ("砂州や堆積で外海から隔てられた沿岸浅水域", "汽水で塩分が混在することが多い", "干拓や開発の歴史的影響も受ける")),
        ),
    ),
    Entry(
        "17179246901312223080.png",
        "PD Yahr II × 就労",
        "身体障害 vs 精神障害手帳",
        ("#C46A4F", "#7B5FA8", "#1A3A5C"),
        "2025-11-05 参考記事を確認",
        (
            Section("Yahr II度", ("両側障害だが軽微で姿勢反射障害はない", "日常生活ADLは自立し身体機能は概ね維持", "身体障害認定では低等級になりやすい")),
            Section("手帳の評価軸", ("身体障害手帳は肢体機能障害の程度を評価", "精神障害保健福祉手帳はうつ、疲労、集中力低下による就労困難を評価", "再就労では障壁の本質を見極める"), (("身体", "6級相当"), ("精神", "3級"))),
            Section("就労支援ルート", ("障害者雇用枠、就労移行支援、リワーク支援を接続", "自立支援医療で精神通院の自己負担を軽減", "合理的配慮を具体化する")),
            Section("実装フロー", ("症状安定化から精神障害手帳申請へ", "リワーク参加後に週20-30時間の障害者雇用枠を検討", "ジョブコーチと定着支援を薬物療法と並行"), (("勤務", "週20-30h"), ("医療費", "1割負担"))),
        ),
    ),
)


def hex_to_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        r"C:\Windows\Fonts\BIZ-UDGothicB.ttc" if bold else r"C:\Windows\Fonts\BIZ-UDGothicR.ttc",
        r"C:\Windows\Fonts\meiryob.ttc" if bold else r"C:\Windows\Fonts\meiryo.ttc",
        r"C:\Windows\Fonts\NotoSansJP-VF.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


F_TITLE = font(54, True)
F_SUB = font(29, True)
F_HEAD = font(31, True)
F_BODY = font(24, True)
F_SMALL = font(20)
F_STAT = font(37, True)
F_STAT_LABEL = font(18, True)
F_NUM = font(52, True)


def text_width(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont) -> int:
    if not text:
        return 0
    box = draw.textbbox((0, 0), text, font=fnt)
    return box[2] - box[0]


def wrap_text(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont, max_width: int) -> list[str]:
    tokens = re.split(r"(\s+)", text)
    if len(tokens) > 1:
        lines: list[str] = []
        line = ""
        for tok in tokens:
            trial = line + tok
            if text_width(draw, trial, fnt) <= max_width:
                line = trial
            else:
                if line.strip():
                    lines.append(line.strip())
                line = tok.strip()
        if line.strip():
            lines.append(line.strip())
        return lines
    lines = []
    line = ""
    for ch in text:
        trial = line + ch
        if text_width(draw, trial, fnt) <= max_width:
            line = trial
        else:
            if line:
                lines.append(line)
            line = ch
    if line:
        lines.append(line)
    return lines


def rounded_rect(draw: ImageDraw.ImageDraw, xy: tuple[int, int, int, int], fill, outline=None, width=1, radius=28) -> None:
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def draw_abstract_icon(draw: ImageDraw.ImageDraw, cx: int, cy: int, accent: tuple[int, int, int], secondary: tuple[int, int, int], idx: int) -> None:
    # Geometric-only icons: no anatomy, no devices, no real medical images.
    r = 34
    if idx % 4 == 0:
        draw.rounded_rectangle((cx - r, cy - r, cx + r, cy + r), radius=12, outline=accent, width=5)
        for off in (-18, 0, 18):
            draw.line((cx - 48, cy + off, cx - 34, cy + off), fill=secondary, width=4)
            draw.line((cx + 34, cy + off, cx + 48, cy + off), fill=secondary, width=4)
        draw.ellipse((cx - 12, cy - 12, cx + 12, cy + 12), fill=accent)
    elif idx % 4 == 1:
        draw.arc((cx - 42, cy - 42, cx + 42, cy + 42), start=20, end=320, fill=accent, width=6)
        draw.polygon([(cx + 42, cy - 8), (cx + 62, cy + 2), (cx + 40, cy + 14)], fill=accent)
        draw.line((cx - 26, cy + 20, cx - 4, cy - 12, cx + 18, cy + 10), fill=secondary, width=6)
    elif idx % 4 == 2:
        for i, h in enumerate((30, 54, 42)):
            x = cx - 45 + i * 34
            draw.rounded_rectangle((x, cy + 34 - h, x + 22, cy + 34), radius=5, fill=accent if i != 1 else secondary)
        draw.line((cx - 54, cy + 36, cx + 54, cy + 36), fill=(205, 214, 224), width=4)
    else:
        for a in range(0, 360, 60):
            x = cx + int(math.cos(math.radians(a)) * 38)
            y = cy + int(math.sin(math.radians(a)) * 38)
            draw.line((cx, cy, x, y), fill=secondary, width=4)
            draw.ellipse((x - 8, y - 8, x + 8, y + 8), fill=accent)
        draw.ellipse((cx - 12, cy - 12, cx + 12, cy + 12), fill=secondary)


def draw_stat_cards(draw: ImageDraw.ImageDraw, stats: tuple[tuple[str, str], ...], x: int, y: int, w: int, accent, secondary) -> int:
    if not stats:
        return 0
    cols = min(len(stats), 3)
    if cols == 3 and w < 500:
        cols = 2
    gap = 12
    card_w = (w - gap * (cols - 1)) // cols
    rows = math.ceil(len(stats) / cols)
    card_h = 74
    for i, (label, value) in enumerate(stats[:6]):
        row = i // cols
        col = i % cols
        cx = x + col * (card_w + gap)
        cy = y + row * (card_h + 10)
        fill = (248, 251, 255) if i % 2 == 0 else (252, 249, 253)
        rounded_rect(draw, (cx, cy, cx + card_w, cy + card_h), fill=fill, outline=(222, 229, 238), width=2, radius=18)
        stat_font = F_STAT
        for size in (37, 33, 30, 27, 24):
            candidate = font(size, True)
            if text_width(draw, value, candidate) <= card_w - 24:
                stat_font = candidate
                break
        draw.text((cx + 12, cy + 9), value, font=stat_font, fill=accent)
        draw.text((cx + 16, cy + 47), label, font=F_STAT_LABEL, fill=secondary)
    return rows * card_h + (rows - 1) * 10 + 12


def blend(c1: tuple[int, int, int], c2: tuple[int, int, int], t: float) -> tuple[int, int, int]:
    return tuple(int(a * (1 - t) + b * t) for a, b in zip(c1, c2))


def draw_bar_chart(
    draw: ImageDraw.ImageDraw,
    stats: tuple[tuple[str, str], ...],
    x: int,
    y: int,
    w: int,
    h: int,
    accent,
    secondary,
    ink,
) -> None:
    if not stats:
        return
    draw.rounded_rectangle((x, y, x + w, y + h), radius=18, fill=(248, 251, 255), outline=(213, 224, 237), width=2)
    vals: list[float] = []
    for _, value in stats[:4]:
        m = re.search(r"(\d+(?:\.\d+)?)", value.replace(",", ""))
        vals.append(float(m.group(1)) if m else 1.0)
    max_v = max(vals) if vals else 1.0
    base = y + h - 42
    bar_w = max(36, (w - 64) // max(1, len(vals)) - 12)
    for i, ((label, value), v) in enumerate(zip(stats[:4], vals)):
        bx = x + 28 + i * ((w - 56) // max(1, len(vals)))
        bh = max(22, int((h - 92) * (v / max_v)))
        color = accent if i % 2 == 0 else secondary
        draw.rounded_rectangle((bx, base - bh, bx + bar_w, base), radius=8, fill=color)
        draw.text((bx, base - bh - 35), value, font=F_STAT_LABEL, fill=color)
        for line_i, line in enumerate(wrap_text(draw, label, F_STAT_LABEL, bar_w + 34)[:2]):
            draw.text((bx - 2, base + 8 + line_i * 18), line, font=F_STAT_LABEL, fill=ink)


def draw_flow(draw: ImageDraw.ImageDraw, words: list[str], x: int, y: int, w: int, accent, secondary, ink) -> None:
    words = words[:4]
    if not words:
        return
    gap = 16
    box_w = (w - gap * (len(words) - 1)) // len(words)
    for i, word in enumerate(words):
        bx = x + i * (box_w + gap)
        draw.rounded_rectangle((bx, y, bx + box_w, y + 66), radius=18, fill=(255, 255, 255), outline=accent if i % 2 == 0 else secondary, width=3)
        tx_lines = wrap_text(draw, word, F_STAT_LABEL, box_w - 20)[:2]
        ty = y + 12 if len(tx_lines) == 2 else y + 23
        for j, line in enumerate(tx_lines):
            tw = text_width(draw, line, F_STAT_LABEL)
            draw.text((bx + (box_w - tw) // 2, ty + j * 19), line, font=F_STAT_LABEL, fill=ink)
        if i < len(words) - 1:
            ax = bx + box_w + 4
            draw.polygon([(ax, y + 24), (ax + 12, y + 33), (ax, y + 42)], fill=secondary)


def draw_section(draw: ImageDraw.ImageDraw, section: Section, box: tuple[int, int, int, int], accent, secondary, ink, idx: int) -> None:
    x1, y1, x2, y2 = box
    shadow = (0, 0, 0, 28)
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sd = ImageDraw.Draw(layer)
    sd.rounded_rectangle((x1 + 5, y1 + 7, x2 + 5, y2 + 7), radius=24, fill=shadow)
    draw.bitmap((0, 0), layer, fill=None)

    rounded_rect(draw, box, fill=(255, 255, 255), outline=blend(accent, (255, 255, 255), 0.48), width=3, radius=24)
    strip_h = 74
    draw.rounded_rectangle((x1, y1, x2, y1 + strip_h), radius=24, fill=accent)
    draw.rectangle((x1, y1 + 34, x2, y1 + strip_h), fill=accent)
    draw.ellipse((x1 + 23, y1 + 14, x1 + 83, y1 + 74), fill=(255, 255, 255))
    num = str(idx + 1)
    nw = text_width(draw, num, F_NUM)
    draw.text((x1 + 53 - nw // 2, y1 + 13), num, font=F_NUM, fill=accent)
    draw.text((x1 + 104, y1 + 20), section.heading, font=F_HEAD, fill=(255, 255, 255))

    icon_x = x2 - 74
    icon_y = y1 + 37
    draw_abstract_icon(draw, icon_x, icon_y, (255, 255, 255), blend(secondary, (255, 255, 255), 0.22), idx)

    content_y = y1 + 94
    left_w = 560 if section.stats else 890
    right_x = x1 + left_w + 56
    max_w = left_w - 28
    for bullet in section.bullets[:4]:
        lines = wrap_text(draw, bullet, F_BODY, max_w)
        draw.rounded_rectangle((x1 + 30, content_y + 5, x1 + 52, content_y + 27), radius=5, fill=secondary)
        for j, line in enumerate(lines[:2]):
            draw.text((x1 + 66, content_y + j * 30), line, font=F_BODY, fill=ink)
        content_y += 34 * min(len(lines), 2) + 15
        if content_y > y2 - 46:
            break

    if section.stats:
        draw_stat_cards(draw, section.stats[:3], right_x, y1 + 102, x2 - right_x - 28, accent, secondary)
        draw_bar_chart(draw, section.stats, right_x, y1 + 270, x2 - right_x - 28, max(106, y2 - y1 - 298), accent, secondary, ink)
    else:
        words = []
        for bullet in section.bullets:
            words.extend(re.split(r"[、・／/]+", bullet))
        draw_flow(draw, [w for w in words if 2 <= len(w) <= 12], x1 + 64, y2 - 94, x2 - x1 - 128, accent, secondary, ink)


def render(entry: Entry) -> None:
    accent = hex_to_rgb(entry.colors[0])
    secondary = hex_to_rgb(entry.colors[1])
    ink = hex_to_rgb(entry.colors[2])
    bg = Image.new("RGB", (W, H), "#F2F6FA")
    draw = ImageDraw.Draw(bg)

    for gx in range(0, W, 54):
        draw.line((gx, 0, gx, H), fill=(230, 236, 243), width=1)
    for gy in range(0, H, 54):
        draw.line((0, gy, W, gy), fill=(230, 236, 243), width=1)

    # Header band inspired by previous generated assets.
    draw.rounded_rectangle((18, 18, W - 18, 158), radius=14, fill=ink)
    draw.rectangle((18, 115, W - 18, 158), fill=ink)
    draw.ellipse((42, 42, 112, 112), fill=(255, 255, 255))
    draw_abstract_icon(draw, 77, 77, accent, secondary, 0)
    draw.text((140, 40), entry.title, font=F_TITLE, fill=(255, 255, 255))
    draw.text((143, 108), entry.subtitle, font=F_SUB, fill=blend(accent, (255, 255, 255), 0.72))
    draw.rounded_rectangle((W - 290, 116, W - 42, 148), radius=16, fill=accent)
    sw = text_width(draw, entry.source, F_SMALL)
    draw.text((W - 166 - sw // 2, 121), entry.source, font=F_SMALL, fill=(255, 255, 255))

    y = 184
    gap = 18
    section_h = 405
    for idx, section in enumerate(entry.sections):
        draw_section(draw, section, (22, y, W - 22, y + section_h), accent, secondary, ink, idx)
        y += section_h + gap

    footer_y = H - 76
    draw.rounded_rectangle((22, footer_y, W - 22, H - 22), radius=16, fill=ink)
    footer = "医学画像・実写・企業ロゴ・薬剤外観は使わず、抽象アイコンと図表で整理"
    fw = text_width(draw, footer, F_SMALL)
    draw.text(((W - fw) // 2, footer_y + 17), footer, font=F_SMALL, fill=(255, 255, 255))
    bg.save(OUT_DIR / entry.file, "PNG")


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def update_db(entry: Entry) -> None:
    if not DB_PATH.exists():
        return
    entry_id = Path(entry.file).stem
    image_path = str((OUT_DIR / entry.file).resolve())
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute(
            """UPDATE generations
            SET image_path=?, created_at=?
            WHERE entry_id=?
            AND gen_id=(
                SELECT gen_id FROM generations
                WHERE entry_id=? ORDER BY attempt DESC LIMIT 1
            )""",
            (image_path, now_iso(), entry_id, entry_id),
        )
        conn.commit()
    finally:
        conn.close()


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    success = 0
    skipped = 0
    failures: list[tuple[str, str]] = []
    for entry in ENTRIES:
        out_path = OUT_DIR / entry.file
        if out_path.exists():
            skipped += 1
            print(f"[SKIP existing] {entry.file}")
            continue
        try:
            render(entry)
            update_db(entry)
            success += 1
            print(f"[OK] {entry.file}")
        except Exception as exc:  # noqa: BLE001
            failures.append((Path(entry.file).stem, repr(exc)))
            print(f"[FAIL] {entry.file}: {exc}")
    print(f"成功: {success} 件 / 既存スキップ: {skipped} 件 / 失敗: {len(failures)} 件")
    for entry_id, reason in failures:
        print(f"  - {entry_id}: {reason}")
    return 0 if not failures else 2


if __name__ == "__main__":
    raise SystemExit(main())
