#!/usr/bin/env python3
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(r"C:\Users\jsber\OneDrive\Documents\Claude_task\blog\symptom-concern-series")
OUT_DIR = ROOT / "infographics"
OUT_PATH = OUT_DIR / "06_autonomic_nerves.png"

W, H = 1620, 2880
BG = "#F8F9FA"
NAVY = "#1B3A5C"
BLUE = "#2C5AA0"
ORANGE = "#E8850C"
RED = "#DC3545"
GREEN = "#28A745"
YELLOW_BG = "#FFF3CD"
YELLOW_BORDER = "#F5C518"
TEXT = "#1E2430"
MUTED = "#5D6A7A"
CARD = "#FFFFFF"
BORDER = "#D9E2EC"
LIGHT_BLUE = "#EAF3FF"
LIGHT_RED = "#FDECEE"
LIGHT_GREEN = "#EAF7EE"
LIGHT_NAVY = "#EEF4FA"

FONT_REG = r"C:\Windows\Fonts\NotoSansJP-VF.ttf"
FONT_BOLD = r"C:\Windows\Fonts\BIZ-UDGothicB.ttc"


def font(size: int, bold: bool = False):
    path = FONT_BOLD if bold else FONT_REG
    return ImageFont.truetype(path, size)


def text_size(draw, text, fnt):
    box = draw.multiline_textbbox((0, 0), text, font=fnt, spacing=6)
    return box[2] - box[0], box[3] - box[1]


def wrap(draw, text, fnt, max_w):
    lines = []
    for para in text.split("\n"):
        if not para:
            lines.append("")
            continue
        cur = ""
        for ch in para:
            trial = cur + ch
            if text_size(draw, trial, fnt)[0] <= max_w or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = ch
        if cur:
            lines.append(cur)
    return "\n".join(lines)


def rounded(draw, xy, fill, outline=None, width=1, radius=28):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def shadowed_card(base, xy, fill=CARD, outline=BORDER, radius=28):
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    x1, y1, x2, y2 = xy
    od.rounded_rectangle((x1 + 8, y1 + 10, x2 + 8, y2 + 10), radius=radius, fill=(18, 38, 63, 24))
    base.alpha_composite(overlay)
    draw = ImageDraw.Draw(base)
    rounded(draw, xy, fill, outline, 2, radius)


def badge(draw, x, y, label, fill=BLUE):
    r = 28
    draw.ellipse((x - r, y - r, x + r, y + r), fill=fill)
    fnt = font(28, True)
    tw, th = text_size(draw, label, fnt)
    draw.text((x - tw / 2, y - th / 2 - 2), label, font=fnt, fill="white")


def bullet_block(draw, x, y, items, width, fnt, color=TEXT, bullet_fill=BLUE, gap=12):
    cy = y
    for item in items:
        bullet_y = cy + 14
        draw.ellipse((x, bullet_y, x + 12, bullet_y + 12), fill=bullet_fill)
        wrapped = wrap(draw, item, fnt, width - 28)
        draw.multiline_text((x + 24, cy), wrapped, font=fnt, fill=color, spacing=6)
        _, th = text_size(draw, wrapped, fnt)
        cy += th + gap
    return cy


def label_chip(draw, x, y, text, fill, text_fill="white", pad_x=18, pad_y=10):
    fnt = font(22, True)
    tw, th = text_size(draw, text, fnt)
    rounded(draw, (x, y, x + tw + pad_x * 2, y + th + pad_y * 2), fill=fill, radius=18)
    draw.text((x + pad_x, y + pad_y - 2), text, font=fnt, fill=text_fill)
    return x + tw + pad_x * 2


def draw_title(draw):
    rounded(draw, (48, 36, W - 48, 230), fill=NAVY, radius=36)
    draw.text((84, 64), "「自律神経失調症」と言われたら", font=font(64, True), fill="white")
    draw.text((84, 138), "本当の原因と正しい受診先", font=font(56, True), fill="#F5C518")
    draw.text((86, 196), "脳神経内科医が解説  その症状、大丈夫？ #06", font=font(26, False), fill="#D8E6F5")


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGBA", (W, H), BG)
    draw = ImageDraw.Draw(img)

    for i in range(H):
        c = 248 - int(i * 8 / H)
        draw.line((0, i, W, i), fill=(248, 249, 250, 255) if i < H / 2 else (244, 247, 250, 255))
    draw.ellipse((-180, -100, 380, 460), fill="#E3EEF9")
    draw.ellipse((W - 500, 120, W + 120, 760), fill="#FCE6D2")
    draw.ellipse((W - 360, H - 420, W + 90, H + 40), fill="#E6F4EA")

    draw_title(draw)

    # 1. role
    draw.text((72, 274), "① 自律神経の役割", font=font(34, True), fill=NAVY)
    shadowed_card(img, (72, 326, 774, 622), fill="#FFF3F4", outline="#F2C9CF")
    shadowed_card(img, (846, 326, 1548, 622), fill="#EEF9F2", outline="#CAE8D1")
    badge(draw, 126, 378, "A", RED)
    badge(draw, 900, 378, "B", GREEN)
    draw.text((168, 348), "交感神経（アクセル）", font=font(34, True), fill=RED)
    draw.text((942, 348), "副交感神経（ブレーキ）", font=font(34, True), fill=GREEN)
    bullet_block(draw, 110, 420, ["心拍数↑・血圧↑", "瞳孔拡大・消化抑制", "緊張・興奮・活動時"], 610, font(26), bullet_fill=RED)
    bullet_block(draw, 884, 420, ["心拍数↓・血圧↓", "消化促進・体の回復", "リラックス・休息時"], 610, font(26), bullet_fill=GREEN)
    rounded(draw, (96, 640, W - 96, 724), fill=LIGHT_BLUE, outline="#C9DBEF", width=2, radius=24)
    note = "この2つがシーソーのようにバランスを取る。崩れると、めまい・動悸・だるさなどが起こる。"
    draw.text((126, 664), wrap(draw, note, font(26, True), W - 252), font=font(26, True), fill=BLUE)

    # 2. symptoms + identity
    draw.text((72, 768), "② よくある症状 / 「自律神経失調症」の正体", font=font(34, True), fill=NAVY)
    shadowed_card(img, (72, 820, W - 72, 1114), fill=CARD)
    symptoms = [
        ("循環器系", "動悸・立ちくらみ"),
        ("消化器系", "吐き気・胃もたれ"),
        ("体温調節", "多汗 / 無汗・微熱"),
        ("全身", "めまい・だるさ"),
        ("睡眠", "不眠・途中で覚醒"),
        ("精神面", "不安感・イライラ"),
    ]
    x1, x2 = 108, 822
    y = 864
    for idx, (head, body) in enumerate(symptoms):
        col_x = x1 if idx % 2 == 0 else x2
        row_y = y + (idx // 2) * 86
        rounded(draw, (col_x, row_y, col_x + 606, row_y + 66), fill=LIGHT_NAVY, outline="#D8E4F0", width=1, radius=18)
        draw.text((col_x + 18, row_y + 14), head, font=font(22, True), fill=NAVY)
        draw.text((col_x + 180, row_y + 14), body, font=font(22), fill=TEXT)
    rounded(draw, (102, 1094, W - 102, 1328), fill=YELLOW_BG, outline=YELLOW_BORDER, width=3, radius=28)
    draw.text((132, 1118), "正式な病名ではない", font=font(34, True), fill="#8A6500")
    body = [
        "ICD・DSMに載っていない",
        "検査で異常が見つからないときの仮の診断名",
        "この言葉で安心すると、治療可能な病気の見逃しにつながる",
    ]
    bullet_block(draw, 132, 1168, body, W - 264, font(26), color=TEXT, bullet_fill=ORANGE, gap=10)

    # 3. hidden causes
    draw.text((72, 1370), "③ 本当の原因  隠れた病気6つ", font=font(34, True), fill=NAVY)
    causes = [
        ("1", "甲状腺疾患", "動悸・発汗・体重変化\n血液検査"),
        ("2", "鉄欠乏性貧血", "めまい・だるさ・息切れ\n月経のある女性に多い"),
        ("3", "うつ病・不安障害", "倦怠感・不眠・食欲低下\n適切な治療で改善"),
        ("4", "起立性調節障害", "立ち上がりでめまい\n若年者に多い"),
        ("5", "不整脈", "動悸の原因が心臓\n心電図 / ホルター心電図"),
        ("6", "副腎不全", "慢性のだるさ・低血圧\nホルモン検査"),
    ]
    top = 1422
    card_w, card_h = 464, 158
    gap_x, gap_y = 42, 26
    for i, (num, head, body) in enumerate(causes):
        cx = 72 + (i % 3) * (card_w + gap_x)
        cy = top + (i // 3) * (card_h + gap_y)
        shadowed_card(img, (cx, cy, cx + card_w, cy + card_h), fill=CARD)
        badge(draw, cx + 42, cy + 38, num, BLUE)
        draw.text((cx + 82, cy + 18), head, font=font(26, True), fill=NAVY)
        draw.multiline_text((cx + 24, cy + 62), wrap(draw, body, font(22), card_w - 48), font=font(22), fill=TEXT, spacing=6)

    # 4. tests
    draw.text((72, 1796), "④ まず受けるべき5つの検査", font=font(34, True), fill=NAVY)
    tests = [
        ("1", "血液検査", "貧血・甲状腺機能・血糖値・肝腎機能・電解質"),
        ("2", "心電図", "不整脈・虚血性心疾患"),
        ("3", "起立試験（シェロンテスト）", "起立性低血圧・起立性調節障害"),
        ("4", "尿検査", "腎臓・副腎の異常"),
        ("5", "頭部MRI（必要時）", "脳幹・小脳の異常"),
    ]
    y0 = 1848
    for idx, (num, name, desc) in enumerate(tests):
        cy = y0 + idx * 94
        shadowed_card(img, (72, cy, W - 72, cy + 74), fill=LIGHT_NAVY, outline="#CFDDEA", radius=22)
        badge(draw, 118, cy + 38, num, BLUE)
        draw.text((164, cy + 14), name, font=font(24, True), fill=NAVY)
        draw.text((520, cy + 14), wrap(draw, desc, font(22), 940), font=font(22), fill=TEXT)

    # 5. warning signs
    draw.text((72, 2328), "⑤ 危険なサイン  すぐ受診", font=font(34, True), fill=RED)
    rounded(draw, (72, 2380, 1004, 2588), fill=LIGHT_RED, outline="#F2B8BF", width=3, radius=28)
    warning_items = [
        "急な激しいめまい＋嘔吐 → 脳卒中の可能性、119番",
        "失神・意識消失 → 不整脈・てんかん",
        "急な体重減少 / 大量の寝汗 / 動悸＋胸痛＋息苦しさ",
    ]
    bullet_block(draw, 104, 2414, warning_items, 854, font(24), color=TEXT, bullet_fill=RED, gap=12)

    # 6. departments
    rounded(draw, (1034, 2380, 1548, 2588), fill=LIGHT_BLUE, outline="#B9D1E8", width=3, radius=28)
    draw.text((1062, 2404), "⑥ 症状で選ぶ受診先", font=font(24, True), fill=NAVY)
    dept_lines = [
        "めまい・しびれ → 脳神経内科",
        "動悸・胸痛 → 循環器内科",
        "不眠・不安 → 心療内科 / 精神科",
        "倦怠感・体重変化 → 内分泌内科",
        "迷うとき → まず内科",
    ]
    bullet_block(draw, 1062, 2450, dept_lines, 430, font(19), color=TEXT, bullet_fill=BLUE, gap=8)

    # 7. habits and summary strip
    rounded(draw, (72, 2620, W - 72, 2832), fill="#F3FBF6", outline="#CBE6D3", width=3, radius=28)
    draw.text((98, 2646), "⑦ 自律神経を整える5つの習慣", font=font(28, True), fill=GREEN)
    habits = [
        "規則正しい睡眠",
        "有酸素運動 週3回30分",
        "38〜40℃で15分入浴",
        "腹式呼吸 4-7-8",
        "就寝1時間前はスマホOFF",
    ]
    hx = 98
    hy = 2692
    for i, item in enumerate(habits, start=1):
        chip = f"{i}. {item}"
        hx_end = label_chip(draw, hx, hy, chip, "#DFF3E5", text_fill="#205C35", pad_x=14, pad_y=8)
        hx = hx_end + 12
        if hx > 1380 and i < len(habits):
            hx = 98
            hy += 52

    draw.text((98, 2770), "まとめ", font=font(24, True), fill=NAVY)
    summary = "1. 正式な病名ではない  2. 隠れた病気を見逃さない  3. 症状に合った診療科へ"
    draw.text((186, 2770), wrap(draw, summary, font(22, True), 1280), font=font(22, True), fill=TEXT)

    img = img.convert("RGB")
    img.save(OUT_PATH, quality=95)
    print(OUT_PATH)


if __name__ == "__main__":
    main()
