#!/usr/bin/env python3
"""Render client journey as a true block-diagram PNG."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "Блок-схема_работы_с_клиентом.png"
OUT_ARTIFACT = Path("/opt/cursor/artifacts/Блок-схема_работы_с_клиентом.png")

FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

# Colors
BG = (245, 248, 252)
HEADER_BG = (30, 58, 76)
WHITE = (255, 255, 255)
DARK = (26, 26, 26)
MUTED = (80, 90, 100)
BLUE = (38, 166, 224)
GREEN = (46, 139, 87)
ORANGE = (217, 119, 6)
BROWN = (180, 83, 9)
LEARN_BG = (232, 244, 251)
TELL_BG = (234, 247, 238)
CTA_BG = (255, 244, 229)
LINE = (150, 170, 190)


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REG, size)


def wrap(draw: ImageDraw.ImageDraw, text: str, fnt, max_width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    cur = ""
    for w in words:
        trial = (cur + " " + w).strip()
        if draw.textlength(trial, font=fnt) <= max_width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


def rounded(draw, box, fill, radius=18, outline=None, width=2):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def draw_text_block(draw, x, y, w, lines, fnt, fill=DARK, line_gap=4, align="left"):
    for i, line in enumerate(lines):
        tw = draw.textlength(line, font=fnt)
        tx = x if align == "left" else x + (w - tw) / 2
        draw.text((tx, y + i * (fnt.size + line_gap)), line, font=fnt, fill=fill)
    return len(lines) * (fnt.size + line_gap)


def arrow_down(draw, x, y1, y2, color=BLUE):
    draw.line((x, y1, x, y2 - 12), fill=color, width=4)
    draw.polygon(
        [(x - 10, y2 - 14), (x + 10, y2 - 14), (x, y2)],
        fill=color,
    )


def build():
    W = 1600
    margin = 60
    content_w = W - 2 * margin

    stages = [
        {
            "num": "1",
            "title": "Первое знакомство",
            "sub": "Холодный звонок · скрипт",
            "color": BLUE,
            "goal": "Понять боль клиента и договориться о следующем шаге",
            "learn": [
                "Кто на линии: руководитель / бухгалтер / кадровик / IT",
                "Как сейчас: бумага или уже электронно",
                "Расчётные листки и подтверждение ознакомления",
                "Отпуска, справки, удалёнка, число сотрудников",
                "ЭДО с контрагентами · командировки",
            ],
            "tell": [
                "Мы из «Форуса», ваш партнёр по 1С",
                "ДОКИ: 3 месяца безлимита по ИТС ПРОФ",
                "Кабинет сотрудника — как Госуслуги для сотрудников",
                "Старт с одного отдела, меньше 30 ₽/чел в месяц",
                "Предложить короткий показ",
            ],
            "cta": "Результат: договорённость о презентации / демо",
        },
        {
            "num": "2",
            "title": "Преддемонстрация",
            "sub": "Презентация · запись на демо",
            "color": GREEN,
            "goal": "Показать ценность и записать на демо со специалистом",
            "learn": [
                "Кто придёт на демо",
                "Что важнее увидеть: расчётки / отпуска / заявления",
                "Сколько сотрудников на пилот (цель — 10)",
                "Что смущает и удобные дата/время",
            ],
            "tell": [
                "Пройти презентацию под боль клиента",
                "Выгоды: меньше рутины, данные сразу в 1С",
                "Мягкий вход: 10 кабинетов ≈ 3 360 ₽/год",
                "Письмо: дата, время, ссылка → ту же ссылку специалисту",
                "Подтвердить участие за 1 день и за 1 час",
            ],
            "cta": "Результат: клиент записан и подтвердил участие",
        },
        {
            "num": "3",
            "title": "Демонстрация сервиса",
            "sub": "Онлайн-демо со специалистом",
            "color": ORANGE,
            "goal": "Показать сервис «вживую» на задачах клиента",
            "learn": [
                "Какой сценарий показать первым",
                "Нужны ли вопросы по безопасности / IT",
                "Решает ли сервис задачу?",
                "Что мешает стартовать и кто согласует оплату",
            ],
            "tell": [
                "Менеджер: знакомство + короткий план демо",
                "Специалист: кабинет, расчётки, заявление → в 1С",
                "Согласование руководителем, ответы по подписи",
                "На демо не давим на покупку",
                "В течение 3 дней — созвон за обратной связью",
            ],
            "cta": "Результат: клиент увидел пользу и готов обсуждать пилот",
        },
        {
            "num": "4",
            "title": "Закрытие на счёт",
            "sub": "Пилотный запуск · 10 кабинетов",
            "color": BROWN,
            "goal": "Перевести интерес в пилот без давления",
            "learn": [
                "Что понравилось на демо",
                "Каких 10 сотрудников / какой отдел берём",
                "С чего стартуем: расчётки или отпуска",
                "Кто оплачивает и удобная дата подключения",
            ],
            "tell": [
                "Пилот на 10 кабинетов — без перевода всей компании",
                "Настройку и обучение берём на себя",
                "Выставить счёт и согласовать старт",
                "Через 2–4 недели — разговор о расширении",
            ],
            "cta": "Результат: счёт выставлен · пилот запущен",
        },
    ]

    # Pre-measure height
    title_h = 120
    start_h = 70
    gap = 36
    stage_heights = []
    f_title = font(28, True)
    f_sub = font(16)
    f_sec = font(15, True)
    f_body = font(15)
    f_goal = font(16, True)
    f_cta = font(16, True)
    f_small = font(13)

    # temp image for measuring
    tmp = Image.new("RGB", (W, 100), BG)
    dtmp = ImageDraw.Draw(tmp)

    for st in stages:
        # header band + goal + two columns + cta
        learn_lines = []
        for item in st["learn"]:
            learn_lines.extend(wrap(dtmp, "• " + item, f_body, content_w // 2 - 50))
        tell_lines = []
        for item in st["tell"]:
            tell_lines.extend(wrap(dtmp, "• " + item, f_body, content_w // 2 - 50))
        col_h = max(len(learn_lines), len(tell_lines)) * (f_body.size + 6) + 50
        goal_lines = wrap(dtmp, st["goal"], f_goal, content_w - 60)
        cta_lines = wrap(dtmp, st["cta"], f_cta, content_w - 60)
        h = 70 + 10 + len(goal_lines) * 24 + 16 + col_h + 16 + 50
        stage_heights.append(h)
        st["_learn_lines"] = learn_lines
        st["_tell_lines"] = tell_lines
        st["_goal_lines"] = goal_lines
        st["_cta_lines"] = cta_lines
        st["_col_h"] = col_h

    end_h = 90
    H = (
        margin
        + title_h
        + gap
        + start_h
        + gap
        + sum(stage_heights)
        + gap * len(stages)
        + end_h
        + margin
        + 40
    )

    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    # Title bar
    rounded(draw, (margin, margin, W - margin, margin + title_h - 10), HEADER_BG, 20)
    draw_text_block(
        draw,
        margin,
        margin + 22,
        content_w,
        ["Блок-схема работы с клиентом", "«1С:Кабинет сотрудника»"],
        font(30, True),
        WHITE,
        6,
        "center",
    )
    draw_text_block(
        draw,
        margin,
        margin + 88,
        content_w,
        ["На каждом этапе: что узнать → что рассказать → следующий шаг"],
        f_small,
        (200, 220, 230),
        2,
        "center",
    )

    y = margin + title_h + gap
    cx = W // 2

    # Start oval
    start_box = (cx - 180, y, cx + 180, y + start_h)
    rounded(draw, start_box, BLUE, 35)
    draw_text_block(
        draw,
        cx - 180,
        y + 22,
        360,
        ["СТАРТ · контакт с клиентом"],
        font(18, True),
        WHITE,
        2,
        "center",
    )
    y2 = y + start_h
    arrow_down(draw, cx, y2 + 4, y2 + gap, BLUE)
    y = y2 + gap

    for i, st in enumerate(stages):
        x0, x1 = margin, W - margin
        y0 = y
        y1 = y + stage_heights[i]
        color = st["color"]

        # Outer stage card
        rounded(draw, (x0, y0, x1, y1), WHITE, 22, outline=color, width=3)

        # Colored header
        rounded(draw, (x0, y0, x1, y0 + 64), color, 22)
        draw.rectangle((x0, y0 + 40, x1, y0 + 64), fill=color)

        # Number circle
        draw.ellipse((x0 + 18, y0 + 12, x0 + 54, y0 + 48), fill=WHITE)
        draw_text_block(
            draw, x0 + 18, y0 + 18, 36, [st["num"]], font(20, True), color, 0, "center"
        )
        draw.text((x0 + 70, y0 + 12), st["title"], font=font(22, True), fill=WHITE)
        draw.text((x0 + 70, y0 + 38), st["sub"], font=f_small, fill=(240, 245, 250))

        yy = y0 + 78
        # Goal
        for line in st["_goal_lines"]:
            draw.text((x0 + 28, yy), line, font=f_goal, fill=DARK)
            yy += 24
        yy += 10

        # Two columns
        col_w = (content_w - 40) // 2
        left = (x0 + 20, yy, x0 + 20 + col_w, yy + st["_col_h"])
        right = (x0 + 30 + col_w, yy, x1 - 20, yy + st["_col_h"])
        rounded(draw, left, LEARN_BG, 14, outline=(180, 210, 230), width=1)
        rounded(draw, right, TELL_BG, 14, outline=(170, 210, 180), width=1)

        draw.text((left[0] + 16, yy + 12), "УЗНАТЬ", font=f_sec, fill=BLUE)
        draw.text((right[0] + 16, yy + 12), "РАССКАЗАТЬ", font=f_sec, fill=GREEN)

        ty = yy + 38
        for line in st["_learn_lines"]:
            draw.text((left[0] + 16, ty), line, font=f_body, fill=DARK)
            ty += f_body.size + 6

        ty = yy + 38
        for line in st["_tell_lines"]:
            draw.text((right[0] + 16, ty), line, font=f_body, fill=DARK)
            ty += f_body.size + 6

        # CTA strip
        cta_y = yy + st["_col_h"] + 12
        rounded(draw, (x0 + 20, cta_y, x1 - 20, y1 - 14), CTA_BG, 12)
        draw_text_block(
            draw,
            x0 + 20,
            cta_y + 12,
            content_w - 40,
            st["_cta_lines"],
            f_cta,
            BROWN if i == 3 else DARK,
            2,
            "center",
        )

        y = y1
        if i < len(stages) - 1:
            arrow_down(draw, cx, y + 4, y + gap, color)
            y += gap
        else:
            arrow_down(draw, cx, y + 4, y + gap, BROWN)
            y += gap

    # End
    end_box = (cx - 260, y, cx + 260, y + end_h)
    rounded(draw, end_box, BROWN, 24)
    draw_text_block(
        draw,
        cx - 260,
        y + 18,
        520,
        ["ФИНИШ", "Счёт на 10 кабинетов · пилот запущен"],
        font(20, True),
        WHITE,
        6,
        "center",
    )

    img.save(OUT, "PNG", optimize=True)
    OUT_ARTIFACT.parent.mkdir(parents=True, exist_ok=True)
    img.save(OUT_ARTIFACT, "PNG", optimize=True)
    print(f"Saved: {OUT} ({img.size[0]}x{img.size[1]})")
    print(f"Artifact: {OUT_ARTIFACT}")


if __name__ == "__main__":
    build()
