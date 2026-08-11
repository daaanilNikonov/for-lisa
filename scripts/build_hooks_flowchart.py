#!/usr/bin/env python3
"""Block-diagram sales script: ДОКИ ЭДО → ДОКИ логистика → КЭДО → Смартвей."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "Блок-схема_заходы_сервисы.png"
ARTIFACT = Path("/opt/cursor/artifacts/Блок-схема_заходы_сервисы.png")

FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

BG = (245, 248, 252)
HEADER = (30, 58, 76)
WHITE = (255, 255, 255)
DARK = (26, 26, 26)
MUTED = (70, 80, 90)
BLUE = (38, 166, 224)
TEAL = (15, 118, 110)
GREEN = (46, 139, 87)
ORANGE = (217, 119, 6)
BROWN = (180, 83, 9)
HOOK_BG = (255, 244, 229)
ARG_BG = (232, 244, 251)
OK_BG = (234, 247, 238)


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REG, size)


def wrap(draw, text: str, fnt, max_w: int) -> list[str]:
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if draw.textlength(trial, font=fnt) <= max_w:
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


def arrow_down(draw, x, y1, y2, color=BLUE):
    draw.line((x, y1, x, y2 - 12), fill=color, width=4)
    draw.polygon([(x - 9, y2 - 14), (x + 9, y2 - 14), (x, y2)], fill=color)


def text_center(draw, box, lines, fnt, fill=WHITE, gap=4):
    x0, y0, x1, y1 = box
    line_h = fnt.size + gap
    total = len(lines) * line_h - gap
    y = y0 + (y1 - y0 - total) / 2
    for line in lines:
        tw = draw.textlength(line, font=fnt)
        draw.text((x0 + (x1 - x0 - tw) / 2, y), line, font=fnt, fill=fill)
        y += line_h


def build():
    W = 1400
    margin = 50
    content_w = W - 2 * margin
    cx = W // 2

    stages = [
        {
            "num": "1",
            "title": "ДОКИ · ЭДО",
            "color": BLUE,
            "hook": "Здравствуйте! Хочу рассказать вам про новый сервис для обмена по ЭДО — бесплатно.",
            "blocks": [
                (
                    "Аргумент 1",
                    "Вам нужен резервный канал связи с контрагентами: операторов ЭДО часто атакуют, "
                    "и вы можете остаться без обмена документами.",
                ),
                (
                    "Аргумент 2",
                    "Предлагаю просто попробовать и сравнить удобство с вашим текущим сервисом.",
                ),
            ],
            "next": "Далее — про логистику и ЭПД",
        },
        {
            "num": "2",
            "title": "ДОКИ · Логистика / ЭПД",
            "color": TEAL,
            "hook": "А вы используете транспортные накладные?",
            "blocks": [
                (
                    "Вопрос / заход",
                    "Вы слышали про закон с 1 сентября об обязательном использовании ЭПД "
                    "(электронных перевозочных документов)?",
                ),
                (
                    "Что сказать",
                    "Если да — кратко рассказываем про возможности «ДОКИ» для логистики и ЭПД.",
                ),
            ],
            "next": "Переход к кадровому контуру",
        },
        {
            "num": "3",
            "title": "КЭДО · 1С:Кабинет сотрудника",
            "color": ORANGE,
            "hook": "ЭПД ввели и реализовали очень быстро. Следующим шагом, скорее всего, будет кадровый документооборот.",
            "blocks": [
                (
                    "Вопрос",
                    "Как сейчас сотрудники подписывают расчётные листки? Никак? "
                    "Это требование законодательства.",
                ),
                (
                    "Презентация",
                    "Рассказываем про сервис «1С:Кабинет сотрудника»: "
                    "расчётные листки с ознакомлением, заявления, кадровые документы в контуре 1С.",
                ),
            ],
            "next": "Далее — командировки",
        },
        {
            "num": "4",
            "title": "Смартвей",
            "color": BROWN,
            "hook": "Ваши сотрудники часто ездят в командировки?",
            "blocks": [
                (
                    "Если да",
                    "Рассказываем про сервис «Смартвей»: организация командировок и связка с 1С.",
                ),
                (
                    "Если нет / редко",
                    "Фиксируем и не продавливаем. Переходим к завершению разговора.",
                ),
            ],
            "next": "Завершение звонка",
        },
    ]

    tmp = Image.new("RGB", (10, 10), BG)
    dtmp = ImageDraw.Draw(tmp)
    f_title = font(26, True)
    f_sub = font(15)
    f_hook = font(16, True)
    f_label = font(14, True)
    f_body = font(15)
    f_next = font(14, True)
    f_num = font(20, True)

    # measure heights
    heights = []
    for st in stages:
        hook_lines = wrap(dtmp, st["hook"], f_hook, content_w - 80)
        block_h = 0
        prepared = []
        for label, text in st["blocks"]:
            lines = wrap(dtmp, text, f_body, content_w // 2 - 60)
            prepared.append((label, lines))
            block_h = max(block_h, 36 + len(lines) * (f_body.size + 5) + 20)
        next_lines = wrap(dtmp, st["next"], f_next, content_w - 80)
        h = 70 + 16 + len(hook_lines) * 24 + 16 + block_h + 16 + 44
        heights.append(h)
        st["_hook_lines"] = hook_lines
        st["_prepared"] = prepared
        st["_block_h"] = block_h
        st["_next_lines"] = next_lines

    title_h = 110
    start_h = 64
    gap = 34
    end_h = 80
    H = margin + title_h + gap + start_h + gap + sum(heights) + gap * len(stages) + end_h + margin

    img = Image.new("RGB", (W, int(H)), BG)
    draw = ImageDraw.Draw(img)

    # Header
    rounded(draw, (margin, margin, W - margin, margin + title_h - 8), HEADER, 20)
    text_center(
        draw,
        (margin, margin + 18, W - margin, margin + 70),
        ["Схематический скрипт звонка", "ДОКИ ЭДО → ДОКИ логистика → КЭДО → Смартвей"],
        font(28, True),
        WHITE,
        6,
    )
    text_center(
        draw,
        (margin, margin + 72, W - margin, margin + title_h - 16),
        ["На каждом шаге: заход → аргументы / вопросы → переход дальше"],
        f_sub,
        (190, 210, 220),
        2,
    )

    y = margin + title_h + gap
    # Start
    start_box = (cx - 200, y, cx + 200, y + start_h)
    rounded(draw, start_box, BLUE, 32)
    text_center(draw, start_box, ["СТАРТ · разговор с клиентом"], font(18, True), WHITE)
    arrow_down(draw, cx, y + start_h + 2, y + start_h + gap, BLUE)
    y += start_h + gap

    for i, st in enumerate(stages):
        x0, x1 = margin, W - margin
        y0 = y
        y1 = y + heights[i]
        color = st["color"]

        rounded(draw, (x0, y0, x1, y1), WHITE, 20, outline=color, width=3)
        rounded(draw, (x0, y0, x1, y0 + 62), color, 20)
        draw.rectangle((x0, y0 + 36, x1, y0 + 62), fill=color)

        # number
        draw.ellipse((x0 + 18, y0 + 12, x0 + 52, y0 + 46), fill=WHITE)
        text_center(draw, (x0 + 18, y0 + 12, x0 + 52, y0 + 46), [st["num"]], f_num, color, 0)
        draw.text((x0 + 66, y0 + 18), st["title"], font=font(22, True), fill=WHITE)

        yy = y0 + 78
        # Hook strip
        hook_h = len(st["_hook_lines"]) * 24 + 20
        rounded(draw, (x0 + 18, yy, x1 - 18, yy + hook_h), HOOK_BG, 12)
        draw.text((x0 + 34, yy + 8), "ЗАХОД", font=f_label, fill=BROWN)
        ty = yy + 28
        for line in st["_hook_lines"]:
            draw.text((x0 + 34, ty), line, font=f_hook, fill=DARK)
            ty += 24

        yy = yy + hook_h + 14
        # Two argument cards
        col_w = (content_w - 46) // 2
        left = (x0 + 18, yy, x0 + 18 + col_w, yy + st["_block_h"])
        right = (x0 + 28 + col_w, yy, x1 - 18, yy + st["_block_h"])
        for box, (label, lines), bg in (
            (left, st["_prepared"][0], ARG_BG),
            (right, st["_prepared"][1], OK_BG if i != 3 else (253, 236, 236)),
        ):
            rounded(draw, box, bg, 12, outline=(200, 210, 220), width=1)
            draw.text((box[0] + 16, box[1] + 12), label.upper(), font=f_label, fill=color if i < 3 or label.startswith("Если да") else BROWN)
            ty = box[1] + 36
            for line in lines:
                draw.text((box[0] + 16, ty), line, font=f_body, fill=DARK)
                ty += f_body.size + 5

        # Next strip
        ny = yy + st["_block_h"] + 12
        rounded(draw, (x0 + 18, ny, x1 - 18, y1 - 14), (240, 244, 248), 10)
        text_center(
            draw,
            (x0 + 18, ny, x1 - 18, y1 - 14),
            st["_next_lines"],
            f_next,
            MUTED,
            2,
        )

        y = y1
        if i < len(stages) - 1:
            arrow_down(draw, cx, y + 2, y + gap, color)
            y += gap
        else:
            arrow_down(draw, cx, y + 2, y + gap, BROWN)
            y += gap

    end_box = (cx - 240, y, cx + 240, y + end_h)
    rounded(draw, end_box, BROWN, 22)
    text_center(
        draw,
        end_box,
        ["ФИНИШ", "Сверка контактов · следующий шаг"],
        font(18, True),
        WHITE,
        6,
    )

    img.save(OUT, "PNG", optimize=True)
    ARTIFACT.parent.mkdir(parents=True, exist_ok=True)
    img.save(ARTIFACT, "PNG", optimize=True)
    print(f"Saved {OUT} ({img.size[0]}x{img.size[1]})")
    print(f"Artifact {ARTIFACT}")


if __name__ == "__main__":
    build()
