#!/usr/bin/env python3
"""
Вертикальная блок-схема: холодный звонок «Блиц-запись» — ветка «Дикиди».
Н → уровни → развилка → К
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
import shutil

OUT = "Блиц-запись_блок-схема.pdf"
ASCII_OUT = "Blitz-zapis-flowchart.pdf"

NAVY = HexColor("#1B3A4B")
TEAL = HexColor("#2A9D8F")
ORANGE = HexColor("#E76F51")
BLUE = HexColor("#3D7EA6")
GREEN = HexColor("#2D6A4F")
LIGHT = HexColor("#F0F4F5")
YELLOW = HexColor("#FFF6E8")
SOFT_ORANGE = HexColor("#FCE8E2")
BORDER = HexColor("#B8C2CA")
DARK = HexColor("#243038")
GRAY = HexColor("#6B7680")

FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"


def setup_fonts():
    pdfmetrics.registerFont(TTFont("Body", FONT_REG))
    pdfmetrics.registerFont(TTFont("BodyBold", FONT_BOLD))
    return "Body", "BodyBold"


def wrap(text, font, size, max_w):
    return simpleSplit(text, font, size, max_w)


def draw_circle(c, cx, cy, r, letter, fonts, fill=TEAL):
    _, bold = fonts
    c.setFillColor(fill)
    c.circle(cx, cy, r, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 13)
    c.drawCentredString(cx, cy - 4.5, letter)


def draw_arrow_v(c, x, y_from, y_to):
    c.setStrokeColor(TEAL)
    c.setFillColor(TEAL)
    c.setLineWidth(1.8)
    tip = y_to + 1
    c.line(x, y_from, x, tip + 8)
    path = c.beginPath()
    path.moveTo(x, tip)
    path.lineTo(x - 4.5, tip + 8)
    path.lineTo(x + 4.5, tip + 8)
    path.close()
    c.drawPath(path, fill=1, stroke=0)


def draw_diamond(c, cx, cy, w, h, text, fonts):
    """Ромб-развилка. cy — центр. Возвращает y_bottom."""
    body, bold = fonts
    path = c.beginPath()
    path.moveTo(cx, cy + h / 2)
    path.lineTo(cx + w / 2, cy)
    path.lineTo(cx, cy - h / 2)
    path.lineTo(cx - w / 2, cy)
    path.close()
    c.setFillColor(YELLOW)
    c.setStrokeColor(ORANGE)
    c.setLineWidth(1.4)
    c.drawPath(path, fill=1, stroke=1)
    c.setFillColor(NAVY)
    c.setFont(bold, 8.5)
    lines = wrap(text, bold, 8.5, w * 0.62)
    ty = cy + (len(lines) * 10) / 2 - 2
    for line in lines:
        c.drawCentredString(cx, ty - 8.5, line)
        ty -= 10.5
    return cy - h / 2


def draw_level(c, x, y_top, w, title, fonts, accent=TEAL):
    _, bold = fonts
    pad = 10
    size = 11.5
    lines = wrap(title, bold, size, w - 2 * pad)
    h = 14 + len(lines) * (size + 3)
    y_bottom = y_top - h
    c.setFillColor(accent)
    c.roundRect(x, y_bottom, w, h, 6, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, size)
    ty = y_top - 10
    for line in lines:
        c.drawCentredString(x + w / 2, ty - size, line)
        ty -= size + 3
    return y_bottom


def draw_questions(c, x, y_top, w, items, fonts, caption="Что говорим / спрашиваем", fill=LIGHT):
    body, bold = fonts
    pad = 8
    cap_size = 8
    body_size = 8.2
    max_w = w - 2 * pad - 8
    prepared = []
    content_h = 8 + cap_size + 4
    for item in items:
        lines = wrap("• " + item, body, body_size, max_w)
        prepared.append(lines)
        content_h += len(lines) * (body_size + 2.3) + 3
    h = content_h + pad
    y_bottom = y_top - h
    c.setFillColor(fill)
    c.setStrokeColor(BORDER)
    c.setLineWidth(0.8)
    c.roundRect(x, y_bottom, w, h, 5, fill=1, stroke=1)
    c.setFillColor(TEAL)
    c.setFont(bold, cap_size)
    c.drawString(x + pad, y_top - cap_size - 4, caption)
    y = y_top - cap_size - 10
    c.setFillColor(DARK)
    c.setFont(body, body_size)
    for lines in prepared:
        for line in lines:
            c.drawString(x + pad, y - body_size, line)
            y -= body_size + 2.3
        y -= 2
    return y_bottom


def page_head(c, page_w, page_h, fonts, n, total, cont=False):
    body, bold = fonts
    c.setFillColor(NAVY)
    c.rect(0, page_h - 16 * mm, page_w, 16 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 12)
    c.drawCentredString(page_w / 2, page_h - 7.5 * mm, "Блиц-запись — блок-схема (ветка «Дикиди»)")
    c.setFont(body, 8)
    c.drawCentredString(
        page_w / 2,
        page_h - 12.5 * mm,
        "бесплатный сервис · уточнение проблем · хранение в Иркутске · демонстрация",
    )
    c.setFillColor(GRAY)
    c.drawCentredString(page_w / 2, 7 * mm, f"{n} / {total}")
    if cont:
        c.setFillColor(TEAL)
        c.setFont(bold, 8)
        c.drawCentredString(page_w / 2, page_h - 20 * mm, "↓ продолжение схемы")


def build():
    fonts = setup_fonts()
    body, bold = fonts
    page_w, page_h = A4
    c = canvas.Canvas(OUT, pagesize=A4)
    margin = 16 * mm
    box_w = page_w - 2 * margin
    cx = page_w / 2
    total = 2

    # ========== PAGE 1 ==========
    page_head(c, page_w, page_h, fonts, 1, total)
    y = page_h - 24 * mm

    c.setFillColor(NAVY)
    c.setFont(bold, 11)
    c.drawCentredString(cx, y, "Холодный звонок: клиент на «Дикиди»")
    y -= 9 * mm

    draw_circle(c, cx, y - 7, 8, "Н", fonts, TEAL)
    y -= 15
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    # Level 1
    y = draw_level(
        c,
        margin,
        y,
        box_w,
        "1. Первичный контакт / приветствие и презентация компании",
        fonts,
        TEAL,
    )
    draw_arrow_v(c, cx, y, y - 3.5 * mm)
    y -= 3.5 * mm
    y = draw_questions(
        c,
        margin + 4 * mm,
        y,
        box_w - 8 * mm,
        [
            "Здравствуйте! Меня зовут [имя], группа компаний «Форус». Мы автоматизируем бизнес и онлайн-запись.",
            "Звоню по вопросу сотрудничества: предлагаем бесплатный сервис для онлайн-записи клиентов.",
            "Подскажите, чем вы сейчас пользуетесь?",
        ],
        fonts,
    )
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    # Level 2 - answer Dikidi
    y = draw_level(
        c,
        margin,
        y,
        box_w,
        "2. Ответ клиента: «Дикиди» — уточняющий вопрос",
        fonts,
        BLUE,
    )
    draw_arrow_v(c, cx, y, y - 3.5 * mm)
    y -= 3.5 * mm
    y = draw_questions(
        c,
        margin + 4 * mm,
        y,
        box_w - 8 * mm,
        [
            "Понял(а) вас! Многие клиенты сейчас замечают проблемы с сервисом.",
            "Подскажите, вы ещё не сталкивались?",
        ],
        fonts,
        "Что говорим",
    )
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    # Diamond
    y_before_diamond = y
    diamond_h = 28 * mm
    diamond_w = 72 * mm
    cy_d = y - diamond_h / 2
    draw_diamond(c, cx, cy_d, diamond_w, diamond_h, "Сталкивались с проблемами в «Дикиди»?", fonts)
    y = cy_d - diamond_h / 2

    # Two branches side by side below diamond
    gap = 6 * mm
    col_w = (box_w - gap) / 2
    left_x = margin
    right_x = margin + col_w + gap
    branch_top = y - 8 * mm

    # arrows to branches
    c.setStrokeColor(ORANGE)
    c.setFillColor(ORANGE)
    c.setLineWidth(1.5)
    # left fork
    c.line(cx - 8, y, left_x + col_w / 2, branch_top + 2)
    c.setFont(bold, 8)
    c.drawString(left_x + 4, branch_top + 4, "ДА →")
    # right fork
    c.setStrokeColor(GREEN)
    c.setFillColor(GREEN)
    c.line(cx + 8, y, right_x + col_w / 2, branch_top + 2)
    c.drawRightString(right_x + col_w - 4, branch_top + 4, "← НЕТ")

    # Left branch - problems
    y_left = draw_level(
        c, left_x, branch_top, col_w, "2А. Проблемы есть", fonts, ORANGE
    )
    y_left = draw_questions(
        c,
        left_x + 2 * mm,
        y_left - 3 * mm,
        col_w - 4 * mm,
        [
            "Очень жаль. Из-за текущих обстоятельств малый бизнес страдает.",
            "Мы можем предложить помощь и решить вашу проблему с записью.",
            "→ переходим к рассказу про бесплатный сервис",
        ],
        fonts,
        "Что говорим",
        SOFT_ORANGE,
    )

    # Right branch - no problems
    y_right = draw_level(
        c, right_x, branch_top, col_w, "2Б. Проблем нет", fonts, GREEN
    )
    y_right = draw_questions(
        c,
        right_x + 2 * mm,
        y_right - 3 * mm,
        col_w - 4 * mm,
        [
            "Приятно слышать, что всё хорошо!",
            "Буду благодарен(на) за обратную связь.",
            "Какими блоками пользуетесь в «Дикиди»?",
            "Что нравится больше всего?",
            "→ затем мягко к бесплатному сервису",
        ],
        fonts,
        "Что говорим / спрашиваем",
        HexColor("#E8F5E9"),
    )

    merge_y = min(y_left, y_right) - 4 * mm
    c.setStrokeColor(TEAL)
    c.setFillColor(TEAL)
    c.setLineWidth(1.5)
    c.line(left_x + col_w / 2, y_left, left_x + col_w / 2, merge_y + 6)
    c.line(right_x + col_w / 2, y_right, right_x + col_w / 2, merge_y + 6)
    c.line(left_x + col_w / 2, merge_y + 6, right_x + col_w / 2, merge_y + 6)
    draw_arrow_v(c, cx, merge_y + 6, merge_y - 2 * mm)

    c.setFillColor(TEAL)
    c.setFont(bold, 8)
    c.drawCentredString(cx, 13 * mm, "↓ обе ветки → рассказ про сервис (след. страница)")
    c.showPage()

    # ========== PAGE 2 ==========
    page_head(c, page_w, page_h, fonts, 2, total, cont=True)
    y = page_h - 26 * mm
    draw_circle(c, cx, y - 5, 5, "↓", fonts, GRAY)
    y -= 10
    draw_arrow_v(c, cx, y, y - 5 * mm)
    y -= 5 * mm

    y = draw_level(
        c,
        margin,
        y,
        box_w,
        "3. Рассказ про бесплатный сервис «Блиц-запись»",
        fonts,
        TEAL,
    )
    draw_arrow_v(c, cx, y, y - 3.5 * mm)
    y -= 3.5 * mm
    y = draw_questions(
        c,
        margin + 4 * mm,
        y,
        box_w - 8 * mm,
        [
            "Самостоятельная онлайн-запись: клиент выбирает услугу и время сам.",
            "Тариф «Старт» — 0 ₽ в месяц: календарь, онлайн-запись, прайс, до пяти сотрудников.",
            "Данные хранятся в Иркутске — рядом поддержка группы компаний «Форус».",
            "Подключается к 2ГИС, сайту и социальным сетям.",
            "Уточнить: сколько сотрудников в записи? Где сейчас стоит онлайн-запись?",
        ],
        fonts,
    )
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    y = draw_level(
        c, margin, y, box_w, "4. Приглашение на демонстрацию", fonts, ORANGE
    )
    draw_arrow_v(c, cx, y, y - 3.5 * mm)
    y -= 3.5 * mm
    y = draw_questions(
        c,
        margin + 4 * mm,
        y,
        box_w - 8 * mm,
        [
            "Давайте покажу, как это выглядит рядом с тем, к чему вы привыкли в «Дикиди».",
            "15–20 минут: онлайн-запись и бесплатный старт.",
            "Когда удобнее: завтра до обеда или после? / вторник или среда?",
            "Если «не хотим менять»: сравните без обязательств, особенно бесплатный «Старт».",
            "Фиксация: почта или мессенджер, дата, время; напомнить за день и за час.",
        ],
        fonts,
    )
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    y = draw_level(
        c, margin, y, box_w, "5. Если демо не ставят / возражения", fonts, NAVY
    )
    draw_arrow_v(c, cx, y, y - 3.5 * mm)
    y -= 3.5 * mm
    y = draw_questions(
        c,
        margin + 4 * mm,
        y,
        box_w - 8 * mm,
        [
            "Скинуть описание с бесплатным «Стартом» + конкретная дата перезвона.",
            "«Всё устраивает» → запасной вариант на случай сбоев, можно параллельно и бесплатно.",
            "«Страшно переезжать» → помогаем с настройкой, данные в Иркутске, демо без обязательства уходить.",
            "«Уже оплатили Дикиди» → подключить «Старт» параллельно до конца оплаченного периода.",
        ],
        fonts,
        "Что говорим",
    )
    draw_arrow_v(c, cx, y, y - 9 * mm)
    y -= 9 * mm

    draw_circle(c, cx, y - 8, 8, "К", fonts, NAVY)
    y -= 20
    c.setFillColor(GRAY)
    c.setFont(body, 8)
    c.drawCentredString(cx, y, "конец звонка")

    y -= 12 * mm
    c.setFillColor(NAVY)
    c.setFont(bold, 8)
    c.drawCentredString(cx, y, "Цепочка:")
    y -= 5 * mm
    c.setFillColor(DARK)
    c.setFont(body, 7.2)
    chain = "Н → 1 Заход → 2 «Дикиди?» → развилка (проблемы / всё ок) → 3 Бесплатный сервис + Иркутск → 4 Демонстрация → К"
    for line in wrap(chain, body, 7.2, box_w - 6 * mm):
        c.drawCentredString(cx, y, line)
        y -= 10

    c.showPage()
    c.save()
    shutil.copy(OUT, ASCII_OUT)
    print("Saved", OUT)
    print("Saved", ASCII_OUT)


if __name__ == "__main__":
    build()
