#!/usr/bin/env python3
"""
Вертикальная блок-схема «Скрипт клиентов без ИТС»
Формат: Н → уровень → стрелка → уровень → … → К
Под каждым уровнем — вопросы / тезисы скрипта.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
import os
import shutil

OUT = "Скрипт_клиентов_без_ИТС_блок-схема.pdf"
ASCII_OUT = "Script-clients-without-ITS-flowchart.pdf"

NAVY = HexColor("#1B3A4B")
TEAL = HexColor("#2A9D8F")
ORANGE = HexColor("#E76F51")
BLUE = HexColor("#3D7EA6")
LIGHT = HexColor("#F0F4F5")
SOFT = HexColor("#E8F5F3")
YELLOW = HexColor("#FFF6E8")
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


def draw_circle_node(c, cx, cy, r, letter, fonts, fill=TEAL):
    body, bold = fonts
    c.setFillColor(fill)
    c.setStrokeColor(fill)
    c.circle(cx, cy, r, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 14)
    c.drawCentredString(cx, cy - 5, letter)


def draw_arrow_v(c, x, y_from, y_to):
    """Стрелка вниз от y_from к y_to."""
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


def draw_level_box(c, x, y_top, w, title, fonts, accent=TEAL):
    """Прямоугольник уровня (как в рукописной схеме). Возвращает y_bottom, cx."""
    body, bold = fonts
    pad_x = 10
    title_size = 12
    lines = wrap(title, bold, title_size, w - 2 * pad_x)
    h = 14 + len(lines) * (title_size + 3)
    y_bottom = y_top - h

    c.setFillColor(accent)
    c.setStrokeColor(accent)
    c.setLineWidth(1.5)
    c.roundRect(x, y_bottom, w, h, 6, fill=1, stroke=0)

    c.setFillColor(white)
    c.setFont(bold, title_size)
    ty = y_top - 10
    for line in lines:
        c.drawCentredString(x + w / 2, ty - title_size, line)
        ty -= title_size + 3
    return y_bottom, x + w / 2


def draw_questions_block(c, x, y_top, w, items, fonts, title="Вопросы / тезисы"):
    """
    Блок под уровнем: список вопросов/тезисов.
    items: list of strings
    Returns y_bottom
    """
    body, bold = fonts
    pad = 8
    title_size = 8
    body_size = 8
    max_w = w - 2 * pad - 10

    prepared = []
    content_h = 6 + title_size + 4
    for item in items:
        lines = wrap("• " + item, body, body_size, max_w)
        prepared.append(lines)
        content_h += len(lines) * (body_size + 2.4) + 3

    h = content_h + pad
    y_bottom = y_top - h

    c.setFillColor(LIGHT)
    c.setStrokeColor(BORDER)
    c.setLineWidth(0.8)
    c.roundRect(x, y_bottom, w, h, 5, fill=1, stroke=1)

    c.setFillColor(TEAL)
    c.setFont(bold, title_size)
    c.drawString(x + pad, y_top - title_size - 4, title)

    y = y_top - title_size - 10
    c.setFillColor(DARK)
    c.setFont(body, body_size)
    for lines in prepared:
        for line in lines:
            c.drawString(x + pad, y - body_size, line)
            y -= body_size + 2.4
        y -= 2

    return y_bottom


def draw_branch_block(c, x, y_top, w, rows, fonts, title="Если клиент говорит → что говорим → куда дальше"):
    """
    Таблица-навигатор под уровнем.
    rows: list of (client, say, next_step)
    """
    body, bold = fonts
    pad = 6
    title_size = 8
    cell_size = 7.2
    col_w = [(w - 2 * pad) * 0.28, (w - 2 * pad) * 0.42, (w - 2 * pad) * 0.30]

    # measure
    prepared = []
    content_h = 8 + title_size + 6
    # header row
    headers = ["Клиент говорит", "Что говорим", "Дальше"]
    header_lines = [wrap(h, bold, cell_size, col_w[i] - 4) for i, h in enumerate(headers)]
    header_h = max(len(hl) for hl in header_lines) * (cell_size + 2) + 4
    content_h += header_h + 2

    for client, say, nxt in rows:
        cells = [
            wrap(client, body, cell_size, col_w[0] - 4),
            wrap(say, body, cell_size, col_w[1] - 4),
            wrap(nxt, body, cell_size, col_w[2] - 4),
        ]
        rh = max(len(x) for x in cells) * (cell_size + 2) + 4
        prepared.append((cells, rh))
        content_h += rh

    h = content_h + pad
    y_bottom = y_top - h

    c.setFillColor(YELLOW)
    c.setStrokeColor(BORDER)
    c.setLineWidth(0.8)
    c.roundRect(x, y_bottom, w, h, 5, fill=1, stroke=1)

    c.setFillColor(ORANGE)
    c.setFont(bold, title_size)
    c.drawString(x + pad, y_top - title_size - 3, title)

    y = y_top - title_size - 8

    # header
    c.setFillColor(HexColor("#F3D9C0"))
    c.rect(x + pad, y - header_h, w - 2 * pad, header_h, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont(bold, cell_size)
    cx = x + pad
    for i, hl in enumerate(header_lines):
        yy = y - 3
        for line in hl:
            c.drawString(cx + 2, yy - cell_size, line)
            yy -= cell_size + 2
        cx += col_w[i]
    y -= header_h

    # rows
    for idx, (cells, rh) in enumerate(prepared):
        if idx % 2 == 0:
            c.setFillColor(white)
        else:
            c.setFillColor(HexColor("#FFFBF3"))
        c.rect(x + pad, y - rh, w - 2 * pad, rh, fill=1, stroke=0)
        # separators
        c.setStrokeColor(BORDER)
        c.setLineWidth(0.4)
        c.line(x + pad, y - rh, x + w - pad, y - rh)

        c.setFillColor(DARK)
        c.setFont(body, cell_size)
        cx = x + pad
        for i, cell_lines in enumerate(cells):
            yy = y - 3
            for line in cell_lines:
                c.drawString(cx + 2, yy - cell_size, line)
                yy -= cell_size + 2
            cx += col_w[i]
        y -= rh

    return y_bottom


def new_page(c, page_w, page_h, fonts, page_num, total, continue_flow=False):
    body, bold = fonts
    c.setFillColor(NAVY)
    c.rect(0, page_h - 16 * mm, page_w, 16 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 13)
    c.drawCentredString(page_w / 2, page_h - 10 * mm, "Скрипт клиентов без ИТС")
    c.setFont(body, 8)
    c.drawCentredString(page_w / 2, page_h - 14 * mm, "блок-схема звонка")

    c.setFillColor(GRAY)
    c.setFont(body, 8)
    c.drawCentredString(page_w / 2, 8 * mm, f"{page_num} / {total}")

    if continue_flow:
        # small note that flow continues
        c.setFillColor(TEAL)
        c.setFont(bold, 8)
        c.drawCentredString(page_w / 2, page_h - 20 * mm, "↓ продолжение схемы")


def build():
    fonts = setup_fonts()
    body, bold = fonts
    page_w, page_h = A4  # portrait — вертикальная схема
    c = canvas.Canvas(OUT, pagesize=A4)

    margin_x = 18 * mm
    box_w = page_w - 2 * margin_x
    cx = page_w / 2
    total_pages = 3

    # ===================== PAGE 1 =====================
    new_page(c, page_w, page_h, fonts, 1, total_pages)
    y = page_h - 24 * mm

    # Title label like "Вечер" in the sketch
    c.setFillColor(NAVY)
    c.setFont(bold, 11)
    c.drawCentredString(cx, y, "Звонок клиенту без ИТС")
    y -= 10 * mm

    # Н — начало
    draw_circle_node(c, cx, y - 7, 8, "Н", fonts, TEAL)
    y -= 7 + 8
    draw_arrow_v(c, cx, y, y - 8 * mm)
    y -= 8 * mm

    # --- Уровень 1: Открытие ---
    y, _ = draw_level_box(c, margin_x, y, box_w, "1. Открытие", fonts, TEAL)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions_block(
        c,
        margin_x + 6 * mm,
        y,
        box_w - 12 * mm,
        [
            "Добрый день! [Имя], компания [Название]. Удобно сейчас на пару минут?",
        ],
        fonts,
        "Что говорим",
    )
    draw_arrow_v(c, cx, y, y - 7 * mm)
    y -= 7 * mm

    # --- Уровень 2: Инфоповод ---
    y, _ = draw_level_box(c, margin_x, y, box_w, "2. Инфоповод", fonts, TEAL)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions_block(
        c,
        margin_x + 6 * mm,
        y,
        box_w - 12 * mm,
        [
            "Вы подключали ЭДО, из других продуктов ничего не брали. Хочу уточнить — как у вас учёт организован, есть 1С?",
        ],
        fonts,
        "Что говорим / спрашиваем",
    )
    draw_arrow_v(c, cx, y, y - 7 * mm)
    y -= 7 * mm

    # --- Уровень 3 header, table starts ---
    y, _ = draw_level_box(
        c, margin_x, y, box_w, "3. Навигатор по ответам клиента", fonts, BLUE
    )
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm

    # First half of navigator rows (fit on page 1)
    rows_p1 = [
        (
            "«Да, есть 1С»",
            "«В каком формате: локально или облако? Сопровождение есть?»",
            "→ уровень 4 «Углубление»",
        ),
        (
            "«1С: Бухгалтерия / УТ / другая»",
            "«Поняла, спасибо. А обновления кто делает?»",
            "→ 2–3 вопроса уровня 4 → уровень 5",
        ),
        (
            "«1С в облаке»",
            "«Когда возникают вопросы — решаете со своими программистами или куда-то обращаетесь?» Слушаем. «Мы официальный партнёр 1С более 30 лет, работаем по всей России. Предлагаю рассмотреть условия сопровождения.»",
            "→ уровень 5 «Предложение»",
        ),
        (
            "«Нет 1С / Excel / другая программа»",
            "«Поняла. А бухгалтерию сами ведёте или на аутсорсе?»",
            "Уточни потребность; если аутсорс — кто отвечает за учёт",
        ),
    ]
    y = draw_branch_block(c, margin_x + 4 * mm, y, box_w - 8 * mm, rows_p1, fonts)

    c.setFillColor(TEAL)
    c.setFont(bold, 8)
    c.drawCentredString(cx, 14 * mm, "↓ продолжение на следующей странице")
    c.showPage()

    # ===================== PAGE 2 =====================
    new_page(c, page_w, page_h, fonts, 2, total_pages, continue_flow=True)
    y = page_h - 26 * mm

    # continuation arrow from previous
    draw_circle_node(c, cx, y - 5, 5, "↓", fonts, GRAY)
    y -= 5 + 5
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    c.setFillColor(BLUE)
    c.setFont(bold, 9)
    c.drawCentredString(cx, y, "уровень 3 — продолжение")
    y -= 5 * mm

    rows_p2 = [
        (
            "«Не знаю / уточню»",
            "«А кто у вас за это отвечает? Может, с ним поговорить?»",
            "Запиши контакт ЛПР → завершай",
        ),
        (
            "«Всё устраивает, ничего не надо»",
            "«Рада слышать! А сопровождение есть? Часто при сбое теряют время и деньги.»",
            "Нет сопровождения → уровень 5; есть → оставь контакты",
        ),
        (
            "«Дорого»",
            "«ИТС от Х руб — дешевле одного вызова специалиста»",
            "Отправь коммерческое предложение",
        ),
        (
            "«Пришлите на почту»",
            "«Конечно отправлю. А вы сами принимаете решение или с кем-то согласуете? Как вам в целом моё предложение?»",
            "Почта + роль → следующий контакт",
        ),
        (
            "«Подумаю»",
            "«Хорошо. Пришлю материал. Как вам в целом моё предложение?»",
            "Отработай возражение → предложи встречу → дата",
        ),
        (
            "«Нет» (отказ)",
            "«Спасибо за честность. Если что, мы на связи. Давайте предложение направлю, ознакомитесь, и позже свяжусь?»",
            "Заверши звонок, сохрани контакт, отправь КП",
        ),
    ]
    y = draw_branch_block(c, margin_x + 4 * mm, y, box_w - 8 * mm, rows_p2, fonts)
    draw_arrow_v(c, cx, y, y - 8 * mm)
    y -= 8 * mm

    # --- Уровень 4 ---
    y, _ = draw_level_box(
        c, margin_x, y, box_w, "4. Углубление (если работает в 1С)", fonts, TEAL
    )
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions_block(
        c,
        margin_x + 6 * mm,
        y,
        box_w - 12 * mm,
        [
            "Обновления кто делает — свой специалист или ищете со стороны?",
            "Бывает, что программа работает не так, как надо, или вопросы возникают?",
            "Пользуетесь сопровождением? Как решаете вопросы?",
        ],
        fonts,
        "Вопросы для выявления боли",
    )
    draw_arrow_v(c, cx, y, y - 8 * mm)
    y -= 8 * mm

    # --- Уровень 5 ---
    y, _ = draw_level_box(c, margin_x, y, box_w, "5. Предложение", fonts, ORANGE)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions_block(
        c,
        margin_x + 6 * mm,
        y,
        box_w - 12 * mm,
        [
            "Нет сопровождения / ИТС → предложить ИТС: официальная подписка 1С (обновления, консультации, сервисы).",
            "Локальная 1С + интерес к облаку → предложить 1С:Фреш: облако из браузера, без установок.",
            "Уже всё есть → оставить контакты: «если что-то понадобится».",
        ],
        fonts,
        "Что предложить по ситуации",
    )

    c.setFillColor(TEAL)
    c.setFont(bold, 8)
    c.drawCentredString(cx, 14 * mm, "↓ возражения и завершение — на следующей странице")
    c.showPage()

    # ===================== PAGE 3 =====================
    new_page(c, page_w, page_h, fonts, 3, total_pages, continue_flow=True)
    y = page_h - 26 * mm

    draw_circle_node(c, cx, y - 5, 5, "↓", fonts, GRAY)
    y -= 5 + 5
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    # --- Уровень 6 ---
    y, _ = draw_level_box(c, margin_x, y, box_w, "6. Отработка возражений", fonts, ORANGE)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm

    obj_rows = [
        (
            "«У нас уже всё есть»",
            "«А сопровождение есть? Часто программа есть, а поддержки нет — и при сбое теряют время.»",
            "Уточни → если нет ИТС, предложи",
        ),
        (
            "«Не нужно»",
            "«А что сейчас закрывает эту задачу?»",
            "Выясни альтернативу → покажи отличие",
        ),
        (
            "«Дорого»",
            "«ИТС от Х руб — дешевле одного вызова специалиста + 1С:Напарник + сервисы.»",
            "Предложи встречу по демонстрации",
        ),
        (
            "«Пришлите на почту»",
            "«Конечно отправлю. А вы сами принимаете решение или с кем-то согласуете? Как вам в целом моё предложение?»",
            "Определи ЛПР → следующий контакт",
        ),
        (
            "«Подумаю»",
            "«Хорошо. Пришлю материал. Как вам в целом моё предложение?»",
            "Отработай → предложи встречу → дата",
        ),
    ]
    y = draw_branch_block(
        c,
        margin_x + 4 * mm,
        y,
        box_w - 8 * mm,
        obj_rows,
        fonts,
        title="Возражение → ответ → что делать дальше",
    )
    draw_arrow_v(c, cx, y, y - 8 * mm)
    y -= 8 * mm

    # --- Уровень 7: фиксация результата ---
    y, _ = draw_level_box(
        c, margin_x, y, box_w, "7. Фиксация результата звонка", fonts, NAVY
    )
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions_block(
        c,
        margin_x + 6 * mm,
        y,
        box_w - 12 * mm,
        [
            "Договорились о следующем шаге: встреча / перезвон / коммерческое предложение?",
            "Зафиксированы почта, роль ЛПР, дата следующего контакта?",
            "Комментарий сохранён в системе учёта?",
        ],
        fonts,
        "Что проверить перед завершением",
    )
    draw_arrow_v(c, cx, y, y - 10 * mm)
    y -= 10 * mm

    # К — конец
    draw_circle_node(c, cx, y - 8, 8, "К", fonts, NAVY)
    y -= 8 + 12
    c.setFillColor(GRAY)
    c.setFont(body, 8)
    c.drawCentredString(cx, y, "конец звонка")

    # Legend of levels at very bottom
    y -= 12 * mm
    c.setFillColor(NAVY)
    c.setFont(bold, 8)
    c.drawCentredString(cx, y, "Цепочка уровней:")
    y -= 5 * mm
    c.setFont(body, 7.5)
    c.setFillColor(DARK)
    chain = "Н → 1 Открытие → 2 Инфоповод → 3 Навигатор → 4 Углубление → 5 Предложение → 6 Возражения → 7 Фиксация → К"
    for line in wrap(chain, body, 7.5, box_w - 10 * mm):
        c.drawCentredString(cx, y, line)
        y -= 10

    c.showPage()
    c.save()
    shutil.copy(OUT, ASCII_OUT)
    print("Saved", OUT)
    print("Saved", ASCII_OUT)


if __name__ == "__main__":
    build()
