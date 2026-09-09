#!/usr/bin/env python3
"""PDF-блок-схема: скрипт клиентов без ИТС."""

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
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
BORDER = HexColor("#C5CDD4")
DARK = HexColor("#243038")
ARROW = HexColor("#2A9D8F")
GRAY = HexColor("#7A8792")

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
]
FONT_BOLD_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
]


def find_font(paths):
    for p in paths:
        if os.path.exists(p):
            return p
    raise SystemExit("No Cyrillic TTF font found")


def setup_fonts():
    pdfmetrics.registerFont(TTFont("Body", find_font(FONT_CANDIDATES)))
    pdfmetrics.registerFont(TTFont("BodyBold", find_font(FONT_BOLD_CANDIDATES)))
    return "Body", "BodyBold"


def wrap(text, font, size, max_w):
    return simpleSplit(text, font, size, max_w)


def draw_round_rect(c, x, y, w, h, fill, stroke=BORDER, radius=8):
    c.setFillColor(fill)
    c.setStrokeColor(stroke)
    c.setLineWidth(1.2)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=1)


def draw_arrow_right(c, x_from, x_to, y, label=None, fonts=None):
    c.setStrokeColor(ARROW)
    c.setFillColor(ARROW)
    c.setLineWidth(2.2)
    tip = x_to - 2
    c.line(x_from, y, tip - 9, y)
    path = c.beginPath()
    path.moveTo(tip, y)
    path.lineTo(tip - 10, y - 5)
    path.lineTo(tip - 10, y + 5)
    path.close()
    c.drawPath(path, fill=1, stroke=0)
    if label and fonts:
        _, bold = fonts
        c.setFillColor(TEAL)
        c.setFont(bold, 7.5)
        c.drawCentredString((x_from + tip) / 2, y + 7, label)


def draw_arrow_down(c, x, y_from, y_to, label=None, fonts=None):
    c.setStrokeColor(ARROW)
    c.setFillColor(ARROW)
    c.setLineWidth(2.2)
    tip = y_to + 2
    c.line(x, y_from, x, tip + 9)
    path = c.beginPath()
    path.moveTo(x, tip)
    path.lineTo(x - 5, tip + 10)
    path.lineTo(x + 5, tip + 10)
    path.close()
    c.drawPath(path, fill=1, stroke=0)
    if label and fonts:
        _, bold = fonts
        c.setFillColor(TEAL)
        c.setFont(bold, 7.5)
        c.drawCentredString(x + 36, (y_from + tip) / 2 - 2, label)


def stage_card(c, x, y_top, w, stage_num, title, script_lines, fonts, accent=TEAL):
    body, bold = fonts
    pad = 9
    title_size = 10.5
    body_size = 8.2
    max_w = w - 2 * pad

    if isinstance(stage_num, int):
        title_full = f"Этап {stage_num}. {title}"
        badge_text = str(stage_num)
    else:
        title_full = title
        badge_text = str(stage_num)[:2]

    title_wrapped = wrap(title_full, bold, title_size, max_w - 28)

    content_blocks = []
    content_h = 0
    for block in script_lines:
        lines = []
        for para in block.split("\n"):
            lines.extend(wrap(para, body, body_size, max_w) or [""])
        content_blocks.append(lines)
        content_h += len(lines) * (body_size + 2.6) + 4.5

    header_h = 11 + len(title_wrapped) * (title_size + 2.4)
    h = header_h + content_h + pad
    y_bottom = y_top - h

    draw_round_rect(c, x, y_bottom, w, h, white, BORDER, 8)

    c.setFillColor(accent)
    c.roundRect(x, y_top - header_h, w, header_h, 8, fill=1, stroke=0)
    c.rect(x, y_top - header_h, w, 10, fill=1, stroke=0)

    badge_r = 8.5
    bx = x + pad + badge_r
    by = y_top - header_h / 2
    c.setFillColor(white)
    c.circle(bx, by, badge_r, fill=1, stroke=0)
    c.setFillColor(accent)
    c.setFont(bold, 9.5)
    c.drawCentredString(bx, by - 3.2, badge_text)

    c.setFillColor(white)
    c.setFont(bold, title_size)
    ty = y_top - 9
    for i, line in enumerate(title_wrapped):
        indent = 26 if i == 0 else 10
        c.drawString(x + pad + indent, ty - title_size, line)
        ty -= title_size + 2.4

    y = y_top - header_h - 6
    c.setFillColor(DARK)
    c.setFont(body, body_size)
    for lines in content_blocks:
        for line in lines:
            c.drawString(x + pad, y - body_size, line)
            y -= body_size + 2.6
        y -= 3

    return y_bottom, x + w / 2


def page_header(c, page_w, page_h, fonts):
    _, bold = fonts
    c.setFillColor(NAVY)
    c.rect(0, page_h - 14 * mm, page_w, 14 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 12)
    c.drawString(12 * mm, page_h - 9 * mm, "Скрипт клиентов без ИТС — блок-схема звонка")


def page_footer(c, page_w, n, total, fonts):
    body, _ = fonts
    c.setFillColor(GRAY)
    c.setFont(body, 8)
    c.drawCentredString(page_w / 2, 6 * mm, f"{n} / {total}")


def flow_strip(c, page_w, fonts, active=None):
    _, bold = fonts
    labels = [
        (1, "Открытие"),
        (2, "Инфоповод"),
        (3, "По ответам"),
        (4, "Углубление"),
        (5, "Предложение"),
        (6, "Возражения"),
    ]
    margin = 14 * mm
    usable = page_w - 2 * margin
    y = 13 * mm
    step = usable / (len(labels) - 1)
    for i, (num, lab) in enumerate(labels):
        x = margin + i * step
        on = active is None or num in active
        c.setFillColor(TEAL if on else GRAY)
        c.circle(x, y + 9, 5.5, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont(bold, 7.5)
        c.drawCentredString(x, y + 6.2, str(num))
        c.setFillColor(NAVY if on else GRAY)
        c.setFont(bold, 6.5)
        c.drawCentredString(x, y - 2, lab)
        if i < len(labels) - 1:
            c.setStrokeColor(ARROW if on else GRAY)
            c.setFillColor(ARROW if on else GRAY)
            c.setLineWidth(1.5)
            c.line(x + 7, y + 9, x + step - 7, y + 9)
            path = c.beginPath()
            path.moveTo(x + step - 7, y + 9)
            path.lineTo(x + step - 12, y + 9 - 3)
            path.lineTo(x + step - 12, y + 9 + 3)
            path.close()
            c.drawPath(path, fill=1, stroke=0)


def build():
    fonts = setup_fonts()
    page_w, page_h = landscape(A4)
    c = canvas.Canvas(OUT, pagesize=landscape(A4))

    margin = 11 * mm
    top = page_h - 20 * mm
    usable = page_w - 2 * margin
    gap = 16 * mm
    col_w = (usable - gap) / 2
    left_x = margin
    right_x = margin + col_w + gap

    # ========== PAGE 1: 1 → 2 ==========
    page_header(c, page_w, page_h, fonts)
    page_footer(c, page_w, 1, 3, fonts)
    flow_strip(c, page_w, fonts, active={1, 2})

    stage_card(
        c,
        left_x,
        top,
        col_w,
        1,
        "Открытие",
        [
            "Добрый день! [Имя], компания [Название]. Удобно сейчас на пару минут?",
        ],
        fonts,
        TEAL,
    )

    stage_card(
        c,
        right_x,
        top,
        col_w,
        2,
        "Инфоповод",
        [
            "Вы подключали ЭДО, из других продуктов ничего не брали. Хочу уточнить — как у вас учёт организован, есть 1С?",
        ],
        fonts,
        TEAL,
    )

    draw_arrow_right(
        c,
        left_x + col_w + 2,
        right_x - 2,
        top - 28 * mm,
        "этап 1 → этап 2",
        fonts,
    )

    # Down arrow hint to page 2
    c.setFillColor(TEAL)
    c.setFont("BodyBold", 9)
    c.drawCentredString(page_w / 2, 22 * mm, "далее → этап 3. Навигатор по ответам клиента")

    c.showPage()

    # ========== PAGE 2: stage 3 navigator ==========
    page_header(c, page_w, page_h, fonts)
    page_footer(c, page_w, 2, 3, fonts)
    flow_strip(c, page_w, fonts, active={3})

    # Full width navigator card
    stage_card(
        c,
        margin,
        top,
        usable,
        3,
        "Навигатор по ответам клиента",
        [
            "«Да, есть 1С» → «В каком формате: локально или облако? Сопровождение есть?» → переходи к этапу 4 «Углубление».",
            "«1С: Бухгалтерия / УТ / другая конфигурация» → «Поняла, спасибо. А обновления кто делает?» → 2–3 вопроса из этапа 4 → затем этап 5, предложение ИТС.",
            "«1С в облаке» → «А вот когда возникают вопросы, вы решаете их с помощью своих программистов или куда-то обращаетесь?» Слушаем ответ. «Поняла вас. В двух словах: мы официальный партнёр 1С более 30 лет, работаем по всей России. У нас опытные специалисты. Предлагаю рассмотреть условия нашего сопровождения.» → этап 5.",
            "«Нет 1С / Excel / другая программа» → «Поняла. А бухгалтерию сами ведёте или на аутсорсе?» Уточни потребность. Если аутсорс — спроси, кто отвечает за учёт.",
            "«Не знаю / уточню» → «А кто у вас за это отвечает? Может, с ним поговорить?» Запиши контакт ЛПР → завершай.",
            "«Всё устраивает, ничего не надо» → «Рада слышать! А сопровождение есть? Часто при сбое теряют время и деньги.» Если нет сопровождения → предложи продукт и ИТС. Если есть → оставь контакты.",
            "«Дорого» → «ИТС от Х руб — дешевле одного вызова специалиста» → отправь коммерческое предложение на почту.",
            "«Пришлите на почту» → «Конечно отправлю. А вы сами принимаете решение или с кем-то согласуете? Как вам в целом моё предложение?» Запиши почту и роль → договорись о следующем контакте.",
            "«Подумаю» → «Хорошо. Пришлю материал. Как вам в целом моё предложение?» Отработай возражение, предложи встречу, зафиксируй дату.",
            "«Нет» (отказ) → «Спасибо за честность. Если что, мы на связи. Давайте предложение направлю, ознакомитесь, и позже свяжусь?» Заверши звонок, сохрани контакт, отправь коммерческое предложение.",
        ],
        fonts,
        BLUE,
    )

    c.setFillColor(TEAL)
    c.setFont("BodyBold", 9)
    c.drawCentredString(
        page_w / 2,
        22 * mm,
        "если подтвердили 1С → этап 4 «Углубление» → этап 5 «Предложение»",
    )

    c.showPage()

    # ========== PAGE 3: 4 → 5, then objections ==========
    page_header(c, page_w, page_h, fonts)
    page_footer(c, page_w, 3, 3, fonts)
    flow_strip(c, page_w, fonts, active={4, 5, 6})

    yb4, _ = stage_card(
        c,
        left_x,
        top,
        col_w,
        4,
        "Углубление (если работает в 1С)",
        [
            "Обновления кто делает — свой специалист или ищете со стороны?",
            "Бывает, что программа работает не так, как надо, или вопросы возникают?",
            "Пользуетесь сопровождением? Как решаете вопросы?",
        ],
        fonts,
        TEAL,
    )

    yb5, _ = stage_card(
        c,
        right_x,
        top,
        col_w,
        5,
        "Предложение",
        [
            "Нет сопровождения / ИТС → ИТС — официальная подписка 1С: обновления, консультации, сервисы.",
            "Локальная 1С + интерес к облаку → 1С:Фреш — облако из браузера, без установок.",
            "Уже всё есть → оставить контакты: «если что-то понадобится».",
        ],
        fonts,
        ORANGE,
    )

    draw_arrow_right(
        c,
        left_x + col_w + 2,
        right_x - 2,
        top - 35 * mm,
        "этап 4 → этап 5",
        fonts,
    )

    # Objections below
    bottom_top = min(yb4, yb5) - 12 * mm
    draw_arrow_down(
        c,
        page_w / 2,
        min(yb4, yb5) - 2,
        bottom_top + 2,
        "этап 5 → этап 6",
        fonts,
    )

    stage_card(
        c,
        margin,
        bottom_top,
        usable,
        6,
        "Отработка возражений",
        [
            "«У нас уже всё есть» → «А сопровождение есть? Часто программа есть, а поддержки нет — и при сбое теряют время.» Уточни → если нет ИТС, предложи.",
            "«Не нужно» → «А что сейчас закрывает эту задачу?» Выясни альтернативу → покажи отличие.",
            "«Дорого» → «ИТС от Х руб — это дешевле одного вызова специалиста + решение вопросов с помощью 1С:Напарник + сервисы.» Предложи встречу по демонстрации.",
            "«Пришлите на почту» → «Конечно отправлю. А вы сами принимаете решение или с кем-то согласуете? Как вам в целом моё предложение?» Определи ЛПР → назначь следующий контакт.",
            "«Подумаю» → «Хорошо. Пришлю материал. Как вам в целом моё предложение?» Отработай возражение, предложи встречу, зафиксируй дату.",
        ],
        fonts,
        ORANGE,
    )

    c.showPage()
    c.save()
    shutil.copy(OUT, ASCII_OUT)
    print("Saved", OUT)
    print("Saved", ASCII_OUT)


if __name__ == "__main__":
    build()
