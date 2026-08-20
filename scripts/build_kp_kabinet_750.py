#!/usr/bin/env python3
"""КП-листовка PDF: 1С:Кабинет сотрудника для корп. клиента (750 кабинетов)."""

from __future__ import annotations

from pathlib import Path

from reportlab.lib.colors import Color, HexColor, white, black
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "КП_Кабинет_сотрудника_750.pdf"
BRAND = ROOT / "assets_forus" / "brand"

# Фирменные цвета Форус
YELLOW = HexColor("#E8B84A")
YELLOW_SOFT = HexColor("#FFF6E0")
DARK = HexColor("#1E1E1E")
GRAY = HexColor("#555555")
LIGHT = HexColor("#F4F4F4")
BAD = HexColor("#F8E6E6")
BAD_TEXT = HexColor("#8B1E1E")
GOOD = HexColor("#E8F5E9")
GOOD_TEXT = HexColor("#1B5E20")
LINE = HexColor("#DDDDDD")

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

PAGE_W, PAGE_H = A4
ML, MR, MT, MB = 12 * mm, 12 * mm, 11 * mm, 11 * mm
CONTENT_W = PAGE_W - ML - MR


def draw_wave(c: canvas.Canvas, x, y, w, h, color=YELLOW, flip=False):
    """Декоративная волна в стиле листовок Форус."""
    c.setFillColor(color)
    c.setStrokeColor(color)
    path = c.beginPath()
    if not flip:
        path.moveTo(x, y)
        path.curveTo(x + w * 0.25, y + h, x + w * 0.5, y - h * 0.3, x + w * 0.75, y + h * 0.6)
        path.curveTo(x + w * 0.9, y + h, x + w, y + h * 0.2, x + w, y + h * 0.4)
        path.lineTo(x + w, y + h * 1.4)
        path.lineTo(x, y + h * 1.4)
        path.close()
    else:
        path.moveTo(x, y + h)
        path.curveTo(x + w * 0.3, y, x + w * 0.55, y + h * 1.2, x + w * 0.8, y + h * 0.3)
        path.curveTo(x + w * 0.9, y, x + w, y + h * 0.5, x + w, y)
        path.lineTo(x + w, y - h * 0.5)
        path.lineTo(x, y - h * 0.5)
        path.close()
    c.drawPath(path, fill=1, stroke=0)


def draw_yellow_bar(c, y, thickness=2.2):
    c.setFillColor(YELLOW)
    c.rect(ML, y, CONTENT_W, thickness, fill=1, stroke=0)


def wrap_text(c, text, font, size, max_width):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def draw_paragraph(c, text, x, y, max_width, font="DejaVu", size=8.5, color=DARK, leading=11):
    c.setFillColor(color)
    c.setFont(font, size)
    for line in wrap_text(c, text, font, size, max_width):
        c.drawString(x, y, line)
        y -= leading
    return y


def draw_bullet(c, text, x, y, max_width, size=8.2, leading=10.5, bullet_color=YELLOW):
    c.setFillColor(bullet_color)
    c.circle(x + 1.6 * mm, y + 1.2, 1.3, fill=1, stroke=0)
    return draw_paragraph(c, text, x + 4.5 * mm, y, max_width - 4.5 * mm, size=size, leading=leading)


def rounded_rect(c, x, y, w, h, r=4, fill_color=None, stroke_color=None, stroke_width=1):
    c.saveState()
    if fill_color:
        c.setFillColor(fill_color)
    if stroke_color:
        c.setStrokeColor(stroke_color)
        c.setLineWidth(stroke_width)
    c.roundRect(x, y, w, h, r, fill=1 if fill_color else 0, stroke=1 if stroke_color else 0)
    c.restoreState()


def header(c):
    logo = BRAND / "forus_logo_word.png"
    if logo.exists():
        c.drawImage(str(logo), ML, PAGE_H - MT - 14 * mm, width=38 * mm, height=14 * mm, mask="auto", preserveAspectRatio=True, anchor="sw")
    else:
        c.setFont("DejaVuBold", 18)
        c.setFillColor(DARK)
        c.drawString(ML, PAGE_H - MT - 8 * mm, "Форус")

    c.setFillColor(GRAY)
    c.setFont("DejaVu", 7.5)
    right = [
        "Группа компаний «Форус»",
        "Центр компетенции по кадровому электронному документообороту",
        "+7 (3952) 78-00-00  ·  www.forus.ru",
    ]
    y = PAGE_H - MT - 4 * mm
    for line in right:
        c.drawRightString(PAGE_W - MR, y, line)
        y -= 9

    draw_yellow_bar(c, PAGE_H - MT - 16 * mm, 2.5)
    # decorative wave top-right
    draw_wave(c, PAGE_W - 55 * mm, PAGE_H - 18 * mm, 55 * mm, 8 * mm, YELLOW)


def footer(c, page_no: int):
    draw_wave(c, 0, 0, 50 * mm, 7 * mm, YELLOW, flip=True)
    draw_yellow_bar(c, MB - 2 * mm, 1.8)
    c.setFont("DejaVu", 7)
    c.setFillColor(GRAY)
    c.drawString(ML, MB - 6 * mm, "www.forus.ru  ·  г. Иркутск, ул. Ямская, 1/1")
    c.drawRightString(PAGE_W - MR, MB - 6 * mm, f"{page_no} / 2")


def page1(c: canvas.Canvas):
    header(c)
    y = PAGE_H - MT - 22 * mm

    c.setFont("DejaVuBold", 16)
    c.setFillColor(DARK)
    c.drawCentredString(PAGE_W / 2, y, "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ")
    y -= 6 * mm
    c.setFont("DejaVuBold", 11)
    c.setFillColor(YELLOW)
    # yellow underline pill
    title = "1С:Кабинет сотрудника — 750 личных кабинетов"
    tw = c.stringWidth(title, "DejaVuBold", 11)
    rounded_rect(c, (PAGE_W - tw) / 2 - 4 * mm, y - 2 * mm, tw + 8 * mm, 7 * mm, r=3, fill_color=DARK)
    c.setFillColor(YELLOW)
    c.drawCentredString(PAGE_W / 2, y, title)
    y -= 9 * mm

    c.setFont("DejaVu", 8.5)
    c.setFillColor(GRAY)
    c.drawCentredString(PAGE_W / 2, y, "Персональное предложение для корпоративного клиента · кадровый электронный документооборот")
    y -= 7 * mm

    # Hero value strip
    rounded_rect(c, ML, y - 18 * mm, CONTENT_W, 20 * mm, r=5, fill_color=YELLOW_SOFT, stroke_color=YELLOW, stroke_width=1.5)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 9)
    c.drawString(ML + 4 * mm, y - 4 * mm, "Почему это предложение выгодно именно сейчас")
    benefits = [
        "Экономия до 70% времени кадровой службы и до 75% затрат на бумагу, печать и курьеров",
        "Акция «Больше, чем кешбэк»: 5 часов линии консультаций в подарок (18 300 ₽)",
        "Работа внутри привычной 1С — без отдельной HR-платформы и двойного ввода данных",
    ]
    yy = y - 8 * mm
    for b in benefits:
        yy = draw_bullet(c, b, ML + 3 * mm, yy, CONTENT_W - 8 * mm, size=7.8, leading=9.5)
    y -= 24 * mm

    # Capabilities - user's "что умею"
    c.setFont("DejaVuBold", 11)
    c.setFillColor(DARK)
    c.drawString(ML, y, "Что умеет 1С:Кабинет сотрудника")
    draw_yellow_bar(c, y - 2 * mm, 1.6)
    y -= 7 * mm

    caps = [
        (
            "Ознакомление с документами в один клик",
            "Больше не нужно распечатывать, собирать подписи и хранить горы бумаг. Отправьте любой документ сотрудникам одной кнопкой — они ознакомятся и подтвердят получение.",
        ),
        (
            "Удалённый приём на работу",
            "Новый сотрудник оформляет документы дистанционно: трудовой договор, заявление на приём, ознакомление с правилами — без визита в офис.",
        ),
        (
            "Ответы на вопросы по отпуску",
            "Сотрудник сам смотрит остаток отпуска и график в личном кабинете и подаёт заявление онлайн. Кадры и бухгалтерия не тратят время на одни и те же вопросы.",
        ),
        (
            "Согласование без походов по кабинетам",
            "Отпуск, командировка, отгул — руководитель согласовывает с телефона, данные сразу уходят в 1С.",
        ),
        (
            "Общение без сторонних мессенджеров",
            "Рабочие вопросы и документы — внутри системы, с учётом требований закона о персональных данных.",
        ),
    ]

    col_w = (CONTENT_W - 3 * mm) / 2
    # first 4 in 2x2, last full width
    for i in range(0, 4, 2):
        for col in range(2):
            idx = i + col
            title_t, body = caps[idx]
            x = ML + col * (col_w + 3 * mm)
            rounded_rect(c, x, y - 22 * mm, col_w, 23 * mm, r=4, fill_color=LIGHT)
            c.setFillColor(YELLOW)
            c.rect(x, y - 22 * mm, 1.8 * mm, 23 * mm, fill=1, stroke=0)
            c.setFillColor(DARK)
            c.setFont("DejaVuBold", 8)
            c.drawString(x + 4 * mm, y - 4 * mm, title_t)
            draw_paragraph(c, body, x + 4 * mm, y - 8 * mm, col_w - 7 * mm, size=7.3, leading=9.2, color=GRAY)
        y -= 25 * mm

    # 5th full
    title_t, body = caps[4]
    rounded_rect(c, ML, y - 14 * mm, CONTENT_W, 15 * mm, r=4, fill_color=LIGHT)
    c.setFillColor(YELLOW)
    c.rect(ML, y - 14 * mm, 1.8 * mm, 15 * mm, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8)
    c.drawString(ML + 4 * mm, y - 4 * mm, title_t)
    draw_paragraph(c, body, ML + 4 * mm, y - 8 * mm, CONTENT_W - 8 * mm, size=7.3, leading=9.2, color=GRAY)
    y -= 18 * mm

    # Pricing block
    c.setFont("DejaVuBold", 11)
    c.setFillColor(DARK)
    c.drawString(ML, y, "Ваш пакет: 750 личных кабинетов")
    draw_yellow_bar(c, y - 2 * mm, 1.6)
    y -= 6 * mm

    # price table
    rows = [
        ("Позиция", "Состав / пояснение", "Сумма, ₽", True),
        ("1С:Кабинет сотрудника, 750 кабинетов на 12 месяцев", "Пакеты: 500 + 200 + 50 кабинетов", "223 200", False),
        ("Линия консультаций — 4 часа (оплачиваете)", "Удалённая помощь по запуску и настройке, тариф от 4 часов", "13 680", False),
        ("ПОДАРОК: 5 часов линии консультаций", "Акция «Больше, чем кешбэк!» — начисляем на баланс бесплатно", "0 (выгода 18 300)", False),
    ]

    row_h = 7.2 * mm
    table_h = row_h * len(rows)
    rounded_rect(c, ML, y - table_h, CONTENT_W, table_h, r=3, fill_color=white, stroke_color=LINE)
    yy = y
    col1, col2, col3 = ML, ML + 78 * mm, ML + 145 * mm
    for i, (a, b, d, is_head) in enumerate(rows):
        yy -= row_h
        if is_head:
            c.setFillColor(DARK)
            c.rect(ML, yy, CONTENT_W, row_h, fill=1, stroke=0)
            c.setFillColor(white)
            c.setFont("DejaVuBold", 7.5)
        else:
            if i % 2 == 0:
                c.setFillColor(YELLOW_SOFT if "ПОДАРОК" in a else LIGHT)
                c.rect(ML, yy, CONTENT_W, row_h, fill=1, stroke=0)
            if "ПОДАРОК" in a:
                c.setFillColor(YELLOW_SOFT)
                c.rect(ML, yy, CONTENT_W, row_h, fill=1, stroke=0)
            c.setFillColor(DARK)
            c.setFont("DejaVuBold" if "ПОДАРОК" in a else "DejaVu", 7.2)
        c.drawString(col1 + 2 * mm, yy + 2.5 * mm, a[:58] + ("…" if len(a) > 58 else ""))
        c.setFont("DejaVu", 6.8)
        if not is_head:
            c.setFillColor(GRAY)
        c.drawString(col2, yy + 2.5 * mm, b[:42] + ("…" if len(b) > 42 else ""))
        c.setFont("DejaVuBold", 7.5)
        c.setFillColor(GOOD_TEXT if "ПОДАРОК" in a else (white if is_head else DARK))
        c.drawRightString(PAGE_W - MR - 2 * mm, yy + 2.5 * mm, d)

    y = yy - 5 * mm

    # totals
    rounded_rect(c, ML, y - 22 * mm, CONTENT_W, 23 * mm, r=5, fill_color=DARK)
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 9)
    c.drawString(ML + 4 * mm, y - 5 * mm, "ИТОГО К ОПЛАТЕ")
    c.setFont("DejaVuBold", 16)
    c.drawRightString(PAGE_W - MR - 4 * mm, y - 6 * mm, "236 880 ₽")
    c.setFillColor(white)
    c.setFont("DejaVu", 7.5)
    c.drawString(ML + 4 * mm, y - 11 * mm, "Кабинеты 223 200 ₽ + 4 часа линии 13 680 ₽")
    c.drawString(ML + 4 * mm, y - 15.5 * mm, "Вы получаете сразу 9 часов поддержки (4 оплаченных + 5 подарочных) на сумму 30 780 ₽")
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 8)
    c.drawString(ML + 4 * mm, y - 20 * mm, "≈ 26 ₽ в месяц за одного сотрудника  ·  подпись для сотрудников — бесплатно")

    footer(c, 1)


def page2(c: canvas.Canvas):
    header(c)
    y = PAGE_H - MT - 22 * mm

    c.setFont("DejaVuBold", 12)
    c.setFillColor(DARK)
    c.drawString(ML, y, "Внедрение у Форус vs другие операторы")
    draw_yellow_bar(c, y - 2 * mm, 1.6)
    y -= 5 * mm
    c.setFont("DejaVu", 7.5)
    c.setFillColor(GRAY)
    c.drawString(ML, y, "Сравнение по ключевым параметрам для компании на 750 сотрудников. Красным выделены самые невыгодные условия у конкурентов.")
    y -= 5 * mm

    # Comparison table
    headers = ["Параметр", "Форус\n+ 1С:Кабинет", "HRlink /\nVK HR Tek", "Контур КЭДО", "Saby"]
    data = [
        ["Работа внутри вашей 1С", "Да — без новой платформы", "Нет — отдельная система", "Нет — отдельная система", "Нет — отдельная система"],
        ["Двойной ввод данных\nкадровиком", "Нет — заявления сразу\nстановятся документами 1С", "Часто нужен перенос\nи сверка", "Часто нужен перенос\nи сверка", "Часто нужен перенос\nи сверка"],
        ["Где хранятся документы", "В вашей базе 1С", "Облако оператора\n(часто платно)", "Облако / доп. хранилище", "Облако оператора"],
        ["Электронная подпись\nсотрудникам", "Бесплатно\n(усиленная неквалифицированная)", "Часто платные\nсертификаты / пакеты", "Часто платные\nсертификаты / пакеты", "Часто платные\nсертификаты / пакеты"],
        ["Сложность внедрения", "Быстрый старт\nв привычной 1С", "Долгая интеграция\nс 1С", "Долгая интеграция\nс 1С", "Долгая интеграция\nс 1С"],
        ["Нагрузка на ИТ", "Без новых клиентских\nлицензий 1С", "Новая система +\nинтеграции", "Новая система +\nинтеграции", "Новая система +\nинтеграции"],
        ["Стоимость владения\nна старте (ориентир)", "Пакет кабинетов +\nчасы линии", "Лицензии + внедрение\n+ хранение", "Лицензии + внедрение\n+ хранение", "Лицензии + внедрение\n+ хранение"],
    ]

    # column widths
    w0 = 38 * mm
    w_rest = (CONTENT_W - w0) / 4
    widths = [w0, w_rest, w_rest, w_rest, w_rest]
    row_h = 11.5 * mm
    head_h = 10 * mm

    # header row
    x = ML
    c.setFillColor(DARK)
    c.rect(ML, y - head_h, CONTENT_W, head_h, fill=1, stroke=0)
    c.setFillColor(YELLOW)
    c.rect(ML + widths[0], y - head_h, widths[1], head_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("DejaVuBold", 6.8)
    xx = ML
    for i, h in enumerate(headers):
        lines = h.split("\n")
        c.setFillColor(DARK if i == 1 else white)
        if i == 1:
            pass  # yellow bg already, dark text
        text_color = DARK if i == 1 else white
        c.setFillColor(text_color)
        for li, line in enumerate(lines):
            c.drawCentredString(xx + widths[i] / 2, y - 4 * mm - li * 3.2 * mm, line)
        xx += widths[i]
    y -= head_h

    bad_cols = {2, 3, 4}  # competitor columns to highlight as disadvantageous when cell is negative

    for ri, row in enumerate(data):
        yy = y - row_h
        xx = ML
        for ci, cell in enumerate(row):
            # background
            if ci == 1:
                bg = GOOD
            elif ci in bad_cols and ri in (0, 1, 2, 3, 4, 5, 6):
                bg = BAD
            elif ri % 2 == 0:
                bg = LIGHT
            else:
                bg = white
            c.setFillColor(bg)
            c.rect(xx, yy, widths[ci], row_h, fill=1, stroke=0)
            c.setStrokeColor(LINE)
            c.setLineWidth(0.4)
            c.rect(xx, yy, widths[ci], row_h, fill=0, stroke=1)

            if ci == 1:
                c.setFillColor(GOOD_TEXT)
                c.setFont("DejaVuBold", 6.2)
            elif ci in bad_cols:
                c.setFillColor(BAD_TEXT)
                c.setFont("DejaVu", 6.0)
            else:
                c.setFillColor(DARK)
                c.setFont("DejaVuBold" if ci == 0 else "DejaVu", 6.2)

            lines = cell.split("\n")
            for li, line in enumerate(lines[:3]):
                if ci == 0:
                    c.drawString(xx + 1.5 * mm, yy + row_h - 3.5 * mm - li * 3 * mm, line)
                else:
                    c.drawCentredString(xx + widths[ci] / 2, yy + row_h - 3.5 * mm - li * 3 * mm, line)
            xx += widths[ci]
        y = yy

    y -= 5 * mm
    rounded_rect(c, ML, y - 12 * mm, CONTENT_W, 13 * mm, r=4, fill_color=YELLOW_SOFT, stroke_color=YELLOW, stroke_width=1.2)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8)
    c.drawString(ML + 3 * mm, y - 4 * mm, "Вывод для корпоративного клиента")
    draw_paragraph(
        c,
        "Отдельные операторы кадрового электронного документооборота почти всегда тянут за собой новую систему, "
        "двойную работу кадровика, платное хранение и долгую интеграцию с 1С. "
        "С Форус вы остаётесь в привычной 1С, получаете подарочные часы поддержки и прозрачную стоимость пакета.",
        ML + 3 * mm,
        y - 7.5 * mm,
        CONTENT_W - 6 * mm,
        size=7.3,
        leading=9.2,
    )
    y -= 17 * mm

    # How we implement + promo reminder
    c.setFont("DejaVuBold", 11)
    c.setFillColor(DARK)
    c.drawString(ML, y, "Как мы запускаем сервис у вас")
    draw_yellow_bar(c, y - 2 * mm, 1.6)
    y -= 6 * mm

    steps = [
        ("01", "Подключение", "Подключаем сервис к вашей 1С"),
        ("02", "Настройка", "Роли, процессы, печатные формы"),
        ("03", "Подписи", "Выпуск подписей сотрудникам"),
        ("04", "Обучение", "Видеоуроки и короткие инструкции"),
        ("05", "Старт", "Заявления, расчётные, согласования"),
    ]
    sw = CONTENT_W / 5
    for i, (num, title, body) in enumerate(steps):
        x = ML + i * sw
        rounded_rect(c, x + 1 * mm, y - 18 * mm, sw - 2 * mm, 19 * mm, r=4, fill_color=LIGHT)
        c.setFillColor(YELLOW)
        c.setFont("DejaVuBold", 10)
        c.drawCentredString(x + sw / 2, y - 5 * mm, num)
        c.setFillColor(DARK)
        c.setFont("DejaVuBold", 7)
        c.drawCentredString(x + sw / 2, y - 9 * mm, title)
        c.setFont("DejaVu", 6.2)
        c.setFillColor(GRAY)
        for li, line in enumerate(wrap_text(c, body, "DejaVu", 6.2, sw - 5 * mm)[:2]):
            c.drawCentredString(x + sw / 2, y - 12.5 * mm - li * 3 * mm, line)
    y -= 23 * mm

    # Why Forus
    c.setFont("DejaVuBold", 11)
    c.setFillColor(DARK)
    c.drawString(ML, y, "Почему Форус")
    draw_yellow_bar(c, y - 2 * mm, 1.6)
    y -= 6 * mm

    why = [
        "Центр компетенции по кадровому электронному документообороту и по управлению персоналом",
        "ТОП-5 дистрибьюторов 1С в России · более 10 000 клиентов на сопровождении",
        "Успешные внедрения, включая крупные сети (в том числе кейс DNS)",
        "Шаблоны документов для перехода, видеоуроки, чат со специалистом и линия консультаций",
    ]
    for w in why:
        y = draw_bullet(c, w, ML, y, CONTENT_W, size=7.8, leading=10)
        y -= 1 * mm

    y -= 3 * mm

    # CTA box
    rounded_rect(c, ML, y - 28 * mm, CONTENT_W, 29 * mm, r=5, fill_color=DARK)
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 11)
    c.drawCentredString(PAGE_W / 2, y - 6 * mm, "Готовы зафиксировать условия и запустить 750 кабинетов?")
    c.setFillColor(white)
    c.setFont("DejaVu", 8)
    c.drawCentredString(PAGE_W / 2, y - 11 * mm, "Оставьте подтверждение — подключим сервис, начислим 5 подарочных часов и начнём настройку.")

    # manager
    c.setFillColor(YELLOW_SOFT)
    c.roundRect(ML + 4 * mm, y - 26 * mm, CONTENT_W - 8 * mm, 12 * mm, 3, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8.5)
    c.drawString(ML + 7 * mm, y - 18 * mm, "Ваш менеджер: Оглоблина Софья")
    c.setFont("DejaVu", 7.5)
    c.setFillColor(GRAY)
    c.drawString(ML + 7 * mm, y - 22 * mm, "sogloblina@forus.ru  ·  +7 (3952) 78-00-00, доб. 1861  ·  Москва +5 часов")
    c.drawRightString(PAGE_W - MR - 7 * mm, y - 18 * mm, "ГК «Форус»")
    c.drawRightString(PAGE_W - MR - 7 * mm, y - 22 * mm, "www.forus.ru")

    footer(c, 2)


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=A4)
    c.setTitle("Коммерческое предложение — 1С:Кабинет сотрудника (750)")
    c.setAuthor("ГК Форус")

    page1(c)
    c.showPage()
    page2(c)
    c.save()

    root_copy = ROOT / "КП_Кабинет_сотрудника_750.pdf"
    root_copy.write_bytes(OUT.read_bytes())
    print(f"Saved: {OUT}")
    print(f"Saved: {root_copy}")
    print(f"Size: {OUT.stat().st_size} bytes")


if __name__ == "__main__":
    build()
