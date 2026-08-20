#!/usr/bin/env python3
"""КП PDF для ТФМ Спецтехника — КЭДО / 1С:Кабинет сотрудника, 750 сотрудников.

Тендер: условия, сроки, бюджет.
Акцент: простота внедрения + ИТ-отдел.
Менеджер: Данил Кургузов.
Линия консультаций: 5 часов в подарок (4 на настройку + 1 в запас) — клиент не платит.
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "КП_ТФМ_Спецтехника_КЭДО_750.pdf"
BRAND = ROOT / "assets_forus" / "brand"

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
ML, MR, MT, MB = 11 * mm, 11 * mm, 10 * mm, 10 * mm
CONTENT_W = PAGE_W - ML - MR

CABINETS = 223_200  # 500+200+50
HOURS_SETUP = 4
HOURS_RESERVE = 1
HOURS_GIFT = 5
HOUR_VALUE = 3_660
GIFT_VALUE = HOURS_GIFT * HOUR_VALUE  # 18300
TOTAL = CABINETS  # клиент платит только за кабинеты

MANAGER_NAME = "Данил Кургузов"
MANAGER_EMAIL = "dkurguzov@forus.ru"
MANAGER_PHONE = "+7 (3952) 78-00-00"


def draw_wave(c, x, y, w, h, color=YELLOW, flip=False):
    c.setFillColor(color)
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


def yellow_bar(c, y, t=2.2):
    c.setFillColor(YELLOW)
    c.rect(ML, y, CONTENT_W, t, fill=1, stroke=0)


def wrap(c, text, font, size, max_w):
    words, lines, cur = text.split(), [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def para(c, text, x, y, max_w, font="DejaVu", size=8, color=DARK, leading=10.5):
    c.setFillColor(color)
    c.setFont(font, size)
    for line in wrap(c, text, font, size, max_w):
        c.drawString(x, y, line)
        y -= leading
    return y


def bullet(c, text, x, y, max_w, size=7.8, leading=10):
    c.setFillColor(YELLOW)
    c.circle(x + 1.5 * mm, y + 1.1, 1.2, fill=1, stroke=0)
    return para(c, text, x + 4.2 * mm, y, max_w - 4.2 * mm, size=size, leading=leading)


def round_rect(c, x, y, w, h, r=4, fill=None, stroke=None, sw=1):
    c.saveState()
    if fill:
        c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(sw)
    c.roundRect(x, y, w, h, r, fill=1 if fill else 0, stroke=1 if stroke else 0)
    c.restoreState()


def section_title(c, text, y):
    c.setFont("DejaVuBold", 10.5)
    c.setFillColor(DARK)
    c.drawString(ML, y, text)
    yellow_bar(c, y - 1.8 * mm, 1.5)
    return y - 6 * mm


def header(c):
    logo = BRAND / "forus_logo_word.png"
    if logo.exists():
        c.drawImage(
            str(logo), ML, PAGE_H - MT - 13 * mm,
            width=36 * mm, height=13 * mm, mask="auto",
            preserveAspectRatio=True, anchor="sw",
        )
    c.setFillColor(GRAY)
    c.setFont("DejaVu", 7)
    lines = [
        "Группа компаний «Форус»",
        "Центр компетенции по кадровому электронному документообороту",
        f"{MANAGER_PHONE}  ·  www.forus.ru",
    ]
    yy = PAGE_H - MT - 3.5 * mm
    for line in lines:
        c.drawRightString(PAGE_W - MR, yy, line)
        yy -= 8.5
    yellow_bar(c, PAGE_H - MT - 15 * mm, 2.4)
    draw_wave(c, PAGE_W - 48 * mm, PAGE_H - 16 * mm, 48 * mm, 7 * mm, YELLOW)


def footer(c, page):
    draw_wave(c, 0, 0, 45 * mm, 6 * mm, YELLOW, flip=True)
    yellow_bar(c, MB - 1.5 * mm, 1.6)
    c.setFont("DejaVu", 6.5)
    c.setFillColor(GRAY)
    c.drawString(ML, MB - 5.5 * mm, "Для: ТФМ Спецтехника  ·  кадры2  ·  конфиденциально")
    c.drawRightString(PAGE_W - MR, MB - 5.5 * mm, f"{page} / 2")


def page1(c):
    header(c)
    y = PAGE_H - MT - 20 * mm

    c.setFont("DejaVuBold", 15)
    c.setFillColor(DARK)
    c.drawCentredString(PAGE_W / 2, y, "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ")
    y -= 5.5 * mm

    round_rect(c, ML, y - 8 * mm, CONTENT_W, 10 * mm, r=4, fill=DARK)
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 9.5)
    c.drawCentredString(PAGE_W / 2, y - 3.5 * mm, "ТФМ Спецтехника  ·  кадровый электронный документооборот")
    c.setFillColor(white)
    c.setFont("DejaVu", 7)
    c.drawCentredString(
        PAGE_W / 2, y - 7 * mm,
        "750 сотрудников  ·  для тендера  ·  кадры2 + ИТ-отдел",
    )
    y -= 12 * mm

    y = para(
        c,
        "Уважаемые коллеги! ГК «Форус» направляет коммерческое предложение по запуску "
        "кадрового электронного документооборота на базе «1С:Кабинет сотрудника». "
        "Ниже — условия, сроки и бюджет для выбора подрядчика.",
        ML, y, CONTENT_W, size=8, leading=10.2,
    )
    y -= 2.5 * mm

    # Простота внедрения — главный акцент
    y = section_title(c, "Главное преимущество — простое внедрение", y)
    round_rect(c, ML, y - 28 * mm, CONTENT_W, 29 * mm, r=4, fill=YELLOW_SOFT, stroke=YELLOW, sw=1.3)
    easy = [
        "Сервис уже внутри 1С: не ставим отдельную HR-платформу и не ломаем привычные процессы.",
        "Типовая настройка занимает около 4 часов работы специалистов — без длинного проектного внедрения.",
        "Сотрудники работают в личном кабинете с телефона или компьютера; в базу 1С их пускать не нужно.",
        "Кадровик продолжает работать в знакомой 1С: заявления сами превращаются в документы учёта.",
        "ИТ получает готовый контур без новых клиентских лицензий 1С и без сложной интеграции «с нуля».",
    ]
    yy = y - 4 * mm
    for t in easy:
        yy = bullet(c, t, ML + 2 * mm, yy, CONTENT_W - 5 * mm, size=7.4, leading=9.3)
    y -= 31 * mm

    # IT block shorter
    y = section_title(c, "Что важно ИТ-отделу", y)
    for t in [
        "Нет прямой нагрузки сотрудников на информационную базу 1С и нет отдельных контуров доступа.",
        "Документы хранятся в вашей 1С; дата-центр 1С аттестован ФСТЭК.",
        "Электронная подпись сотрудникам выпускается бесплатно (усиленная неквалифицированная).",
    ]:
        y = bullet(c, t, ML, y, CONTENT_W, size=7.4, leading=9.4)
        y -= 0.3 * mm
    y -= 2 * mm

    # Capabilities compact 2x3
    y = section_title(c, "Что получает компания", y)
    caps = [
        ("Ознакомление в 1 клик", "Документы без печати и сбора подписей"),
        ("Удалённый приём", "Оформление без визита в офис"),
        ("Отпуска без звонков", "Остаток, график и заявление в кабинете"),
        ("Согласование online", "Руководитель утверждает с телефона"),
        ("Без мессенджеров", "Обмен с учётом персональных данных"),
        ("Меньше рутины", "Заявления сразу становятся документами 1С"),
    ]
    cw = (CONTENT_W - 4 * mm) / 3
    for i in range(0, 6, 3):
        for col in range(3):
            title, body = caps[i + col]
            x = ML + col * (cw + 2 * mm)
            round_rect(c, x, y - 14.5 * mm, cw, 15.5 * mm, r=3, fill=LIGHT)
            c.setFillColor(YELLOW)
            c.rect(x, y - 14.5 * mm, 1.5 * mm, 15.5 * mm, fill=1, stroke=0)
            c.setFillColor(DARK)
            c.setFont("DejaVuBold", 7)
            c.drawString(x + 3.2 * mm, y - 4 * mm, title)
            para(c, body, x + 3.2 * mm, y - 8 * mm, cw - 5.5 * mm, size=6.5, leading=8.2, color=GRAY)
        y -= 17 * mm

    y -= 1 * mm
    y = section_title(c, "1. Условия предложения", y)
    for t in [
        "Объект: «1С:Кабинет сотрудника» — 750 личных кабинетов на 12 месяцев (пакеты 500 + 200 + 50).",
        "Внедрение типовое и быстрое: подключение к вашей 1С, настройка ролей и правил, выпуск подписей, запуск пилота.",
        "Акция «Больше, чем кешбэк»: при оплате пакета дарим 5 часов линии консультаций (выгода 18 300 ₽).",
        "Из подарочных часов: 4 часа — на настройку и запуск; ещё 1 час остаётся у вас на вопросы, консультации "
        "и решение возникающих задач по 1С с нашими специалистами.",
        "За часы линии консультаций вы не платите — они полностью покрываются подарком по акции.",
        "Шаблоны документов для перехода на кадровый электронный документооборот предоставляем.",
        "Срок действия предложения: 30 календарных дней. Работы — удалённо; выезд и нетиповые доработки — отдельно.",
    ]:
        y = bullet(c, t, ML, y, CONTENT_W, size=7.15, leading=9.1)

    footer(c, 1)


def page2(c):
    header(c)
    y = PAGE_H - MT - 20 * mm

    y = section_title(c, "2. Сроки внедрения", y)
    c.setFont("DejaVu", 7.2)
    c.setFillColor(GRAY)
    c.drawString(ML, y, "Простой план без «проекта на полгода». Сроки уточняются после согласования доступов.")
    y -= 4 * mm

    headers = ["Этап", "Срок", "Результат"]
    col_ws = [58 * mm, 38 * mm, CONTENT_W - 96 * mm]
    rows = [
        ["Договор и доступы к 1С", "2–4 рабочих дня", "Можно начинать настройку"],
        ["Настройка сервиса (4 подарочных часа)", "1 рабочий день", "Подключение, роли, правила, подписи"],
        ["Пилот 30–50 сотрудников", "3–7 рабочих дней", "Проверенный сценарий обмена"],
        ["Обучение ключевых пользователей", "параллельно", "Видеоуроки и короткие инструкции"],
        ["Тиражирование на 750 сотрудников", "2–3 недели", "Массовый запуск кабинетов"],
        ["Резерв 1 подарочный час", "по запросу", "Вопросы и помощь по 1С после запуска"],
    ]

    head_h = 6.5 * mm
    c.setFillColor(DARK)
    c.rect(ML, y - head_h, CONTENT_W, head_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("DejaVuBold", 7)
    xx = ML
    for i, h in enumerate(headers):
        c.drawString(xx + 1.5 * mm, y - head_h + 2 * mm, h)
        xx += col_ws[i]
    y -= head_h

    for ri, row in enumerate(rows):
        rh = 7.4 * mm
        yy = y - rh
        c.setFillColor(YELLOW_SOFT if ri in (1, 5) else (LIGHT if ri % 2 == 0 else white))
        c.rect(ML, yy, CONTENT_W, rh, fill=1, stroke=0)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.3)
        c.rect(ML, yy, CONTENT_W, rh, fill=0, stroke=1)
        xx = ML
        colors = [DARK, HexColor("#8A6A1A"), GRAY]
        fonts = ["DejaVuBold", "DejaVuBold", "DejaVu"]
        for ci, cell in enumerate(row):
            c.setFont(fonts[ci], 6.7)
            c.setFillColor(colors[ci])
            c.drawString(xx + 1.5 * mm, yy + 2.5 * mm, cell[:52])
            xx += col_ws[ci]
        y = yy

    y -= 3 * mm
    round_rect(c, ML, y - 10 * mm, CONTENT_W, 11 * mm, r=3, fill=GOOD, stroke=GOOD_TEXT, sw=0.8)
    c.setFillColor(GOOD_TEXT)
    c.setFont("DejaVuBold", 7.8)
    c.drawString(ML + 3 * mm, y - 4 * mm, "Ориентир: настройка за 1 день · пилот за 1–2 недели · полный охват 750 сотрудников за 1–1,5 месяца")
    c.setFont("DejaVu", 6.8)
    c.drawString(
        ML + 3 * mm, y - 8 * mm,
        "Не требуется отдельная платформа, долгая интеграция и выделение большой проектной команды со стороны клиента.",
    )
    y -= 14 * mm

    # Budget
    y = section_title(c, "3. Бюджет", y)

    headers = ["Статья", "Пояснение", "Сумма, ₽"]
    col_ws = [72 * mm, 78 * mm, CONTENT_W - 150 * mm]
    rows = [
        ["1С:Кабинет сотрудника, 750 / 12 мес.", "Пакеты 500 + 200 + 50 кабинетов", f"{CABINETS:,}".replace(",", " ")],
        ["Настройка и запуск — 4 часа", "Покрывается подарком по акции", "0"],
        ["Резерв на вопросы по 1С — 1 час", "Покрывается подарком по акции", "0"],
        ["ПОДАРОК: 5 часов линии консультаций", "Акция «Больше, чем кешбэк!» (выгода 18 300 ₽)", "0"],
    ]

    head_h = 6.5 * mm
    c.setFillColor(DARK)
    c.rect(ML, y - head_h, CONTENT_W, head_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("DejaVuBold", 7)
    xx = ML
    for i, h in enumerate(headers):
        c.drawString(xx + 1.5 * mm, y - head_h + 2 * mm, h)
        xx += col_ws[i]
    y -= head_h

    for ri, row in enumerate(rows):
        rh = 7 * mm
        yy = y - rh
        bg = YELLOW_SOFT if ri >= 1 else LIGHT
        c.setFillColor(bg)
        c.rect(ML, yy, CONTENT_W, rh, fill=1, stroke=0)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.3)
        c.rect(ML, yy, CONTENT_W, rh, fill=0, stroke=1)
        xx = ML
        for ci, cell in enumerate(row):
            if ci == 2:
                c.setFont("DejaVuBold", 7.2)
                c.setFillColor(GOOD_TEXT if ri >= 1 else DARK)
                c.drawRightString(xx + col_ws[ci] - 1.5 * mm, yy + 2.3 * mm, cell)
            else:
                c.setFont("DejaVuBold" if ci == 0 else "DejaVu", 6.7)
                c.setFillColor(DARK)
                c.drawString(xx + 1.5 * mm, yy + 2.3 * mm, cell[:46])
            xx += col_ws[ci]
        y = yy

    y -= 3 * mm
    round_rect(c, ML, y - 22 * mm, CONTENT_W, 23 * mm, r=5, fill=DARK)
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 9)
    c.drawString(ML + 4 * mm, y - 5 * mm, "ИТОГО К ОПЛАТЕ")
    c.setFont("DejaVuBold", 15)
    c.drawRightString(PAGE_W - MR - 4 * mm, y - 5.5 * mm, f"{TOTAL:,}".replace(",", " ") + " ₽")
    c.setFillColor(white)
    c.setFont("DejaVu", 7.3)
    c.drawString(ML + 4 * mm, y - 10.5 * mm, "Только пакет кабинетов. Часы линии консультаций — бесплатно по акции.")
    c.drawString(
        ML + 4 * mm, y - 14.5 * mm,
        "4 часа — настройка и запуск; 1 час остаётся на консультации и решение вопросов по 1С.",
    )
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 7.6)
    c.drawString(
        ML + 4 * mm, y - 19.5 * mm,
        "≈ 25 ₽ в месяц за сотрудника  ·  подпись сотрудникам бесплатно  ·  без оплаты внедрения отдельно",
    )
    y -= 26 * mm

    # Comparison short
    y = section_title(c, "Для тендера: Форус vs другие операторы", y)
    headers = ["Критерий", "Форус + 1С:Кабинет", "Другие операторы"]
    col_ws = [48 * mm, 72 * mm, CONTENT_W - 120 * mm]
    head_h = 6.2 * mm
    c.setFillColor(DARK)
    c.rect(ML, y - head_h, CONTENT_W, head_h, fill=1, stroke=0)
    c.setFillColor(YELLOW)
    c.rect(ML + col_ws[0], y - head_h, col_ws[1], head_h, fill=1, stroke=0)
    c.setFont("DejaVuBold", 6.8)
    xx = ML
    for i, h in enumerate(headers):
        c.setFillColor(DARK if i == 1 else white)
        c.drawCentredString(xx + col_ws[i] / 2, y - head_h + 1.9 * mm, h)
        xx += col_ws[i]
    y -= head_h

    cmp_rows = [
        ["Сложность внедрения", "Типовая настройка ~4 часа", "Долгая интеграция новой системы"],
        ["Платформа", "Внутри вашей 1С", "Отдельная HR-платформа"],
        ["Нагрузка на ИТ", "Без новых лицензий 1С", "Новые доступы и интеграции"],
        ["Хранение документов", "В вашей базе 1С", "Облако оператора, часто платно"],
        ["Оплата запуска", "Часы линии — в подарок", "Внедрение оплачивается отдельно"],
    ]
    for ri, row in enumerate(cmp_rows):
        rh = 6.5 * mm
        yy = y - rh
        xx = ML
        for ci, cell in enumerate(row):
            if ci == 1:
                bg, tc, f = GOOD, GOOD_TEXT, "DejaVuBold"
            elif ci == 2:
                bg, tc, f = BAD, BAD_TEXT, "DejaVu"
            else:
                bg, tc, f = (LIGHT if ri % 2 == 0 else white), DARK, "DejaVuBold"
            c.setFillColor(bg)
            c.rect(xx, yy, col_ws[ci], rh, fill=1, stroke=0)
            c.setStrokeColor(LINE)
            c.setLineWidth(0.3)
            c.rect(xx, yy, col_ws[ci], rh, fill=0, stroke=1)
            c.setFillColor(tc)
            c.setFont(f, 6.3)
            if ci == 0:
                c.drawString(xx + 1.3 * mm, yy + 2 * mm, cell)
            else:
                c.drawCentredString(xx + col_ws[ci] / 2, yy + 2 * mm, cell)
            xx += col_ws[ci]
        y = yy

    y -= 4 * mm
    round_rect(c, ML, y - 26 * mm, CONTENT_W, 27 * mm, r=5, fill=DARK)
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 9)
    c.drawCentredString(PAGE_W / 2, y - 5 * mm, "Готовы войти в тендер ТФМ Спецтехника")
    c.setFillColor(white)
    c.setFont("DejaVu", 7.2)
    c.drawCentredString(
        PAGE_W / 2, y - 9.5 * mm,
        "Простая настройка, прозрачный бюджет и поддержка ИТ без новой платформы.",
    )

    c.setFillColor(YELLOW_SOFT)
    c.roundRect(ML + 3 * mm, y - 24 * mm, CONTENT_W - 6 * mm, 12.5 * mm, 3, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8.2)
    c.drawString(ML + 5 * mm, y - 16 * mm, f"Ваш менеджер: {MANAGER_NAME}")
    c.setFont("DejaVu", 7)
    c.setFillColor(GRAY)
    c.drawString(ML + 5 * mm, y - 20 * mm, f"{MANAGER_EMAIL}  ·  {MANAGER_PHONE}")
    c.drawString(ML + 5 * mm, y - 23 * mm, "Получатель КП: кадры2  ·  копия для ИТ-отдела по запросу")
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 7.5)
    c.drawRightString(PAGE_W - MR - 5 * mm, y - 16 * mm, "ГК «Форус»")
    c.setFont("DejaVu", 7)
    c.setFillColor(GRAY)
    c.drawRightString(PAGE_W - MR - 5 * mm, y - 20 * mm, "www.forus.ru")
    c.drawRightString(PAGE_W - MR - 5 * mm, y - 23 * mm, "г. Иркутск, ул. Ямская, 1/1")

    footer(c, 2)


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=A4)
    c.setTitle("КП — ТФМ Спецтехника — КЭДО 750")
    c.setAuthor("ГК Форус · Данил Кургузов")
    page1(c)
    c.showPage()
    page2(c)
    c.save()

    for p in [
        ROOT / "КП_ТФМ_Спецтехника_КЭДО_750.pdf",
        ROOT / "КП_Кабинет_сотрудника_750.pdf",
    ]:
        p.write_bytes(OUT.read_bytes())
        print(f"Saved: {p}")
    print(f"Total to pay: {TOTAL} RUB (cabinets only); gift hours: {HOURS_GIFT}")


if __name__ == "__main__":
    build()
