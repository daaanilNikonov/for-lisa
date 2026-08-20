#!/usr/bin/env python3
"""КП PDF для ТФМ Спецтехника — КЭДО / 1С:Кабинет сотрудника, 750 сотрудников.

Формат для тендера: условия, сроки, бюджет. Акцент на ИТ-отдел.
Получатель: кадры2.
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

# Бюджет
CABINETS = 223_200  # 500+200+50
HOURS_PAID = 4
HOUR_RATE = 3_420  # от 4 часов
HOURS_GIFT = 5
HOUR_VALUE = 3_660
PAID_HOURS_SUM = HOURS_PAID * HOUR_RATE  # 13680
GIFT_VALUE = HOURS_GIFT * HOUR_VALUE  # 18300
TOTAL = CABINETS + PAID_HOURS_SUM  # 236880


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
        "+7 (3952) 78-00-00  ·  www.forus.ru",
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


def simple_table(c, y, headers, rows, col_ws, row_h=7 * mm, head_h=7 * mm, highlight_gift_row=None):
    """Draw table; returns new y (bottom)."""
    # header
    x = ML
    c.setFillColor(DARK)
    c.rect(ML, y - head_h, CONTENT_W, head_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("DejaVuBold", 7)
    xx = ML
    for i, h in enumerate(headers):
        c.drawString(xx + 1.5 * mm, y - head_h + 2.2 * mm, h)
        xx += col_ws[i]
    y -= head_h

    for ri, row in enumerate(rows):
        yy = y - row_h
        bg = YELLOW_SOFT if highlight_gift_row is not None and ri == highlight_gift_row else (LIGHT if ri % 2 == 0 else white)
        c.setFillColor(bg)
        c.rect(ML, yy, CONTENT_W, row_h, fill=1, stroke=0)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.3)
        c.rect(ML, yy, CONTENT_W, row_h, fill=0, stroke=1)
        xx = ML
        for ci, cell in enumerate(row):
            if ci == len(row) - 1:
                c.setFont("DejaVuBold", 7)
                c.setFillColor(GOOD_TEXT if highlight_gift_row == ri else DARK)
                c.drawRightString(xx + col_ws[ci] - 1.5 * mm, yy + 2.3 * mm, cell)
            else:
                c.setFont("DejaVuBold" if ci == 0 else "DejaVu", 6.8)
                c.setFillColor(DARK)
                # truncate
                max_w = col_ws[ci] - 3 * mm
                t = cell
                while c.stringWidth(t, "DejaVu", 6.8) > max_w and len(t) > 3:
                    t = t[:-2]
                if t != cell:
                    t = t[:-1] + "…"
                c.drawString(xx + 1.5 * mm, yy + 2.3 * mm, t)
            xx += col_ws[ci]
        y = yy
    return y


def page1(c):
    header(c)
    y = PAGE_H - MT - 20 * mm

    c.setFont("DejaVuBold", 15)
    c.setFillColor(DARK)
    c.drawCentredString(PAGE_W / 2, y, "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ")
    y -= 5.5 * mm

    # client badge
    round_rect(c, ML, y - 8 * mm, CONTENT_W, 10 * mm, r=4, fill=DARK)
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 9.5)
    c.drawCentredString(PAGE_W / 2, y - 3.5 * mm, "ТФМ Спецтехника  ·  кадровый электронный документооборот")
    c.setFillColor(white)
    c.setFont("DejaVu", 7)
    c.drawCentredString(PAGE_W / 2, y - 7 * mm, "750 сотрудников  ·  решение для тендера  ·  для кадровой службы и ИТ-отдела")
    y -= 13 * mm

    c.setFont("DejaVu", 8)
    c.setFillColor(DARK)
    y = para(
        c,
        "Уважаемые коллеги! ГК «Форус» направляет коммерческое предложение по внедрению "
        "кадрового электронного документооборота на базе сервиса «1С:Кабинет сотрудника». "
        "Предложение подготовлено для участия в выборе подрядчика: ниже — условия, сроки и бюджет.",
        ML, y, CONTENT_W, size=8, leading=10.5,
    )
    y -= 2 * mm

    # IT-focused value
    y = section_title(c, "Почему это удобно ИТ-отделу", y)
    it_points = [
        "Сервис встроен в 1С: не нужна отдельная HR-платформа, лишние интеграции и новые контуры доступа.",
        "Сотрудники не получают прямой доступ к информационной базе 1С — работают через личный кабинет.",
        "Не требуются дополнительные клиентские лицензии 1С на каждого сотрудника.",
        "Документы хранятся в вашей базе 1С (есть облачный и локальный сценарий) — проще контроль и аудит.",
        "Дата-центр 1С аттестован ФСТЭК; усиленная неквалифицированная подпись сотрудникам — бесплатно.",
    ]
    for p in it_points:
        y = bullet(c, p, ML, y, CONTENT_W, size=7.5, leading=9.5)
        y -= 0.5 * mm
    y -= 2 * mm

    # Capabilities compact
    y = section_title(c, "Что получает компания", y)
    caps = [
        ("Ознакомление в 1 клик", "Документы сотрудникам без печати и сбора подписей"),
        ("Удалённый приём", "Трудовой договор и заявления без визита в офис"),
        ("Отпуска без звонков", "Остаток, график и заявление — в личном кабинете"),
        ("Согласование online", "Руководитель утверждает с телефона, данные в 1С"),
        ("Без мессенджеров", "Обмен внутри системы, с учётом персональных данных"),
        ("Меньше рутины", "Заявления автоматически становятся документами 1С"),
    ]
    cw = (CONTENT_W - 4 * mm) / 3
    for i in range(0, 6, 3):
        for col in range(3):
            title, body = caps[i + col]
            x = ML + col * (cw + 2 * mm)
            round_rect(c, x, y - 16 * mm, cw, 17 * mm, r=3, fill=LIGHT)
            c.setFillColor(YELLOW)
            c.rect(x, y - 16 * mm, 1.5 * mm, 17 * mm, fill=1, stroke=0)
            c.setFillColor(DARK)
            c.setFont("DejaVuBold", 7.2)
            c.drawString(x + 3.5 * mm, y - 4 * mm, title)
            para(c, body, x + 3.5 * mm, y - 8 * mm, cw - 6 * mm, size=6.6, leading=8.5, color=GRAY)
        y -= 18.5 * mm

    y -= 1 * mm
    y = section_title(c, "1. Условия предложения", y)

    conditions = [
        "Объект: подключение «1С:Кабинет сотрудника» на 750 личных кабинетов сроком 12 месяцев.",
        "Состав лицензий: пакеты 500 + 200 + 50 кабинетов (официальная линейка 1С).",
        "Включено в запуск: помощь в подключении, настройка ролей и правил, выпуск подписей, обучение по видеоурокам и инструкциям.",
        "Акция «Больше, чем кешбэк»: при оплате пакета начисляем 5 часов линии консультаций бесплатно (выгода 18 300 ₽).",
        "Дополнительно оплачиваются 4 часа линии консультаций для сопровождения внедрения (тариф от 4 часов — 3 420 ₽/час).",
        "Шаблоны документов для перехода на кадровый электронный документооборот (положение, согласия) — предоставляем.",
        "Срок действия коммерческого предложения: 30 календарных дней с даты направления.",
        "Работы выполняются удалённо; выезд и нетиповые доработки печатных форм — по отдельному согласованию.",
    ]
    for p in conditions:
        y = bullet(c, p, ML, y, CONTENT_W, size=7.3, leading=9.3)
    footer(c, 1)


def page2(c):
    header(c)
    y = PAGE_H - MT - 20 * mm

    y = section_title(c, "2. Сроки внедрения", y)
    c.setFont("DejaVu", 7.2)
    c.setFillColor(GRAY)
    c.drawString(ML, y, "Ориентир для планирования тендера. Календарный план уточняется после согласования дорожной карты.")
    y -= 4 * mm

    # timeline table
    headers = ["Этап", "Срок", "Результат"]
    col_ws = [55 * mm, 35 * mm, CONTENT_W - 90 * mm]
    rows = [
        ["Согласование условий и договор", "3–5 рабочих дней", "Зафиксированы объём, бюджет, ответственные"],
        ["Подключение сервиса к 1С", "1–2 рабочих дня", "Сервис активирован, базовые настройки"],
        ["Настройка ролей, процессов, форм", "3–5 рабочих дней", "Готов контур кадровой службы и руководителей"],
        ["Выпуск подписей, пилот 30–50 чел.", "5–10 рабочих дней", "Пилотная группа работает в кабинете"],
        ["Обучение ключевых пользователей", "параллельно", "Видеоуроки + короткие инструкции"],
        ["Тиражирование на 750 сотрудников", "2–4 недели", "Массовое подключение и обмен документами"],
        ["Сопровождение на линии консультаций", "в рамках 9 часов", "Оперативные ответы и помощь по запуску"],
    ]
    # custom taller rows for timeline
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
        rh = 7.2 * mm
        yy = y - rh
        c.setFillColor(LIGHT if ri % 2 == 0 else white)
        c.rect(ML, yy, CONTENT_W, rh, fill=1, stroke=0)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.3)
        c.rect(ML, yy, CONTENT_W, rh, fill=0, stroke=1)
        xx = ML
        fonts = [("DejaVuBold", 6.8), ("DejaVuBold", 6.8), ("DejaVu", 6.6)]
        colors = [DARK, HexColor("#8A6A1A"), GRAY]
        for ci, cell in enumerate(row):
            c.setFont(*fonts[ci])
            c.setFillColor(colors[ci])
            c.drawString(xx + 1.5 * mm, yy + 2.4 * mm, cell[:48])
            xx += col_ws[ci]
        y = yy

    y -= 3.5 * mm
    round_rect(c, ML, y - 8 * mm, CONTENT_W, 9 * mm, r=3, fill=YELLOW_SOFT, stroke=YELLOW, sw=1)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 7.5)
    c.drawString(ML + 3 * mm, y - 3.5 * mm, "Итоговый ориентир запуска:")
    c.setFont("DejaVu", 7.5)
    c.drawString(ML + 48 * mm, y - 3.5 * mm, "пилот за 2–3 недели, полный охват 750 сотрудников — в течение 1–1,5 месяца.")
    c.setFont("DejaVu", 6.5)
    c.setFillColor(GRAY)
    c.drawString(ML + 3 * mm, y - 6.5 * mm, "Внедрение не ломает привычную работу в 1С. Срок зависит от скорости предоставления доступов и готовности пилотной группы.")
    y -= 12 * mm

    # Budget
    y = section_title(c, "3. Бюджет", y)

    headers = ["Статья бюджета", "Состав / пояснение", "Сумма, ₽"]
    col_ws = [70 * mm, 75 * mm, CONTENT_W - 145 * mm]
    rows = [
        ["1С:Кабинет сотрудника, 750 кабинетов / 12 мес.", "Пакеты 500 + 200 + 50", f"{CABINETS:,}".replace(",", " ")],
        ["Линия консультаций — 4 часа (оплата)", "Тариф от 4 часов: 3 420 ₽/час", f"{PAID_HOURS_SUM:,}".replace(",", " ")],
        ["ПОДАРОК: 5 часов линии консультаций", "Акция «Больше, чем кешбэк!»", "0 (выгода 18 300)"],
    ]
    y = simple_table(c, y, headers, rows, col_ws, row_h=7 * mm, head_h=6.5 * mm, highlight_gift_row=2)
    y -= 3 * mm

    round_rect(c, ML, y - 20 * mm, CONTENT_W, 21 * mm, r=5, fill=DARK)
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 9)
    c.drawString(ML + 4 * mm, y - 5 * mm, "ИТОГО БЮДЖЕТ К ОПЛАТЕ")
    c.setFont("DejaVuBold", 15)
    c.drawRightString(PAGE_W - MR - 4 * mm, y - 5.5 * mm, f"{TOTAL:,}".replace(",", " ") + " ₽")
    c.setFillColor(white)
    c.setFont("DejaVu", 7.2)
    c.drawString(ML + 4 * mm, y - 10 * mm, "Без НДС по тарифам 1С (уточняется в счёте). Цены рекомендованные розничные.")
    c.drawString(ML + 4 * mm, y - 14 * mm, f"Вы получаете 9 часов поддержки (4 оплаченных + 5 подарочных) на сумму {(PAID_HOURS_SUM + GIFT_VALUE):,}".replace(",", " ") + " ₽")
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 7.5)
    c.drawString(ML + 4 * mm, y - 18 * mm, "≈ 26 ₽ в месяц за сотрудника  ·  электронная подпись сотрудникам — бесплатно")
    y -= 24 * mm

    # Comparison compact for tender
    y = section_title(c, "Сравнение для тендера: Форус vs другие операторы", y)
    headers = ["Критерий", "Форус + 1С:Кабинет", "Другие операторы"]
    col_ws = [48 * mm, 70 * mm, CONTENT_W - 118 * mm]
    # manual 2-col competitor highlight
    head_h = 6.5 * mm
    c.setFillColor(DARK)
    c.rect(ML, y - head_h, CONTENT_W, head_h, fill=1, stroke=0)
    c.setFillColor(YELLOW)
    c.rect(ML + col_ws[0], y - head_h, col_ws[1], head_h, fill=1, stroke=0)
    c.setFont("DejaVuBold", 7)
    xx = ML
    for i, h in enumerate(headers):
        c.setFillColor(DARK if i == 1 else white)
        c.drawCentredString(xx + col_ws[i] / 2, y - head_h + 2.1 * mm, h)
        xx += col_ws[i]
    y -= head_h

    cmp_rows = [
        ["Платформа", "Работа внутри вашей 1С", "Отдельная HR-система"],
        ["Интеграция", "Без двойного ввода данных", "Часто перенос и сверка с 1С"],
        ["Хранение", "В вашей базе 1С", "Облако оператора, часто платно"],
        ["Подписи сотрудников", "Бесплатно", "Часто платные пакеты"],
        ["Нагрузка на ИТ", "Без новых клиентских лицензий 1С", "Новая система + интеграции"],
        ["Срок старта", "Пилот за 2–3 недели", "Долгая интеграция с 1С"],
    ]
    for ri, row in enumerate(cmp_rows):
        rh = 6.8 * mm
        yy = y - rh
        xx = ML
        for ci, cell in enumerate(row):
            if ci == 1:
                bg = GOOD
                tc = GOOD_TEXT
                f = "DejaVuBold"
            elif ci == 2:
                bg = BAD
                tc = BAD_TEXT
                f = "DejaVu"
            else:
                bg = LIGHT if ri % 2 == 0 else white
                tc = DARK
                f = "DejaVuBold"
            c.setFillColor(bg)
            c.rect(xx, yy, col_ws[ci], rh, fill=1, stroke=0)
            c.setStrokeColor(LINE)
            c.setLineWidth(0.3)
            c.rect(xx, yy, col_ws[ci], rh, fill=0, stroke=1)
            c.setFillColor(tc)
            c.setFont(f, 6.4)
            if ci == 0:
                c.drawString(xx + 1.5 * mm, yy + 2.2 * mm, cell)
            else:
                c.drawCentredString(xx + col_ws[ci] / 2, yy + 2.2 * mm, cell)
            xx += col_ws[ci]
        y = yy

    y -= 4 * mm
    # CTA + manager for kadry2
    round_rect(c, ML, y - 28 * mm, CONTENT_W, 29 * mm, r=5, fill=DARK)
    c.setFillColor(YELLOW)
    c.setFont("DejaVuBold", 9.5)
    c.drawCentredString(PAGE_W / 2, y - 5 * mm, "Готовы войти в тендер ТФМ Спецтехника с прозрачными условиями")
    c.setFillColor(white)
    c.setFont("DejaVu", 7.3)
    c.drawCentredString(PAGE_W / 2, y - 9.5 * mm, "Направим счёт, дорожную карту и проведём демонстрацию для кадровой службы и ИТ-отдела.")

    c.setFillColor(YELLOW_SOFT)
    c.roundRect(ML + 3 * mm, y - 26 * mm, CONTENT_W - 6 * mm, 14 * mm, 3, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8)
    c.drawString(ML + 5 * mm, y - 17 * mm, "Ваш менеджер: Оглоблина Софья")
    c.setFont("DejaVu", 7)
    c.setFillColor(GRAY)
    c.drawString(ML + 5 * mm, y - 21 * mm, "sogloblina@forus.ru  ·  +7 (3952) 78-00-00, доб. 1861")
    c.drawString(ML + 5 * mm, y - 24.5 * mm, "Получатель КП: кадры2  ·  копия для ИТ-отдела по запросу")
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 7.5)
    c.drawRightString(PAGE_W - MR - 5 * mm, y - 17 * mm, "ГК «Форус»")
    c.setFont("DejaVu", 7)
    c.setFillColor(GRAY)
    c.drawRightString(PAGE_W - MR - 5 * mm, y - 21 * mm, "www.forus.ru")
    c.drawRightString(PAGE_W - MR - 5 * mm, y - 24.5 * mm, "г. Иркутск, ул. Ямская, 1/1")

    footer(c, 2)


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT), pagesize=A4)
    c.setTitle("КП — ТФМ Спецтехника — КЭДО 750")
    c.setAuthor("ГК Форус")
    page1(c)
    c.showPage()
    page2(c)
    c.save()

    copies = [
        ROOT / "КП_ТФМ_Спецтехника_КЭДО_750.pdf",
        ROOT / "КП_Кабинет_сотрудника_750.pdf",  # keep alias updated for client pack
    ]
    data = OUT.read_bytes()
    for p in copies:
        p.write_bytes(data)
        print(f"Saved: {p}")
    print(f"Total budget: {TOTAL} RUB")


if __name__ == "__main__":
    build()
