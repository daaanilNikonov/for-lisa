#!/usr/bin/env python3
"""Листовка КП: 1С:Кабинет сотрудника — 750 кабинетов.

Жёлтый фирменный стиль Форус, без повторов, понятное ценообразование.
PDF + DOCX. Менеджер: Данил Кургузов.
5 часов в подарок (≈4 настройка + 1 вопросы). Клиент за часы не платит.
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib.colors import HexColor, white, Color
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parents[1]
BRAND = ROOT / "assets_forus" / "brand"
OUT_PDF = ROOT / "output" / "КП_ТФМ_Спецтехника_КЭДО_750.pdf"
OUT_DOCX = ROOT / "output" / "КП_ТФМ_Спецтехника_КЭДО_750.docx"

# Палитра Форус — больше жёлтого, меньше чёрного
Y = HexColor("#F0C14A")
Y2 = HexColor("#FFE08A")
YSOFT = HexColor("#FFF6D8")
YCARD = HexColor("#FFFBEA")
DARK = HexColor("#2B2B2B")
GRAY = HexColor("#666666")
MUTED = HexColor("#8A8A8A")
LINE = HexColor("#F0C14A")
WHITE = white
OK = HexColor("#2E7D32")
OKBG = HexColor("#E8F5E9")
BAD = HexColor("#C62828")
BADBG = HexColor("#FFEBEE")

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

W, H = A4
L, R = 13 * mm, 13 * mm
CW = W - L - R

CABINETS = 223_200
GIFT_VALUE = 18_300
PER_MONTH = 25  # approx

MANAGER = "Данил Кургузов"
EMAIL = "dkurguzov@forus.ru"
PHONE = "+7 (3952) 78-00-00"


def wrap(c, text, font, size, max_w):
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = f"{cur} {w}".strip()
        if c.stringWidth(t, font, size) <= max_w:
            cur = t
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def draw_waves(c):
    tr = BRAND / "wave_tr.png"
    bl = BRAND / "wave_bl.png"
    if tr.exists():
        c.drawImage(str(tr), W - 55 * mm, H - 28 * mm, width=55 * mm, height=18 * mm, mask="auto")
    if bl.exists():
        c.drawImage(str(bl), 0, 0, width=50 * mm, height=16 * mm, mask="auto")


def header(c):
    draw_waves(c)
    logo = BRAND / "forus_logo_word.png"
    if logo.exists():
        c.drawImage(str(logo), L, H - 18 * mm, width=32 * mm, height=11 * mm, mask="auto", preserveAspectRatio=True)
    c.setFillColor(GRAY)
    c.setFont("DejaVu", 7)
    c.drawRightString(W - R, H - 8 * mm, "Группа компаний «Форус»")
    c.drawRightString(W - R, H - 12 * mm, f"{PHONE}  ·  www.forus.ru")
    # yellow line
    c.setFillColor(Y)
    c.roundRect(L, H - 21 * mm, CW, 2.8, 1.4, fill=1, stroke=0)


def footer(c, n):
    c.setFillColor(Y)
    c.roundRect(L, 8 * mm, CW, 2, 1, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.setFont("DejaVu", 6.5)
    c.drawString(L, 4.5 * mm, "www.forus.ru  ·  г. Иркутск, ул. Ямская, 1/1")
    c.drawRightString(W - R, 4.5 * mm, f"{n} / 2")


def yellow_chip(c, x, y, text, size=7.5):
    c.setFont("DejaVuBold", size)
    tw = c.stringWidth(text, "DejaVuBold", size) + 8 * mm
    c.setFillColor(Y)
    c.roundRect(x, y - 1.5 * mm, tw, 6.5 * mm, 3.2, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.drawString(x + 4 * mm, y, text)
    return tw


def callout(c, x, y, w, h, title, body, fill=YSOFT):
    """Eye-catching info block with yellow left bar."""
    c.setFillColor(fill)
    c.setStrokeColor(Y)
    c.setLineWidth(1.4)
    c.roundRect(x, y - h, w, h, 6, fill=1, stroke=1)
    c.setFillColor(Y)
    c.roundRect(x, y - h, 3.2 * mm, h, 1.5, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8.5)
    c.drawString(x + 6 * mm, y - 5.5 * mm, title)
    c.setFont("DejaVu", 7.4)
    c.setFillColor(GRAY)
    yy = y - 10.5 * mm
    for line in wrap(c, body, "DejaVu", 7.4, w - 10 * mm):
        c.drawString(x + 6 * mm, yy, line)
        yy -= 9.5
    return y - h


def num_circle(c, x, y, num):
    c.setFillColor(Y)
    c.circle(x, y, 6.5, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 9)
    c.drawCentredString(x, y - 3, str(num))


def build_pdf():
    OUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT_PDF), pagesize=A4)
    c.setTitle("1С:Кабинет сотрудника — коммерческое предложение")
    c.setAuthor(f"ГК Форус · {MANAGER}")

    # ========== PAGE 1 ==========
    header(c)
    y = H - 28 * mm

    # Title block
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 20)
    c.drawString(L, y, "1С:Кабинет сотрудника")
    y -= 7 * mm
    c.setFont("DejaVu", 10)
    c.setFillColor(GRAY)
    c.drawString(L, y, "Коммерческое предложение  ·  750 личных кабинетов")
    y -= 5 * mm
    yellow_chip(c, L, y, "самое простое внедрение  ·  без отдельной платформы")
    y -= 10 * mm

    # Big value callout
    y = callout(
        c, L, y, CW, 22 * mm,
        "Выгода, которую сразу видно",
        "Вы платите только за личные кабинеты. Настройка и поддержка на старте — в подарок. "
        "Сервис уже внутри 1С: не нужно покупать и внедрять отдельную кадровую систему.",
        fill=YSOFT,
    )
    y -= 5 * mm

    # Two columns: benefits + pricing tease
    left_w = CW * 0.55
    right_w = CW * 0.42
    gap = CW - left_w - right_w

    # Left: what changes for the company
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 11)
    c.drawString(L, y, "Что меняется в работе")
    c.setFillColor(Y)
    c.roundRect(L, y - 2.2 * mm, 42 * mm, 2, 1, fill=1, stroke=0)
    y_left = y - 7 * mm

    benefits = [
        ("Расчётные листки и справки — в телефоне",
         "Сотрудник сам открывает документы в личном кабинете. Меньше очередей у кадровика."),
        ("Заявления и согласования online",
         "Отпуск, отгул, командировка: сотрудник подал — руководитель утвердил с телефона — данные в 1С."),
        ("Удалённый приём без визита в офис",
         "Трудовой договор и ознакомление с правилами подписываются дистанционно."),
        ("Документы без бумаги и мессенджеров",
         "Ознакомление в один клик, подтверждение получения, обмен внутри системы — по закону о персональных данных."),
        ("Кадровик остаётся в привычной 1С",
         "Заявления автоматически становятся приказами. Меньше ручного ввода и ошибок."),
    ]
    for title, body in benefits:
        # mini yellow card
        h = 17 * mm
        c.setFillColor(YCARD)
        c.setStrokeColor(Y2)
        c.setLineWidth(0.9)
        c.roundRect(L, y_left - h, left_w, h, 4, fill=1, stroke=1)
        c.setFillColor(Y)
        c.circle(L + 4.5 * mm, y_left - 5 * mm, 2.2, fill=1, stroke=0)
        c.setFillColor(DARK)
        c.setFont("DejaVuBold", 7.6)
        c.drawString(L + 9 * mm, y_left - 4.5 * mm, title)
        c.setFont("DejaVu", 6.8)
        c.setFillColor(GRAY)
        yy = y_left - 9 * mm
        for line in wrap(c, body, "DejaVu", 6.8, left_w - 12 * mm):
            c.drawString(L + 9 * mm, yy, line)
            yy -= 8.2
        y_left -= h + 2.2 * mm

    # Right column sticky pricing
    rx = L + left_w + gap
    ry = y
    c.setFillColor(Y)
    c.roundRect(rx, ry - 78 * mm, right_w, 78 * mm, 8, fill=1, stroke=0)
    # white inner
    c.setFillColor(WHITE)
    c.roundRect(rx + 2.5 * mm, ry - 75.5 * mm, right_w - 5 * mm, 62 * mm, 6, fill=1, stroke=0)

    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 9)
    c.drawCentredString(rx + right_w / 2, ry - 6 * mm, "Цена пакета")
    c.setFont("DejaVuBold", 18)
    c.drawCentredString(rx + right_w / 2, ry - 18 * mm, f"{CABINETS:,}".replace(",", " ") + " ₽")
    c.setFont("DejaVu", 7.5)
    c.setFillColor(GRAY)
    c.drawCentredString(rx + right_w / 2, ry - 24 * mm, "за 750 кабинетов / 12 месяцев")
    c.drawCentredString(rx + right_w / 2, ry - 29 * mm, f"≈ {PER_MONTH} ₽ в месяц за сотрудника")

    # divider
    c.setStrokeColor(Y)
    c.setLineWidth(1)
    c.line(rx + 6 * mm, ry - 33 * mm, rx + right_w - 6 * mm, ry - 33 * mm)

    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8)
    c.drawCentredString(rx + right_w / 2, ry - 39 * mm, "В подарок — 5 часов поддержки")
    c.setFont("DejaVu", 7)
    c.setFillColor(GRAY)
    for i, line in enumerate([
        "выгода 18 300 ₽",
        "≈4 часа — настройка и запуск",
        "1 час — ваши вопросы по 1С",
        "за часы вы не платите",
    ]):
        c.drawCentredString(rx + right_w / 2, ry - 45 * mm - i * 4.2 * mm, line)

    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8)
    c.drawCentredString(rx + right_w / 2, ry - 72 * mm, "Итого к оплате = цена пакета")

    y = min(y_left, ry - 82 * mm) - 4 * mm

    # Одна сноска-акцент без повтора цены
    c.setFillColor(Y)
    c.roundRect(L, y - 12 * mm, CW, 12 * mm, 6, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8.5)
    c.drawString(L + 4 * mm, y - 5 * mm, "Сноска для ИТ-отдела")
    c.setFont("DejaVu", 7.5)
    c.drawString(
        L + 4 * mm, y - 9.5 * mm,
        "Сотрудники не получают прямой доступ к базе 1С. Отдельные клиентские лицензии 1С на каждого сотрудника не нужны.",
    )

    footer(c, 1)

    # ========== PAGE 2 ==========
    c.showPage()
    header(c)
    y = H - 28 * mm

    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 14)
    c.drawString(L, y, "Почему это легче и выгоднее других решений")
    c.setFillColor(Y)
    c.roundRect(L, y - 2.2 * mm, 55 * mm, 2, 1, fill=1, stroke=0)
    y -= 8 * mm

    # 3 yellow argument cards
    cards = [
        ("01", "Без новой платформы",
         "Сотрудники и кадры работают через 1С и личный кабинет. ИТ-отделу не нужно поднимать отдельную систему и долгую интеграцию."),
        ("02", "Запуск за часы, не месяцы",
         "Типовая настройка занимает около одного рабочего дня. Пилот — за 1–2 недели, все 750 кабинетов — за 1–1,5 месяца."),
        ("03", "Прозрачная цена",
         "В счёте — только пакет кабинетов. Часы на настройку и стартовые вопросы уже внутри подарка по акции."),
    ]
    cw = (CW - 6 * mm) / 3
    for i, (num, title, body) in enumerate(cards):
        x = L + i * (cw + 3 * mm)
        c.setFillColor(YCARD)
        c.setStrokeColor(Y)
        c.setLineWidth(1.2)
        c.roundRect(x, y - 42 * mm, cw, 42 * mm, 7, fill=1, stroke=1)
        num_circle(c, x + 8 * mm, y - 8 * mm, num)
        c.setFillColor(DARK)
        c.setFont("DejaVuBold", 8)
        c.drawString(x + 15 * mm, y - 9.5 * mm, title)
        c.setFont("DejaVu", 6.9)
        c.setFillColor(GRAY)
        yy = y - 16 * mm
        for line in wrap(c, body, "DejaVu", 6.9, cw - 8 * mm):
            c.drawString(x + 4 * mm, yy, line)
            yy -= 9
    y -= 48 * mm

    # Comparison - yellow header not black
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 11)
    c.drawString(L, y, "Сравнение для выбора подрядчика")
    c.setFillColor(Y)
    c.roundRect(L, y - 2.2 * mm, 48 * mm, 2, 1, fill=1, stroke=0)
    y -= 7 * mm

    cols = [40 * mm, 72 * mm, CW - 112 * mm]
    rows = [
        ("Запуск", "Около 4 часов настройки", "Часто долгое внедрение"),
        ("Система", "Внутри вашей 1С", "Отдельная кадровая платформа"),
        ("ИТ-нагрузка", "Без новых клиентских лицензий 1С", "Новые доступы и интеграции"),
        ("Документы", "В вашей базе 1С", "Облако оператора, часто платно"),
        ("Оплата старта", "Часы поддержки — подарок", "Внедрение обычно отдельной строкой"),
    ]
    hh = 7 * mm
    # yellow header
    c.setFillColor(Y)
    c.roundRect(L, y - hh, CW, hh, 3, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 7.5)
    xs = [L, L + cols[0], L + cols[0] + cols[1]]
    for i, htxt in enumerate(["Критерий", "Предложение Форус", "Типичные альтернативы"]):
        c.drawCentredString(xs[i] + cols[i] / 2, y - hh + 2.2 * mm, htxt)
    y -= hh
    for ri, (a, b, d) in enumerate(rows):
        rh = 7.2 * mm
        for ci, (val, w) in enumerate(zip((a, b, d), cols)):
            x = xs[ci]
            if ci == 1:
                bg, tc = OKBG, OK
            elif ci == 2:
                bg, tc = BADBG, BAD
            else:
                bg, tc = (YSOFT if ri % 2 == 0 else WHITE), DARK
            c.setFillColor(bg)
            c.rect(x, y - rh, w, rh, fill=1, stroke=0)
            c.setStrokeColor(Y2)
            c.setLineWidth(0.5)
            c.rect(x, y - rh, w, rh, fill=0, stroke=1)
            c.setFillColor(tc)
            c.setFont("DejaVuBold" if ci == 0 else "DejaVu", 6.6)
            if ci == 0:
                c.drawString(x + 2 * mm, y - rh + 2.4 * mm, val)
            else:
                c.drawCentredString(x + w / 2, y - rh + 2.4 * mm, val)
        y -= rh

    y -= 6 * mm

    # Timeline as yellow steps (not dense table)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 11)
    c.drawString(L, y, "Как проходит запуск")
    c.setFillColor(Y)
    c.roundRect(L, y - 2.2 * mm, 36 * mm, 2, 1, fill=1, stroke=0)
    y -= 8 * mm

    steps = [
        ("1", "Договор\nи доступы", "2–4 дня"),
        ("2", "Настройка\n≈4 часа", "1 день"),
        ("3", "Пилот\nна группе", "до 1 недели"),
        ("4", "Все\n750 кабинетов", "2–3 недели"),
        ("5", "1 час\nв запасе", "по запросу"),
    ]
    sw = CW / 5
    for i, (n, title, timing) in enumerate(steps):
        x = L + i * sw + sw / 2
        num_circle(c, x, y - 2 * mm, n)
        c.setFillColor(DARK)
        c.setFont("DejaVuBold", 7)
        lines = title.split("\n")
        for li, line in enumerate(lines):
            c.drawCentredString(x, y - 11 * mm - li * 3.2 * mm, line)
        c.setFillColor(HexColor("#9A7A10"))
        c.setFont("DejaVu", 6.5)
        c.drawCentredString(x, y - 20 * mm, timing)
        if i < len(steps) - 1:
            c.setStrokeColor(Y)
            c.setLineWidth(2)
            c.line(x + 8 * mm, y - 2 * mm, x + sw - 8 * mm, y - 2 * mm)
    y -= 26 * mm

    # Pricing breakdown - crystal clear once
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 11)
    c.drawString(L, y, "Ценообразование — одной таблицей")
    c.setFillColor(Y)
    c.roundRect(L, y - 2.2 * mm, 48 * mm, 2, 1, fill=1, stroke=0)
    y -= 7 * mm

    price_rows = [
        ("750 личных кабинетов на 12 месяцев (пакеты 500+200+50)", f"{CABINETS:,} ₽".replace(",", " "), False),
        ("Настройка и запуск сервиса (≈4 часа поддержки)", "0 ₽ — из подарка", True),
        ("Запас на вопросы по 1С после запуска (1 час)", "0 ₽ — из подарка", True),
        ("Подарок по акции «Больше, чем кешбэк!» — 5 часов", f"выгода {GIFT_VALUE:,} ₽".replace(",", " "), True),
    ]
    for label, amount, gift in price_rows:
        rh = 8 * mm
        c.setFillColor(YSOFT if gift else YCARD)
        c.setStrokeColor(Y)
        c.setLineWidth(0.9)
        c.roundRect(L, y - rh, CW, rh, 3, fill=1, stroke=1)
        c.setFillColor(DARK)
        c.setFont("DejaVu", 7.4)
        c.drawString(L + 3 * mm, y - rh + 2.8 * mm, label)
        c.setFont("DejaVuBold", 8)
        c.setFillColor(OK if gift else DARK)
        c.drawRightString(W - R - 3 * mm, y - rh + 2.8 * mm, amount)
        y -= rh + 1.5 * mm

    y -= 2 * mm
    # Total - yellow not black
    c.setFillColor(Y)
    c.roundRect(L, y - 16 * mm, CW, 16 * mm, 7, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 11)
    c.drawString(L + 4 * mm, y - 6.5 * mm, "Итого к оплате")
    c.setFont("DejaVuBold", 16)
    c.drawRightString(W - R - 4 * mm, y - 7 * mm, f"{CABINETS:,} ₽".replace(",", " "))
    c.setFont("DejaVu", 7.5)
    c.drawString(L + 4 * mm, y - 12.5 * mm, "Только пакет кабинетов. Настройка и стартовая поддержка — бесплатно. Подпись сотрудникам — бесплатно.")
    y -= 20 * mm

    # Важно знать — 2×2 жёлтые сноски (без повтора цены)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 10)
    c.drawString(L, y, "Важно знать")
    c.setFillColor(Y)
    c.roundRect(L, y - 2 * mm, 24 * mm, 2, 1, fill=1, stroke=0)
    y -= 6 * mm

    notes = [
        ("Подпись бесплатно", "Электронная подпись сотрудникам выпускается бесплатно."),
        ("Шаблоны документов", "Положение, уведомления и согласия для перехода — в комплекте."),
        ("Формат работ", "Удалённо. Выезд и нетиповые доработки — отдельно."),
        ("Срок предложения", "Действует 30 дней с даты отправки."),
    ]
    nw = (CW - 3 * mm) / 2
    nh = 14 * mm
    for i, (title, body) in enumerate(notes):
        col = i % 2
        row = i // 2
        x = L + col * (nw + 3 * mm)
        yy = y - row * (nh + 2.5 * mm)
        c.setFillColor(YCARD)
        c.setStrokeColor(Y)
        c.setLineWidth(1)
        c.roundRect(x, yy - nh, nw, nh, 5, fill=1, stroke=1)
        c.setFillColor(Y)
        c.circle(x + 4 * mm, yy - 4.5 * mm, 2, fill=1, stroke=0)
        c.setFillColor(DARK)
        c.setFont("DejaVuBold", 7.5)
        c.drawString(x + 8 * mm, yy - 5 * mm, title)
        c.setFont("DejaVu", 6.8)
        c.setFillColor(GRAY)
        c.drawString(x + 4 * mm, yy - 10.5 * mm, body[:52])

    # Manager — жёстко у низа страницы
    c.setFillColor(Y)
    c.roundRect(L, 14 * mm, CW, 20 * mm, 7, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 10)
    c.drawString(L + 4 * mm, 27 * mm, f"Ваш менеджер: {MANAGER}")
    c.setFont("DejaVu", 8)
    c.drawString(L + 4 * mm, 21 * mm, f"{EMAIL}   ·   {PHONE}")
    c.setFont("DejaVuBold", 8)
    c.drawRightString(W - R - 4 * mm, 27 * mm, "Готовы подключить и показать на демо")
    c.setFont("DejaVu", 7.2)
    c.drawRightString(W - R - 4 * mm, 21 * mm, "для кадровой службы и ИТ-отдела")

    footer(c, 2)
    c.save()
    print("PDF OK", OUT_PDF)


# ---------------- DOCX ----------------

def set_run(run, size=10, bold=False, color=None):
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def shade(cell, hexfill):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hexfill)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_borders(table, color="F0C14A", sz="8"):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def clear(cell):
    cell.paragraphs[0].clear()


def p_add(doc, text, *, size=10, bold=False, color=RGBColor(0x2B, 0x2B, 0x2B), after=6, center=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(0)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run(r, size=size, bold=bold, color=color)
    return p


def yellow_title(doc, text):
    p = p_add(doc, text, size=13, bold=True, after=2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "20")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), "F0C14A")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def info_box(doc, title, body, fill="FFF6D8"):
    t = doc.add_table(rows=1, cols=1)
    set_borders(t, "F0C14A", "12")
    cell = t.rows[0].cells[0]
    shade(cell, fill)
    clear(cell)
    r = cell.paragraphs[0].add_run(title)
    set_run(r, size=11, bold=True, color=RGBColor(0x2B, 0x2B, 0x2B))
    p = cell.add_paragraph()
    r = p.add_run(body)
    set_run(r, size=9, color=RGBColor(0x66, 0x66, 0x66))
    doc.add_paragraph().paragraph_format.space_after = Pt(8)


def build_docx():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    for m in ("left_margin", "right_margin"):
        setattr(sec, m, Mm(14))
    sec.top_margin = Mm(12)
    sec.bottom_margin = Mm(12)

    # Header
    ht = doc.add_table(rows=1, cols=2)
    a, b = ht.rows[0].cells
    clear(a)
    logo = BRAND / "forus_logo_word.png"
    if logo.exists():
        a.paragraphs[0].add_run().add_picture(str(logo), width=Cm(3.4))
    clear(b)
    b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for i, line in enumerate(["Группа компаний «Форус»", f"{PHONE}  ·  www.forus.ru"]):
        p = b.paragraphs[0] if i == 0 else b.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(0)
        set_run(p.add_run(line), size=8, bold=(i == 0), color=RGBColor(0x66, 0x66, 0x66))

    rule = doc.add_paragraph()
    rule.paragraph_format.space_after = Pt(10)
    pPr = rule._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "24")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "F0C14A")
    pBdr.append(bottom)
    pPr.append(pBdr)

    p_add(doc, "1С:Кабинет сотрудника", size=20, bold=True, center=True, after=2)
    p_add(doc, "Коммерческое предложение  ·  750 личных кабинетов", size=11, color=RGBColor(0x66, 0x66, 0x66), center=True, after=4)
    p_add(doc, "Самое простое внедрение  ·  без отдельной платформы", size=10, bold=True, color=RGBColor(0x9A, 0x7A, 0x10), center=True, after=10)

    info_box(
        doc,
        "Выгода, которую сразу видно",
        "Вы платите только за личные кабинеты. Настройка и поддержка на старте — в подарок. "
        "Сервис уже внутри 1С: не нужно покупать и внедрять отдельную кадровую систему.",
    )

    yellow_title(doc, "Что меняется в работе")
    for title, body in [
        ("Расчётные листки и справки — в телефоне",
         "Сотрудник сам открывает документы в личном кабинете. Меньше очередей у кадровика."),
        ("Заявления и согласования online",
         "Отпуск, отгул, командировка: сотрудник подал — руководитель утвердил с телефона — данные в 1С."),
        ("Удалённый приём без визита в офис",
         "Трудовой договор и ознакомление с правилами подписываются дистанционно."),
        ("Документы без бумаги и мессенджеров",
         "Ознакомление в один клик, подтверждение получения, обмен внутри системы — по закону о персональных данных."),
        ("Кадровик остаётся в привычной 1С",
         "Заявления автоматически становятся приказами. Меньше ручного ввода и ошибок."),
    ]:
        t = doc.add_table(rows=1, cols=1)
        set_borders(t, "FFE08A", "6")
        cell = t.rows[0].cells[0]
        shade(cell, "FFFBEA")
        clear(cell)
        set_run(cell.paragraphs[0].add_run(title), size=10, bold=True)
        p = cell.add_paragraph()
        set_run(p.add_run(body), size=9, color=RGBColor(0x66, 0x66, 0x66))
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Price block
    yellow_title(doc, "Цена пакета")
    pt = doc.add_table(rows=3, cols=1)
    set_borders(pt, "F0C14A", "14")
    cells = pt.rows
    shade(cells[0].cells[0], "F0C14A")
    clear(cells[0].cells[0])
    cells[0].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(cells[0].cells[0].paragraphs[0].add_run(f"{CABINETS:,} ₽".replace(",", " ")), size=22, bold=True)
    shade(cells[1].cells[0], "FFFFFF")
    clear(cells[1].cells[0])
    cells[1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(cells[1].cells[0].paragraphs[0].add_run("за 750 кабинетов на 12 месяцев  ·  ≈ 25 ₽ в месяц за сотрудника"), size=10, color=RGBColor(0x66, 0x66, 0x66))
    shade(cells[2].cells[0], "FFF6D8")
    clear(cells[2].cells[0])
    cells[2].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(
        cells[2].cells[0].paragraphs[0].add_run(
            "В подарок 5 часов поддержки (выгода 18 300 ₽): ≈4 часа — настройка и запуск, 1 час — вопросы по 1С. За часы вы не платите."
        ),
        size=9, bold=True,
    )

    p_add(doc, "", after=8)
    yellow_title(doc, "Почему это легче и выгоднее")
    for num, title, body in [
        ("01", "Без новой платформы", "Сотрудники и кадры работают через 1С и личный кабинет. ИТ-отделу не нужно поднимать отдельную систему."),
        ("02", "Запуск за часы, не месяцы", "Типовая настройка — около одного рабочего дня. Пилот за 1–2 недели, все 750 кабинетов за 1–1,5 месяца."),
        ("03", "Прозрачная цена", "В счёте — только пакет кабинетов. Часы на настройку и стартовые вопросы уже внутри подарка."),
    ]:
        t = doc.add_table(rows=1, cols=2)
        set_borders(t, "F0C14A", "8")
        n, body_cell = t.rows[0].cells
        shade(n, "F0C14A")
        shade(body_cell, "FFFBEA")
        clear(n)
        n.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(n.paragraphs[0].add_run(num), size=14, bold=True)
        n.width = Cm(1.5)
        clear(body_cell)
        set_run(body_cell.paragraphs[0].add_run(title), size=10, bold=True)
        p = body_cell.add_paragraph()
        set_run(p.add_run(body), size=9, color=RGBColor(0x66, 0x66, 0x66))
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    yellow_title(doc, "Сравнение для выбора подрядчика")
    ct = doc.add_table(rows=1, cols=3)
    set_borders(ct, "F0C14A")
    for i, h in enumerate(["Критерий", "Предложение Форус", "Типичные альтернативы"]):
        cell = ct.rows[0].cells[i]
        shade(cell, "F0C14A")
        clear(cell)
        set_run(cell.paragraphs[0].add_run(h), size=9, bold=True)
    for a, b, d in [
        ("Запуск", "Около 4 часов настройки", "Часто долгое внедрение"),
        ("Система", "Внутри вашей 1С", "Отдельная кадровая платформа"),
        ("ИТ-нагрузка", "Без новых клиентских лицензий 1С", "Новые доступы и интеграции"),
        ("Документы", "В вашей базе 1С", "Облако оператора, часто платно"),
        ("Оплата старта", "Часы поддержки — подарок", "Внедрение обычно отдельной строкой"),
    ]:
        row = ct.add_row().cells
        shade(row[1], "E8F5E9")
        shade(row[2], "FFEBEE")
        for i, val in enumerate((a, b, d)):
            clear(row[i])
            col = RGBColor(0x2E, 0x7D, 0x32) if i == 1 else (RGBColor(0xC6, 0x28, 0x28) if i == 2 else RGBColor(0x2B, 0x2B, 0x2B))
            set_run(row[i].paragraphs[0].add_run(val), size=8, bold=(i == 0), color=col)

    p_add(doc, "", after=8)
    yellow_title(doc, "Как проходит запуск")
    st = doc.add_table(rows=1, cols=5)
    set_borders(st, "F0C14A")
    steps = [
        ("1", "Договор и доступы", "2–4 дня"),
        ("2", "Настройка ≈4 часа", "1 день"),
        ("3", "Пилот на группе", "до 1 недели"),
        ("4", "Все 750 кабинетов", "2–3 недели"),
        ("5", "1 час в запасе", "по запросу"),
    ]
    for i, (n, title, timing) in enumerate(steps):
        cell = st.rows[0].cells[i]
        shade(cell, "FFFBEA")
        clear(cell)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(cell.paragraphs[0].add_run(n), size=12, bold=True, color=RGBColor(0x9A, 0x7A, 0x10))
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(title), size=8, bold=True)
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(timing), size=8, color=RGBColor(0x66, 0x66, 0x66))

    p_add(doc, "", after=8)
    yellow_title(doc, "Ценообразование — одной таблицей")
    bt = doc.add_table(rows=1, cols=2)
    set_borders(bt, "F0C14A")
    for i, h in enumerate(["Статья", "Сумма"]):
        cell = bt.rows[0].cells[i]
        shade(cell, "F0C14A")
        clear(cell)
        set_run(cell.paragraphs[0].add_run(h), size=9, bold=True)
    for label, amount, gift in [
        ("750 личных кабинетов на 12 месяцев (пакеты 500+200+50)", f"{CABINETS:,} ₽".replace(",", " "), False),
        ("Настройка и запуск сервиса (≈4 часа поддержки)", "0 ₽ — из подарка", True),
        ("Запас на вопросы по 1С после запуска (1 час)", "0 ₽ — из подарка", True),
        ("Подарок по акции — 5 часов линии консультаций", f"выгода {GIFT_VALUE:,} ₽".replace(",", " "), True),
    ]:
        row = bt.add_row().cells
        shade(row[0], "FFF6D8" if gift else "FFFBEA")
        shade(row[1], "FFF6D8" if gift else "FFFBEA")
        clear(row[0])
        clear(row[1])
        set_run(row[0].paragraphs[0].add_run(label), size=9)
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_run(row[1].paragraphs[0].add_run(amount), size=10, bold=True, color=RGBColor(0x2E, 0x7D, 0x32) if gift else RGBColor(0x2B, 0x2B, 0x2B))

    tot = doc.add_table(rows=1, cols=2)
    set_borders(tot, "F0C14A", "14")
    a, b = tot.rows[0].cells
    shade(a, "F0C14A")
    shade(b, "F0C14A")
    clear(a)
    clear(b)
    set_run(a.paragraphs[0].add_run("Итого к оплате"), size=12, bold=True)
    p = a.add_paragraph()
    set_run(p.add_run("Только пакет кабинетов. Настройка и стартовая поддержка — бесплатно."), size=8)
    b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run(b.paragraphs[0].add_run(f"{CABINETS:,} ₽".replace(",", " ")), size=18, bold=True)

    p_add(doc, "", after=8)
    yellow_title(doc, "Важно знать")
    notes = [
        ("Подпись бесплатно", "Электронная подпись сотрудникам выпускается бесплатно."),
        ("Шаблоны документов", "Положение, уведомления и согласия для перехода — в комплекте."),
        ("Формат работ", "Удалённо. Выезд и нетиповые доработки — отдельно."),
        ("Срок предложения", "Действует 30 дней с даты отправки."),
    ]
    nt = doc.add_table(rows=2, cols=2)
    set_borders(nt, "F0C14A", "8")
    for i, (title, body) in enumerate(notes):
        cell = nt.rows[i // 2].cells[i % 2]
        shade(cell, "FFFBEA")
        clear(cell)
        set_run(cell.paragraphs[0].add_run(title), size=10, bold=True)
        p = cell.add_paragraph()
        set_run(p.add_run(body), size=9, color=RGBColor(0x66, 0x66, 0x66))

    mt = doc.add_table(rows=1, cols=2)
    set_borders(mt, "F0C14A", "12")
    a, b = mt.rows[0].cells
    shade(a, "FFF6D8")
    shade(b, "FFF6D8")
    clear(a)
    clear(b)
    set_run(a.paragraphs[0].add_run("Ваш менеджер"), size=8, color=RGBColor(0x66, 0x66, 0x66))
    p = a.add_paragraph()
    set_run(p.add_run(MANAGER), size=13, bold=True)
    for line in [EMAIL, PHONE]:
        p = a.add_paragraph()
        set_run(p.add_run(line), size=9)
    set_run(b.paragraphs[0].add_run("ГК «Форус»"), size=11, bold=True)
    for line in ["Готовы подключить и показать на демо", "для кадровой службы и ИТ-отдела", "www.forus.ru"]:
        p = b.add_paragraph()
        set_run(p.add_run(line), size=9, color=RGBColor(0x66, 0x66, 0x66))

    doc.save(OUT_DOCX)
    print("DOCX OK", OUT_DOCX)


def main():
    build_pdf()
    build_docx()
    for src in (OUT_PDF, OUT_DOCX):
        (ROOT / src.name).write_bytes(src.read_bytes())
        print("->", ROOT / src.name)
    (ROOT / "КП_Кабинет_сотрудника_750.pdf").write_bytes(OUT_PDF.read_bytes())


if __name__ == "__main__":
    main()
