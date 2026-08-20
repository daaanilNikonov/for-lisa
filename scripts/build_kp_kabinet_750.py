#!/usr/bin/env python3
"""Листовка КП: 1С:Кабинет сотрудника — 750 кабинетов.

Клиентский взгляд: самое выгодное и самое лёгкое внедрение на рынке.
Цена — один раз и кристально ясно. Больше жёлтого, фигуры Форус, сноски-блоки.
PDF + DOCX. Менеджер: Данил Кургузов.
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
BRAND = ROOT / "assets_forus" / "brand"
OUT_PDF = ROOT / "output" / "КП_ТФМ_Спецтехника_КЭДО_750.pdf"
OUT_DOCX = ROOT / "output" / "КП_ТФМ_Спецтехника_КЭДО_750.docx"

# Палитра Форус — жёлтый доминирует
Y = HexColor("#F0C14A")
Y2 = HexColor("#FFE08A")
YSOFT = HexColor("#FFF6D8")
YCARD = HexColor("#FFFBEA")
YGLOW = HexColor("#FFD966")
YPAGE = HexColor("#FFF9E6")
DARK = HexColor("#2B2B2B")
GRAY = HexColor("#5C5C5C")
MUTED = HexColor("#8A8A8A")
WHITE = white
OK = HexColor("#2E7D32")
OKBG = HexColor("#E8F5E9")
BAD = HexColor("#B71C1C")
BADBG = HexColor("#FFEBEE")

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

W, H = A4
L, R = 12 * mm, 12 * mm
CW = W - L - R

CABINETS = 223_200
GIFT_VALUE = 18_300
PER_MONTH = 25

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


def decor_shapes(c, page=1):
    """Corporate yellow shapes — page wash, side bar, waves, circles."""
    # soft yellow page atmosphere
    c.setFillColor(YPAGE)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.roundRect(5 * mm, 10 * mm, W - 10 * mm, H - 18 * mm, 8, fill=1, stroke=0)

    # left yellow accent stripe (Forus leaflet language)
    c.setFillColor(Y)
    c.rect(0, 0, 3.2 * mm, H, fill=1, stroke=0)

    tr = BRAND / "wave_tr.png"
    bl = BRAND / "wave_bl.png"
    if tr.exists():
        c.drawImage(str(tr), W - 62 * mm, H - 28 * mm, width=62 * mm, height=20 * mm, mask="auto")
    if bl.exists():
        c.drawImage(str(bl), 0, 0, width=56 * mm, height=18 * mm, mask="auto")

    # Extra yellow circles / pills (Forus style accents)
    c.setFillColor(Y2)
    c.circle(W - 16 * mm, H - 88 * mm, 11 * mm, fill=1, stroke=0)
    c.setFillColor(YGLOW)
    c.circle(W - 8 * mm, H - 102 * mm, 6 * mm, fill=1, stroke=0)
    c.setFillColor(Y)
    c.circle(W - 22 * mm, H - 112 * mm, 3.5 * mm, fill=1, stroke=0)
    c.setFillColor(YSOFT)
    c.ellipse(-8 * mm, H * 0.42, 18 * mm, H * 0.58, fill=1, stroke=0)
    if page == 2:
        c.setFillColor(Y2)
        c.circle(W - 10 * mm, 68 * mm, 13 * mm, fill=1, stroke=0)
        c.setFillColor(YGLOW)
        c.circle(W - 24 * mm, 84 * mm, 5.5 * mm, fill=1, stroke=0)
        c.setFillColor(Y)
        c.circle(W - 6 * mm, 92 * mm, 3 * mm, fill=1, stroke=0)
        c.setFillColor(YSOFT)
        c.ellipse(-6 * mm, 105 * mm, 20 * mm, 135 * mm, fill=1, stroke=0)


def header(c):
    logo = BRAND / "forus_logo_word.png"
    if logo.exists():
        c.drawImage(str(logo), L, H - 17 * mm, width=30 * mm, height=10 * mm, mask="auto", preserveAspectRatio=True)
    c.setFillColor(GRAY)
    c.setFont("DejaVu", 7)
    c.drawRightString(W - R, H - 7 * mm, "Группа компаний «Форус»")
    c.drawRightString(W - R, H - 11 * mm, f"{PHONE}  ·  www.forus.ru")
    c.setFillColor(Y)
    c.roundRect(L, H - 20 * mm, CW, 3.2, 1.6, fill=1, stroke=0)


def footer(c, n):
    c.setFillColor(Y)
    c.roundRect(L, 7.5 * mm, CW, 2.2, 1.1, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.setFont("DejaVu", 6.5)
    c.drawString(L, 4 * mm, "www.forus.ru  ·  г. Иркутск, ул. Ямская, 1/1")
    c.drawRightString(W - R, 4 * mm, f"{n} / 2")


def pill(c, x, y, text, size=7.2):
    c.setFont("DejaVuBold", size)
    tw = c.stringWidth(text, "DejaVuBold", size) + 7 * mm
    c.setFillColor(Y)
    c.roundRect(x, y - 1.8 * mm, tw, 6.8 * mm, 3.4, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.drawString(x + 3.5 * mm, y, text)
    return tw


def section_title(c, x, y, text, bar_w=36 * mm):
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 11)
    c.drawString(x, y, text)
    c.setFillColor(Y)
    c.roundRect(x, y - 2.4 * mm, bar_w, 2.2, 1.1, fill=1, stroke=0)


def callout_block(c, x, y, w, h, eyebrow, title, body_lines, fill=YSOFT):
    """Eye-catching footnote / info block."""
    c.setFillColor(fill)
    c.setStrokeColor(Y)
    c.setLineWidth(1.6)
    c.roundRect(x, y - h, w, h, 7, fill=1, stroke=1)
    # left accent bar
    c.setFillColor(Y)
    c.roundRect(x, y - h, 3.5 * mm, h, 1.6, fill=1, stroke=0)
    # top yellow chip
    c.setFont("DejaVuBold", 6.5)
    chip_w = c.stringWidth(eyebrow, "DejaVuBold", 6.5) + 5 * mm
    c.roundRect(x + 6 * mm, y - 5.5 * mm, chip_w, 4.2 * mm, 2.1, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.drawString(x + 8.5 * mm, y - 4.3 * mm, eyebrow)
    c.setFont("DejaVuBold", 9)
    c.drawString(x + 6 * mm, y - 11 * mm, title)
    c.setFont("DejaVu", 7.2)
    c.setFillColor(GRAY)
    yy = y - 16 * mm
    for line in body_lines:
        c.drawString(x + 6 * mm, yy, line)
        yy -= 3.6 * mm


def num_dot(c, x, y, num, r=5.8):
    c.setFillColor(Y)
    c.circle(x, y, r, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 8)
    c.drawCentredString(x, y - 2.6, str(num))


def build_pdf():
    OUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT_PDF), pagesize=A4)
    c.setTitle("1С:Кабинет сотрудника — коммерческое предложение")
    c.setAuthor(f"ГК Форус · {MANAGER}")

    # ========== PAGE 1: оффер + выгода ==========
    decor_shapes(c, 1)
    header(c)

    y = H - 27 * mm
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 19)
    c.drawString(L, y, "1С:Кабинет сотрудника")
    y -= 6.5 * mm
    c.setFont("DejaVu", 9.5)
    c.setFillColor(GRAY)
    c.drawString(L, y, "Электронные кадровые документы внутри вашей 1С")
    y -= 7 * mm
    x = L
    x += pill(c, x, y, "самое лёгкое внедрение") + 2.5 * mm
    x += pill(c, x, y, "без новой платформы") + 2.5 * mm
    pill(c, x, y, "750 кабинетов")

    # HERO OFFER — одна большая жёлтая карточка с ценой (новый выгодный оффер)
    y = H - 58 * mm
    hero_h = 48 * mm
    c.setFillColor(Y)
    c.roundRect(L, y - hero_h, CW, hero_h, 10, fill=1, stroke=0)
    # white inner panel for price clarity
    c.setFillColor(WHITE)
    c.roundRect(L + 3 * mm, y - hero_h + 3 * mm, CW * 0.42, hero_h - 6 * mm, 8, fill=1, stroke=0)

    px = L + 3 * mm + (CW * 0.42) / 2
    c.setFillColor(MUTED)
    c.setFont("DejaVu", 7.5)
    c.drawCentredString(px, y - 9 * mm, "вы платите только за кабинеты")
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 22)
    c.drawCentredString(px, y - 20 * mm, f"{CABINETS:,}".replace(",", " ") + " ₽")
    c.setFont("DejaVu", 8)
    c.setFillColor(GRAY)
    c.drawCentredString(px, y - 27 * mm, "за 750 кабинетов · 12 месяцев")
    c.drawCentredString(px, y - 32 * mm, f"≈ {PER_MONTH} ₽ / мес. на сотрудника")
    c.setFillColor(Y)
    c.roundRect(px - 28 * mm, y - 42 * mm, 56 * mm, 7 * mm, 3.5, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 7.5)
    c.drawCentredString(px, y - 39.5 * mm, "пакеты 500 + 200 + 50")

    # Right side of hero — gift & total in one place (no repeat later)
    rx = L + CW * 0.45
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 10)
    c.drawString(rx, y - 9 * mm, "Акция «Больше, чем кешбэк!»")
    c.setFont("DejaVu", 8)
    gift_lines = [
        "5 часов линии консультаций — в подарок",
        f"(рыночная выгода {GIFT_VALUE:,} ₽)".replace(",", " "),
        "",
        "≈ 4 часа — настройка и запуск сервиса",
        "1 час — запас на вопросы по 1С",
        "",
        "За часы вы не платите отдельно.",
        "В счёте — только пакет кабинетов.",
    ]
    yy = y - 15 * mm
    for line in gift_lines:
        if line:
            c.drawString(rx, yy, line)
        yy -= 3.8 * mm

    # WHAT YOU GET — 2×3 grid of yellow cards (from leaflet, no fluff)
    y = y - hero_h - 7 * mm
    section_title(c, L, y, "Что получает компания", 42 * mm)
    y -= 6 * mm

    benefits = [
        ("До 70% меньше рутины", "Расчётные листки, справки и отпуска — у сотрудника в телефоне."),
        ("Меньше ручного ввода", "Заявления сами становятся приказами и записями в 1С."),
        ("Защита от штрафов", "Документы вовремя, ознакомление подтверждено, архив электронный."),
        ("Удалённый приём", "Трудовой договор и локальные акты — без визита в офис."),
        ("Кадровик в привычной 1С", "Не нужно учить новую кадровую систему."),
        ("ИТ без лишней нагрузки", "Сотрудникам не нужны клиентские лицензии 1С."),
    ]
    cols, rows = 2, 3
    card_w = (CW - 3 * mm) / cols
    card_h = 18 * mm
    for i, (title, body) in enumerate(benefits):
        col, row = i % cols, i // cols
        x = L + col * (card_w + 3 * mm)
        cy = y - row * (card_h + 2.5 * mm)
        c.setFillColor(YCARD)
        c.setStrokeColor(Y)
        c.setLineWidth(1.2)
        c.roundRect(x, cy - card_h, card_w, card_h, 6, fill=1, stroke=1)
        c.setFillColor(Y)
        c.circle(x + 5 * mm, cy - 5.5 * mm, 2.4, fill=1, stroke=0)
        c.setFillColor(DARK)
        c.setFont("DejaVuBold", 8)
        c.drawString(x + 9.5 * mm, cy - 6.2 * mm, title)
        c.setFont("DejaVu", 7)
        c.setFillColor(GRAY)
        yy = cy - 11 * mm
        for line in wrap(c, body, "DejaVu", 7, card_w - 12 * mm):
            c.drawString(x + 5 * mm, yy, line)
            yy -= 3.3 * mm

    # Bottom callout — one IT footnote (eye-catcher)
    y = y - rows * (card_h + 2.5 * mm) - 2 * mm
    callout_block(
        c, L, y, CW, 22 * mm,
        "сноска для выбора",
        "Почему это выгоднее типичных предложений рынка",
        [
            "Отдельная кадровая платформа не покупается. Внедрение не тянется месяцами.",
            "Настройка и стартовая поддержка уже внутри подарка — не отдельная строка счёта.",
            "Документы остаются в вашей базе 1С, а не только в облаке стороннего оператора.",
        ],
        fill=YSOFT,
    )

    footer(c, 1)

    # ========== PAGE 2: сроки, сравнение, условия ==========
    c.showPage()
    decor_shapes(c, 2)
    header(c)

    y = H - 27 * mm
    section_title(c, L, y, "Сроки и порядок запуска", 48 * mm)
    y -= 8 * mm

    steps = [
        ("1", "Договор\nи доступы", "2–4 дня"),
        ("2", "Настройка\nсервиса", "≈ 4 часа"),
        ("3", "Пилот\nна группе", "до 1 недели"),
        ("4", "Все\n750 кабинетов", "2–3 недели"),
        ("5", "Запас\nна вопросы", "1 час"),
    ]
    # yellow track bar
    c.setFillColor(Y2)
    c.roundRect(L + 6 * mm, y - 4 * mm, CW - 12 * mm, 4 * mm, 2, fill=1, stroke=0)
    sw = CW / 5
    for i, (n, title, timing) in enumerate(steps):
        x = L + i * sw + sw / 2
        num_dot(c, x, y - 2 * mm, n, r=6.2)
        c.setFillColor(DARK)
        c.setFont("DejaVuBold", 7)
        for li, line in enumerate(title.split("\n")):
            c.drawCentredString(x, y - 11 * mm - li * 3.3 * mm, line)
        c.setFillColor(HexColor("#9A7A10"))
        c.setFont("DejaVuBold", 7)
        c.drawCentredString(x, y - 20.5 * mm, timing)

    y -= 28 * mm

    # Comparison
    section_title(c, L, y, "С чем сравниваете на тендере", 52 * mm)
    y -= 6 * mm
    cols_w = [38 * mm, 78 * mm, CW - 116 * mm]
    xs = [L, L + cols_w[0], L + cols_w[0] + cols_w[1]]
    hh = 7.5 * mm
    c.setFillColor(Y)
    c.roundRect(L, y - hh, CW, hh, 4, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 7.5)
    for i, htxt in enumerate(["Критерий", "Предложение Форус", "Типичные альтернативы"]):
        c.drawCentredString(xs[i] + cols_w[i] / 2, y - hh + 2.5 * mm, htxt)
    y -= hh

    rows_data = [
        ("Запуск", "Типовая настройка ≈ 4 часа", "Часто долгое внедрение"),
        ("Архитектура", "Работает внутри вашей 1С", "Отдельная кадровая платформа"),
        ("Лицензии", "Без клиентских лицензий 1С на сотрудников", "Новые доступы и интеграции"),
        ("Хранение", "Документы в вашей базе 1С", "Облако оператора, часто платно"),
        ("Старт работ", "Часы поддержки — подарок", "Внедрение отдельной строкой"),
    ]
    for ri, (a, b, d) in enumerate(rows_data):
        rh = 8 * mm
        for ci, (val, w) in enumerate(zip((a, b, d), cols_w)):
            x = xs[ci]
            if ci == 1:
                bg, tc = OKBG, OK
            elif ci == 2:
                bg, tc = BADBG, BAD
            else:
                bg, tc = (YSOFT if ri % 2 == 0 else WHITE), DARK
            c.setFillColor(bg)
            c.roundRect(x, y - rh, w, rh, 0, fill=1, stroke=0)
            c.setStrokeColor(Y2)
            c.setLineWidth(0.6)
            c.rect(x, y - rh, w, rh, fill=0, stroke=1)
            c.setFillColor(tc)
            c.setFont("DejaVuBold" if ci == 0 else "DejaVu", 6.7)
            if ci == 0:
                c.drawString(x + 2 * mm, y - rh + 2.8 * mm, val)
            else:
                c.drawCentredString(x + w / 2, y - rh + 2.8 * mm, val)
        y -= rh

    y -= 7 * mm

    # Budget reminder — compact yellow strip (ссылка на стр.1, без повторного разбора часов)
    c.setFillColor(Y)
    c.roundRect(L, y - 14 * mm, CW, 14 * mm, 7, fill=1, stroke=0)
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 10)
    c.drawString(L + 4 * mm, y - 5.5 * mm, "Бюджет к оплате")
    c.setFont("DejaVuBold", 14)
    c.drawRightString(W - R - 4 * mm, y - 6 * mm, f"{CABINETS:,} ₽ / год".replace(",", " "))
    c.setFont("DejaVu", 7.5)
    c.drawString(
        L + 4 * mm, y - 11 * mm,
        "Только пакет кабинетов. Настройка, стартовая поддержка и подпись сотрудникам — без доплаты.",
    )
    y -= 20 * mm

    # 4 eye-catching footnotes
    section_title(c, L, y, "Условия — коротко и по делу", 48 * mm)
    y -= 5 * mm
    notes = [
        ("подпись", "Электронная подпись", "Выпускается сотрудникам бесплатно."),
        ("комплект", "Шаблоны документов", "Положение, уведомления и согласия — в комплекте."),
        ("формат", "Формат работ", "Удалённо. Выезд и нетиповые доработки — отдельно."),
        ("срок", "Срок предложения", "Действует 30 дней с даты отправки."),
    ]
    nw = (CW - 3 * mm) / 2
    nh = 18 * mm
    for i, (chip, title, body) in enumerate(notes):
        col, row = i % 2, i // 2
        x = L + col * (nw + 3 * mm)
        yy = y - row * (nh + 3 * mm)
        c.setFillColor(YCARD)
        c.setStrokeColor(Y)
        c.setLineWidth(1.4)
        c.roundRect(x, yy - nh, nw, nh, 6, fill=1, stroke=1)
        c.setFillColor(Y)
        c.roundRect(x + 3 * mm, yy - 5.2 * mm, c.stringWidth(chip, "DejaVuBold", 6.2) + 4 * mm, 3.8 * mm, 1.9, fill=1, stroke=0)
        c.setFillColor(DARK)
        c.setFont("DejaVuBold", 6.2)
        c.drawString(x + 5 * mm, yy - 4.2 * mm, chip)
        c.setFont("DejaVuBold", 8.5)
        c.drawString(x + 3 * mm, yy - 10 * mm, title)
        c.setFont("DejaVu", 7.2)
        c.setFillColor(GRAY)
        c.drawString(x + 3 * mm, yy - 14.5 * mm, body)

    # Manager — fixed bottom zone, no overlap with notes
    c.setFillColor(Y)
    c.roundRect(L, 14 * mm, CW, 22 * mm, 8, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.roundRect(L + 2.5 * mm, 16 * mm, CW * 0.48, 18 * mm, 6, fill=1, stroke=0)
    c.setFillColor(MUTED)
    c.setFont("DejaVu", 7)
    c.drawString(L + 5 * mm, 30 * mm, "ваш менеджер")
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 11)
    c.drawString(L + 5 * mm, 24 * mm, MANAGER)
    c.setFont("DejaVu", 8)
    c.drawString(L + 5 * mm, 19 * mm, f"{EMAIL}  ·  {PHONE}")
    c.setFillColor(DARK)
    c.setFont("DejaVuBold", 9)
    c.drawRightString(W - R - 4 * mm, 28 * mm, "Покажем на демо")
    c.setFont("DejaVu", 7.5)
    c.drawRightString(W - R - 4 * mm, 22 * mm, "кадровой службе и ИТ-отделу")
    c.setFont("DejaVuBold", 7.5)
    c.drawRightString(W - R - 4 * mm, 17 * mm, "готовы подключить под тендер")

    footer(c, 2)
    c.save()
    print("PDF OK", OUT_PDF)


# ---------------- DOCX (stable tables, no floating layout) ----------------

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
    p = p_add(doc, text, size=13, bold=True, after=4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "22")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), "F0C14A")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def build_docx():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.left_margin = Mm(14)
    sec.right_margin = Mm(14)
    sec.top_margin = Mm(12)
    sec.bottom_margin = Mm(12)

    # Header
    ht = doc.add_table(rows=1, cols=2)
    a, b = ht.rows[0].cells
    clear(a)
    logo = BRAND / "forus_logo_word.png"
    if logo.exists():
        a.paragraphs[0].add_run().add_picture(str(logo), width=Cm(3.2))
    clear(b)
    for i, line in enumerate(["Группа компаний «Форус»", f"{PHONE}  ·  www.forus.ru"]):
        p = b.paragraphs[0] if i == 0 else b.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(0)
        set_run(p.add_run(line), size=8, bold=(i == 0), color=RGBColor(0x5C, 0x5C, 0x5C))

    rule = doc.add_paragraph()
    rule.paragraph_format.space_after = Pt(8)
    pPr = rule._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "28")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "F0C14A")
    pBdr.append(bottom)
    pPr.append(pBdr)

    p_add(doc, "1С:Кабинет сотрудника", size=20, bold=True, center=True, after=2)
    p_add(
        doc,
        "Электронные кадровые документы внутри вашей 1С  ·  750 личных кабинетов",
        size=10, color=RGBColor(0x5C, 0x5C, 0x5C), center=True, after=4,
    )
    p_add(
        doc,
        "Самое лёгкое внедрение  ·  без новой платформы",
        size=10, bold=True, color=RGBColor(0x9A, 0x7A, 0x10), center=True, after=10,
    )

    # ONE pricing offer block
    offer = doc.add_table(rows=2, cols=2)
    set_borders(offer, "F0C14A", "16")
    left, right = offer.rows[0].cells
    shade(left, "F0C14A")
    shade(right, "F0C14A")
    clear(left)
    clear(right)
    left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(left.paragraphs[0].add_run("вы платите только за кабинеты"), size=8, color=RGBColor(0x2B, 0x2B, 0x2B))
    p = left.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(f"{CABINETS:,} ₽".replace(",", " ")), size=22, bold=True)
    p = left.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(f"750 кабинетов / 12 мес.  ·  ≈ {PER_MONTH} ₽/мес. на сотрудника"), size=8, color=RGBColor(0x5C, 0x5C, 0x5C))
    p = left.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("пакеты 500 + 200 + 50"), size=9, bold=True)

    set_run(right.paragraphs[0].add_run("Акция «Больше, чем кешбэк!»"), size=11, bold=True)
    for line in [
        "5 часов линии консультаций — в подарок",
        f"(рыночная выгода {GIFT_VALUE:,} ₽)".replace(",", " "),
        "≈ 4 часа — настройка и запуск",
        "1 час — запас на вопросы по 1С",
        "За часы вы не платите отдельно.",
    ]:
        p = right.add_paragraph()
        set_run(p.add_run(line), size=9)

    # merge second row as total
    bot_l, bot_r = offer.rows[1].cells
    shade(bot_l, "FFF6D8")
    shade(bot_r, "FFF6D8")
    clear(bot_l)
    clear(bot_r)
    set_run(bot_l.paragraphs[0].add_run("Итого к оплате = только пакет кабинетов"), size=10, bold=True)
    bot_r.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run(bot_r.paragraphs[0].add_run(f"{CABINETS:,} ₽".replace(",", " ")), size=14, bold=True)

    p_add(doc, "", after=8)
    yellow_title(doc, "Что получает компания")

    benefits = [
        ("До 70% меньше рутины", "Расчётные листки, справки и отпуска — у сотрудника в телефоне."),
        ("Меньше ручного ввода", "Заявления сами становятся приказами и записями в 1С."),
        ("Защита от штрафов", "Документы вовремя, ознакомление подтверждено, архив электронный."),
        ("Удалённый приём", "Трудовой договор и локальные акты — без визита в офис."),
        ("Кадровик в привычной 1С", "Не нужно учить новую кадровую систему."),
        ("ИТ без лишней нагрузки", "Сотрудникам не нужны клиентские лицензии 1С."),
    ]
    bt = doc.add_table(rows=3, cols=2)
    set_borders(bt, "F0C14A", "8")
    for i, (title, body) in enumerate(benefits):
        cell = bt.rows[i // 2].cells[i % 2]
        shade(cell, "FFFBEA")
        clear(cell)
        set_run(cell.paragraphs[0].add_run(title), size=10, bold=True)
        p = cell.add_paragraph()
        set_run(p.add_run(body), size=9, color=RGBColor(0x5C, 0x5C, 0x5C))

    p_add(doc, "", after=8)
    # Callout
    tip = doc.add_table(rows=1, cols=1)
    set_borders(tip, "F0C14A", "14")
    cell = tip.rows[0].cells[0]
    shade(cell, "FFF6D8")
    clear(cell)
    set_run(cell.paragraphs[0].add_run("Почему это выгоднее типичных предложений рынка"), size=11, bold=True)
    for line in [
        "Отдельная кадровая платформа не покупается. Внедрение не тянется месяцами.",
        "Настройка и стартовая поддержка уже внутри подарка — не отдельная строка счёта.",
        "Документы остаются в вашей базе 1С, а не только в облаке стороннего оператора.",
    ]:
        p = cell.add_paragraph()
        set_run(p.add_run("• " + line), size=9, color=RGBColor(0x5C, 0x5C, 0x5C))

    p_add(doc, "", after=10)
    yellow_title(doc, "Сроки и порядок запуска")
    st = doc.add_table(rows=1, cols=5)
    set_borders(st, "F0C14A", "8")
    steps = [
        ("1", "Договор и доступы", "2–4 дня"),
        ("2", "Настройка сервиса", "≈ 4 часа"),
        ("3", "Пилот на группе", "до 1 недели"),
        ("4", "Все 750 кабинетов", "2–3 недели"),
        ("5", "Запас на вопросы", "1 час"),
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
        set_run(p.add_run(timing), size=8, color=RGBColor(0x5C, 0x5C, 0x5C))

    p_add(doc, "", after=8)
    yellow_title(doc, "С чем сравниваете на тендере")
    ct = doc.add_table(rows=1, cols=3)
    set_borders(ct, "F0C14A")
    for i, h in enumerate(["Критерий", "Предложение Форус", "Типичные альтернативы"]):
        cell = ct.rows[0].cells[i]
        shade(cell, "F0C14A")
        clear(cell)
        set_run(cell.paragraphs[0].add_run(h), size=9, bold=True)
    for a, b, d in [
        ("Запуск", "Типовая настройка ≈ 4 часа", "Часто долгое внедрение"),
        ("Архитектура", "Работает внутри вашей 1С", "Отдельная кадровая платформа"),
        ("Лицензии", "Без клиентских лицензий 1С на сотрудников", "Новые доступы и интеграции"),
        ("Хранение", "Документы в вашей базе 1С", "Облако оператора, часто платно"),
        ("Старт работ", "Часы поддержки — подарок", "Внедрение отдельной строкой"),
    ]:
        row = ct.add_row().cells
        shade(row[1], "E8F5E9")
        shade(row[2], "FFEBEE")
        for i, val in enumerate((a, b, d)):
            clear(row[i])
            col = RGBColor(0x2E, 0x7D, 0x32) if i == 1 else (RGBColor(0xB7, 0x1C, 0x1C) if i == 2 else RGBColor(0x2B, 0x2B, 0x2B))
            set_run(row[i].paragraphs[0].add_run(val), size=8, bold=(i == 0), color=col)

    p_add(doc, "", after=8)
    yellow_title(doc, "Условия — коротко и по делу")
    notes = [
        ("Электронная подпись", "Выпускается сотрудникам бесплатно."),
        ("Шаблоны документов", "Положение, уведомления и согласия — в комплекте."),
        ("Формат работ", "Удалённо. Выезд и нетиповые доработки — отдельно."),
        ("Срок предложения", "Действует 30 дней с даты отправки."),
    ]
    nt = doc.add_table(rows=2, cols=2)
    set_borders(nt, "F0C14A", "10")
    for i, (title, body) in enumerate(notes):
        cell = nt.rows[i // 2].cells[i % 2]
        shade(cell, "FFFBEA")
        clear(cell)
        set_run(cell.paragraphs[0].add_run(title), size=10, bold=True)
        p = cell.add_paragraph()
        set_run(p.add_run(body), size=9, color=RGBColor(0x5C, 0x5C, 0x5C))

    p_add(doc, "", after=10)
    mt = doc.add_table(rows=1, cols=2)
    set_borders(mt, "F0C14A", "14")
    a, b = mt.rows[0].cells
    shade(a, "F0C14A")
    shade(b, "FFF6D8")
    clear(a)
    clear(b)
    set_run(a.paragraphs[0].add_run("Ваш менеджер"), size=8)
    p = a.add_paragraph()
    set_run(p.add_run(MANAGER), size=13, bold=True)
    for line in [EMAIL, PHONE]:
        p = a.add_paragraph()
        set_run(p.add_run(line), size=9)
    set_run(b.paragraphs[0].add_run("ГК «Форус»"), size=11, bold=True)
    for line in ["Покажем на демо кадровой службе и ИТ-отделу", "Готовы подключить под тендер", "www.forus.ru"]:
        p = b.add_paragraph()
        set_run(p.add_run(line), size=9, color=RGBColor(0x5C, 0x5C, 0x5C))

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
