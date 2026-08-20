#!/usr/bin/env python3
"""КП-листовка DOCX: 1С:Кабинет сотрудника — 750 кабинетов.

Информационно = первый прототип (полное сравнение с конкурентами).
Часы = актуальная логика: 5 часов в подарок (≈4 настройка + 1 запас), к оплате только кабинеты.
Оформление = жёлтый фирменный стиль Форус (последняя версия).
Менеджер: Данил Кургузов.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
BRAND = ROOT / "assets_forus" / "brand"
OUT_DOCX = ROOT / "output" / "КП_ТФМ_Спецтехника_КЭДО_750.docx"

CABINETS = 223_200
GIFT_VALUE = 18_300
PER_MONTH = 25

MANAGER = "Данил Кургузов"
EMAIL = "dkurguzov@forus.ru"
PHONE = "+7 (3952) 78-00-00"

DARK = RGBColor(0x2B, 0x2B, 0x2B)
GRAY = RGBColor(0x5C, 0x5C, 0x5C)
MUTED = RGBColor(0x8A, 0x8A, 0x8A)
OK = RGBColor(0x2E, 0x7D, 0x32)
BAD = RGBColor(0xB7, 0x1C, 0x1C)
GOLD = RGBColor(0x9A, 0x7A, 0x10)


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


def no_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def clear(cell):
    cell.paragraphs[0].clear()


def keep_together(table):
    """Prevent table from splitting across pages (keeps ИТОГО on page 1)."""
    for row in table.rows:
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        cantSplit = OxmlElement("w:cantSplit")
        trPr.append(cantSplit)


def set_cell_margins(cell, top=40, bottom=40, left=60, right=60):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def p_add(doc, text, *, size=10, bold=False, color=DARK, after=6, before=0, center=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run(r, size=size, bold=bold, color=color)
    return p


def yellow_title(doc, text, *, size=12, after=2, before=4):
    p = p_add(doc, text, size=size, bold=True, after=after, before=before)
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


def yellow_box(doc, title, lines, fill="FFF6D8", title_size=10):
    t = doc.add_table(rows=1, cols=1)
    set_borders(t, "F0C14A", "12")
    cell = t.rows[0].cells[0]
    shade(cell, fill)
    set_cell_margins(cell, 40, 40, 60, 60)
    clear(cell)
    set_run(cell.paragraphs[0].add_run(title), size=title_size, bold=True, color=DARK)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        set_run(p.add_run(line), size=8, color=GRAY)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(3)


def feature_card(cell, title, body):
    shade(cell, "FFFBEA")
    set_cell_margins(cell, 28, 28, 40, 40)
    clear(cell)
    set_run(cell.paragraphs[0].add_run(title), size=9, bold=True, color=DARK)
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(1)
    set_run(p.add_run(body), size=8, color=GRAY)


def build_docx():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.left_margin = Mm(14)
    sec.right_margin = Mm(14)
    sec.top_margin = Mm(10)
    sec.bottom_margin = Mm(10)

    # ----- Header: logo fully visible (padded asset) -----
    ht = doc.add_table(rows=1, cols=2)
    no_borders(ht)
    a, b = ht.rows[0].cells
    clear(a)
    logo = BRAND / "forus_logo_docx.png"
    if not logo.exists():
        logo = BRAND / "forus_logo_word.png"
    if logo.exists():
        run = a.paragraphs[0].add_run()
        # width keeps aspect; padding in PNG prevents crop look
        run.add_picture(str(logo), width=Cm(3.1))
    clear(b)
    for i, (line, bold) in enumerate([
        ("Группа компаний «Форус»", True),
        ("Центр компетенции по кадровому электронному документообороту", False),
        (f"{PHONE}  ·  www.forus.ru", False),
    ]):
        p = b.paragraphs[0] if i == 0 else b.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(0)
        set_run(p.add_run(line), size=8 if i else 9, bold=bold, color=GRAY)

    # yellow rule
    rule = doc.add_paragraph()
    rule.paragraph_format.space_before = Pt(2)
    rule.paragraph_format.space_after = Pt(4)
    pPr = rule._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "28")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "F0C14A")
    pBdr.append(bottom)
    pPr.append(pBdr)

    p_add(doc, "1С:Кабинет сотрудника — 750 личных кабинетов", size=14, bold=True, center=True, after=1)
    p_add(
        doc,
        "Персональное предложение · кадровый электронный документооборот",
        size=9, color=GRAY, center=True, after=4,
    )

    # Why now
    yellow_box(
        doc,
        "Почему это предложение выгодно именно сейчас",
        [
            "• Экономия до 70% времени кадровой службы и до 75% затрат на бумагу, печать и курьеров",
            "• Акция «Больше, чем кешбэк»: 5 часов линии консультаций в подарок (выгода 18 300 ₽)",
            "• Работа внутри привычной 1С — без отдельной HR-платформы и двойного ввода данных",
        ],
    )

    yellow_title(doc, "Что умеет 1С:Кабинет сотрудника", size=11, after=2, before=3)
    caps = [
        (
            "Ознакомление с документами в один клик",
            "Отправьте документ одной кнопкой — сотрудник ознакомится и подтвердит получение. Без печати и сбора подписей.",
        ),
        (
            "Удалённый приём на работу",
            "Трудовой договор, заявление и ознакомление с правилами — без визита в офис.",
        ),
        (
            "Ответы на вопросы по отпуску",
            "Сотрудник сам смотрит остаток отпуска и подаёт заявление онлайн — меньше повторяющихся вопросов кадрам.",
        ),
        (
            "Согласование без походов по кабинетам",
            "Отпуск, командировка, отгул — руководитель согласовывает с телефона, данные сразу в 1С.",
        ),
        (
            "Общение без сторонних мессенджеров",
            "Рабочие вопросы и документы — внутри системы, с учётом требований закона о персональных данных.",
        ),
        (
            "Простое внедрение для ИТ",
            "Сервис уже внутри 1С — без отдельной платформы и долгой интеграции.",
        ),
    ]
    ft = doc.add_table(rows=3, cols=2)
    set_borders(ft, "F0C14A", "8")
    for i, (title, body) in enumerate(caps):
        feature_card(ft.rows[i // 2].cells[i % 2], title, body)

    yellow_title(doc, "Ваш пакет: 750 личных кабинетов", size=11, after=2, before=4)
    p_add(
        doc,
        f"Состав лицензий: 500 + 200 + 50 · ≈ {PER_MONTH} ₽ в месяц за сотрудника",
        size=8, color=GRAY, after=3,
    )

    # Pricing — hours logic UPDATED vs first prototype
    pt = doc.add_table(rows=1, cols=3)
    set_borders(pt, "F0C14A", "10")
    for i, h in enumerate(["Позиция", "Состав / пояснение", "Сумма, ₽"]):
        cell = pt.rows[0].cells[i]
        shade(cell, "F0C14A")
        clear(cell)
        set_cell_margins(cell, 20, 20, 30, 30)
        set_run(cell.paragraphs[0].add_run(h), size=8, bold=True, color=DARK)

    price_rows = [
        (
            "1С:Кабинет сотрудника, 750 кабинетов на 12 месяцев",
            "Пакеты: 500 + 200 + 50 кабинетов",
            f"{CABINETS:,}".replace(",", " "),
            False,
        ),
        (
            "Настройка и запуск сервиса (≈4 часа)",
            "Покрывается подарочными часами линии консультаций",
            "0 ₽",
            True,
        ),
        (
            "Запас на вопросы по 1С (1 час)",
            "Покрывается подарочными часами линии консультаций",
            "0 ₽",
            True,
        ),
        (
            "ПОДАРОК: 5 часов линии консультаций",
            f"Акция «Больше, чем кешбэк!» — выгода {GIFT_VALUE:,} ₽".replace(",", " "),
            "0 ₽",
            True,
        ),
    ]
    for label, note, amount, gift in price_rows:
        row = pt.add_row().cells
        fill = "FFF6D8" if gift else "FFFBEA"
        for cell in row:
            shade(cell, fill)
            clear(cell)
            set_cell_margins(cell, 18, 18, 30, 30)
        set_run(row[0].paragraphs[0].add_run(label), size=8, bold=gift, color=DARK)
        set_run(row[1].paragraphs[0].add_run(note), size=7.5, color=GRAY)
        row[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_run(row[2].paragraphs[0].add_run(amount), size=9, bold=True, color=OK if gift else DARK)
    keep_together(pt)

    # Total — compact yellow bar (smaller sum so it stays on page 1)
    tot = doc.add_table(rows=1, cols=2)
    set_borders(tot, "F0C14A", "12")
    left, right = tot.rows[0].cells
    shade(left, "F0C14A")
    shade(right, "F0C14A")
    set_cell_margins(left, 30, 30, 50, 50)
    set_cell_margins(right, 30, 30, 50, 50)
    clear(left)
    clear(right)
    set_run(left.paragraphs[0].add_run("ИТОГО К ОПЛАТЕ"), size=9, bold=True, color=DARK)
    p = left.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(0)
    set_run(
        p.add_run("Только пакет кабинетов. Настройка и поддержка — из подарка. Подпись сотрудникам — бесплатно."),
        size=7.5, color=DARK,
    )
    right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run(right.paragraphs[0].add_run(f"{CABINETS:,} ₽".replace(",", " ")), size=14, bold=True, color=DARK)
    keep_together(tot)

    # Page break before comparison (like first prototype page 2)
    doc.add_page_break()

    # Header again lite
    ht2 = doc.add_table(rows=1, cols=2)
    no_borders(ht2)
    a, b = ht2.rows[0].cells
    clear(a)
    if logo.exists():
        a.paragraphs[0].add_run().add_picture(str(logo), width=Cm(3.2))
    clear(b)
    b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run(b.paragraphs[0].add_run(f"{PHONE}  ·  www.forus.ru"), size=8, color=GRAY)

    yellow_title(doc, "Внедрение у Форус vs другие операторы")
    p_add(
        doc,
        "Сравнение по ключевым параметрам для компании на 750 сотрудников. "
        "Красным выделены самые невыгодные условия у конкурентов.",
        size=9, color=GRAY, after=6,
    )

    headers = ["Параметр", "Форус + 1С:Кабинет", "HRlink / VK HR Tek", "Контур КЭДО", "Saby"]
    data = [
        ["Работа внутри вашей 1С", "Да — без новой платформы", "Нет — отдельная система", "Нет — отдельная система", "Нет — отдельная система"],
        ["Двойной ввод данных кадровиком", "Нет — заявления сразу становятся документами 1С", "Часто нужен перенос и сверка", "Часто нужен перенос и сверка", "Часто нужен перенос и сверка"],
        ["Где хранятся документы", "В вашей базе 1С", "Облако оператора (часто платно)", "Облако / доп. хранилище", "Облако оператора"],
        ["Электронная подпись сотрудникам", "Бесплатно (усиленная неквалифицированная)", "Часто платные сертификаты / пакеты", "Часто платные сертификаты / пакеты", "Часто платные сертификаты / пакеты"],
        ["Сложность внедрения", "Быстрый старт в привычной 1С", "Долгая интеграция с 1С", "Долгая интеграция с 1С", "Долгая интеграция с 1С"],
        ["Нагрузка на ИТ", "Без новых клиентских лицензий 1С", "Новая система + интеграции", "Новая система + интеграции", "Новая система + интеграции"],
        ["Стоимость владения на старте", "Пакет кабинетов + подарочные часы линии", "Лицензии + внедрение + хранение", "Лицензии + внедрение + хранение", "Лицензии + внедрение + хранение"],
    ]

    ct = doc.add_table(rows=1, cols=5)
    set_borders(ct, "F0C14A", "8")
    ct.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = ct.rows[0].cells[i]
        shade(cell, "F0C14A" if i != 1 else "FFE08A")
        clear(cell)
        set_cell_margins(cell, 30, 30, 30, 30)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
        set_run(cell.paragraphs[0].add_run(h), size=8, bold=True, color=DARK)

    for row_vals in data:
        row = ct.add_row().cells
        for i, val in enumerate(row_vals):
            clear(row[i])
            set_cell_margins(row[i], 30, 30, 30, 30)
            if i == 1:
                shade(row[i], "E8F5E9")
                col = OK
                bold = True
            elif i >= 2:
                shade(row[i], "FFEBEE")
                col = BAD
                bold = False
            else:
                shade(row[i], "FFFBEA")
                col = DARK
                bold = True
            row[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
            set_run(row[i].paragraphs[0].add_run(val), size=7.5, bold=bold, color=col)

    p_add(doc, "", after=8)
    yellow_box(
        doc,
        "Вывод для корпоративного клиента",
        [
            "Отдельные операторы кадрового электронного документооборота почти всегда тянут за собой новую систему, "
            "двойную работу кадровика, платное хранение и долгую интеграцию с 1С.",
            "С Форус вы остаётесь в привычной 1С, получаете подарочные часы поддержки и прозрачную стоимость пакета.",
        ],
        fill="FFF6D8",
    )

    yellow_title(doc, "Как мы запускаем сервис у вас")
    steps = [
        ("01", "Подключение", "Подключаем сервис к вашей 1С"),
        ("02", "Настройка", "Роли, процессы, печатные формы"),
        ("03", "Подписи", "Выпуск подписей сотрудникам"),
        ("04", "Обучение", "Видеоуроки и короткие инструкции"),
        ("05", "Старт", "Заявления, расчётные, согласования"),
    ]
    st = doc.add_table(rows=1, cols=5)
    set_borders(st, "F0C14A", "8")
    for i, (num, title, body) in enumerate(steps):
        cell = st.rows[0].cells[i]
        shade(cell, "FFFBEA")
        set_cell_margins(cell, 40, 40, 30, 30)
        clear(cell)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(cell.paragraphs[0].add_run(num), size=12, bold=True, color=GOLD)
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(title), size=9, bold=True, color=DARK)
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(body), size=8, color=GRAY)

    p_add(doc, "", after=8)
    yellow_title(doc, "Почему Форус")
    for line in [
        "Центр компетенции по кадровому электронному документообороту и по управлению персоналом",
        "ТОП-5 дистрибьюторов 1С в России · более 10 000 клиентов на сопровождении",
        "Успешные внедрения, включая крупные сети (в том числе кейс DNS)",
        "Шаблоны документов для перехода, видеоуроки и линия консультаций",
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        set_run(p.add_run("●  "), size=10, bold=True, color=GOLD)
        set_run(p.add_run(line), size=10, color=DARK)

    p_add(doc, "", after=10)

    # CTA + manager — yellow style
    cta = doc.add_table(rows=2, cols=1)
    set_borders(cta, "F0C14A", "14")
    top = cta.rows[0].cells[0]
    shade(top, "F0C14A")
    set_cell_margins(top, 70, 50, 80, 80)
    clear(top)
    top.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(top.paragraphs[0].add_run("Готовы зафиксировать условия и запустить 750 кабинетов?"), size=12, bold=True, color=DARK)
    p = top.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(
        p.add_run("Подключим сервис, начислим 5 подарочных часов и начнём настройку."),
        size=9, color=DARK,
    )

    bot = cta.rows[1].cells[0]
    shade(bot, "FFF6D8")
    set_cell_margins(bot, 60, 60, 80, 80)
    clear(bot)
    set_run(bot.paragraphs[0].add_run("Ваш менеджер"), size=8, color=MUTED)
    p = bot.add_paragraph()
    set_run(p.add_run(MANAGER), size=13, bold=True, color=DARK)
    for line in [EMAIL, PHONE, "www.forus.ru  ·  г. Иркутск, ул. Ямская, 1/1"]:
        p = bot.add_paragraph()
        set_run(p.add_run(line), size=9, color=GRAY)

    OUT_DOCX.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT_DOCX)
    root_copy = ROOT / OUT_DOCX.name
    root_copy.write_bytes(OUT_DOCX.read_bytes())
    print("DOCX OK", OUT_DOCX)
    print("->", root_copy)


if __name__ == "__main__":
    build_docx()
