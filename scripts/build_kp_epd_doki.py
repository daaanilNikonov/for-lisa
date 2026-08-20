#!/usr/bin/env python3
"""Коммерческое предложение / листовка: 1С-ЭПД и Доки.Логистика (стиль ГК Форус)."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from docx.shared import Cm, Mm, Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "КП_1С-ЭПД_и_Доки_Логистика.docx"
ASSETS = ROOT / "assets_forus"

# Фирменные цвета Форус
YELLOW = RGBColor(0xE8, 0xB8, 0x4A)
YELLOW_HEX = "E8B84A"
DARK = RGBColor(0x1E, 0x1E, 0x1E)
GRAY = RGBColor(0x55, 0x55, 0x55)
LIGHT_GRAY = "F5F5F5"
WHITE = "FFFFFF"
SOFT_YELLOW = "FFF8E7"
HEADER_BG = "1E1E1E"


def set_run_font(run, name="Arial", size=10, bold=False, color=DARK):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def set_cell_shading(cell, hex_color: str):
    tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_borders(cell, color=YELLOW_HEX, sz="12", sides=("top", "left", "bottom", "right")):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in sides:
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)


def set_table_borders(table, color="CCCCCC", sz="4"):
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


def clear_cell(cell):
    for p in cell.paragraphs:
        p.clear()


def add_para(cell_or_doc, text, *, size=10, bold=False, color=DARK, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4, space_before=0):
    if hasattr(cell_or_doc, "paragraphs") and hasattr(cell_or_doc, "add_paragraph") and not hasattr(cell_or_doc, "rows"):
        # Document
        p = cell_or_doc.add_paragraph()
    elif hasattr(cell_or_doc, "add_paragraph"):
        p = cell_or_doc.add_paragraph()
    else:
        # cell
        p = cell_or_doc.paragraphs[0] if not cell_or_doc.paragraphs[0].text else cell_or_doc.add_paragraph()
        if cell_or_doc.paragraphs[0].text == "" and len(cell_or_doc.paragraphs) == 1:
            p = cell_or_doc.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_bullet(doc, text, size=10):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(0)
    p.clear()
    run = p.add_run(text)
    set_run_font(run, size=size, color=DARK)
    # yellow bullet via numbering is hard; prefix with mark
    return p


def add_check_item(doc, text, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.left_indent = Cm(0.3)
    mark = p.add_run("●  ")
    set_run_font(mark, size=size, color=YELLOW, bold=True)
    run = p.add_run(text)
    set_run_font(run, size=size, color=DARK)
    return p


def section_title(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_run_font(run, size=14, bold=True, color=DARK)
    # yellow underline via border
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), YELLOW_HEX)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def fill_header_cell(cell, text, center=True):
    clear_cell(cell)
    set_cell_shading(cell, HEADER_BG)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    set_run_font(run, size=9, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))


def fill_cell(cell, text, *, bold=False, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, shade=None):
    clear_cell(cell)
    if shade:
        set_cell_shading(cell, shade)
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=DARK)


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()

    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(14)
    section.right_margin = Mm(14)
    section.top_margin = Mm(12)
    section.bottom_margin = Mm(12)

    # --- Шапка ---
    header_table = doc.add_table(rows=1, cols=2)
    header_table.autofit = True
    left, right = header_table.rows[0].cells
    clear_cell(left)
    clear_cell(right)

    logo_path = ASSETS / "brand" / "forus_logo_word.png"
    if not logo_path.exists():
        logo_path = ASSETS / "brand" / "forus_logo_clean.png"
    if not logo_path.exists():
        logo_path = ASSETS / "brand" / "forus_logo_word.png"
    if logo_path.exists():
        p = left.paragraphs[0]
        run = p.add_run()
        run.add_picture(str(logo_path), width=Cm(4.2))
    else:
        add_para(left, "Форус", size=22, bold=True)

    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for i, line in enumerate(
        [
            "Группа компаний «Форус»",
            "г. Иркутск, ул. Ямская, 1/1",
            "+7 (3952) 78-00-00  ·  www.forus.ru",
        ]
    ):
        if i:
            p = right.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
        else:
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(0)
        run = p.add_run(line)
        set_run_font(run, size=9, bold=(i == 0), color=GRAY if i else DARK)

    # жёлтая линия
    line = doc.add_paragraph()
    line.paragraph_format.space_before = Pt(4)
    line.paragraph_format.space_after = Pt(8)
    pPr = line._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "24")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), YELLOW_HEX)
    pBdr.append(bottom)
    pPr.append(pBdr)

    # --- Заголовок ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(2)
    run = title.add_run("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ")
    set_run_font(run, size=18, bold=True, color=DARK)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.paragraph_format.space_after = Pt(4)
    run = sub.add_run("сервисы электронных перевозочных документов")
    set_run_font(run, size=11, color=GRAY)

    products = doc.add_paragraph()
    products.alignment = WD_ALIGN_PARAGRAPH.CENTER
    products.paragraph_format.space_after = Pt(8)
    run = products.add_run("1С-ЭПД")
    set_run_font(run, size=13, bold=True, color=DARK)
    run = products.add_run("  ·  ")
    set_run_font(run, size=13, color=YELLOW)
    run = products.add_run("Доки.Логистика")
    set_run_font(run, size=13, bold=True, color=DARK)

    intro = doc.add_paragraph()
    intro.paragraph_format.space_after = Pt(8)
    run = intro.add_run(
        "С 01.09.2026 г. вступает в силу Федеральный закон от 07.06.2025 № 140-ФЗ: "
        "переход на электронные перевозочные документы (ЭТрН, ЭЗЗ, экспедиторские документы, ЭПЛ) "
        "становится обязательным. ГК «Форус» предлагает два готовых решения для быстрого и "
        "законного старта обмена ЭПД — "
    )
    set_run_font(run, size=10, color=DARK)
    run = intro.add_run("1С-ЭПД")
    set_run_font(run, size=10, bold=True, color=DARK)
    run = intro.add_run(" и ")
    set_run_font(run, size=10, color=DARK)
    run = intro.add_run("Доки.Логистика")
    set_run_font(run, size=10, bold=True, color=DARK)
    run = intro.add_run(".")
    set_run_font(run, size=10, color=DARK)

    # --- Плюсы двух продуктов в две колонки ---
    section_title(doc, "Плюсы каждого сервиса")

    pros = doc.add_table(rows=2, cols=2)
    pros.autofit = True
    set_table_borders(pros, color=YELLOW_HEX, sz="10")

    h1, h2 = pros.rows[0].cells
    fill_header_cell(h1, "1С-ЭПД")
    fill_header_cell(h2, "Доки.Логистика")

    c1, c2 = pros.rows[1].cells
    set_cell_shading(c1, SOFT_YELLOW)
    set_cell_shading(c2, "F7F7FB")

    clear_cell(c1)
    items_epd = [
        "Встроено в типовые конфигурации 1С — без доп. расширений",
        "Поддержка всех форматов ЭПД, требуемых ФНС",
        "Роуминг с другими операторами ИС ЭПД",
        "Мобильное подписание водителем (УКЭП) в ЭТрН",
        "Работа из привычной учётной системы 1С",
        "Бесплатная настройка рабочего места при покупке пакета от 1 000 титулов (акция)",
        "Можно начать с бесплатной конфигурации «1С:Клиент ЭДО»",
    ]
    for i, t in enumerate(items_epd):
        p = c1.paragraphs[0] if i == 0 else c1.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(2)
        m = p.add_run("●  ")
        set_run_font(m, size=9, color=YELLOW, bold=True)
        r = p.add_run(t)
        set_run_font(r, size=9, color=DARK)

    clear_cell(c2)
    items_doki = [
        "Работа в 1С, веб-кабинете и мобильном приложении независимо",
        "ЭПД, УПД, акты, счета и договоры — все документы в одном сервисе",
        "Подходит без 1С, без ПК или при редко обновляемой 1С",
        "Облачный архив: документы не потеряются при сбое базы",
        "Разделение доступа по сотрудникам и подразделениям",
        "Промотариф: 3 месяца безлимита для новых клиентов",
        "Мобильное подписание водителем (ПЭП / УКЭП / Госключ)",
    ]
    for i, t in enumerate(items_doki):
        p = c2.paragraphs[0] if i == 0 else c2.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(2)
        m = p.add_run("●  ")
        set_run_font(m, size=9, color=YELLOW, bold=True)
        r = p.add_run(t)
        set_run_font(r, size=9, color=DARK)

    # --- Отличия ---
    section_title(doc, "Основные отличия")

    diff = doc.add_table(rows=1, cols=3)
    set_table_borders(diff, color="DDDDDD", sz="4")
    fill_header_cell(diff.rows[0].cells[0], "Критерий", center=False)
    fill_header_cell(diff.rows[0].cells[1], "1С-ЭПД")
    fill_header_cell(diff.rows[0].cells[2], "Доки.Логистика")

    rows_data = [
        ("Где работает", "Внутри типовой 1С (+ моб. приложение)", "1С + веб + мобильное приложение"),
        ("Кому подходит", "Учёт в 1С, регулярные обновления, уже есть 1С-ЭДО", "Нет 1С / нет ПК / логисты вне базы / редкие обновления"),
        ("Установка", "Типовой функционал, без расширений", "Расширение в 1С (веб и мобильный — сразу)"),
        ("Типы документов", "Фокус на ЭПД (ЭТрН, ЭЗЗ, ЭПЛ и др.)", "ЭПД + полный ЭДО с контрагентами"),
        ("Хранение", "В информационной базе 1С", "Облачный архив + синхронизация"),
        ("Старт для новичков", "1С:Клиент ЭДО (бесплатно)", "Промотариф 3 мес. безлимит*"),
        ("Модель оплаты", "Пакеты титулов / постоплата 7 ₽", "Годовые пакеты исходящих документов"),
    ]
    for i, (a, b, c) in enumerate(rows_data):
        row = diff.add_row().cells
        shade = LIGHT_GRAY if i % 2 == 0 else WHITE
        fill_cell(row[0], a, bold=True, size=8, align=WD_ALIGN_PARAGRAPH.LEFT, shade=shade)
        fill_cell(row[1], b, size=8, align=WD_ALIGN_PARAGRAPH.LEFT, shade=shade)
        fill_cell(row[2], c, size=8, align=WD_ALIGN_PARAGRAPH.LEFT, shade=shade)

    note = doc.add_paragraph()
    note.paragraph_format.space_before = Pt(4)
    note.paragraph_format.space_after = Pt(4)
    run = note.add_run(
        "* Промотариф Доки недоступен клиентам с учётной записью в сервисе Астрал.ЭДО. "
        "Ключевое отличие: 1С-ЭПД — типовое решение внутри 1С; Доки.Логистика работает "
        "независимо в 1С, вебе и мобильном приложении."
    )
    set_run_font(run, size=8, color=GRAY)

    # --- Тарифы 1С-ЭПД ---
    section_title(doc, "Стоимость пакетов документов")

    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(2)
    h.paragraph_format.space_after = Pt(4)
    run = h.add_run("1С-ЭПД — предоплатные пакеты титулов (12 месяцев)")
    set_run_font(run, size=11, bold=True, color=DARK)

    note2 = doc.add_paragraph()
    note2.paragraph_format.space_after = Pt(4)
    run = note2.add_run(
        "Постоплатная модель: 7 ₽ за 1 титул. Титулы в ЭЗЗ бесплатны для ГО и перевозчика; "
        "в ЭТрН оплачиваются Т1 (ГО) и Т2 (перевозчик); в ЭПЛ — Т4. Актуально до 31.12.2026."
    )
    set_run_font(run, size=8, color=GRAY)

    t_epd = doc.add_table(rows=1, cols=3)
    set_table_borders(t_epd, color="DDDDDD", sz="4")
    fill_header_cell(t_epd.rows[0].cells[0], "Пакет", center=False)
    fill_header_cell(t_epd.rows[0].cells[1], "Цена за 1 документ, ₽")
    fill_header_cell(t_epd.rows[0].cells[2], "Стоимость пакета, ₽")

    epd_packages = [
        ("«1С-ЭДО. ЭПД-600»", "6,00", "3 600"),
        ("«1С-ЭДО. ЭПД-1000»", "5,00", "5 000"),
        ("«1С-ЭДО. ЭПД-5000»", "4,50", "22 500"),
        ("«1С-ЭДО. ЭПД-10000»", "4,00", "40 000"),
        ("«1С-ЭДО. ЭПД-50000»", "3,00", "150 000"),
        ("«1С-ЭДО. ЭПД-100000»", "2,50", "250 000"),
    ]
    for i, (name, per, total) in enumerate(epd_packages):
        row = t_epd.add_row().cells
        shade = SOFT_YELLOW if i % 2 == 0 else WHITE
        fill_cell(row[0], name, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.LEFT, shade=shade)
        fill_cell(row[1], per, size=9, shade=shade)
        fill_cell(row[2], total, bold=True, size=9, shade=shade)

    promo = doc.add_paragraph()
    promo.paragraph_format.space_before = Pt(6)
    promo.paragraph_format.space_after = Pt(8)
    run = promo.add_run("Акция: ")
    set_run_font(run, size=9, bold=True, color=DARK)
    run = promo.add_run(
        "при покупке пакета от 1 000 титулов — бесплатная настройка 1 рабочего места "
        "для работы с ЭПД (активация сервиса, КриптоПро, УКЭП, МЧД, до 3 контрагентов). "
        "Лицензия СКЗИ и сертификаты 1С:Подпись оплачиваются отдельно."
    )
    set_run_font(run, size=9, color=DARK)

    # --- Тарифы Доки ---
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(4)
    h.paragraph_format.space_after = Pt(4)
    run = h.add_run("Доки.Логистика — тарифы на исходящие документы (12 месяцев)")
    set_run_font(run, size=11, bold=True, color=DARK)

    note3 = doc.add_paragraph()
    note3.paragraph_format.space_after = Pt(4)
    run = note3.add_run(
        "Входящие документы подписываются без тарифа. Чем больше пакет — тем ниже цена "
        "одного документа. Неизрасходованный остаток переносится при своевременном продлении. "
        "Для сравнения: стоимость бумажного документа ≈ 50 ₽."
    )
    set_run_font(run, size=8, color=GRAY)

    t_doki = doc.add_table(rows=1, cols=3)
    set_table_borders(t_doki, color="DDDDDD", sz="4")
    fill_header_cell(t_doki.rows[0].cells[0], "Кол-во документов / год", center=False)
    fill_header_cell(t_doki.rows[0].cells[1], "Стоимость 1 док-та, ₽")
    fill_header_cell(t_doki.rows[0].cells[2], "Стоимость тарифа, ₽ / год")

    doki_packages = [
        ("200", "9,00", "1 800"),
        ("500", "8,40", "4 200"),
        ("1 200", "7,67", "9 200"),
        ("1 500", "7,60", "11 400"),
        ("3 000", "5,70", "17 100"),
        ("5 000", "4,90", "24 500"),
        ("10 000", "3,90", "39 000"),
        ("20 000", "3,65", "73 000"),
        ("30 000", "3,50", "105 000"),
        ("50 000", "3,40", "170 000"),
        ("100 000", "3,20", "320 000"),
    ]
    for i, (qty, per, total) in enumerate(doki_packages):
        row = t_doki.add_row().cells
        shade = "F0EEF8" if i % 2 == 0 else WHITE
        fill_cell(row[0], qty, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.LEFT, shade=shade)
        fill_cell(row[1], per, size=9, shade=shade)
        fill_cell(row[2], total, bold=True, size=9, shade=shade)

    promo2 = doc.add_paragraph()
    promo2.paragraph_format.space_before = Pt(6)
    promo2.paragraph_format.space_after = Pt(6)
    run = promo2.add_run("Акция «15 месяцев по цене 12»: ")
    set_run_font(run, size=9, bold=True, color=DARK)
    run = promo2.add_run(
        "при покупке любого пакета Доки + 1 часа линии консультаций (3 660 ₽) — "
        "дополнительно 3 месяца демо-доступа. Оплаченный пакет активируется после демо-периода."
    )
    set_run_font(run, size=9, color=DARK)

    # --- Дополнительно ---
    section_title(doc, "Что может понадобиться дополнительно")

    extra = doc.add_table(rows=1, cols=3)
    set_table_borders(extra, color="DDDDDD", sz="4")
    fill_header_cell(extra.rows[0].cells[0], "Наименование", center=False)
    fill_header_cell(extra.rows[0].cells[1], "Цена")
    fill_header_cell(extra.rows[0].cells[2], "Комментарий", center=False)

    extras = [
        ("Линия консультаций 1С (1–3 ч)", "3 660 ₽/час", "Удалённая техподдержка и настройка"),
        ("Подготовка 1 рабочего места 1С-ЭПД", "7 320 ₽", "Обучение + настройка / экспресс-анализ"),
        ("Лицензия КриптоПро CSP (бессрочно)", "3 700 ₽", "СКЗИ для работы с ЭП"),
        ("УКЭП (1С:Подпись)", "1 050 ₽", "На каждого подписанта"),
        ("Рутокен 3.0 / NFC-токен", "от 2 700 ₽", "Носитель для мобильного подписания"),
        ("Договор 1С:ИТС", "от 3 273 ₽/мес.", "Сопровождение и обновления 1С"),
    ]
    for i, (n, p, c) in enumerate(extras):
        row = extra.add_row().cells
        shade = LIGHT_GRAY if i % 2 == 0 else WHITE
        fill_cell(row[0], n, bold=True, size=8, align=WD_ALIGN_PARAGRAPH.LEFT, shade=shade)
        fill_cell(row[1], p, size=8, shade=shade)
        fill_cell(row[2], c, size=8, align=WD_ALIGN_PARAGRAPH.LEFT, shade=shade)

    # --- Контакты менеджера ---
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(10)
    spacer.paragraph_format.space_after = Pt(0)
    pPr = spacer._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "24")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), YELLOW_HEX)
    pBdr.append(bottom)
    pPr.append(pBdr)

    contact_box = doc.add_table(rows=1, cols=2)
    set_table_borders(contact_box, color=YELLOW_HEX, sz="12")
    left_c, right_c = contact_box.rows[0].cells
    set_cell_shading(left_c, SOFT_YELLOW)
    set_cell_shading(right_c, SOFT_YELLOW)

    clear_cell(left_c)
    p = left_c.paragraphs[0]
    run = p.add_run("Ваш менеджер")
    set_run_font(run, size=9, color=GRAY)
    p = left_c.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run("Оглоблина Софья")
    set_run_font(run, size=13, bold=True, color=DARK)
    p = left_c.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run("Менеджер по продаже 1С-ЭПД")
    set_run_font(run, size=9, color=DARK)
    for line in [
        "E-mail: sogloblina@forus.ru",
        "Тел.: +7 (3952) 78-00-00, доб. 1861",
        "Москва +5 часов",
    ]:
        p = left_c.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(line)
        set_run_font(run, size=9, color=DARK)

    clear_cell(right_c)
    p = right_c.paragraphs[0]
    run = p.add_run("Контакты ГК «Форус»")
    set_run_font(run, size=9, color=GRAY)
    p = right_c.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    run = p.add_run("ООО НПФ «Форус»")
    set_run_font(run, size=12, bold=True, color=DARK)
    for line in [
        "664047, г. Иркутск, ул. Ямская, 1/1, офис 1",
        "Тел.: +7 (3952) 78-00-00, 72-87-02",
        "E-mail: info@forus.ru",
        "Сайт: www.forus.ru",
        "ИНН 3812023430  ·  ОГРН 1023801752633",
    ]:
        p = right_c.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(line)
        set_run_font(run, size=9, color=DARK)

    closing = doc.add_paragraph()
    closing.alignment = WD_ALIGN_PARAGRAPH.CENTER
    closing.paragraph_format.space_before = Pt(12)
    run = closing.add_run(
        "Готовы подобрать оптимальный пакет и провести демонстрацию сервиса.\n"
        "Ожидаем обратную связь и надеемся на плодотворное сотрудничество!"
    )
    set_run_font(run, size=9, color=GRAY)

    doc.save(OUT)
    print(f"Saved: {OUT}")
    return OUT


if __name__ == "__main__":
    # Ensure logo exists
    from PIL import Image, ImageDraw, ImageFont
    import os

    ASSETS.mkdir(parents=True, exist_ok=True)
    logo_path = ASSETS / "brand" / "forus_logo_clean.png"
    if not logo_path.exists():
        font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
        img = Image.new("RGBA", (420, 120), (255, 255, 255, 0))
        d = ImageDraw.Draw(img)
        font = ImageFont.truetype(font_path, 64)
        text = "Форус"
        x, y = 10, 30
        d.text((x, y), text, fill=(30, 30, 30, 255), font=font)
        bbox_f = d.textbbox((0, 0), "Ф", font=font)
        fw = bbox_f[2] - bbox_f[0]
        bbox_o = d.textbbox((0, 0), "о", font=font)
        ow = bbox_o[2] - bbox_o[0]
        ox = x + fw + 2
        oy = y - 6
        d.rounded_rectangle([ox, oy, ox + ow - 2, oy + 9], radius=2, fill=(232, 184, 74, 255))
        img.save(logo_path)

    build()
