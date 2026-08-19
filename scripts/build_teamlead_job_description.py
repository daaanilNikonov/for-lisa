#!/usr/bin/env python3
"""Generate Forus-branded job description for Product Launch Team Lead."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "Должностная_инструкция_тимлид_продуктовый_запуск.docx"
OUT_ROOT = ROOT / "Должностная_инструкция_тимлид_продуктовый_запуск.docx"
ART_RU = Path("/opt/cursor/artifacts/Должностная_инструкция_тимлид_продуктовый_запуск.docx")
ART_EN = Path("/opt/cursor/artifacts/Dolzhnostnaya_instrukciya_timlid_produktovyy_zapusk.docx")

YELLOW = RGBColor(0xFC, 0xCD, 0x68)
NEAR_BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x76, 0x76, 0x77)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LABEL_BG = RGBColor(0xFF, 0xF3, 0xD6)
ROW_ALT = RGBColor(0xFF, 0xF8, 0xE8)
HEADER_BG = YELLOW
PRIORITY_BG = RGBColor(0xF0, 0xF7, 0xFB)
FONT = "Verdana"

# Ранг 1 = наиболее важно / сильнее влияет на результат роли.
# Норма месяца ≈ 168 ч (21 р.д. × 8 ч). Сумма ориентиров ≈ 164 ч (+ буфер на срочное).
# Блоки 1–4 (работа с МПП) ≈ 60% времени — ядро роли.
FUNCTIONS_TABLE: list[dict] = [
    {
        "rank": 1,
        "function": "Операционная работа с МПП: разбор канбанов, сделок, следующих шагов и «красных» зон по каждому менеджеру",
        "hours": "45 ч / мес (~27%)",
        "result": "Актуальная картина канбанов; список проблемных сделок; согласованные next steps; прогноз выполнения плана по МПП",
        "needs": "Доступ к CRM/канбану МПП; стадии воронки; дашборд сделок; возможность видеть историю касаний и комментарии",
    },
    {
        "rank": 2,
        "function": "Контроль качества продаж и коучинг: скрипты холодной / тёплой / горячей базы, разбор звонков, стандарты диалога",
        "hours": "26 ч / мес (~15%)",
        "result": "Чек-листы качества; протоколы разборов; скорректированный скрипт; % соответствия стандарту; индивидуальные точки роста МПП",
        "needs": "Записи звонков; актуальные скрипты; критерии оценки; доступ к карточке сделки",
    },
    {
        "rank": 3,
        "function": "Обучение МПП новым продуктам и закрепление навыка на реальных сделках",
        "hours": "18 ч / мес (~11%)",
        "result": "Обучающий пакет; протокол обучения; допуск МПП к продукту; срез знаний; рост конверсии на новом продукте",
        "needs": "Карточка продукта (ценность / ограничения / УТП); демо-доступ; кейсы; эталонные звонки; согласованный оффер",
    },
    {
        "rank": 4,
        "function": "1-1, мотивация, оценка эффективности МПП и планы развития",
        "hours": "12 ч / мес (~7%)",
        "result": "Протоколы 1-1; ИПР; решения по мотивации / кадровым рискам; прозрачная обратная связь сотруднику",
        "needs": "KPI и факт по менеджеру; данные канбана; карта компетенций; правила мотивации",
    },
    {
        "rank": 5,
        "function": "Проверка продуктовых и коммерческих гипотез через команду и рынок",
        "hours": "10 ч / мес (~6%)",
        "result": "Отчёт по гипотезам; вывод go / iterate / stop; правки оффера и скрипта на основе практики",
        "needs": "Сводка возражений МПП; выгрузка воронки из CRM; обратная связь с демо; вводные от продукта / маркетинга",
    },
    {
        "rank": 6,
        "function": "Формирование и актуализация КП, позиционирования и скриптов под запуск",
        "hours": "10 ч / мес (~6%)",
        "result": "КП / one-pager; документ позиционирования; скрипт под сегмент; краткая шпаргалка для МПП",
        "needs": "Вводные от ЦРП / продукта; ограничения лицензирования; конкурентные материалы; согласование формулировок",
    },
    {
        "rank": 7,
        "function": "Взаимодействие со смежными отделами (маркетинг, продукт / ЦРП, поддержка, финблок)",
        "hours": "10 ч / мес (~6%)",
        "result": "Согласованные стыки процессов; протоколы встреч; закрытые эскалации; актуальный статус запуска для команды",
        "needs": "Владельцы процессов; каналы эскалации; статусы рекламных активностей; ограничения продукта",
    },
    {
        "rank": 8,
        "function": "Контроль плана продаж группы и операционных показателей воронки",
        "hours": "8 ч / мес (~5%)",
        "result": "Еженедельный / месячный отчёт по плану; отклонения и план действий; валидация гипотез цифрами",
        "needs": "План продаж; дашборд воронки; выгрузки из CRM; при необходимости сверка с УТ по отгрузкам / оплатам",
    },
    {
        "rank": 9,
        "function": "Бюджет группы, KPI, расчёт мотивации / ЗП",
        "hours": "6 ч / мес (~4%)",
        "result": "Финплан проекта / группы; утверждённые KPI; расчёт мотивации; понятные правила для МПП",
        "needs": "Выгрузки продаж из УТ; правила мотивации; табель / активности CRM; согласованные ставки и пороги",
    },
    {
        "rank": 10,
        "function": "Анализ конкурентов и эффективности маркетинговых подходов на уровне команды",
        "hours": "5 ч / мес (~3%)",
        "result": "Карта конкурентов / входа на рынок; рекомендации в скрипт и оффер; выводы для гипотез",
        "needs": "Открытые источники; данные маркетинга; feedback МПП с линии; материалы конкурентов",
    },
    {
        "rank": 11,
        "function": "Развитие партнёрской / агентской сети совместно с ЦРП",
        "hours": "4 ч / мес (~2%)",
        "result": "Статус партнёрского канала; план касаний; сделки / гипотезы через партнёров",
        "needs": "Условия партнёрок; контакты ЦРП; CRM по партнёрским сделкам",
    },
    {
        "rank": 12,
        "function": "Участие в подборе и онбординге МПП",
        "hours": "4 ч / мес (~2%)",
        "result": "Оценка кандидата; решение по найму; план адаптации на 30 / 60 / 90 дней",
        "needs": "Профиль должности; кейсы собеседования; доступы и чек-лист онбординга",
    },
    {
        "rank": 13,
        "function": "Формализация и оптимизация процессов группы и регламентов взаимодействия",
        "hours": "3 ч / мес (~2%)",
        "result": "Регламент / схема процесса; обновлённые правила эскалации; меньше потерь на стыках",
        "needs": "Описание текущего as-is; боли МПП; согласование смежных отделов",
    },
    {
        "rank": 14,
        "function": "Собственное развитие и передача лучших практик команде",
        "hours": "3 ч / мес (~2%)",
        "result": "Личный план развития; короткие внутренние практики / памятки для МПП",
        "needs": "Доступ к Корпоративному университету / курсам 1С; база лучших звонков",
    },
]


def set_run_font(run, size=10, bold=False, color=NEAR_BLACK, font=FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), font)


def set_cell_shading(cell, color: RGBColor):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    hex_color = f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    for old in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(old)
    tc_pr.append(shd)


def set_cell_borders(cell, color="D0D0D0", sz="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    for old in tc_pr.findall(qn("w:tcBorders")):
        tc_pr.remove(old)
    tc_pr.append(borders)


def set_cell_margins(cell, top=40, bottom=40, left=50, right=50):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for name, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    for old in tc_pr.findall(qn("w:tcMar")):
        tc_pr.remove(old)
    tc_pr.append(tc_mar)


def set_vertical_align(cell, val="center"):
    tc_pr = cell._tc.get_or_add_tcPr()
    v_align = OxmlElement("w:vAlign")
    v_align.set(qn("w:val"), val)
    for old in tc_pr.findall(qn("w:vAlign")):
        tc_pr.remove(old)
    tc_pr.append(v_align)


def clear_paragraph(paragraph):
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.1


def write_cell(
    cell,
    text,
    size=8.5,
    bold=False,
    color=NEAR_BLACK,
    align=WD_ALIGN_PARAGRAPH.LEFT,
    fill=None,
):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    clear_paragraph(paragraph)
    paragraph.alignment = align
    run = paragraph.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    if fill is not None:
        set_cell_shading(cell, fill)
    set_cell_borders(cell)
    set_cell_margins(cell)
    set_vertical_align(cell, "center")


def add_horizontal_line(paragraph, color="FCCD68", size="24"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def set_table_full_width(table):
    tbl_pr = table._tbl.tblPr
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:w"), "5000")
    tbl_w.set(qn("w:type"), "pct")
    for old in tbl_pr.findall(qn("w:tblW")):
        tbl_pr.remove(old)
    tbl_pr.append(tbl_w)


def set_col_widths(table, widths_cm):
    for row in table.rows:
        for idx, width in enumerate(widths_cm):
            row.cells[idx].width = Cm(width)


def add_heading(doc, text):
    p = doc.add_paragraph()
    clear_paragraph(p)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    set_run_font(p.add_run(text), size=12, bold=True)
    add_horizontal_line(p, color="26A6E0", size="16")


def add_body(doc, text, size=9.5, space_after=4, bold=False):
    p = doc.add_paragraph()
    clear_paragraph(p)
    p.paragraph_format.space_after = Pt(space_after)
    set_run_font(p.add_run(text), size=size, bold=bold)


def add_bullet(doc, text, size=9.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.1
    if p.runs:
        p.runs[0].text = text
        set_run_font(p.runs[0], size=size)
    else:
        set_run_font(p.add_run(text), size=size)


def build() -> Path:
    doc = Document()
    section = doc.sections[0]
    # Landscape — таблица функций читается лучше
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)

    brand = doc.add_table(rows=1, cols=2)
    set_table_full_width(brand)
    write_cell(brand.rows[0].cells[0], "ФОРУС", size=16, bold=True, fill=YELLOW)
    write_cell(
        brand.rows[0].cells[1],
        "Группа компаний  ·  Отдел продуктового запуска",
        size=9,
        fill=YELLOW,
        align=WD_ALIGN_PARAGRAPH.RIGHT,
    )
    set_col_widths(brand, [6.0, 21.0])

    title = doc.add_paragraph()
    clear_paragraph(title)
    title.paragraph_format.space_before = Pt(10)
    title.paragraph_format.space_after = Pt(3)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(title.add_run("ДОЛЖНОСТНАЯ ИНСТРУКЦИЯ"), size=14, bold=True)
    add_horizontal_line(title, color="FCCD68", size="28")

    subtitle = doc.add_paragraph()
    clear_paragraph(subtitle)
    subtitle.paragraph_format.space_before = Pt(4)
    subtitle.paragraph_format.space_after = Pt(8)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(subtitle.add_run("Тимлид группы продуктового запуска"), size=12, bold=True)

    meta = doc.add_table(rows=4, cols=4)
    set_table_full_width(meta)
    meta_rows = [
        [("Должность", True), ("Тимлид группы продуктового запуска", False), ("Подразделение", True), ("Продуктовый запуск", False)],
        [("Категория", True), ("Руководитель", False), ("Подчинение", True), ("Руководитель направления", False)],
        [("В подчинении", True), ("Менеджеры по продажам (МПП)", False), ("Версия", True), ("1.1", False)],
        [("ФИО сотрудника", True), ("________________________________", False), ("Дата введения", True), ("«____» ________ 20____ г.", False)],
    ]
    for i, row_data in enumerate(meta_rows):
        for j, (text, is_label) in enumerate(row_data):
            cell = meta.rows[i].cells[j]
            if is_label:
                write_cell(cell, text, size=8, bold=True, fill=LABEL_BG)
            else:
                color = GRAY if "___" in text or text.startswith("«") else NEAR_BLACK
                write_cell(cell, text, size=8.5, color=color, fill=WHITE)
    set_col_widths(meta, [3.8, 9.7, 3.8, 9.7])

    note = doc.add_paragraph()
    clear_paragraph(note)
    note.paragraph_format.space_before = Pt(6)
    note.paragraph_format.space_after = Pt(2)
    set_run_font(
        note.add_run(
            "Основание: карта функций / компетенций группы продуктового запуска (зона Тимлида). "
            "Приоритет роли — управление менеджерами и их сделками."
        ),
        size=8,
        color=GRAY,
    )

    add_heading(doc, "1. Общие положения")
    for t in [
        "1.1. Настоящая должностная инструкция определяет функциональные обязанности, права и ответственность Тимлида группы продуктового запуска ГК Форус.",
        "1.2. Тимлид относится к категории руководителей и назначается / освобождается от должности приказом руководителя в установленном порядке.",
        "1.3. Тимлид непосредственно подчиняется руководителю направления / куратору группы продуктового запуска.",
        "1.4. В непосредственном подчинении Тимлида находятся менеджеры по продажам (МПП) группы продуктового запуска.",
        "1.5. В своей работе Тимлид руководствуется законодательством РФ, локальными нормативными актами ГК Форус, регламентами взаимодействия подразделений, скриптами и стандартами продаж, а также настоящей инструкцией.",
    ]:
        add_body(doc, t)

    add_heading(doc, "2. Цель должности и принцип приоритетов")
    add_body(
        doc,
        "Основная цель деятельности Тимлида — обеспечить результат группы через работу с менеджерами по продажам: "
        "их канбанами, сделками, качеством исполнения стандартов и развитием навыков. "
        "Большую часть рабочего времени Тимлид должен тратить на работу с МПП, а не на самостоятельное «закрытие» операционки вместо команды.",
        bold=False,
    )
    add_body(doc, "Принцип распределения внимания:")
    for item in [
        "Ядро роли (≥ 55–60% времени): канбаны и сделки МПП, коучинг, контроль качества, обучение, 1-1 и мотивация.",
        "Поддержка ядра: гипотезы, КП / скрипты, стыки со смежными отделами, контроль плана.",
        "Обеспечивающие функции (меньший вес по времени, но обязательные): бюджет / KPI / ЗП, конкурентный анализ, найм, регламенты, саморазвитие.",
        "Продуктовый результат достигается через сильную команду МПП: чем лучше Тимлид управляет людьми и сделками, тем быстрее проверяются гипотезы запуска.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "3. Матрица функций: приоритет, время, результат, входы")
    add_body(
        doc,
        "Функции проранжированы от наиболее важных (сильнее влияют на результат роли) к менее важным. "
        "Затраты времени — ориентир на месяц при норме ≈ 168 рабочих часов. Итого ориентиров в таблице ≈ 164 ч; "
        "остаток — буфер на срочные эскалации и запуски.",
        size=8.5,
        space_after=6,
    )

    table = doc.add_table(rows=1 + len(FUNCTIONS_TABLE), cols=5)
    set_table_full_width(table)
    headers = [
        "Ранг\nважности",
        "Функция тимлида",
        "Затрата времени\n(ориентир / мес)",
        "Результат функции\n(документ / артефакт / показатель)",
        "Что требуется для выполнения\n(данные, доступы, входы)",
    ]
    for idx, header in enumerate(headers):
        write_cell(
            table.rows[0].cells[idx],
            header,
            size=8,
            bold=True,
            fill=HEADER_BG,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )

    for i, item in enumerate(FUNCTIONS_TABLE):
        row = table.rows[i + 1]
        bg = WHITE if i % 2 == 0 else ROW_ALT
        # top ranks subtly highlighted
        rank_fill = PRIORITY_BG if item["rank"] <= 4 else bg
        write_cell(row.cells[0], str(item["rank"]), size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, fill=rank_fill)
        write_cell(row.cells[1], item["function"], size=7.5, fill=rank_fill if item["rank"] <= 4 else bg)
        write_cell(row.cells[2], item["hours"], size=7.5, align=WD_ALIGN_PARAGRAPH.CENTER, fill=bg)
        write_cell(row.cells[3], item["result"], size=7.5, fill=bg)
        write_cell(row.cells[4], item["needs"], size=7.5, fill=bg)
        row.height = Cm(1.55)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    set_col_widths(table, [2.0, 7.2, 3.6, 7.0, 7.2])

    legend = doc.add_paragraph()
    clear_paragraph(legend)
    legend.paragraph_format.space_before = Pt(6)
    legend.paragraph_format.space_after = Pt(4)
    set_run_font(
        legend.add_run(
            "Подсветка рангов 1–4: ядро роли — работа с менеджерами (канбаны, качество, обучение, 1-1). "
            "Именно эти функции сильнее всего влияют на результат Тимлида."
        ),
        size=8,
        color=GRAY,
    )

    add_heading(doc, "4. Ключевые зоны ответственности (кратко)")
    add_body(doc, "4.1. Управление МПП и сделками (приоритет №1)", bold=True)
    for item in [
        "Ежедневно / еженедельно разбирать канбаны МПП: стадии, зависшие сделки, качество next step.",
        "Понимать логику каждой ключевой сделки: ЛПР, потребность, продукт, риск, следующий шаг.",
        "Не подменять собой МПП в типовых звонках; усиливать через разбор, коучинг и контроль стандарта.",
        "Держать прозрачную картину выполнения плана по каждому менеджеру и по группе.",
    ]:
        add_bullet(doc, item)

    add_body(doc, "4.2. Продукт, гипотезы и рынок", bold=True)
    for item in [
        "Через практику команды проверять гипотезы запуска и быстро возвращать выводы в оффер / скрипт / продукт.",
        "Поддерживать актуальные КП, позиционирование и материалы для МПП.",
    ]:
        add_bullet(doc, item)

    add_body(doc, "4.3. Стыки, финансы и обеспеченность процесса", bold=True)
    for item in [
        "Синхронизировать маркетинг, продукт / ЦРП, поддержку и финблок так, чтобы МПП не теряли скорость.",
        "Для финконтроля и мотивации использовать выгрузки продаж из УТ, данные CRM и утверждённые правила KPI.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "5. Требования к компетенциям")
    comp = doc.add_table(rows=1, cols=2)
    set_table_full_width(comp)
    write_cell(comp.rows[0].cells[0], "Блок компетенций", size=8.5, bold=True, fill=HEADER_BG, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(comp.rows[0].cells[1], "Ожидаемый уровень", size=8.5, bold=True, fill=HEADER_BG, align=WD_ALIGN_PARAGRAPH.CENTER)
    rows = [
        ("Управление МПП через канбан и сделки", "Свободно читает воронку, видит риски, ставит точные next steps"),
        ("Коучинг и контроль стандартов продаж", "Системно разбирает звонки и удерживает скрипт в команде"),
        ("Обучение продуктам и вывод МПП на новый оффер", "Готовит материалы и доводит до уверенной практики"),
        ("Продуктовые гипотезы и адаптация оффера", "Переводит рынок в решения по продукту / скрипту / КП"),
        ("Взаимодействие со смежными отделами", "Закрывает стыки без потери скорости команды"),
        ("Финансовый и операционный контроль", "Сводит план, УТ, CRM и мотивацию в управляемые решения"),
    ]
    for i, (left, right) in enumerate(rows):
        row = comp.add_row()
        bg = WHITE if i % 2 == 0 else ROW_ALT
        write_cell(row.cells[0], left, size=8, fill=bg)
        write_cell(row.cells[1], right, size=8, fill=bg)
    set_col_widths(comp, [12.0, 15.0])

    add_heading(doc, "6. Права")
    for item in [
        "Запрашивать у смежных подразделений данные и доступы, необходимые для управления командой и запусками (включая выгрузки УТ / CRM).",
        "Вносить предложения по изменению продукта, оффера, скриптов, процессов и KPI группы.",
        "Распределять задачи внутри команды МПП и контролировать их исполнение.",
        "Инициировать обучение, наставничество, оценку и кадровые решения в рамках полномочий.",
        "Эскалировать риски по сделкам, продукту, срокам, качеству исполнения и финансовым отклонениям.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "7. Ответственность")
    for item in [
        "За приоритет работы с МПП и качество управления их сделками / канбанами.",
        "За качество и своевременность выполнения функций настоящей инструкции.",
        "За соблюдение стандартов продаж, регламентов и корпоративных правил ГК Форус.",
        "За достоверность управленческой и операционной отчётности группы.",
        "За результат обучения и контроля МПП в части стандартов, скриптов и работы с продуктом.",
        "За соблюдение конфиденциальности коммерческой и внутренней информации.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "8. Взаимодействия")
    inter = doc.add_table(rows=1, cols=2)
    set_table_full_width(inter)
    write_cell(inter.rows[0].cells[0], "Контрагент", size=8.5, bold=True, fill=HEADER_BG, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(inter.rows[0].cells[1], "Предмет взаимодействия", size=8.5, bold=True, fill=HEADER_BG, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, (who, what) in enumerate([
        ("МПП группы", "Канбаны, сделки, обучение, качество, мотивация, разбор кейсов — основной фокус времени"),
        ("Руководитель направления", "Цели, приоритеты запусков, KPI, эскалации, кадровые решения"),
        ("Маркетинг", "Гипотезы, материалы, рекламные активности, сегментные кампании"),
        ("Продуктовая разработка / ЦРП", "Требования к продукту, ограничения, партнёрская сеть, доработки"),
        ("Финансовая аналитика / УТ", "Выгрузки продаж, финпланы, бюджеты, мотивация, сверка факта"),
        ("Смежные продажи / поддержка", "Передача клиентов, качество сервиса, стыковка процессов"),
    ]):
        row = inter.add_row()
        bg = WHITE if i % 2 == 0 else ROW_ALT
        write_cell(row.cells[0], who, size=8, bold=True, fill=bg)
        write_cell(row.cells[1], what, size=8, fill=bg)
    set_col_widths(inter, [7.0, 20.0])

    add_heading(doc, "9. Результат работы (ключевые показатели)")
    for item in [
        "Доля рабочего времени, фактически потраченная на работу с МПП / сделками / канбанами (целевой ориентир ≥ 55–60%).",
        "Качество ведения канбанов МПП и доля сделок с понятным next step.",
        "Выполнение плана продаж группы и динамика конверсий по этапам воронки.",
        "Уровень соблюдения стандартов / скриптов МПП и качества ведения CRM.",
        "Скорость и качество выхода команды на новый продукт после обучения.",
        "Качество проверки продуктовых и коммерческих гипотез на рынке.",
        "Своевременность стыков со смежными отделами и управляемость бюджета / KPI / мотивации.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "10. Ознакомление")
    sign = doc.add_table(rows=3, cols=2)
    set_table_full_width(sign)
    write_cell(sign.rows[0].cells[0], "Инструкцию утвердил:\n\n________________ / ________________", size=9, fill=WHITE)
    write_cell(sign.rows[0].cells[1], "С инструкцией ознакомлен(а):\n\n________________ / ________________", size=9, fill=WHITE)
    write_cell(sign.rows[1].cells[0], "Должность: _______________________", size=8, color=GRAY, fill=WHITE)
    write_cell(sign.rows[1].cells[1], "Тимлид группы продуктового запуска", size=8, color=GRAY, fill=WHITE)
    write_cell(sign.rows[2].cells[0], "Дата: «____» ______________ 20____ г.", size=8, color=GRAY, fill=WHITE)
    write_cell(sign.rows[2].cells[1], "Дата: «____» ______________ 20____ г.", size=8, color=GRAY, fill=WHITE)
    set_col_widths(sign, [13.5, 13.5])

    foot = doc.add_paragraph()
    clear_paragraph(foot)
    foot.paragraph_format.space_before = Pt(10)
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(
        foot.add_run("© ГК Форус  ·  Внутренний документ  ·  Должностная инструкция тимлида группы продуктового запуска"),
        size=7,
        color=GRAY,
    )

    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(10)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    ART_RU.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    doc.save(OUT_ROOT)
    doc.save(ART_RU)
    doc.save(ART_EN)
    return OUT


if __name__ == "__main__":
    path = build()
    print(f"Saved: {path}")
