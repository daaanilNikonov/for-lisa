#!/usr/bin/env python3
"""Generate Forus-branded job description for Product Launch Team Lead."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "Должностная_инструкция_тимлид_продуктовый_запуск.docx"
OUT_ROOT = ROOT / "Должностная_инструкция_тимлид_продуктовый_запуск.docx"

YELLOW = RGBColor(0xFC, 0xCD, 0x68)
NEAR_BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GRAY = RGBColor(0x76, 0x76, 0x77)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LABEL_BG = RGBColor(0xFF, 0xF3, 0xD6)
ROW_ALT = RGBColor(0xFF, 0xF8, 0xE8)
HEADER_BG = YELLOW
FONT = "Verdana"


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


def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
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
    paragraph.paragraph_format.line_spacing = 1.15


def write_cell(cell, text, size=9, bold=False, color=NEAR_BLACK, align=WD_ALIGN_PARAGRAPH.LEFT, fill=None):
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
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    set_run_font(p.add_run(text), size=12, bold=True)
    add_horizontal_line(p, color="26A6E0", size="16")


def add_body(doc, text, size=9.5, space_after=4):
    p = doc.add_paragraph()
    clear_paragraph(p)
    p.paragraph_format.space_after = Pt(space_after)
    set_run_font(p.add_run(text), size=size)


def add_bullet(doc, text, size=9.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    if p.runs:
        p.runs[0].text = text
        set_run_font(p.runs[0], size=size)
    else:
        set_run_font(p.add_run(text), size=size)


def build() -> Path:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)

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
    set_col_widths(brand, [5.5, 12.0])

    title = doc.add_paragraph()
    clear_paragraph(title)
    title.paragraph_format.space_before = Pt(14)
    title.paragraph_format.space_after = Pt(4)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(title.add_run("ДОЛЖНОСТНАЯ ИНСТРУКЦИЯ"), size=14, bold=True)
    add_horizontal_line(title, color="FCCD68", size="28")

    subtitle = doc.add_paragraph()
    clear_paragraph(subtitle)
    subtitle.paragraph_format.space_before = Pt(6)
    subtitle.paragraph_format.space_after = Pt(10)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(subtitle.add_run("Тимлид группы продуктового запуска"), size=12, bold=True)

    meta = doc.add_table(rows=4, cols=4)
    set_table_full_width(meta)
    meta_rows = [
        [("Должность", True), ("Тимлид группы продуктового запуска", False), ("Подразделение", True), ("Продуктовый запуск", False)],
        [("Категория", True), ("Руководитель", False), ("Подчинение", True), ("Руководитель направления", False)],
        [("В подчинении", True), ("Менеджеры по продажам (МПП)", False), ("Версия", True), ("1.0", False)],
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
    set_col_widths(meta, [3.4, 6.0, 3.4, 5.0])

    note = doc.add_paragraph()
    clear_paragraph(note)
    note.paragraph_format.space_before = Pt(8)
    note.paragraph_format.space_after = Pt(4)
    set_run_font(
        note.add_run(
            "Основание: карта функций / компетенций группы продуктового запуска "
            "(зона ответственности Тимлида), дополненная блоками продуктовой разработки, "
            "взаимодействия со смежными отделами, контроля и обучения МПП."
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

    add_heading(doc, "2. Цель должности")
    add_body(
        doc,
        "Обеспечить успешный вывод и развитие продуктов на рынке: сформировать продуктовую и коммерческую гипотезу, "
        "выстроить продажи через команду МПП, обеспечить качество исполнения стандартов и устойчивый финансовый результат группы.",
    )

    add_heading(doc, "3. Должностные обязанности")
    add_body(doc, "Тимлид группы продуктового запуска выполняет следующие функции:", space_after=6)

    blocks = [
        (
            "3.1. Продуктовая разработка и адаптация продукта",
            [
                "Участвует в продуктовой разработке: формулирует требования к продукту с опорой на рынок, обратную связь клиентов и результаты продаж.",
                "Адаптирует продукт под потребности сегмента: прорабатывает продуктовые гипотезы, проверяет ценность и ограничения предложения.",
                "Формирует и актуализирует коммерческие предложения (КП) под сегменты и сценарии продаж.",
                "Разрабатывает позиционирование продукта для сегментов и готовит документы по позиционированию.",
                "Собирает и оптимизирует продуктовый портфель группы; формирует комплексные предложения из портфеля ГК для работы с клиентом.",
                "Участвует в подготовке маркетинговых материалов и отслеживает новости сегмента, влияющие на продукт и оффер.",
            ],
        ),
        (
            "3.2. Анализ рынка, конкурентов и маркетинговых гипотез",
            [
                "Анализирует конкурентную среду, формирует карту входа продукта на рынок и использует выводы для КП и скриптов.",
                "Анализирует успешность маркетинговых стратегий и внедряет рабочие подходы на уровне команды.",
                "Организует сбор и интерпретацию обратной связи с рынка для корректировки гипотез, оффера и продуктовых решений.",
            ],
        ),
        (
            "3.3. Взаимодействие со смежными отделами",
            [
                "Выстраивает и формализует процессы взаимодействия со смежными подразделениями: маркетинг, финансовая аналитика, продуктовая разработка / ЦРП, поддержка и другие связанные функции.",
                "Определяет и поддерживает актуальные регламенты взаимодействия МПП со смежными отделами.",
                "Согласовывает изменения процессов внутри группы и на стыке с другими подразделениями.",
                "Обеспечивает своевременный обмен информацией по гипотезам, рекламным активностям, ограничениям продукта и статусам запусков.",
                "Участвует в развитии партнёрской / агентской сети совместно с ЦРП по закреплённым продуктам.",
            ],
        ),
        (
            "3.4. Управление командой МПП: контроль, обучение и развитие",
            [
                "Внедряет и контролирует стандарты продаж в команде (скрипты холодной / тёплой / горячей базы, качество диалога, фиксация в CRM).",
                "Обучает МПП: готовит и актуализирует обучающие материалы по новым продуктам, проводит разборы, наставничество и практику.",
                "Контролирует работу МПП по качеству и результативности: аналитика, дашборды, правила заполнения карточек, соблюдение процессов.",
                "Проводит оценочные и мотивационные интервью, выявляет зоны роста и формирует индивидуальные планы развития.",
                "Отвечает за мотивацию команды и поддержание рабочего ритма при запусках и работе с холодной базой.",
                "Участвует в формировании команды и подборе сотрудников в группу.",
                "Оптимизирует рабочие процессы группы и обеспечивает прозрачность ролей, зон ответственности и правил эскалации.",
            ],
        ),
        (
            "3.5. Операционный и финансовый контроль",
            [
                "Контролирует выполнение плана продаж команды для проверки продуктовых и коммерческих гипотез.",
                "Контролирует финансовые показатели группы и качество клиентского сервиса.",
                "Планирует бюджет направления / группы, формирует финпланы по проектам, участвует в сборе данных для кросс-распределения.",
                "Согласовывает рекламные бюджеты, связанные с запусками продуктов группы.",
                "Утверждает KPI группы и обеспечивает понятность метрик для команды.",
                "Участвует в расчёте заработной платы / мотивации группы в рамках установленных правил.",
            ],
        ),
        (
            "3.6. Собственное развитие",
            [
                "Развивает управленческие и продуктовые компетенции: обучение 1С, программы Корпоративного университета, управленческая литература и практика.",
                "Систематизирует лучшие практики запусков и передаёт их команде.",
            ],
        ),
    ]
    for title_text, items in blocks:
        p = doc.add_paragraph()
        clear_paragraph(p)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        set_run_font(p.add_run(title_text), size=10.5, bold=True)
        for item in items:
            add_bullet(doc, item)

    add_heading(doc, "4. Требования к компетенциям")
    comp = doc.add_table(rows=1, cols=2)
    set_table_full_width(comp)
    write_cell(comp.rows[0].cells[0], "Блок компетенций", size=9, bold=True, fill=HEADER_BG, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(comp.rows[0].cells[1], "Ожидаемый уровень", size=9, bold=True, fill=HEADER_BG, align=WD_ALIGN_PARAGRAPH.CENTER)
    rows = [
        ("Продуктовая разработка и адаптация продукта под сегмент", "Уверенное владение: гипотезы, ценность, ограничения, КП"),
        ("Позиционирование и сборка портфельных предложений", "Самостоятельно формирует и актуализирует материалы"),
        ("Анализ конкурентов и маркетинговых гипотез", "Применяет выводы в продажах и запусках"),
        ("Взаимодействие со смежными отделами и регламентация стыков", "Выстраивает понятные процессы без потерь на передаче"),
        ("Контроль МПП: стандарты, CRM, качество, план", "Системный контроль через аналитику и регулярные разборы"),
        ("Обучение и развитие МПП по новым продуктам", "Готовит материалы, проводит обучение, закрепляет навык"),
        ("Мотивация команды и кадровые решения", "Поддерживает результат и вовлечённость команды"),
        ("Финансовое планирование, KPI и бюджетирование группы", "Планирует, контролирует и объясняет показатели команде"),
        ("Управленческое саморазвитие", "Регулярно усиливает управленческую и продуктовую экспертизу"),
    ]
    for i, (left, right) in enumerate(rows):
        row = comp.add_row()
        bg = WHITE if i % 2 == 0 else ROW_ALT
        write_cell(row.cells[0], left, size=8.5, fill=bg)
        write_cell(row.cells[1], right, size=8.5, fill=bg)
    set_col_widths(comp, [9.0, 8.5])

    add_heading(doc, "5. Права")
    for item in [
        "Запрашивать у смежных подразделений информацию, необходимую для запусков, продаж и анализа гипотез.",
        "Вносить предложения по изменению продукта, оффера, скриптов, процессов и KPI группы.",
        "Распределять задачи внутри команды МПП и контролировать их исполнение.",
        "Инициировать обучение, наставничество, оценку и кадровые решения в рамках полномочий.",
        "Эскалировать риски по продукту, срокам, качеству исполнения и финансовым отклонениям.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "6. Ответственность")
    for item in [
        "За качество и своевременность выполнения функций, указанных в настоящей инструкции.",
        "За соблюдение стандартов продаж, регламентов и корпоративных правил ГК Форус.",
        "За достоверность управленческой и операционной отчётности группы.",
        "За результат обучения и контроля МПП в части стандартов, скриптов и работы с продуктом.",
        "За соблюдение конфиденциальности коммерческой и внутренней информации.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "7. Взаимодействия")
    inter = doc.add_table(rows=1, cols=2)
    set_table_full_width(inter)
    write_cell(inter.rows[0].cells[0], "Контрагент", size=9, bold=True, fill=HEADER_BG, align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(inter.rows[0].cells[1], "Предмет взаимодействия", size=9, bold=True, fill=HEADER_BG, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, (who, what) in enumerate([
        ("Руководитель направления", "Цели, приоритеты запусков, KPI, эскалации, кадровые решения"),
        ("МПП группы", "Постановка задач, обучение, контроль качества, мотивация, разбор кейсов"),
        ("Маркетинг", "Гипотезы, материалы, рекламные активности, сегментные кампании"),
        ("Продуктовая разработка / ЦРП", "Требования к продукту, ограничения, партнёрская сеть, доработки"),
        ("Финансовая аналитика", "Финпланы, бюджеты, показатели, мотивация"),
        ("Смежные продажи / поддержка", "Передача клиентов, качество сервиса, стыковка процессов"),
    ]):
        row = inter.add_row()
        bg = WHITE if i % 2 == 0 else ROW_ALT
        write_cell(row.cells[0], who, size=8.5, bold=True, fill=bg)
        write_cell(row.cells[1], what, size=8.5, fill=bg)
    set_col_widths(inter, [6.0, 11.5])

    add_heading(doc, "8. Результат работы (ключевые показатели)")
    for item in [
        "Качество проверки продуктовых и коммерческих гипотез на рынке.",
        "Выполнение плана продаж группы и динамика конверсий по этапам воронки.",
        "Уровень соблюдения стандартов / скриптов МПП и качества ведения CRM.",
        "Скорость и качество выхода команды на новый продукт после обучения.",
        "Своевременность и прозрачность взаимодействия со смежными отделами.",
        "Исполнение бюджета / KPI группы и управляемость финансовой мотивации.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "9. Ознакомление")
    sign = doc.add_table(rows=3, cols=2)
    set_table_full_width(sign)
    write_cell(sign.rows[0].cells[0], "Инструкцию утвердил:\n\n________________ / ________________", size=9, fill=WHITE)
    write_cell(sign.rows[0].cells[1], "С инструкцией ознакомлен(а):\n\n________________ / ________________", size=9, fill=WHITE)
    write_cell(sign.rows[1].cells[0], "Должность: _______________________", size=8, color=GRAY, fill=WHITE)
    write_cell(sign.rows[1].cells[1], "Тимлид группы продуктового запуска", size=8, color=GRAY, fill=WHITE)
    write_cell(sign.rows[2].cells[0], "Дата: «____» ______________ 20____ г.", size=8, color=GRAY, fill=WHITE)
    write_cell(sign.rows[2].cells[1], "Дата: «____» ______________ 20____ г.", size=8, color=GRAY, fill=WHITE)
    set_col_widths(sign, [8.75, 8.75])

    foot = doc.add_paragraph()
    clear_paragraph(foot)
    foot.paragraph_format.space_before = Pt(14)
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
    doc.save(OUT)
    doc.save(OUT_ROOT)
    return OUT


if __name__ == "__main__":
    path = build()
    print(f"Saved: {path}")
