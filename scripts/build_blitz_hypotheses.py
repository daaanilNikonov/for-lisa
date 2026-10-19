#!/usr/bin/env python3
"""Отчёт по гипотезам проекта «Блиц-запись» — шаблон ГК Форус (тёмный 16×9)."""

from __future__ import annotations

import shutil
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "Презентация ГК Форус темный шаблон 16х9 (1).pptx"
OUT = ROOT / "presentation" / "Блиц-запись_Отчет_по_гипотезам.pptx"
ICONS = ROOT / "presentation" / "assets" / "icons"
CHARTS = ROOT / "presentation" / "assets" / "charts"

BLUE = RGBColor(0x26, 0xA6, 0xE0)
CARD = RGBColor(0x3F, 0x3F, 0x3F)
CARD_LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
GRAY = RGBColor(0x76, 0x76, 0x77)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xBF, 0xBF, 0xBF)
NEAR_BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GREEN_OK = RGBColor(0x4C, 0xAF, 0x50)
ORANGE = RGBColor(0xE0, 0xA0, 0x26)
RED_MUTED = RGBColor(0xC0, 0x60, 0x60)

FONT = "Verdana"

L_TITLE = 0
L_CONTENT = 3
L_CONTENT2 = 4
L_BG = 6
L_EMPTY = 22


def emu(inches: float) -> int:
    return int(Inches(inches))


def delete_all_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        r_id = sld_id.get(qn("r:id"))
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def set_run(run, text, size_pt, bold=False, color=WHITE, font_name=FONT):
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = r_pr.find(qn(tag))
        if el is None:
            el = etree.SubElement(r_pr, qn(tag))
        el.set("typeface", font_name)


def set_anchor(text_frame, anchor=MSO_ANCHOR.TOP):
    body_pr = text_frame._txBody.find(qn("a:bodyPr"))
    if body_pr is not None:
        body_pr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(
                anchor, "t"
            ),
        )


def add_textbox(
    slide,
    left,
    top,
    width,
    height,
    text,
    size_pt=14,
    bold=False,
    color=WHITE,
    align=PP_ALIGN.LEFT,
    font_name=FONT,
    anchor=MSO_ANCHOR.TOP,
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    set_anchor(tf, anchor)
    tf.paragraphs[0].alignment = align
    set_run(tf.paragraphs[0].add_run(), text, size_pt, bold, color, font_name)
    return box


def add_multiline(
    slide,
    left,
    top,
    width,
    height,
    lines,
    size_pt=13,
    color=WHITE,
    bold=False,
    align=PP_ALIGN.LEFT,
    line_spacing=1.15,
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    set_anchor(tf, MSO_ANCHOR.TOP)
    first = True
    for text, *opts in lines:
        c = opts[0] if opts else color
        b = opts[1] if len(opts) > 1 else bold
        s = opts[2] if len(opts) > 2 else size_pt
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        set_run(p.add_run(), text, s, b, c)
    return box


def add_card(slide, left, top, width, height, fill=CARD, corner=0.08):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    try:
        shape.adjustments[0] = corner
    except Exception:
        pass
    return shape


def add_icon(slide, name: str, left, top, width, height):
    path = ICONS / name
    if not path.exists():
        return None
    return slide.shapes.add_picture(str(path), left, top, width, height)


def add_chart(slide, name: str, left, top, width, height):
    path = CHARTS / name
    if not path.exists():
        return None
    return slide.shapes.add_picture(str(path), left, top, width, height)


def clear_body_placeholders(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx in (13, 14, 1, 2):
            if ph.has_text_frame:
                ph.text_frame.clear()


def fill_title(slide, text, size=26):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            set_run(p.add_run(), text, size, True, WHITE)
            return ph
    return add_textbox(slide, emu(0.97), emu(0.48), emu(11.4), emu(1.0), text, size, True, WHITE)


def style_table(table, header_blue=True, highlight_last=False):
    n = len(table.rows)
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            for child in list(tc_pr):
                if any(x in child.tag for x in ("solidFill", "blipFill", "gradFill")):
                    tc_pr.remove(child)
            solid = etree.SubElement(tc_pr, qn("a:solidFill"))
            srgb = etree.SubElement(solid, qn("a:srgbClr"))
            if r == 0 and header_blue:
                srgb.set("val", "26A6E0")
                color, bold, size = WHITE, True, 10
            elif highlight_last and r == n - 1:
                srgb.set("val", "3F3F3F")
                color, bold, size = WHITE, True, 11
            else:
                srgb.set("val", "2A2A2A" if r % 2 else "1F1F1F")
                color, bold, size = WHITE, c == 0, 10

            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
                text = p.text
                for el in list(p._p):
                    if el.tag.endswith("}r"):
                        p._p.remove(el)
                set_run(p.add_run(), text, size, bold, color)


def status_color(status: str):
    s = status.lower()
    if "подтверд" in s:
        return GREEN_OK
    if "не подтверд" in s:
        return RED_MUTED
    if "проверена" in s or "подтверждена" in s:
        return GREEN_OK
    if "не начали" in s or "проработ" in s or "обсужд" in s:
        return ORANGE
    return SOFT


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------

def build_slide_1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_TITLE])
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Блиц-запись", 30, True, BLUE)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "Отчёт по проработанным гипотезам", 24, True, WHITE)
            p3 = tf.add_paragraph()
            p3.space_before = Pt(12)
            set_run(p3.add_run(), "Группа продуктового запуска · сентябрь – октябрь 2026", 14, False, SOFT)
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(
                p.add_run(),
                "Статус проверки · способ проверки · цифры по базам · конверсии в демонстрацию",
                13,
                False,
                SOFT,
            )
    add_icon(slide, "icon_69.png", emu(10.55), emu(0.5), emu(1.8), emu(1.8))
    return slide


def build_slide_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Сводный статус гипотез", 26)
    clear_body_placeholders(slide)

    rows = [
        ["№", "Гипотеза (кратко)", "Статус"],
        ["1", "Холодная база: бесплатное сотрудничество → рост конверсии в демо", "Подтверждена"],
        ["2", "Тёплые клиенты (АБУ): конверсия в демо выше холодной на 30%", "Не подтверждена"],
        ["3", "Юклаенс: упор на техподдержку и штат → конверсия >40%", "Проверена"],
        ["4", "ЦА шире разовых услуг: спорт, образование, кабинеты специалистов", "Не начата"],
        ["5", "Продвижение через крупные франшизы целевого сегмента", "Не начата"],
    ]

    table_shape = slide.shapes.add_table(len(rows), 3, emu(0.97), emu(1.55), emu(11.5), emu(4.4))
    table = table_shape.table
    table.columns[0].width = emu(0.7)
    table.columns[1].width = emu(8.0)
    table.columns[2].width = emu(2.8)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table)

    # Перекраска статусов
    for r in range(1, len(rows)):
        cell = table.cell(r, 2)
        for p in cell.text_frame.paragraphs:
            text = p.text
            for el in list(p._p):
                if el.tag.endswith("}r"):
                    p._p.remove(el)
            set_run(p.add_run(), text, 11, True, status_color(rows[r][2]))

    add_textbox(
        slide,
        emu(0.97),
        emu(6.15),
        emu(11.5),
        emu(0.5),
        "Целевое действие во всех проверках — демонстрация сервиса. Все проведённые демо — по холодной базе.",
        11,
        False,
        SOFT,
    )
    return slide


def build_slide_3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 1. Холодная база — бесплатное сотрудничество", 22)
    clear_body_placeholders(slide)

    add_card(slide, emu(0.97), emu(1.5), emu(11.5), emu(1.55), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.65), emu(11.0), emu(0.35), "Формулировка", 12, True, BLUE)
    add_textbox(
        slide,
        emu(1.2),
        emu(2.05),
        emu(11.0),
        emu(0.8),
        "При работе с холодной базой (салоны красоты) конверсия в демонстрацию растёт, "
        "если предлагать бесплатное сотрудничество с последующим переходом на платный пакет.",
        13,
        False,
        WHITE,
    )

    # Способ проверки + результат cards
    add_card(slide, emu(0.97), emu(3.25), emu(5.55), emu(3.0), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(3.4), emu(5.1), emu(0.35), "Способ проверки", 13, True, BLUE)
    add_multiline(
        slide,
        emu(1.2),
        emu(3.9),
        emu(5.1),
        emu(2.2),
        [
            ("• Составление скрипта продаж с оффером бесплатного старта", WHITE, False, 12),
            ("• Прозвон холодной базы салонов красоты", WHITE, False, 12),
            ("• Фиксация воронки: обзвон → сделка → демо", WHITE, False, 12),
            ("• Сравнение конверсии до и после смены оффера", WHITE, False, 12),
        ],
    )

    add_card(slide, emu(6.75), emu(3.25), emu(5.7), emu(3.0), BLUE, 0.06)
    add_textbox(slide, emu(7.0), emu(3.4), emu(5.2), emu(0.35), "Результат", 13, True, WHITE)
    add_multiline(
        slide,
        emu(7.0),
        emu(3.9),
        emu(5.2),
        emu(2.2),
        [
            ("Подтверждена", WHITE, True, 18),
            ("", WHITE, False, 8),
            ("Конверсия сделка → демо: 20%", WHITE, True, 14),
            ("(было 0% до смены оффера)", WHITE, False, 12),
            ("", WHITE, False, 8),
            ("Все демонстрации — с холодной базы", WHITE, False, 12),
        ],
    )
    return slide


def build_slide_4(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 1. Цифры по периодам (холодная база)", 24)
    clear_body_placeholders(slide)

    rows = [
        ["Период", "Обработано", "Сделок", "Демо", "Обзвон→сделка", "Сделка→демо"],
        ["14.09–19.09", "102", "13", "0", "12,75%", "0%"],
        ["21.09–28.09", "62", "5", "0", "8,06%", "0%"],
        ["28.09–05.10", "52", "12", "3", "23,08%", "25%"],
        ["05.10–12.10", "47", "10", "2", "21,28%", "20%"],
        ["Итого", "263", "40", "5", "15,2%", "12,5%*"],
    ]

    table_shape = slide.shapes.add_table(len(rows), 6, emu(0.97), emu(1.45), emu(11.5), emu(3.35))
    table = table_shape.table
    widths = [1.9, 1.7, 1.5, 1.3, 2.5, 2.6]
    for i, w in enumerate(widths):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)

    add_textbox(
        slide,
        emu(0.97),
        emu(4.95),
        emu(11.5),
        emu(0.55),
        "* Итоговая сделка→демо по всем периодам (5/40). Актуальный рабочий показатель после смены оффера — 20% "
        "(период 05–12.10). До оффера (14–28.09) — 0 демо при 164 обработанных.",
        11,
        False,
        SOFT,
    )

    # KPI chips
    chips = [
        ("263", "обработано"),
        ("5", "демонстраций"),
        ("20%", "конверсия сейчас"),
        ("0% → 20%", "динамика"),
    ]
    left0 = emu(0.97)
    for i, (val, label) in enumerate(chips):
        left = left0 + i * emu(2.95)
        add_card(slide, left, emu(5.6), emu(2.75), emu(1.0), CARD, 0.08)
        add_textbox(slide, left + emu(0.1), emu(5.68), emu(2.55), emu(0.45), val, 18, True, BLUE, PP_ALIGN.CENTER)
        add_textbox(slide, left + emu(0.1), emu(6.15), emu(2.55), emu(0.35), label, 11, False, SOFT, PP_ALIGN.CENTER)
    return slide


def build_slide_5(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_BG])
    fill_title(slide, "Гипотеза 1. Динамика конверсии и сравнение баз", 22)
    clear_body_placeholders(slide)

    add_chart(slide, "cold_conversion_dynamics.png", emu(0.7), emu(1.45), emu(7.0), emu(3.15))
    add_chart(slide, "cold_vs_abu.png", emu(7.6), emu(1.45), emu(5.3), emu(3.0))

    add_card(slide, emu(0.97), emu(4.85), emu(11.5), emu(1.7), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(5.0), emu(11.0), emu(0.35), "Вывод по гипотезе 1", 13, True, BLUE)
    add_textbox(
        slide,
        emu(1.2),
        emu(5.45),
        emu(11.0),
        emu(0.9),
        "После внедрения оффера «бесплатное сотрудничество → платный пакет» конверсия сделка→демо "
        "выросла с 0% до 20–25%. На периоде 05–12.10: 47 обработано, 10 сделок, 2 демо (20%). "
        "Демонстраций с базы АБУ за тот же горизонт — 0.",
        13,
        False,
        WHITE,
    )
    return slide


def build_slide_6(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 1. Воронка холодной базы, 01–15.10", 24)
    clear_body_placeholders(slide)

    add_chart(slide, "oct_funnel.png", emu(0.7), emu(1.4), emu(7.4), emu(3.3))

    add_card(slide, emu(8.3), emu(1.5), emu(4.2), emu(4.5), CARD, 0.06)
    add_textbox(slide, emu(8.55), emu(1.7), emu(3.7), emu(0.35), "Сводка за 01–15.10", 13, True, BLUE)
    metrics = [
        ("Звонков", "242"),
        ("Контакт с ЛПР", "123"),
        ("Интерес / КП", "22"),
        ("Презентаций", "6"),
        ("Тестовый доступ", "2"),
        ("Онбординг / ЛК", "2"),
    ]
    y = emu(2.2)
    for label, val in metrics:
        add_textbox(slide, emu(8.55), y, emu(2.4), emu(0.35), label, 12, False, SOFT)
        add_textbox(slide, emu(10.9), y, emu(1.3), emu(0.35), val, 14, True, WHITE, PP_ALIGN.RIGHT)
        y += emu(0.48)

    add_textbox(
        slide,
        emu(0.97),
        emu(5.9),
        emu(11.5),
        emu(0.7),
        "Источник: ежедневная воронка прозвона холодной базы салонов красоты. "
        "За октябрь по сделкам (срез на 12.10): 5 демо проведено, 1 клиент подключён; конверсия демо→подключение — 20%.",
        11,
        False,
        SOFT,
    )
    return slide


def build_slide_7(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 2. Тёплые клиенты (база АБУ)", 24)
    clear_body_placeholders(slide)

    add_card(slide, emu(0.97), emu(1.45), emu(11.5), emu(1.35), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.55), emu(11.0), emu(0.3), "Формулировка", 12, True, BLUE)
    add_textbox(
        slide,
        emu(1.2),
        emu(1.95),
        emu(11.0),
        emu(0.7),
        "При проработке текущих клиентов ГК Форус конверсия в демонстрацию будет выше на 30%, "
        "чем на холодной базе: клиенты тёплые и уже знают компанию.",
        13,
        False,
        WHITE,
    )

    # Method + status
    add_card(slide, emu(0.97), emu(3.0), emu(5.55), emu(2.0), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(3.15), emu(5.1), emu(0.3), "Способ проверки", 13, True, BLUE)
    add_multiline(
        slide,
        emu(1.2),
        emu(3.55),
        emu(5.1),
        emu(1.3),
        [
            ("• Скрипт для базы АБУ", WHITE, False, 12),
            ("• Прозвон тёплых клиентов", WHITE, False, 12),
            ("• Сравнение конверсии с холодной базой", WHITE, False, 12),
        ],
    )

    add_card(slide, emu(6.75), emu(3.0), emu(5.7), emu(2.0), CARD, 0.06)
    add_textbox(slide, emu(7.0), emu(3.15), emu(5.2), emu(0.3), "Статус", 13, True, BLUE)
    add_multiline(
        slide,
        emu(7.0),
        emu(3.55),
        emu(5.2),
        emu(1.3),
        [
            ("Не подтверждена", RED_MUTED, True, 16),
            ("Маленькая база и мало звонков", WHITE, False, 12),
            ("Меняем заход, повторим проверку", WHITE, False, 12),
        ],
    )

    rows = [
        ["Период", "Обработано", "Сделок", "Демо", "Обзвон→сделка", "Сделка→демо"],
        ["25.09–28.09", "3", "0", "0", "0%", "—"],
        ["28.09–05.10", "43", "6", "0", "13,95%", "0%"],
        ["05.10–12.10", "2", "0", "0", "0%", "—"],
        ["Итого", "48", "6", "0", "12,5%", "0%"],
    ]
    table_shape = slide.shapes.add_table(len(rows), 6, emu(0.97), emu(5.2), emu(11.5), emu(1.55))
    table = table_shape.table
    widths = [1.9, 1.7, 1.5, 1.3, 2.5, 2.6]
    for i, w in enumerate(widths):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)
    return slide


def build_slide_8(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 3. Юклаенс — техподдержка и штат", 24)
    clear_body_placeholders(slide)

    add_card(slide, emu(0.97), emu(1.45), emu(11.5), emu(1.55), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.55), emu(11.0), emu(0.3), "Формулировка", 12, True, BLUE)
    add_textbox(
        slide,
        emu(1.2),
        emu(1.95),
        emu(11.0),
        emu(0.9),
        "На этапе сбора потребности у клиентов Юклаенс важно уточнять линию технической поддержки "
        "и количество сотрудников — именно здесь типичная боль. Презентация Блиц-записи с акцентом "
        "на эти точки даёт высокую конверсию в демо и в сделку (более 40%).",
        13,
        False,
        WHITE,
    )

    add_card(slide, emu(0.97), emu(3.2), emu(5.55), emu(1.7), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(3.35), emu(5.1), emu(0.3), "Способ проверки", 13, True, BLUE)
    add_multiline(
        slide,
        emu(1.2),
        emu(3.8),
        emu(5.1),
        emu(1.0),
        [
            ("• Доработка скрипта сбора потребности", WHITE, False, 12),
            ("• Акцент в презентации на ТП и штат", WHITE, False, 12),
            ("• Замер конверсии в демо и сделку", WHITE, False, 12),
        ],
    )

    add_card(slide, emu(6.75), emu(3.2), emu(5.7), emu(1.7), BLUE, 0.06)
    add_textbox(slide, emu(7.0), emu(3.35), emu(5.2), emu(0.3), "Статус", 13, True, WHITE)
    add_multiline(
        slide,
        emu(7.0),
        emu(3.8),
        emu(5.2),
        emu(1.0),
        [
            ("Проверена — подтверждается", WHITE, True, 15),
            ("Конверсия в демо/сделку >40%", WHITE, False, 13),
            ("Цифры по сегментам — ниже", WHITE, False, 12),
        ],
    )

    add_textbox(
        slide,
        emu(0.97),
        emu(5.1),
        emu(11.5),
        emu(0.35),
        "Распределение прозвоненных клиентов по текущим сервисам записи (заполнить)",
        13,
        True,
        BLUE,
    )

    rows = [
        ["Сервис", "Кол-во прозвоненных", "Демо", "Конверсия"],
        ["Дикиди", "", "", ""],
        ["Юклаенс", "", "", ""],
        ["Мастерс", "", "", ""],
        ["Итого", "", "", ""],
    ]
    table_shape = slide.shapes.add_table(len(rows), 4, emu(0.97), emu(5.5), emu(11.5), emu(1.35))
    table = table_shape.table
    for i, w in enumerate([3.5, 3.0, 2.5, 2.5]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)
    return slide


def build_slide_9(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотезы 4 и 5 — новые, проверка не начата", 24)
    clear_body_placeholders(slide)

    # H4
    add_card(slide, emu(0.97), emu(1.5), emu(11.5), emu(2.35), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.65), emu(8.5), emu(0.35), "Гипотеза 4. Расширение ЦА", 15, True, BLUE)
    add_card(slide, emu(9.5), emu(1.7), emu(2.7), emu(0.4), ORANGE, 0.4)
    add_textbox(
        slide, emu(9.5), emu(1.7), emu(2.7), emu(0.4),
        "Не начата", 12, True, NEAR_BLACK, PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        emu(1.2),
        emu(2.2),
        emu(11.0),
        emu(0.9),
        "Продукт нужен не только сегментам с постоянными/разовыми услугами (салоны, автомойки), "
        "но и тем, кто ведёт запись на занятия и разовые сессии: спорт-клубы, образовательные центры, "
        "кабинеты психолога, нутрициолога и др.",
        13,
        False,
        WHITE,
    )
    add_textbox(
        slide,
        emu(1.2),
        emu(3.2),
        emu(11.0),
        emu(0.4),
        "План: активная проработка с ноября. Способ — скрипт + прозвон новых сегментов + замер конверсии в демо.",
        12,
        False,
        SOFT,
    )

    # H5
    add_card(slide, emu(0.97), emu(4.05), emu(11.5), emu(2.5), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(4.2), emu(8.5), emu(0.35), "Гипотеза 5. Крупные франшизы", 15, True, BLUE)
    add_card(slide, emu(9.5), emu(4.25), emu(2.7), emu(0.4), ORANGE, 0.4)
    add_textbox(
        slide, emu(9.5), emu(4.25), emu(2.7), emu(0.4),
        "Требует обсуждения", 11, True, NEAR_BLACK, PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        emu(1.2),
        emu(4.8),
        emu(11.0),
        emu(0.9),
        "Для текущей позиции продукта продвижение возможно через работу с крупными франшизами "
        "подходящего целевого сегмента (образовательные центры, фитнес-клубы и др.).",
        13,
        False,
        WHITE,
    )
    add_textbox(
        slide,
        emu(1.2),
        emu(5.8),
        emu(11.0),
        emu(0.5),
        "Статус: проверка не начата. Нужно согласовать подход, приоритетные сети и критерии успеха.",
        12,
        False,
        SOFT,
    )
    return slide


def build_slide_10(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Итоги и следующие шаги", 26)
    clear_body_placeholders(slide)

    items = [
        (
            "1",
            "Закрепить оффер холодной базы",
            "Бесплатное сотрудничество → платный пакет. Рабочая конверсия сделка→демо — 20%.",
            BLUE,
        ),
        (
            "2",
            "Перезапустить базу АБУ",
            "Гипотеза не подтвердилась на малой выборке (48 обработано, 0 демо). Меняем заход и повторяем.",
            ORANGE,
        ),
        (
            "3",
            "Юклаенс — зафиксировать цифры",
            "Гипотеза проверена. Дозаполнить таблицу: Дикиди / Юклаенс / Мастерс.",
            GREEN_OK,
        ),
        (
            "4",
            "Ноябрь: новые сегменты ЦА",
            "Старт гипотезы 4 (спорт, образование, кабинеты специалистов).",
            SOFT,
        ),
        (
            "5",
            "Обсудить франшизы",
            "Гипотеза 5 — без старта до согласования подхода и приоритетных сетей.",
            SOFT,
        ),
    ]

    y = emu(1.5)
    for num, title, body, accent in items:
        add_card(slide, emu(0.97), y, emu(11.5), emu(0.85), CARD, 0.06)
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(1.15), y + emu(0.22), emu(0.42), emu(0.42))
        circ.fill.solid()
        circ.fill.fore_color.rgb = accent
        circ.line.fill.background()
        add_textbox(
            slide, emu(1.15), y + emu(0.22), emu(0.42), emu(0.42),
            num, 12, True, WHITE if accent != SOFT else NEAR_BLACK,
            PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_textbox(slide, emu(1.8), y + emu(0.12), emu(10.3), emu(0.32), title, 13, True, WHITE)
        add_textbox(slide, emu(1.8), y + emu(0.45), emu(10.3), emu(0.35), body, 11, False, SOFT)
        y += emu(0.95)
    return slide


def main():
    if not TEMPLATE.exists():
        raise SystemExit(f"Template not found: {TEMPLATE}")
    if not CHARTS.exists():
        raise SystemExit(f"Charts not found: {CHARTS}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(TEMPLATE, OUT)

    prs = Presentation(str(OUT))
    delete_all_slides(prs)

    build_slide_1(prs)
    build_slide_2(prs)
    build_slide_3(prs)
    build_slide_4(prs)
    build_slide_5(prs)
    build_slide_6(prs)
    build_slide_7(prs)
    build_slide_8(prs)
    build_slide_9(prs)
    build_slide_10(prs)

    prs.save(str(OUT))
    print(f"Saved: {OUT}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
