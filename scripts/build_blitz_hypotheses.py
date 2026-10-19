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
L_BG = 6


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
    """lines: list of str or (text, color, bold, size)."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    set_anchor(tf, MSO_ANCHOR.TOP)
    first = True
    for item in lines:
        if isinstance(item, str):
            text, c, b, s = item, color, bold, size_pt
        else:
            text = item[0]
            c = item[1] if len(item) > 1 else color
            b = item[2] if len(item) > 2 else bold
            s = item[3] if len(item) > 3 else size_pt
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


def fill_title(slide, text, size=24):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            set_run(p.add_run(), text, size, True, WHITE)
            return ph
    return add_textbox(slide, emu(0.97), emu(0.48), emu(11.4), emu(1.0), text, size, True, WHITE)


def style_table(table, highlight_last=False):
    n = len(table.rows)
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            for child in list(tc_pr):
                if any(x in child.tag for x in ("solidFill", "blipFill", "gradFill")):
                    tc_pr.remove(child)
            solid = etree.SubElement(tc_pr, qn("a:solidFill"))
            srgb = etree.SubElement(solid, qn("a:srgbClr"))
            if r == 0:
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


def status_badge(slide, left, top, width, height, text, fill):
    add_card(slide, left, top, width, height, fill, 0.4)
    fg = NEAR_BLACK if fill in (ORANGE, GREEN_OK) else WHITE
    add_textbox(
        slide, left, top, width, height, text, 11, True, fg,
        PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------

def build_slide_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_TITLE])
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            set_run(tf.paragraphs[0].add_run(), "Блиц-запись", 30, True, BLUE)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "Отчёт по продуктовым гипотезам", 24, True, WHITE)
            p3 = tf.add_paragraph()
            p3.space_before = Pt(12)
            set_run(
                p3.add_run(),
                "Группа продуктового запуска · сентябрь – октябрь 2026",
                14, False, SOFT,
            )
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            set_run(
                tf.paragraphs[0].add_run(),
                "Формулировка · способ проверки · метрики · решение по каждой гипотезе",
                13, False, SOFT,
            )
    add_icon(slide, "icon_69.png", emu(10.55), emu(0.5), emu(1.8), emu(1.8))
    return slide


def build_slide_agenda(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Содержание отчёта", 26)
    clear_body_placeholders(slide)

    items = [
        ("01", "Как мы читаем цифры", "Словарь воронки и что означает каждая метрика"),
        ("02", "Гипотеза 1", "Оффер бесплатного старта на холодной базе салонов красоты"),
        ("03", "Гипотеза 2", "Прозвон тёплых клиентов базы АБУ"),
        ("04", "Гипотеза 3", "Скрипт потребности для клиентов Юклаенс"),
        ("05", "Гипотезы 4 и 5", "Новые сегменты ЦА и канал франшиз — ещё не в проверке"),
        ("06", "Решения", "Что оставляем в работе и что меняем"),
    ]
    y = emu(1.45)
    for num, title, body in items:
        add_card(slide, emu(0.97), y, emu(11.5), emu(0.75), CARD, 0.06)
        add_textbox(slide, emu(1.15), y + emu(0.18), emu(0.7), emu(0.4), num, 16, True, BLUE)
        add_textbox(slide, emu(2.0), y + emu(0.1), emu(9.8), emu(0.32), title, 14, True, WHITE)
        add_textbox(slide, emu(2.0), y + emu(0.4), emu(9.8), emu(0.3), body, 12, False, SOFT)
        y += emu(0.82)
    return slide


def build_slide_metrics_intro(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Как мы измеряем результат", 26)
    clear_body_placeholders(slide)

    add_textbox(
        slide, emu(0.97), emu(1.4), emu(11.5), emu(0.7),
        "Целевое действие в проверках — демонстрация сервиса: клиент согласился "
        "на показ продукта и демо состоялось. Все проведённые демо в отчёте относятся "
        "к холодной базе салонов красоты.",
        13, False, WHITE,
    )

    blocks = [
        (
            "Обработано",
            "Сколько контактов из базы менеджер реально прошёл за период: дозвонился "
            "или зафиксировал итог обработки. Это объём выборки, а не «сырые» попытки набора.",
        ),
        (
            "Сделка",
            "Контакт, по которому появился интерес к продукту: квалификация пройдена, "
            "потребность зафиксирована, часто отправлено коммерческое предложение. "
            "Это ещё не продажа, а вход в коммерческую воронку.",
        ),
        (
            "Демонстрация",
            "Проведённый показ сервиса. Именно эту метрику считаем успехом проверки гипотезы "
            "на этапе привлечения: клиент готов смотреть продукт.",
        ),
        (
            "Подключение",
            "Клиент начал пользоваться сервисом (тестовый доступ, онбординг, настройка кабинета). "
            "Следующий этап после демо — уже про удержание и монетизацию.",
        ),
    ]
    y = emu(2.2)
    for title, body in blocks:
        add_card(slide, emu(0.97), y, emu(11.5), emu(1.0), CARD, 0.06)
        add_textbox(slide, emu(1.2), y + emu(0.12), emu(2.4), emu(0.35), title, 13, True, BLUE)
        add_textbox(slide, emu(3.7), y + emu(0.15), emu(8.5), emu(0.75), body, 12, False, WHITE)
        y += emu(1.08)
    return slide


def build_slide_metrics_conversions(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Две конверсии, которые смотрим в каждой проверке", 22)
    clear_body_placeholders(slide)

    add_card(slide, emu(0.97), emu(1.45), emu(5.55), emu(4.6), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.65), emu(5.1), emu(0.4), "Обзвон в сделку", 16, True, BLUE)
    add_multiline(
        slide, emu(1.2), emu(2.25), emu(5.1), emu(3.5),
        [
            ("Формула", BLUE, True, 12),
            ("Сделки ÷ Обработано × 100%", WHITE, True, 13),
            ("", WHITE, False, 8),
            ("Что показывает", BLUE, True, 12),
            ("Насколько оффер и скрипт цепляют "
             "интерес уже на первом контакте. "
             "Высокая цифра без демо ещё не значит успех: "
             "интерес может «зависать» без показа продукта.", WHITE, False, 12),
            ("", WHITE, False, 8),
            ("Пример", BLUE, True, 12),
            ("47 обработано, 10 сделок — это 21,3%. "
             "Каждый пятый контакт дошёл до коммерческого интереса.", WHITE, False, 12),
        ],
        line_spacing=1.12,
    )

    add_card(slide, emu(6.75), emu(1.45), emu(5.7), emu(4.6), CARD, 0.06)
    add_textbox(slide, emu(7.0), emu(1.65), emu(5.2), emu(0.4), "Сделка в демонстрацию", 16, True, BLUE)
    add_multiline(
        slide, emu(7.0), emu(2.25), emu(5.2), emu(3.5),
        [
            ("Формула", BLUE, True, 12),
            ("Демо ÷ Сделки × 100%", WHITE, True, 13),
            ("", WHITE, False, 8),
            ("Что показывает", BLUE, True, 12),
            ("Главный показатель для гипотез привлечения. "
             "Доля заинтересованных клиентов, которых удалось "
             "довести до показа продукта. "
             "Именно эту конверсию сравниваем «до» и «после» смены оффера.", WHITE, False, 12),
            ("", WHITE, False, 8),
            ("Пример", BLUE, True, 12),
            ("10 сделок, 2 демо — это 20%. "
             "Каждая пятая сделка заканчивается демонстрацией.", WHITE, False, 12),
        ],
        line_spacing=1.12,
    )
    return slide


def build_slide_h1_formulation(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 1. Формулировка", 26)
    clear_body_placeholders(slide)
    status_badge(slide, emu(10.0), emu(0.55), emu(2.5), emu(0.38), "Подтверждена", GREEN_OK)

    add_card(slide, emu(0.97), emu(1.4), emu(11.5), emu(2.5), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.55), emu(11.0), emu(0.3), "Продуктовая гипотеза", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(2.0), emu(11.0), emu(1.7),
        "Если при холодном прозвоне салонов красоты предлагать бесплатное сотрудничество "
        "с последующим переходом на платный пакет, то доля сделок, доведённых до демонстрации "
        "сервиса, вырастет относительно работы без такого оффера и закрепится на уровне "
        "не ниже 20%, потому что бесплатный старт снимает барьер входа у холодного контакта, "
        "который ещё не доверяет продукту и не готов сразу обсуждать оплату.",
        14, False, WHITE,
    )

    cols = [
        ("Сегмент", "Холодная база салонов красоты — контакты без истории работы с ГК Форус по этому продукту."),
        ("Изменение", "В скрипт продаж добавлен оффер: бесплатное сотрудничество, затем продление на платный пакет."),
        ("Критерий успеха", "Конверсия «сделка в демонстрацию» не ниже 20% на рабочем периоде после смены оффера."),
    ]
    x = emu(0.97)
    for title, body in cols:
        add_card(slide, x, emu(4.15), emu(3.7), emu(2.35), CARD, 0.06)
        add_textbox(slide, x + emu(0.2), emu(4.3), emu(3.3), emu(0.35), title, 13, True, BLUE)
        add_textbox(slide, x + emu(0.2), emu(4.8), emu(3.3), emu(1.5), body, 12, False, WHITE)
        x += emu(3.9)
    return slide


def build_slide_h1_method(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 1. Как проверяли", 26)
    clear_body_placeholders(slide)

    steps = [
        ("1", "Скрипт", "Собрали продажный скрипт с оффером бесплатного старта и логикой перехода на платный пакет."),
        ("2", "База", "Прозванивали холодную базу салонов красоты по неделям, фиксируя обработку в одной таблице показателей."),
        ("3", "Сравнение", "Сравнивали конверсию «сделка в демо» до появления оффера (сентябрь) и после (конец сентября — октябрь)."),
        ("4", "Решение", "Гипотезу считали подтверждённой, если на рабочем периоде конверсия держится около 20% и выше, а демо появляются стабильно."),
    ]
    y = emu(1.45)
    for num, title, body in steps:
        add_card(slide, emu(0.97), y, emu(11.5), emu(1.15), CARD, 0.06)
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(1.2), y + emu(0.35), emu(0.45), emu(0.45))
        circ.fill.solid()
        circ.fill.fore_color.rgb = BLUE
        circ.line.fill.background()
        add_textbox(
            slide, emu(1.2), y + emu(0.35), emu(0.45), emu(0.45),
            num, 14, True, WHITE, PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_textbox(slide, emu(1.95), y + emu(0.2), emu(10.0), emu(0.35), title, 14, True, WHITE)
        add_textbox(slide, emu(1.95), y + emu(0.55), emu(10.0), emu(0.5), body, 12, False, SOFT)
        y += emu(1.25)
    return slide


def build_slide_h1_table(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 1. Результаты по периодам", 24)
    clear_body_placeholders(slide)

    rows = [
        ["Период", "Обработано", "Сделок", "Демо", "Обзвон в сделку", "Сделка в демо"],
        ["14–19 сентября", "102", "13", "0", "12,8%", "0%"],
        ["21–28 сентября", "62", "5", "0", "8,1%", "0%"],
        ["28 сентября – 5 октября", "52", "12", "3", "23,1%", "25%"],
        ["5–12 октября", "47", "10", "2", "21,3%", "20%"],
        ["Всего за проверку", "263", "40", "5", "15,2%", "12,5%"],
    ]
    table_shape = slide.shapes.add_table(len(rows), 6, emu(0.7), emu(1.4), emu(12.1), emu(3.5))
    table = table_shape.table
    for i, w in enumerate([2.6, 1.6, 1.4, 1.2, 2.2, 2.1]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)

    add_card(slide, emu(0.97), emu(5.15), emu(11.5), emu(1.5), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(5.3), emu(11.0), emu(0.3), "Как читать таблицу", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(5.7), emu(11.0), emu(0.8),
        "До оффера (две первые недели) при 164 обработанных контактах не было ни одной демонстрации: "
        "конверсия «сделка в демо» равна 0%. После смены оффера появились стабильные демо: 3 и 2 за две недели, "
        "конверсия 25% и 20%. Строка «всего» смешивает оба режима работы, поэтому для решения по гипотезе "
        "ориентируемся на рабочие периоды после оффера — там целевой уровень 20% достигнут.",
        12, False, WHITE,
    )
    return slide


def build_slide_h1_reading(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 1. Разбор ключевых цифр", 24)
    clear_body_placeholders(slide)

    cards = [
        ("263", "обработано за проверку",
         "Сумма контактов холодной базы по четырём отчётным периодам менеджера. "
         "Это объём, на котором смотрели гипотезу."),
        ("40", "сделок",
         "Контакты с зафиксированным интересом и входом в коммерческую воронку. "
         "Без сделок демо не из чего считать."),
        ("5", "демонстраций",
         "Все демо за периоды менеджерских срезов — только холодная база. "
         "На базе АБУ за то же время демо не было."),
        ("20%", "рабочая конверсия",
         "На периоде 5–12 октября: 10 сделок и 2 демо. "
         "Это актуальный показатель после внедрения оффера."),
    ]
    positions = [
        (0.97, 1.45), (6.75, 1.45),
        (0.97, 4.0), (6.75, 4.0),
    ]
    for (val, label, body), (x, y) in zip(cards, positions):
        add_card(slide, emu(x), emu(y), emu(5.7), emu(2.35), CARD, 0.06)
        add_textbox(slide, emu(x + 0.25), emu(y + 0.2), emu(5.2), emu(0.5), val, 28, True, BLUE)
        add_textbox(slide, emu(x + 0.25), emu(y + 0.75), emu(5.2), emu(0.35), label, 13, True, WHITE)
        add_textbox(slide, emu(x + 0.25), emu(y + 1.15), emu(5.2), emu(1.0), body, 12, False, SOFT)
    return slide


def build_slide_h1_charts(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_BG])
    fill_title(slide, "Гипотеза 1. Динамика и сравнение с базой АБУ", 22)
    clear_body_placeholders(slide)

    add_chart(slide, "cold_conversion_dynamics.png", emu(0.55), emu(1.35), emu(7.1), emu(3.2))
    add_chart(slide, "cold_vs_abu.png", emu(7.5), emu(1.4), emu(5.3), emu(3.0))

    add_card(slide, emu(0.97), emu(4.75), emu(11.5), emu(1.85), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(4.9), emu(11.0), emu(0.3), "Вывод", 13, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(5.35), emu(11.0), emu(1.05),
        "Левый график: конверсия «сделка в демо» поднимается с нуля до 20–25% после смены оффера. "
        "Правый: за весь горизонт проверки холодная база дала 5 демонстраций при 263 обработанных; "
        "база АБУ при 48 обработанных — ноль демонстраций. Гипотеза 1 подтверждена на холодном сегменте. "
        "Оффер бесплатного старта оставляем в скрипте как рабочий стандарт для холодного прозвона.",
        13, False, WHITE,
    )
    return slide


def build_slide_h1_funnel(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 1. Воронка звонков за 1–15 октября", 22)
    clear_body_placeholders(slide)

    add_chart(slide, "oct_funnel.png", emu(0.55), emu(1.35), emu(7.3), emu(3.35))

    add_card(slide, emu(8.05), emu(1.4), emu(4.5), emu(4.6), CARD, 0.06)
    add_textbox(slide, emu(8.3), emu(1.55), emu(4.0), emu(0.35), "Что в столбцах", 13, True, BLUE)
    add_multiline(
        slide, emu(8.3), emu(2.05), emu(4.0), emu(3.7),
        [
            ("242 звонка", WHITE, True, 13),
            ("Ежедневные исходящие по холодной базе.", SOFT, False, 11),
            ("", WHITE, False, 6),
            ("123 контакта с ЛПР", WHITE, True, 13),
            ("Дозвонились до лица, принимающего решение, или получили его номер.", SOFT, False, 11),
            ("", WHITE, False, 6),
            ("22 интереса / КП", WHITE, True, 13),
            ("Квалификация и коммерческий интерес.", SOFT, False, 11),
            ("", WHITE, False, 6),
            ("6 презентаций", WHITE, True, 13),
            ("Проведённые демо за половину октября.", SOFT, False, 11),
            ("", WHITE, False, 6),
            ("2 тестовых доступа", WHITE, True, 13),
            ("Клиенты, которые пошли пробовать сервис после показа.", SOFT, False, 11),
        ],
        line_spacing=1.05,
    )

    add_textbox(
        slide, emu(0.97), emu(6.0), emu(11.5), emu(0.55),
        "Срез по сделкам на 12 октября: 5 демо проведено, 1 клиент подключён. "
        "Конверсия «демо в подключение» на этом срезе — 20%.",
        12, False, SOFT,
    )
    return slide


def build_slide_h2_formulation(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 2. Формулировка", 26)
    clear_body_placeholders(slide)
    status_badge(slide, emu(9.5), emu(0.55), emu(3.0), emu(0.38), "Не подтверждена", RED_MUTED)

    add_card(slide, emu(0.97), emu(1.4), emu(11.5), emu(2.6), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.55), emu(11.0), emu(0.3), "Продуктовая гипотеза", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(2.0), emu(11.0), emu(1.8),
        "Если прозванивать тёплую базу текущих клиентов ГК Форус (база АБУ), то конверсия "
        "«сделка в демонстрацию» будет минимум на 30% выше, чем на холодной базе салонов красоты "
        "при сопоставимом оффере, потому что эти клиенты уже знают компанию, им проще доверить "
        "время на демо и ниже стоимость первого контакта по доверию.",
        14, False, WHITE,
    )

    cols = [
        ("Сегмент", "База АБУ — действующие или знакомые с ГК Форус клиенты, не «холодный» рынок."),
        ("Ожидание", "Конверсия в демо выше холодной минимум на 30%. При холодных 20% цель для АБУ — от 26%."),
        ("Зачем проверяли", "Понять, куда выгоднее направлять ресурс менеджера: в холодный рынок или во внутреннюю базу."),
    ]
    x = emu(0.97)
    for title, body in cols:
        add_card(slide, x, emu(4.25), emu(3.7), emu(2.25), CARD, 0.06)
        add_textbox(slide, x + emu(0.2), emu(4.4), emu(3.3), emu(0.35), title, 13, True, BLUE)
        add_textbox(slide, x + emu(0.2), emu(4.9), emu(3.3), emu(1.4), body, 12, False, WHITE)
        x += emu(3.9)
    return slide


def build_slide_h2_results(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 2. Результаты и решение", 24)
    clear_body_placeholders(slide)

    rows = [
        ["Период", "Обработано", "Сделок", "Демо", "Обзвон в сделку", "Сделка в демо"],
        ["25–28 сентября", "3", "0", "0", "0%", "—"],
        ["28 сентября – 5 октября", "43", "6", "0", "14,0%", "0%"],
        ["5–12 октября", "2", "0", "0", "0%", "—"],
        ["Всего", "48", "6", "0", "12,5%", "0%"],
    ]
    table_shape = slide.shapes.add_table(len(rows), 6, emu(0.7), emu(1.35), emu(12.1), emu(2.55))
    table = table_shape.table
    for i, w in enumerate([2.6, 1.6, 1.4, 1.2, 2.2, 2.1]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)

    add_card(slide, emu(0.97), emu(4.15), emu(5.55), emu(2.4), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(4.3), emu(5.1), emu(0.3), "Что получилось", 13, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(4.75), emu(5.1), emu(1.6),
        "За проверку обработали 48 контактов АБУ, открыли 6 сделок и не провели ни одного демо. "
        "Конверсия «сделка в демо» — 0%. На фоне холодной базы с рабочими 20% гипотеза "
        "«тёплые конвертятся лучше» на этих данных не держится.",
        12, False, WHITE,
    )

    add_card(slide, emu(6.75), emu(4.15), emu(5.7), emu(2.4), CARD, 0.06)
    add_textbox(slide, emu(7.0), emu(4.3), emu(5.2), emu(0.3), "Решение", 13, True, BLUE)
    add_textbox(
        slide, emu(7.0), emu(4.75), emu(5.2), emu(1.6),
        "Гипотезу не подтверждаем. Выборки мало: основной объём — одна неделя на 43 контакта. "
        "Меняем заход в скрипте и повторяем проверку на большей базе, иначе нельзя честно "
        "сравнить тёплый и холодный канал.",
        12, False, WHITE,
    )
    return slide


def build_slide_h3_formulation(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 3. Формулировка", 26)
    clear_body_placeholders(slide)
    status_badge(slide, emu(10.0), emu(0.55), emu(2.5), emu(0.38), "Проверена", GREEN_OK)

    add_card(slide, emu(0.97), emu(1.4), emu(11.5), emu(2.7), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.55), emu(11.0), emu(0.3), "Продуктовая гипотеза", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(2.0), emu(11.0), emu(1.9),
        "Если на этапе сбора потребности у клиентов, которые работают на Юклаенс, "
        "целенаправленно уточнять наличие и качество линии технической поддержки "
        "и фактическое количество сотрудников, а в презентации Блиц-записи опираться "
        "именно на эти болевые точки, то конверсия в демонстрацию и дальнейший перевод "
        "в сделку превысит 40%, потому что в этом сегменте типовой разрыв — слабая поддержка "
        "и неудобство при росте штата, и наш продукт закрывает обе проблемы предметно.",
        14, False, WHITE,
    )

    cols = [
        ("Сегмент", "Клиенты и лиды, у которых текущий инструмент записи — Юклаенс (и смежные сервисы для сравнения)."),
        ("Изменение", "В скрипт discovery добавлены обязательные вопросы про техподдержку и численность сотрудников; презентация строится вокруг ответов."),
        ("Критерий успеха", "Конверсия в демонстрацию и в сделку выше 40% на проверенной выборке."),
    ]
    x = emu(0.97)
    for title, body in cols:
        add_card(slide, x, emu(4.35), emu(3.7), emu(2.15), CARD, 0.06)
        add_textbox(slide, x + emu(0.2), emu(4.5), emu(3.3), emu(0.35), title, 13, True, BLUE)
        add_textbox(slide, x + emu(0.2), emu(4.95), emu(3.3), emu(1.35), body, 12, False, WHITE)
        x += emu(3.9)
    return slide


def build_slide_h3_status(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 3. Статус проверки и таблица для цифр", 22)
    clear_body_placeholders(slide)

    add_card(slide, emu(0.97), emu(1.4), emu(11.5), emu(1.7), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.55), emu(11.0), emu(0.3), "Статус", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(2.0), emu(11.0), emu(0.9),
        "Гипотеза проверена в работе: заход через техподдержку и штат действительно усиливает интерес "
        "к демонстрации, конверсия в демо и сделку на практике выше 40%. Ниже — шаблон, чтобы "
        "зафиксировать объём прозвона по текущим сервисам записи. Цифры заполняются отдельно.",
        13, False, WHITE,
    )

    add_textbox(
        slide, emu(0.97), emu(3.35), emu(11.5), emu(0.35),
        "Распределение прозвоненных клиентов по текущему сервису записи",
        13, True, BLUE,
    )

    rows = [
        ["Текущий сервис", "Прозвонено", "Демо проведено", "Сделок", "Конверсия в демо"],
        ["Дикиди", "", "", "", ""],
        ["Юклаенс", "", "", "", ""],
        ["Мастерс", "", "", "", ""],
        ["Итого", "", "", "", ""],
    ]
    table_shape = slide.shapes.add_table(len(rows), 5, emu(0.97), emu(3.8), emu(11.5), emu(2.5))
    table = table_shape.table
    for i, w in enumerate([2.8, 2.2, 2.3, 2.0, 2.2]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)
    return slide


def build_slide_h4(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 4. Расширение целевой аудитории", 24)
    clear_body_placeholders(slide)
    status_badge(slide, emu(9.7), emu(0.55), emu(2.8), emu(0.38), "Не начата", ORANGE)

    add_card(slide, emu(0.97), emu(1.4), emu(11.5), emu(2.55), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.55), emu(11.0), emu(0.3), "Продуктовая гипотеза", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(2.0), emu(11.0), emu(1.75),
        "Если расширить прозвон и презентацию за пределы сегментов с постоянными и разовыми услугами "
        "(салоны красоты, автомойки) на организации, которые ведут запись на занятия и разовые сессии — "
        "спорт-клубы, образовательные центры, кабинеты психолога, нутрициолога и аналогичные форматы — "
        "то мы получим целевые демонстрации и сделки в этих сегментах, потому что у них та же операционная "
        "задача: управляемая онлайн-запись клиентов, а продукт Блиц-запись эту задачу закрывает.",
        13, False, WHITE,
    )

    cols = [
        ("Что проверяем", "Есть ли платёжеспособный спрос на Блиц-запись вне «классических» салонов и моек."),
        ("Как будем проверять", "С ноября: отдельные базы, адаптированный скрипт, замер конверсии в демо по каждому подсегменту."),
        ("Критерий успеха", "Стабильные демо и конверсия не хуже рабочего уровня холодных салонов (ориентир 20% сделка в демо), плюс качественный фидбек по сценариям записи."),
    ]
    x = emu(0.97)
    for title, body in cols:
        add_card(slide, x, emu(4.2), emu(3.7), emu(2.3), CARD, 0.06)
        add_textbox(slide, x + emu(0.2), emu(4.35), emu(3.3), emu(0.4), title, 13, True, BLUE)
        add_textbox(slide, x + emu(0.2), emu(4.85), emu(3.3), emu(1.45), body, 12, False, WHITE)
        x += emu(3.9)
    return slide


def build_slide_h5(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Гипотеза 5. Канал крупных франшиз", 24)
    clear_body_placeholders(slide)
    status_badge(slide, emu(9.3), emu(0.55), emu(3.2), emu(0.38), "Требует обсуждения", ORANGE)

    add_card(slide, emu(0.97), emu(1.4), emu(11.5), emu(2.55), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.55), emu(11.0), emu(0.3), "Продуктовая гипотеза", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(2.0), emu(11.0), emu(1.75),
        "Если выстраивать продажи Блиц-записи через крупные франшизы подходящего целевого сегмента "
        "(образовательные центры, фитнес-клубы и другие сети с единым стандартом записи), "
        "то стоимость привлечения точки и скорость масштабирования будут лучше, чем при точечном "
        "холодном обзвоне независимых точек, потому что решение на уровне управляющей компании "
        "открывает сразу несколько площадок с одним договором и одним онбордингом.",
        13, False, WHITE,
    )

    cols = [
        ("Что проверяем", "Даёт ли франшизный вход масштабируемый канал, а не разовые пилоты на 1–2 точки."),
        ("Почему ещё не стартовали", "Нужно согласовать приоритетные сети, лицо принятия решения внутри франшизы и критерии успеха пилота."),
        ("До старта", "Список целевых франшиз, оффер для управляющей компании, план пилота на N точек и метрики: демо, подключения, срок сделки."),
    ]
    x = emu(0.97)
    for title, body in cols:
        add_card(slide, x, emu(4.2), emu(3.7), emu(2.3), CARD, 0.06)
        add_textbox(slide, x + emu(0.2), emu(4.35), emu(3.3), emu(0.4), title, 13, True, BLUE)
        add_textbox(slide, x + emu(0.2), emu(4.85), emu(3.3), emu(1.45), body, 12, False, WHITE)
        x += emu(3.9)
    return slide


def build_slide_summary(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Сводка по пяти гипотезам", 26)
    clear_body_placeholders(slide)

    rows = [
        ["№", "Суть проверки", "Статус", "Главный результат"],
        ["1", "Бесплатный старт на холодной базе салонов", "Подтверждена", "Сделка в демо: с 0% до 20%"],
        ["2", "Тёплые клиенты АБУ конвертятся лучше холодных", "Не подтверждена", "48 обработано, 0 демо"],
        ["3", "Discovery Юклаенс через ТП и штат", "Проверена", "Конверсия >40%; цифры — в таблице"],
        ["4", "Новые сегменты записи (спорт, образование, кабинеты)", "Не начата", "Старт с ноября"],
        ["5", "Продажи через крупные франшизы", "Не начата", "Нужно обсуждение подхода"],
    ]
    table_shape = slide.shapes.add_table(len(rows), 4, emu(0.7), emu(1.45), emu(12.1), emu(4.4))
    table = table_shape.table
    for i, w in enumerate([0.7, 5.2, 2.3, 3.9]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table)

    for r in range(1, len(rows)):
        cell = table.cell(r, 2)
        for p in cell.text_frame.paragraphs:
            text = p.text
            for el in list(p._p):
                if el.tag.endswith("}r"):
                    p._p.remove(el)
            color = GREEN_OK if "Подтверж" in text or text == "Проверена" else (
                RED_MUTED if "Не подтверж" in text else ORANGE
            )
            set_run(p.add_run(), text, 10, True, color)
    return slide


def build_slide_next(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Что делаем дальше", 26)
    clear_body_placeholders(slide)

    items = [
        ("1", GREEN_OK, "Оставляем оффер холодной базы",
         "Бесплатное сотрудничество с переходом на платный пакет — стандарт скрипта. "
         "Держим конверсию «сделка в демо» на уровне 20% и выше."),
        ("2", ORANGE, "Перезапускаем АБУ с новым заходом",
         "Выборка была маленькой, демо не получили. Меняем скрипт, набираем объём и только потом "
         "снова сравниваем с холодной базой."),
        ("3", GREEN_OK, "Закрываем цифры по гипотезе 3",
         "Заполняем таблицу Дикиди / Юклаенс / Мастерс: прозвон, демо, сделки, конверсия."),
        ("4", BLUE, "Ноябрь — гипотеза 4",
         "Отдельные базы и скрипты для спорта, образования и кабинетов специалистов."),
        ("5", SOFT, "Согласовываем гипотезу 5",
         "До прозвона франшиз — приоритетные сети, ЛПР и критерии пилота."),
    ]
    y = emu(1.4)
    for num, accent, title, body in items:
        add_card(slide, emu(0.97), y, emu(11.5), emu(0.9), CARD, 0.06)
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(1.15), y + emu(0.24), emu(0.42), emu(0.42))
        circ.fill.solid()
        circ.fill.fore_color.rgb = accent
        circ.line.fill.background()
        add_textbox(
            slide, emu(1.15), y + emu(0.24), emu(0.42), emu(0.42),
            num, 12, True, WHITE if accent != SOFT else NEAR_BLACK,
            PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_textbox(slide, emu(1.8), y + emu(0.12), emu(10.3), emu(0.3), title, 13, True, WHITE)
        add_textbox(slide, emu(1.8), y + emu(0.45), emu(10.3), emu(0.4), body, 11, False, SOFT)
        y += emu(0.98)
    return slide


def main():
    if not TEMPLATE.exists():
        raise SystemExit(f"Template not found: {TEMPLATE}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(TEMPLATE, OUT)

    prs = Presentation(str(OUT))
    delete_all_slides(prs)

    builders = [
        build_slide_title,
        build_slide_agenda,
        build_slide_metrics_intro,
        build_slide_metrics_conversions,
        build_slide_h1_formulation,
        build_slide_h1_method,
        build_slide_h1_table,
        build_slide_h1_reading,
        build_slide_h1_charts,
        build_slide_h1_funnel,
        build_slide_h2_formulation,
        build_slide_h2_results,
        build_slide_h3_formulation,
        build_slide_h3_status,
        build_slide_h4,
        build_slide_h5,
        build_slide_summary,
        build_slide_next,
    ]
    for build in builders:
        build(prs)

    prs.save(str(OUT))
    print(f"Saved: {OUT}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
