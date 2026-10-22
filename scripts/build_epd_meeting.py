#!/usr/bin/env python3
"""Встреча по стратегии ЭПД в корпоративном шаблоне ГК Форус.

Титул и цель — в начале. Блок «Актуальный анализ целевой аудитории»
собран только из цифр первых рабочих слайдов файла «Стратегия ЭПД».
Остальные вопросы повестки — пустые слайды с темой, чтобы пройтись на встрече.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "Стратегия ЭПД.pptx"
OUT = ROOT / "presentation" / "Встреча_Стратегия_ЭПД.pptx"

# Схема «Форус»: шафран + тёмный сине-серый, шрифт Verdana.
INK = RGBColor(0x1A, 0x1A, 0x1A)
NAVY = RGBColor(0x39, 0x5F, 0x75)
MUTED = RGBColor(0x5C, 0x65, 0x6E)
GOLD = RGBColor(0xE6, 0xA8, 0x17)
GOLD_DEEP = RGBColor(0xC4, 0x86, 0x0A)
PALE = RGBColor(0xFB, 0xF6, 0xEA)
CARD = RGBColor(0xF4, 0xF6, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xE3, 0xE6, 0xEA)
TRACK = RGBColor(0xEE, 0xF0, 0xF2)

FONT = "Verdana"

# Макет «Заголовок и текст» — белое поле, знак Форус в правом нижнем углу.
L_CONTENT = 4
# Макет «2_Заголовок и текст» — чистое рабочее поле и нижний колонтитул.
L_CLEAN = 5
L_TITLE = 0
L_DIVIDER = 14  # шафранный разделитель
L_CLOSE = 27

# Города и объёмы — слайд «Города и объём рынка» того же файла.
CITIES = [
    ("Челябинск", 23286),
    ("Новосибирск", 11324),
    ("Самара", 11203),
    ("Красноярск", 9098),
    ("Екатеринбург", 2810),
]
SAM_TOTAL = sum(n for _, n in CITIES)


def emu(inches: float) -> Emu:
    return Inches(inches)


def delete_all_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        r_id = sld_id.get(qn("r:id"))
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def set_run(run, text, size_pt, bold=False, color=INK, font_name=FONT):
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = color
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = r_pr.find(qn(tag))
        if el is None:
            el = etree.SubElement(r_pr, qn(tag))
        el.set("typeface", font_name)


def set_anchor(text_frame, anchor):
    body_pr = text_frame._txBody.find(qn("a:bodyPr"))
    if body_pr is not None:
        body_pr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(
                anchor, "t"
            ),
        )


def tighten(text_frame, margin=0.06):
    text_frame.word_wrap = True
    text_frame.auto_size = None
    text_frame.margin_left = emu(margin)
    text_frame.margin_right = emu(margin)
    text_frame.margin_top = emu(0.04)
    text_frame.margin_bottom = emu(0.04)


def add_text(
    slide,
    left,
    top,
    width,
    height,
    lines,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    margin=0.02,
):
    """lines: sequence of (text, size, bold, color, space_before_pt)."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tighten(tf, margin)
    set_anchor(tf, anchor)
    for i, line in enumerate(lines):
        text, size, bold, color = line[:4]
        before = line[4] if len(line) > 4 else 0
        after = line[5] if len(line) > 5 else 0
        paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        paragraph.alignment = align
        paragraph.space_before = Pt(before)
        paragraph.space_after = Pt(after)
        paragraph.line_spacing = 1.0
        set_run(paragraph.add_run(), text, size, bold, color)
    return box


def add_shape(slide, left, top, width, height, fill, corner=None, line=None):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if corner is not None else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    if corner is not None:
        try:
            shape.adjustments[0] = corner
        except Exception:
            pass
    return shape


def drop_placeholders(slide, keep_idx=(4,)):
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx not in keep_idx:
            el = ph._element
            el.getparent().remove(el)


def fill_title_block(slide, lines):
    """Заполняет титульный плейсхолдер списком (text, size, bold, color)."""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.word_wrap = True
            tf.clear()
            for i, (text, size, bold, color) in enumerate(lines):
                paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                paragraph.alignment = PP_ALIGN.LEFT
                paragraph.space_before = Pt(0 if i == 0 else 6)
                paragraph.space_after = Pt(0)
                set_run(paragraph.add_run(), text, size, bold, color)
            return ph
    return None


def fmt_num(value: int) -> str:
    # Неразрывный пробел, чтобы «57 721» не рвалось на две строки.
    return f"{value:,}".replace(",", "\u00a0")


def slide_title(prs, title, kicker=None):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CLEAN])
    drop_placeholders(slide)
    if kicker:
        add_text(
            slide, emu(0.72), emu(0.32), emu(10.5), emu(0.28),
            [(kicker, 12, True, GOLD_DEEP)],
        )
        add_text(
            slide, emu(0.72), emu(0.56), emu(11.6), emu(0.62),
            [(title, 26, True, INK)],
        )
    else:
        add_text(
            slide, emu(0.72), emu(0.38), emu(11.6), emu(0.7),
            [(title, 28, True, INK)],
        )
    add_shape(slide, emu(0.74), emu(1.28), emu(1.15), emu(0.055), GOLD)
    return slide


def build_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_TITLE])
    fill_title_block(
        slide,
        [
            ("Стратегия ЭПД", 36, True, INK),
            ("Очная встреча", 22, False, NAVY),
            ("ЦСВ   ·   ЦП   ·   ЦКС   ·   СМАК", 16, True, INK),
        ],
    )
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.word_wrap = True
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Блохина Елизавета, Юнусова Юлиана", 14, True, INK)
            p2 = tf.add_paragraph()
            p2.space_before = Pt(2)
            set_run(p2.add_run(), "группа продуктового запуска", 13, False, MUTED)
    return slide


def build_goal(prs):
    slide = slide_title(prs, "Цель сегодняшней встречи")
    goals = [
        (
            "01",
            "Аудитория и конкуренты",
            "Синхронизировать понимание целевой аудитории и конкурентной среды.",
        ),
        (
            "02",
            "УТП двух направлений",
            "Сформировать уникальное предложение для маленьких перевозчиков и генераторов трафика.",
        ),
        (
            "03",
            "Критерии генераторов трафика",
            "Определить, кого считаем таким клиентом среди наших клиентов и на общем рынке.",
        ),
        (
            "04",
            "План и ответственность",
            "Зафиксировать план действий и зоны ответственности по проекту.",
        ),
    ]
    card_w, card_h = emu(5.85), emu(2.15)
    gap_x, gap_y = emu(0.2), emu(0.16)
    left0, top0 = emu(0.72), emu(1.62)
    for i, (num, title, body) in enumerate(goals):
        col, row = i % 2, i // 2
        left = left0 + col * (card_w + gap_x)
        top = top0 + row * (card_h + gap_y)
        add_shape(slide, left, top, card_w, card_h, WHITE, corner=0.08, line=LINE)
        add_shape(slide, left, top, emu(0.08), card_h, GOLD)
        add_text(
            slide, left + emu(0.28), top + emu(0.2), emu(1.2), emu(0.4),
            [(num, 18, True, GOLD_DEEP)],
        )
        add_text(
            slide, left + emu(0.28), top + emu(0.62), card_w - emu(0.5), emu(0.42),
            [(title, 16, True, INK)],
        )
        add_text(
            slide, left + emu(0.28), top + emu(1.12), card_w - emu(0.52), emu(0.85),
            [(body, 13, False, MUTED)],
        )
    return slide


def build_agenda(prs):
    slide = slide_title(prs, "Повестка", "Как идём по встрече")
    items = [
        ("01", "Анализ конкурентов", "Гурков · выжимка реестра", True),
        ("02", "Актуальный анализ целевой аудитории", "Блохина, Юнусова", True),
        ("03", "Генераторы трафика: определение и критерии", "весь состав", False),
        ("04", "Формирование УТП для обоих направлений", "весь состав", False),
        ("05", "Анализ отвалов и проблем сервиса", "весь состав", False),
        ("06", "Обновлённая стратегия рекламной кампании", "Лопатенко, Блохина, Юнусова, Савинская, Крутиков", True),
        ("07", "Дальнейшие шаги развития", "весь состав", False),
        ("08", "Зоны ответственности по ведению проекта", "весь состав", False),
    ]
    col_w = emu(5.9)
    row_h = emu(1.14)
    card_h = emu(1.04)
    left0, top0 = emu(0.72), emu(1.55)
    for i, (num, title, owner, ready) in enumerate(items):
        col, row = i // 4, i % 4
        left = left0 + col * (col_w + emu(0.18))
        top = top0 + row * row_h
        fill = PALE if ready else WHITE
        add_shape(slide, left, top, col_w, card_h, fill, corner=0.1, line=GOLD if ready else LINE)
        add_text(
            slide, left + emu(0.14), top, emu(0.62), card_h,
            [(num, 15, True, GOLD_DEEP)],
            anchor=MSO_ANCHOR.MIDDLE,
            margin=0.0,
        )
        add_text(
            slide, left + emu(0.78), top + emu(0.1), col_w - emu(0.96), card_h - emu(0.16),
            [
                (title, 13, True, INK, 0, 0),
                (owner, 11, False, NAVY if ready else MUTED, 2, 0),
            ],
            anchor=MSO_ANCHOR.MIDDLE,
            margin=0.02,
        )
    return slide


def build_discussion(prs, number, title, owner):
    """Пустой слайд: только тема вопроса, чтобы пройтись по нему на встрече."""
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    drop_placeholders(slide)
    add_text(
        slide, emu(0.85), emu(1.85), emu(8.5), emu(0.32),
        [(f"Вопрос повестки   ·   {number}", 13, True, GOLD_DEEP)],
    )
    add_shape(slide, emu(0.88), emu(2.26), emu(1.15), emu(0.055), GOLD)
    add_text(
        slide, emu(0.85), emu(2.48), emu(10.8), emu(2.9),
        [
            (title, 32, True, INK, 0, 0),
            (owner, 18, False, NAVY, 18, 0),
        ],
    )
    return slide


def build_two_directions(prs):
    slide = slide_title(
        prs,
        "Два направления целевой аудитории",
        "Актуальный анализ  ·  Блохина, Юнусова",
    )
    add_text(
        slide, emu(0.72), emu(1.48), emu(11.8), emu(0.55),
        [(
            "В рабочей постановке зафиксированы два контура. Цифры ниже есть только по маленьким перевозчикам.",
            14, False, MUTED,
        )],
    )

    cards = [
        (
            NAVY,
            WHITE,
            RGBColor(0xD6, 0xE4, 0xEC),
            "01",
            "Маленькие перевозчики",
            "TAM  ·  SAM  ·  SOM",
            "Рынок, пять городов с наибольшим индексом интереса и наша текущая база. Этот контур разбираем на следующих слайдах.",
        ),
        (
            WHITE,
            INK,
            MUTED,
            "02",
            "Генераторы трафика",
            "Анализ клиентов",
            "Направление названо. Определения, критериев и цифр в исходных слайдах нет — их фиксируем отдельным вопросом повестки.",
        ),
    ]
    left0, top = emu(0.72), emu(2.2)
    card_w, card_h, gap = emu(5.85), emu(3.75), emu(0.2)
    for i, (bg, title_c, body_c, num, title, label, body) in enumerate(cards):
        left = left0 + i * (card_w + gap)
        line = None if bg == NAVY else LINE
        add_shape(slide, left, top, card_w, card_h, bg, corner=0.06, line=line)
        if bg != NAVY:
            add_shape(slide, left, top, card_w, emu(0.08), GOLD)
        add_text(
            slide, left + emu(0.35), top + emu(0.35), emu(2), emu(0.4),
            [(num, 18, True, GOLD if bg == NAVY else GOLD_DEEP)],
        )
        add_text(
            slide, left + emu(0.35), top + emu(1.0), card_w - emu(0.7), emu(0.85),
            [(title, 22, True, title_c)],
        )
        add_text(
            slide, left + emu(0.35), top + emu(1.9), card_w - emu(0.7), emu(0.4),
            [(label, 14, True, GOLD if bg == NAVY else NAVY)],
        )
        add_text(
            slide, left + emu(0.35), top + emu(2.45), card_w - emu(0.7), emu(1.05),
            [(body, 14, False, body_c)],
        )
    return slide


def build_funnel(prs):
    slide = slide_title(
        prs,
        "TAM–SAM–SOM. Маленькие перевозчики",
        "Актуальный анализ  ·  Блохина, Юнусова",
    )
    cards = [
        (
            "TAM",
            "Весь рынок",
            "200 тыс.+",
            "перевозчиков",
            "Объём рынка на текущий момент.",
            "К ГИС ЭПД подключено 25 тыс.\nпо неподтверждённым источникам.",
        ),
        (
            "SAM",
            "Где спрос заметнее",
            "5",
            "городов",
            "Наибольший индекс интереса: Красноярск, Новосибирск, Екатеринбург, Самара, Челябинск.",
            f"В сумме — {fmt_num(SAM_TOTAL)}.",
        ),
        (
            "SOM",
            "Наша база",
            "245",
            "текущих клиентов",
            "Анализ действующей базы. Это контур, от которого считаем достижимое.",
            "Как эти 245 распределены по городам, в исходных слайдах не указано.",
        ),
    ]
    left0, top = emu(0.72), emu(1.58)
    card_w, gap = emu(3.86), emu(0.18)
    card_h = emu(4.35)
    for i, (code, label, number, unit, lead, note) in enumerate(cards):
        left = left0 + i * (card_w + gap)
        add_shape(slide, left, top, card_w, card_h, WHITE, corner=0.06, line=LINE)
        add_shape(slide, left, top, card_w, emu(0.1), GOLD if i == 1 else NAVY)
        add_text(
            slide, left + emu(0.24), top + emu(0.28), card_w - emu(0.48), emu(0.32),
            [(code, 14, True, GOLD_DEEP)],
        )
        add_text(
            slide, left + emu(0.24), top + emu(0.58), card_w - emu(0.48), emu(0.36),
            [(label, 13, False, MUTED)],
        )
        add_text(
            slide, left + emu(0.24), top + emu(1.1), card_w - emu(0.4), emu(0.7),
            [(number, 32, True, INK)],
        )
        add_text(
            slide, left + emu(0.24), top + emu(1.8), card_w - emu(0.48), emu(0.4),
            [(unit, 15, True, NAVY)],
        )
        add_shape(slide, left + emu(0.24), top + emu(2.32), emu(0.7), emu(0.035), GOLD)
        add_text(
            slide, left + emu(0.24), top + emu(2.5), card_w - emu(0.48), emu(0.85),
            [(lead, 12, False, INK)],
        )
        add_text(
            slide, left + emu(0.24), top + emu(3.4), card_w - emu(0.48), emu(0.8),
            [(note, 12, False, MUTED)],
        )
    add_text(
        slide, emu(0.74), emu(6.05), emu(11.6), emu(0.42),
        [(
            "В исходной формулировке: «подключено 25 тыс. рублей». Источник не подтверждён — на встрече уточняем единицу.",
            11, False, MUTED,
        )],
    )
    return slide


def build_cities(prs):
    slide = slide_title(
        prs,
        "Пять городов: объём рынка",
        "SAM  ·  города с наибольшим индексом интереса",
    )
    add_text(
        slide, emu(0.72), emu(1.46), emu(11.5), emu(0.4),
        [(
            "Объёмы взяты как есть. Доли — от суммы пяти городов, это не доля Форус.",
            13, False, MUTED,
        )],
    )
    max_n = CITIES[0][1]
    bar_max = emu(6.15)
    label_x = emu(0.72)
    bar_x = emu(3.05)
    row_h = emu(0.62)
    top0 = emu(2.05)
    for i, (city, count) in enumerate(CITIES):
        top = top0 + i * emu(0.7)
        share = count / SAM_TOTAL
        add_text(
            slide, label_x, top, emu(2.2), row_h,
            [(city, 15, True, INK)],
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_shape(slide, bar_x, top + emu(0.16), bar_max, emu(0.3), TRACK, corner=0.5)
        width = max(int(bar_max * (count / max_n)), int(emu(0.18)))
        add_shape(slide, bar_x, top + emu(0.16), width, emu(0.3), GOLD, corner=0.5)
        add_text(
            slide, emu(9.35), top, emu(1.7), row_h,
            [(fmt_num(count), 15, True, INK)],
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            slide, emu(11.0), top, emu(1.45), row_h,
            [(f"{share * 100:.1f}%".replace(".", ","), 14, False, NAVY)],
            anchor=MSO_ANCHOR.MIDDLE,
        )
    # итог
    band_top = emu(5.6)
    add_shape(slide, emu(0.72), band_top, emu(11.9), emu(0.72), PALE, corner=0.08)
    add_text(
        slide, emu(0.95), band_top + emu(0.08), emu(7.2), emu(0.56),
        [
            ("Итого по пяти городам", 12, False, MUTED, 0, 0),
            (f"{fmt_num(SAM_TOTAL)} перевозчиков", 16, True, INK, 0, 0),
        ],
    )
    add_text(
        slide, emu(8.3), band_top + emu(0.12), emu(4.0), emu(0.5),
        [("Челябинск — около 40% объёма", 13, True, NAVY)],
        anchor=MSO_ANCHOR.MIDDLE,
        align=PP_ALIGN.RIGHT,
    )
    return slide


def build_reading(prs):
    slide = slide_title(
        prs,
        "Как читаем эти цифры",
        "Актуальный анализ  ·  что уже видно и что ещё открыто",
    )
    points = [
        (
            "01",
            "Рынок шире, чем подключение к ГИС ЭПД",
            "200 тыс.+ перевозчиков на рынке и 25 тыс. подключённых к ГИС ЭПД. Вторую цифру берём осторожно: источник не подтверждён, в исходнике единица записана как рубли.",
        ),
        (
            "02",
            "Рабочий фокус SAM — пять городов",
            f"Красноярск, Новосибирск, Екатеринбург, Самара и Челябинск. Суммарный объём: {fmt_num(SAM_TOTAL)} перевозчиков. Больше всего в Челябинске, меньше всего в Екатеринбурге.",
        ),
        (
            "03",
            "Достижимая база сегодня — 245 клиентов",
            "Это SOM из исходных слайдов. Долю в пяти городах считать рано: география этих клиентов не разложена.",
        ),
    ]
    left, top0 = emu(0.72), emu(1.58)
    card_w, card_h, gap = emu(7.55), emu(1.38), emu(0.1)
    for i, (num, title, body) in enumerate(points):
        top = top0 + i * (card_h + gap)
        add_shape(slide, left, top, card_w, card_h, WHITE, corner=0.08, line=LINE)
        add_text(
            slide, left + emu(0.2), top + emu(0.16), emu(0.7), emu(0.36),
            [(num, 14, True, GOLD_DEEP)],
        )
        add_text(
            slide, left + emu(0.95), top + emu(0.14), card_w - emu(1.2), emu(0.36),
            [(title, 14, True, INK)],
        )
        add_text(
            slide, left + emu(0.95), top + emu(0.52), card_w - emu(1.2), emu(0.74),
            [(body, 12, False, MUTED)],
        )

    panel_left = emu(8.5)
    add_shape(slide, panel_left, top0, emu(4.12), emu(4.34), NAVY, corner=0.06)
    add_text(
        slide, panel_left + emu(0.28), top0 + emu(0.22), emu(3.55), emu(0.7),
        [
            ("Открыто на обсуждение", 14, True, WHITE),
            ("ответов в цифрах нет", 12, False, RGBColor(0xD6, 0xE4, 0xEC), 4),
        ],
    )
    questions = [
        "Подтвердить 25 тыс. и единицу измерения по ГИС ЭПД.",
        "Разложить 245 клиентов по пяти городам.",
        "Как заходим в эти города.",
        "Кто для нас генератор трафика и по каким критериям — следующий вопрос повестки.",
    ]
    q_top = top0 + emu(1.15)
    for i, question in enumerate(questions):
        y = q_top + i * emu(0.75)
        add_text(
            slide, panel_left + emu(0.28), y, emu(0.4), emu(0.32),
            [(f"{i + 1}", 13, True, GOLD)],
        )
        add_text(
            slide, panel_left + emu(0.65), y, emu(3.2), emu(0.7),
            [(question, 12, False, WHITE)],
        )
    return slide


def build_competitors(prs):
    slide = slide_title(prs, "Анализ конкурентов", "Гурков  ·  реестр операторов ИС ЭПД")
    cards = [
        (
            "01",
            "Цена входа",
            "Бесплатно только ЭПД Лайт: нет модуля 1С и мобильного приложения. Платный минимум — DropCat, 4\u00a0500\u00a0₽ за\u00a01\u00a0000 титулов. Контур — 6\u00a0500, Сбер — 7\u00a0000.",
        ),
        (
            "02",
            "1С-ЭПД",
            "5\u00a0000\u00a0₽ за\u00a01\u00a0000 титулов и 250\u00a0000\u00a0₽ за\u00a0100\u00a0000. На большом пакете дешевле Контура и Доки, на одном уровне с Такскомом. Веб-версии нет, только 1С:Фреш. API есть.",
        ),
        (
            "03",
            "Доки",
            "Тот же оператор, Калуга Астрал. Веб, 1С и мобильное есть, API нет. Пакета на\u00a01\u00a0000 титулов нет. С\u00a05\u00a0000 титулов дороже 1С-ЭПД: 24\u00a0500 против 22\u00a0500.",
        ),
        (
            "04",
            "Чем берут другие",
            "Контур — Диадок и TMS, партнёры OZON и DPD. DropCat — Госключ и тест на 100 дней. Saby — УКЭП в телефоне и Деловые линии. Сбер — ГигаЧат.",
        ),
    ]
    card_w, card_h = emu(5.85), emu(2.15)
    gap_x, gap_y = emu(0.2), emu(0.16)
    left0, top0 = emu(0.72), emu(1.55)
    for i, (num, title, body) in enumerate(cards):
        col, row = i % 2, i // 2
        left = left0 + col * (card_w + gap_x)
        top = top0 + row * (card_h + gap_y)
        add_shape(slide, left, top, card_w, card_h, WHITE, corner=0.08, line=LINE)
        add_shape(slide, left, top, emu(0.08), card_h, GOLD if i == 1 else NAVY)
        add_text(
            slide, left + emu(0.28), top + emu(0.16), emu(1.1), emu(0.32),
            [(num, 14, True, GOLD_DEEP)],
        )
        add_text(
            slide, left + emu(1.15), top + emu(0.14), card_w - emu(1.4), emu(0.36),
            [(title, 16, True, INK)],
        )
        add_text(
            slide, left + emu(0.28), top + emu(0.62), card_w - emu(0.5), emu(1.4),
            [(body, 13, False, MUTED)],
        )
    return slide


def build_competitor_prices(prs):
    slide = slide_title(
        prs,
        "Цены пакетов, ₽ в год",
        "Гурков  ·  1С-ЭПД и ближайшие операторы",
    )
    headers = ("Сервис", "1 000 титулов", "100 000 титулов", "Чем отличается")
    rows = [
        ("1С-ЭПД", "5 000", "250 000", "Модуль 1С и API. Веб только во Фреше", True),
        ("Доки", "нет в реестре", "320 000", "Веб есть, API нет. Тот же оператор", True),
        ("Контур.Логистика", "6 500", "400 000", "Диадок, TMS, тест 30 дней", False),
        ("DropCat", "4 500", "200 000", "Госключ, тест 100 дней и 100 документов", False),
        ("Такском", "5 000", "250 000", "Тот же ценовой коридор, что у 1С-ЭПД", False),
        ("Saby", "5 000", "индивидуально", "УКЭП в телефоне, Деловые линии", False),
        ("Сбер2B", "7 000", "индивидуально", "ГигаЧат, экспедиторские документы", False),
        ("ЭПД Лайт", "бесплатно", "бесплатно", "Нет модуля 1С и мобильного приложения", False),
    ]
    widths = (emu(2.35), emu(1.85), emu(2.15), emu(5.45))
    left0, top = emu(0.72), emu(1.52)
    row_h = emu(0.46)
    x = left0
    add_shape(slide, left0, top, sum(widths, emu(0)), row_h, NAVY)
    for header, width in zip(headers, widths):
        add_text(
            slide, x + emu(0.08), top, width - emu(0.1), row_h,
            [(header, 11, True, WHITE)],
            anchor=MSO_ANCHOR.MIDDLE, margin=0.02,
        )
        x += width
    for i, (name, p1, p2, diff, ours) in enumerate(rows):
        y = top + row_h + i * row_h
        fill = PALE if ours else (WHITE if i % 2 == 0 else CARD)
        add_shape(slide, left0, y, sum(widths, emu(0)), row_h, fill)
        x = left0
        for text, width, bold in (
            (name, widths[0], True),
            (p1, widths[1], ours),
            (p2, widths[2], ours),
            (diff, widths[3], False),
        ):
            add_text(
                slide, x + emu(0.08), y, width - emu(0.12), row_h,
                [(text, 12, bold, INK if bold else MUTED)],
                anchor=MSO_ANCHOR.MIDDLE, margin=0.0,
            )
            x += width
    add_text(
        slide, emu(0.74), emu(6.05), emu(11.7), emu(0.36),
        [(
            "Источник — реестр операторов ИС ЭПД. LeraData только с индивидуальными ценами. АТИ-Доки и НТЦ СТЭК в файле без характеристик.",
            11, False, MUTED,
        )],
    )
    return slide


def build_campaign_frame(prs):
    slide = slide_title(
        prs,
        "Рекламная кампания на перевозчиков",
        "Гипотезы  ·  Лопатенко, Блохина, Юнусова, Савинская, Крутиков",
    )
    steps = [
        ("1", "Сегмент", "Уже в законе, верит в отсрочку до марта или ещё не работает."),
        ("2", "Где показываем", "Реклама по запросам, визитам сайтов и софта."),
        ("3", "Что объясняем", "Почему, зачем и на каких условиях переходить."),
        ("4", "Целевое действие", "В схеме не заполнено: его фиксируем на встрече."),
    ]
    left, top0 = emu(0.72), emu(1.52)
    card_w, card_h, gap = emu(7.35), emu(0.9), emu(0.1)
    for i, (num, title, body) in enumerate(steps):
        top = top0 + i * (card_h + gap)
        add_shape(slide, left, top, card_w, card_h, WHITE, corner=0.1, line=LINE)
        add_text(
            slide, left + emu(0.16), top, emu(0.5), card_h,
            [(num, 18, True, GOLD_DEEP)],
            anchor=MSO_ANCHOR.MIDDLE, margin=0.0,
        )
        add_text(
            slide, left + emu(0.7), top + emu(0.1), card_w - emu(0.9), emu(0.32),
            [(title, 14, True, INK)],
        )
        add_text(
            slide, left + emu(0.7), top + emu(0.42), card_w - emu(0.9), emu(0.4),
            [(body, 12, False, MUTED)],
        )

    panel = emu(8.25)
    add_shape(slide, panel, top0, emu(4.35), emu(3.9), NAVY, corner=0.06)
    add_text(
        slide, panel + emu(0.24), top0 + emu(0.16), emu(3.9), emu(0.55),
        [
            ("Кастдев · Витя", 14, True, WHITE),
            ("что выясняем до креатива", 11, False, RGBColor(0xD6, 0xE4, 0xEC), 2),
        ],
    )
    questions = [
        "Как пришёл к ЭТрН и на что опирался, выбирая оператора.",
        "Крупный заказчик сказал или выбирал сам.",
        "Если сказал — какая выгода выбрать именно этого оператора.",
        "Почему перешёл на ЭПД.",
        "Знает ли, что закон уже в силе. Постановление — не отсрочка.",
    ]
    y = top0 + emu(0.9)
    for i, question in enumerate(questions):
        add_text(
            slide, panel + emu(0.24), y, emu(0.35), emu(0.5),
            [(str(i + 1), 12, True, GOLD)],
        )
        add_text(
            slide, panel + emu(0.55), y, emu(3.55), emu(0.55),
            [(question, 11, False, WHITE)],
        )
        y += emu(0.56)
    add_shape(slide, emu(0.72), emu(5.55), emu(11.9), emu(0.7), PALE, corner=0.08)
    add_text(
        slide, emu(0.95), emu(5.62), emu(11.4), emu(0.55),
        [(
            "Ограничение схемы: страница сейчас настроена на отправителя, не на перевозчика.",
            14, True, NAVY,
        )],
        anchor=MSO_ANCHOR.MIDDLE,
    )
    return slide


def build_campaign_segments(prs):
    slide = slide_title(
        prs,
        "Три сегмента перевозчиков",
        "Гипотезы кампании",
    )
    add_text(
        slide, emu(0.72), emu(1.42), emu(11.8), emu(0.36),
        [(
            "На схеме сегмент один: перевозчик между DNS и покупателем-юриком.",
            13, False, MUTED,
        )],
    )
    cards = [
        (
            "Гипотеза 1а",
            "Уже работает и знает, что закон вступил в силу",
            [
                ("Почему", "Выгодная цена. На УСН можно без 1С."),
                ("Зачем", "Крупный заказчик может не дать заказы."),
                ("Важно", "На дороге перевозчика не штрафуют. Это не значит, что закон перенесли."),
            ],
        ),
        (
            "Гипотеза 1б",
            "Знает закон, но считает, что его отсрочили до марта",
            [
                ("Почему", "Закон уже действует. На схеме: штрафы более 3–5 раз без ЭТрН."),
                ("Зачем", "Тот же риск потерять заказы крупного заказчика."),
                ("Важно", "Отсутствие штрафа на дороге закон не переносит."),
            ],
        ),
        (
            "Гипотеза 2",
            "Ещё не работает. Хочет заказы крупного заказчика",
            [
                ("Оффер", "Как пройти требования, почему сейчас, а не в марте, кейс, цена, как подключить."),
                ("Почему", "Выгодная цена. На УСН можно без 1С."),
                ("Зачем", "Не попасть под требования по ЭТрН и ЭДО и остаться без заказов."),
            ],
        ),
    ]
    left0, top = emu(0.72), emu(1.86)
    card_w, card_h, gap = emu(3.86), emu(4.2), emu(0.18)
    for i, (label, title, blocks) in enumerate(cards):
        left = left0 + i * (card_w + gap)
        add_shape(slide, left, top, card_w, card_h, WHITE, corner=0.06, line=LINE)
        add_shape(slide, left, top, card_w, emu(0.08), GOLD if i == 2 else NAVY)
        add_text(
            slide, left + emu(0.2), top + emu(0.2), card_w - emu(0.4), emu(0.28),
            [(label, 12, True, GOLD_DEEP)],
        )
        add_text(
            slide, left + emu(0.2), top + emu(0.5), card_w - emu(0.4), emu(0.95),
            [(title, 14, True, INK)],
        )
        y = top + emu(1.55)
        for head, body in blocks:
            add_text(
                slide, left + emu(0.2), y, card_w - emu(0.4), emu(0.24),
                [(head, 12, True, NAVY)],
            )
            add_text(
                slide, left + emu(0.2), y + emu(0.24), card_w - emu(0.4), emu(0.7),
                [(body, 12, False, MUTED)],
            )
            y += emu(0.85)
    return slide


def build_close(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CLOSE])
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        tf = ph.text_frame
        tf.word_wrap = True
        tf.clear()
        if idx == 0:
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Фиксируем", 36, True, INK)
            p2 = tf.add_paragraph()
            p2.space_before = Pt(2)
            set_run(p2.add_run(), "решения встречи", 32, True, NAVY)
        elif idx == 1:
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Блохина Елизавета", 16, True, INK)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "Юнусова Юлиана", 16, True, INK)
            p3 = tf.add_paragraph()
            p3.space_before = Pt(4)
            set_run(p3.add_run(), "группа продуктового запуска", 13, False, MUTED)
        elif idx == 10:
            p = tf.paragraphs[0]
            set_run(p.add_run(), "info@forus.ru", 14, False, INK)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "www.forus.ru", 14, False, NAVY)
    return slide


def build():
    if not TEMPLATE.exists():
        raise SystemExit(f"Не найден шаблон: {TEMPLATE}")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(TEMPLATE, OUT)
    prs = Presentation(str(OUT))
    delete_all_slides(prs)

    build_title(prs)
    build_goal(prs)
    build_agenda(prs)

    build_competitors(prs)
    build_competitor_prices(prs)

    build_discussion(
        prs, "02", "Актуальный анализ целевой аудитории", "Блохина, Юнусова"
    )
    build_two_directions(prs)
    build_funnel(prs)
    build_cities(prs)
    build_reading(prs)

    build_discussion(
        prs,
        "03",
        "Генераторы трафика: кто это для нас и по каким критериям мы их определяем",
        "весь состав",
    )
    build_discussion(
        prs,
        "04",
        "Формирование УТП для обоих направлений",
        "весь состав",
    )
    build_discussion(
        prs,
        "05",
        "Анализ отвалов и проблем сервиса",
        "весь состав",
    )
    build_campaign_frame(prs)
    build_campaign_segments(prs)
    build_discussion(prs, "07", "Дальнейшие шаги развития", "весь состав")
    build_discussion(
        prs, "08", "Зоны ответственности по ведению проекта", "весь состав"
    )
    build_close(prs)

    prs.core_properties.title = "Стратегия ЭПД. Очная встреча"
    prs.core_properties.author = "Блохина Елизавета, Юнусова Юлиана"
    prs.core_properties.subject = "Цель встречи, анализ целевой аудитории, повестка"
    prs.save(str(OUT))
    print(f"saved {OUT} slides={len(prs.slides)}")


if __name__ == "__main__":
    build()
