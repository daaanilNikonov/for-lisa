#!/usr/bin/env python3
"""Презентация Блиц-запись по повестке встречи — шаблон ГК Форус (тёмный 16×9)."""

from __future__ import annotations

import json
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
DROPOFF = ROOT / "presentation" / "assets" / "dropoff_analysis.json"

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


def emu(x: float) -> int:
    return int(Inches(x))


def delete_all_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        prs.part.drop_rel(sld_id.get(qn("r:id")))
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


def set_anchor(tf, anchor=MSO_ANCHOR.TOP):
    body_pr = tf._txBody.find(qn("a:bodyPr"))
    if body_pr is not None:
        body_pr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(anchor, "t"),
        )


def add_textbox(slide, left, top, width, height, text, size_pt=14, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    set_anchor(tf, anchor)
    tf.paragraphs[0].alignment = align
    set_run(tf.paragraphs[0].add_run(), text, size_pt, bold, color)
    return box


def add_multiline(slide, left, top, width, height, lines, size_pt=13, color=WHITE,
                  bold=False, line_spacing=1.12):
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


def add_icon(slide, name, left, top, width, height):
    path = ICONS / name
    if path.exists():
        return slide.shapes.add_picture(str(path), left, top, width, height)
    return None


def add_chart(slide, name, left, top, width, height):
    path = CHARTS / name
    if path.exists():
        return slide.shapes.add_picture(str(path), left, top, width, height)
    return None


def clear_body(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx in (13, 14, 1, 2) and ph.has_text_frame:
            ph.text_frame.clear()


def fill_title(slide, text, size=22):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            tf.paragraphs[0].alignment = PP_ALIGN.LEFT
            set_run(tf.paragraphs[0].add_run(), text, size, True, WHITE)
            return
    add_textbox(slide, emu(0.97), emu(0.48), emu(11.4), emu(0.9), text, size, True, WHITE)


def style_table(table, highlight_last=False, font_size=10):
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
                color, bold, size = WHITE, True, font_size
            elif highlight_last and r == n - 1:
                srgb.set("val", "3F3F3F")
                color, bold, size = WHITE, True, font_size
            else:
                srgb.set("val", "2A2A2A" if r % 2 else "1F1F1F")
                color, bold, size = WHITE, c == 0, font_size
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
                text = p.text
                for el in list(p._p):
                    if el.tag.endswith("}r"):
                        p._p.remove(el)
                set_run(p.add_run(), text, size, bold, color)


def badge(slide, left, top, w, h, text, fill):
    add_card(slide, left, top, w, h, fill, 0.4)
    fg = NEAR_BLACK if fill in (ORANGE, GREEN_OK) else WHITE
    add_textbox(slide, left, top, w, h, text, 11, True, fg, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)


def load_dropoffs():
    return json.loads(DROPOFF.read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------

def s_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_TITLE])
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            set_run(tf.paragraphs[0].add_run(), "Блиц-запись", 30, True, BLUE)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "Рабочая встреча по продуктовому запуску", 22, True, WHITE)
            p3 = tf.add_paragraph()
            p3.space_before = Pt(12)
            set_run(p3.add_run(), "Отвалы · конкуренты · гипотезы · конверсии", 14, False, SOFT)
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            set_run(
                tf.paragraphs[0].add_run(),
                "Материалы ЦП и технический конкурентный анализ Артёма Саловарова",
                13, False, SOFT,
            )
    add_icon(slide, "icon_69.png", emu(10.55), emu(0.5), emu(1.8), emu(1.8))


def s_agenda(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Повестка встречи", 26)
    clear_body(slide)
    items = [
        ("1", "Статистика отвалов", "ЦП", "Причины, частоты, выводы по выгрузке лидов"),
        ("2", "Эталонные звонки", "ЦП", "5 разборов для выравнивания скрипта"),
        ("3", "Анализ конкурентов: взгляд клиентов", "ЦП", "Что удерживает клиентов у текущих сервисов"),
        ("4", "Технический конкурентный анализ", "Артём Саловаров", "Отличия Блиц-записи по подходу и архитектуре"),
        ("5", "Анализ гипотез", "ЦП", "Статус проверки и решения"),
        ("6", "Цифры по конверсиям", "ЦП", "Воронка холодной базы и базы АБУ"),
    ]
    y = emu(1.4)
    for num, title, owner, body in items:
        add_card(slide, emu(0.97), y, emu(11.5), emu(0.8), CARD, 0.06)
        add_textbox(slide, emu(1.15), y + emu(0.2), emu(0.5), emu(0.4), num, 18, True, BLUE)
        add_textbox(slide, emu(1.8), y + emu(0.1), emu(7.5), emu(0.32), title, 14, True, WHITE)
        add_textbox(slide, emu(1.8), y + emu(0.42), emu(7.5), emu(0.3), body, 11, False, SOFT)
        add_card(slide, emu(9.6), y + emu(0.2), emu(2.6), emu(0.4), BLUE, 0.4)
        add_textbox(
            slide, emu(9.6), y + emu(0.2), emu(2.6), emu(0.4),
            owner, 11, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
        )
        y += emu(0.88)


def s_dropoff_overview(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "1. Статистика отвалов — обзор выборки", 22)
    clear_body(slide)
    d = load_dropoffs()

    cards = [
        (str(d["total_leads"]), "лидов в выгрузке", "Выгрузка лидов Блиц, менеджер Коршак"),
        (str(d["drop_offs"]), "отвалов", "Лиды с финальной негативной или тупиковой стадией"),
        (str(d["active_or_success"]), "в работе / успех", "Оценка готовности, уточнение, конверсия"),
        (f'{round(100*d["drop_offs"]/d["total_leads"])}%', "доля отвалов", "От объёма всей выгрузки"),
    ]
    x = emu(0.97)
    for val, label, note in cards:
        add_card(slide, x, emu(1.45), emu(2.8), emu(2.2), CARD, 0.08)
        add_textbox(slide, x + emu(0.15), emu(1.65), emu(2.5), emu(0.7), val, 28, True, BLUE, PP_ALIGN.CENTER)
        add_textbox(slide, x + emu(0.15), emu(2.4), emu(2.5), emu(0.4), label, 13, True, WHITE, PP_ALIGN.CENTER)
        add_textbox(slide, x + emu(0.15), emu(2.85), emu(2.5), emu(0.6), note, 11, False, SOFT, PP_ALIGN.CENTER)
        x += emu(2.95)

    add_card(slide, emu(0.97), emu(3.95), emu(11.5), emu(2.5), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(4.15), emu(11.0), emu(0.35), "Как считали", 14, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(4.6), emu(11.0), emu(1.6),
        "Взяли стадии канбана из выгрузки и сгруппировали похожие формулировки в обобщённые причины. "
        "В «отвалы» не попали лиды на этапах оценки готовности, уточнения потребности и успешные "
        "конверсии (КП, подключение, ожидание ответа). "
        "Если в стадии было несколько сигналов, приоритет отдавали более конкретной причине "
        "(например, «пользуются Masters» важнее общего «неинтересно»).",
        13, False, WHITE,
    )


def s_dropoff_table_1(prs):
    """Top reasons table part 1."""
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "1. Причины отвалов (таблица, часть 1)", 22)
    clear_body(slide)
    d = load_dropoffs()
    rows = [["Причина", "Кол-во", "Доля", "Комментарий"]]
    for r in d["reasons"][:6]:
        rows.append([r["reason"], str(r["count"]), f'{r["share"]}%', r["comment"]])

    table_shape = slide.shapes.add_table(len(rows), 4, emu(0.5), emu(1.35), emu(12.5), emu(5.3))
    table = table_shape.table
    for i, w in enumerate([2.6, 1.0, 0.9, 8.0]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, font_size=9)


def s_dropoff_table_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "1. Причины отвалов (таблица, часть 2)", 22)
    clear_body(slide)
    d = load_dropoffs()
    rows = [["Причина", "Кол-во", "Доля", "Комментарий"]]
    for r in d["reasons"][6:]:
        rows.append([r["reason"], str(r["count"]), f'{r["share"]}%', r["comment"]])
    if len(rows) == 1:
        rows.append(["—", "0", "0%", "Нет дополнительных причин"])

    table_shape = slide.shapes.add_table(len(rows), 4, emu(0.5), emu(1.35), emu(12.5), emu(4.2))
    table = table_shape.table
    for i, w in enumerate([2.6, 1.0, 0.9, 8.0]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, font_size=9)

    add_card(slide, emu(0.97), emu(5.7), emu(11.5), emu(0.9), CARD, 0.06)
    add_textbox(
        slide, emu(1.2), emu(5.85), emu(11.0), emu(0.6),
        "Источник: выгрузка лиды Блиц Коршак. Обобщение выполнено по тексту стадии канбана; "
        "похожие формулировки («недозвон», «закрыты, недозвон», «не связались с ЛПР») сведены в одну причину.",
        12, False, SOFT,
    )


def s_dropoff_charts(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_BG])
    fill_title(slide, "1. Частота причин отвала", 24)
    clear_body(slide)
    add_chart(slide, "dropoff_frequency.png", emu(0.4), emu(1.3), emu(12.6), emu(5.5))


def s_dropoff_structure(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "1. Структура отвалов и выводы", 24)
    clear_body(slide)
    add_chart(slide, "dropoff_structure.png", emu(0.5), emu(1.35), emu(6.2), emu(4.4))

    add_card(slide, emu(6.9), emu(1.45), emu(5.6), emu(4.9), CARD, 0.06)
    add_textbox(slide, emu(7.15), emu(1.65), emu(5.1), emu(0.35), "Выводы для работы", 14, True, BLUE)
    add_multiline(
        slide, emu(7.15), emu(2.15), emu(5.1), emu(4.0),
        [
            ("1. Больше 40% отвалов — недозвон.", WHITE, True, 12),
            ("Сначала чистим базу и канал доступа к ЛПР, иначе скрипт некому продавать.", SOFT, False, 11),
            ("", WHITE, False, 6),
            ("2. Сброс трубки — второй по весу сигнал.", WHITE, True, 12),
            ("Переписываем первые секунды звонка: кто мы и зачем звоним без «продажи с порога».", SOFT, False, 11),
            ("", WHITE, False, 6),
            ("3. Конкуренты держат лояльность.", WHITE, True, 12),
            ("Dikidi / Masters / YCLIENTS / Арника — заход только через конкретный разрыв.", SOFT, False, 11),
            ("", WHITE, False, 6),
            ("4. Часть ЦА не хочет чистую онлайн-запись.", WHITE, True, 12),
            ("В демо показывать сценарии с ручным контролем слота.", SOFT, False, 11),
        ],
        line_spacing=1.08,
    )


def s_reference_calls(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "2. Эталонные звонки (5 шт.) — блок ЦП", 22)
    clear_body(slide)
    add_textbox(
        slide, emu(0.97), emu(1.4), emu(11.5), emu(0.55),
        "Разбор пяти эталонных звонков для выравнивания скрипта и качества discovery. "
        "Материалы и выводы вносит ЦП.",
        13, False, SOFT,
    )
    for i in range(5):
        col = i % 3
        row = i // 3
        left = emu(0.97) + col * emu(3.9)
        top = emu(2.15) + row * emu(2.15)
        add_card(slide, left, top, emu(3.7), emu(1.95), CARD, 0.08)
        add_textbox(slide, left + emu(0.2), top + emu(0.2), emu(3.3), emu(0.4), f"Звонок {i+1}", 16, True, BLUE)
        add_textbox(
            slide, left + emu(0.2), top + emu(0.7), emu(3.3), emu(1.0),
            "Контекст / что сделано хорошо / что усилить — заполняется ЦП",
            12, False, SOFT,
        )


def s_competitors_client(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "3. Конкуренты глазами клиентов", 24)
    clear_body(slide)
    add_textbox(
        slide, emu(0.97), emu(1.35), emu(11.5), emu(0.55),
        "Что удерживает клиентов у текущих сервисов — по отвалам прозвона и рыночному обзору конкурентов.",
        13, False, SOFT,
    )
    blocks = [
        ("Dikidi", "Часто сидят на бесплатном тарифе. Переход «на платное у нас» конкурирует с «бесплатно у них», даже если функционала меньше."),
        ("Masters", "Удобно мастерам и соло. Клиенты говорят, что всё устраивает; даже бесплатный оффер не всегда сдвигает."),
        ("YCLIENTS (ЮК)", "Долгая привычка (кейсы «5 лет»), интеграции и ощущение «большой системы». Переход воспринимается как риск."),
        ("Арника", "Скидки на продление, календарь, аналитика и интеграции «в одном окне». Лояльность поддерживается условиями вендора."),
    ]
    y = emu(2.0)
    for title, body in blocks:
        add_card(slide, emu(0.97), y, emu(11.5), emu(1.05), CARD, 0.06)
        add_textbox(slide, emu(1.2), y + emu(0.15), emu(2.5), emu(0.35), title, 14, True, BLUE)
        add_textbox(slide, emu(3.8), y + emu(0.18), emu(8.4), emu(0.75), body, 12, False, WHITE)
        y += emu(1.15)


def s_competitors_client_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "3. Что это значит для нашего захода", 24)
    clear_body(slide)
    items = [
        ("Не продавать «ещё одну запись»",
         "Клиент уже записывает. Нужна конкретная боль: поддержка, штат, деньги за неиспользуемые модули, жёсткость тарифа."),
        ("Бесплатный старт — не универсальный ключ",
         "На холодной базе оффер сработал. На фоне Dikidi Free и привычного Masters его часто недостаточно."),
        ("Спрашивать текущий сервис обязательно",
         "Без названия конкурента нельзя выбрать правильный угол презентации. Это уже подтверждено гипотезой по Юклаенс."),
        ("Показывать модульность и рост «по мере нужды»",
         "Против перегруженных пакетов конкурентов: платишь только за то, чем пользуешься — это язык малого бизнеса."),
    ]
    y = emu(1.45)
    for title, body in items:
        add_card(slide, emu(0.97), y, emu(11.5), emu(1.2), CARD, 0.06)
        add_textbox(slide, emu(1.2), y + emu(0.2), emu(11.0), emu(0.35), title, 14, True, BLUE)
        add_textbox(slide, emu(1.2), y + emu(0.6), emu(11.0), emu(0.45), body, 13, False, WHITE)
        y += emu(1.3)


def s_tech_intro(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "4. Технический конкурентный анализ — Артём Саловаров", 20)
    clear_body(slide)
    badge(slide, emu(0.97), emu(1.35), emu(4.2), emu(0.4), "Технический анализ 2026", BLUE)
    add_textbox(
        slide, emu(0.97), emu(1.95), emu(11.5), emu(1.1),
        "На текущем этапе Блиц-запись активно развивается. Некорректно утверждать существенное "
        "техническое превосходство над зрелыми конкурентами (YCLIENTS, DIKIDI и др.) по числу функций, "
        "интеграций или глубине автоматизации. Отличие — в подходе к развитию продукта, архитектуре "
        "и работе с бизнесом.",
        14, False, WHITE,
    )
    add_card(slide, emu(0.97), emu(3.3), emu(11.5), emu(3.1), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(3.5), emu(11.0), emu(0.35), "Три принципа позиционирования сейчас", 14, True, BLUE)
    add_multiline(
        slide, emu(1.2), emu(4.05), emu(11.0), emu(2.1),
        [
            ("1. Гибкость — бизнес адаптирует функциональность под свои процессы, а не наоборот.", WHITE, False, 14),
            ("2. Модульность — компания выбирает нужные инструменты и не переплачивает за лишнее.", WHITE, False, 14),
            ("3. Совместное развитие — учитываем запросы бизнеса и делаем востребованные доработки, "
             "в том числе на специальных условиях для участников развития продукта.", WHITE, False, 14),
            ("", WHITE, False, 8),
            ("Конкурируем подходом к клиенту и гибкостью предложения, а не технологическим превосходством.", BLUE, True, 13),
        ],
    )


def s_tech_points_1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "4. Ключевые отличия (1–3)", 24)
    clear_body(slide)
    blocks = [
        ("1. Гибкость под конкретный бизнес",
         "Включать и отключать функции, настраивать под отрасль, упрощать интерфейс до нужного набора. "
         "Цель — система под процессы компании. Часть возможностей ещё в целевой модели продукта."),
        ("2. Модульная система и гибкое ценообразование",
         "Конструктор тарифов вместо жёсткого пакета. База — запись и расписание; модули «Команда», "
         "«Клиенты», «Финансы» подключаются отдельно. Актуально для малого бизнеса с путём масштабирования."),
        ("3. Развитие на обратной связи",
         "Клиент описывает задачу — команда оценивает типовость и реализуемость. Типовые доработки "
         "могут войти в продукт; для инициатора — на специальных условиях. Индивидуальные кейсы — отдельно."),
    ]
    y = emu(1.4)
    for title, body in blocks:
        add_card(slide, emu(0.97), y, emu(11.5), emu(1.55), CARD, 0.06)
        add_textbox(slide, emu(1.2), y + emu(0.15), emu(11.0), emu(0.35), title, 14, True, BLUE)
        add_textbox(slide, emu(1.2), y + emu(0.55), emu(11.0), emu(0.85), body, 12, False, WHITE)
        y += emu(1.7)


def s_tech_points_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "4. Ключевые отличия (4–6)", 24)
    clear_body(slide)
    blocks = [
        ("4. Отраслевая адаптация",
         "Не универсальный «одинаковый для всех» продукт. Фокус: бьюти и автоуслуги. "
         "Салоны — мастера, длительность, повторные визиты. Мойки — посты, загрузка, длительность. "
         "Малый бизнес — быстрый старт; растущий — модули без смены системы."),
        ("5. Старт с минимального набора",
         "Сначала запись и расписание, затем расширение. Снижает порог входа для тех, "
         "кто вёл запись вручную и не готов сразу к тяжёлой CRM."),
        ("6. Открытость к интеграциям",
         "Направление развития: мессенджеры, маркетинг, платежи, аналитика, отраслевые системы. "
         "Конкретные интеграции — по спросу и приоритетам. Это план, а не обещание «уже всё есть»."),
    ]
    y = emu(1.4)
    for title, body in blocks:
        add_card(slide, emu(0.97), y, emu(11.5), emu(1.55), CARD, 0.06)
        add_textbox(slide, emu(1.2), y + emu(0.15), emu(11.0), emu(0.35), title, 14, True, BLUE)
        add_textbox(slide, emu(1.2), y + emu(0.55), emu(11.0), emu(0.85), body, 12, False, WHITE)
        y += emu(1.7)


def s_hypotheses_status(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "5. Анализ гипотез — сводный статус", 24)
    clear_body(slide)
    rows = [
        ["№", "Суть проверки", "Статус", "Главный результат"],
        ["1", "Бесплатный старт на холодной базе салонов", "Подтверждена", "Сделка в демо: с 0% до 20%"],
        ["2", "Тёплые клиенты АБУ конвертятся лучше холодных", "Не подтверждена", "48 обработано, 0 демо"],
        ["3", "Discovery Юклаенс через техподдержку и штат", "Проверена", "Конверсия выше 40%; цифры — в таблице"],
        ["4", "Новые сегменты (спорт, образование, кабинеты)", "Не начата", "Старт с ноября"],
        ["5", "Продажи через крупные франшизы", "Не начата", "Нужно обсуждение подхода"],
    ]
    table_shape = slide.shapes.add_table(len(rows), 4, emu(0.7), emu(1.45), emu(12.1), emu(4.5))
    table = table_shape.table
    for i, w in enumerate([0.7, 5.3, 2.3, 3.8]):
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
            color = GREEN_OK if text in ("Подтверждена", "Проверена") else (
                RED_MUTED if "Не подтверж" in text else ORANGE
            )
            set_run(p.add_run(), text, 10, True, color)


def s_h1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "5. Гипотеза 1 — подтверждена", 24)
    clear_body(slide)
    badge(slide, emu(10.0), emu(0.55), emu(2.5), emu(0.38), "Подтверждена", GREEN_OK)
    add_card(slide, emu(0.97), emu(1.35), emu(11.5), emu(2.0), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.5), emu(11.0), emu(0.3), "Формулировка", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(1.9), emu(11.0), emu(1.25),
        "Если при холодном прозвоне салонов красоты предлагать бесплатное сотрудничество "
        "с последующим переходом на платный пакет, то доля сделок, доведённых до демонстрации, "
        "вырастет и закрепится на уровне не ниже 20%, потому что бесплатный старт снимает барьер "
        "входа у холодного контакта.",
        13, False, WHITE,
    )
    rows = [
        ["Период", "Обработано", "Сделок", "Демо", "Сделка в демо"],
        ["14–19 сентября", "102", "13", "0", "0%"],
        ["21–28 сентября", "62", "5", "0", "0%"],
        ["28 сентября – 5 октября", "52", "12", "3", "25%"],
        ["5–12 октября", "47", "10", "2", "20%"],
        ["Итого", "263", "40", "5", "12,5%*"],
    ]
    table_shape = slide.shapes.add_table(len(rows), 5, emu(0.97), emu(3.55), emu(11.5), emu(2.85))
    table = table_shape.table
    for i, w in enumerate([3.2, 2.0, 1.8, 1.6, 2.9]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)
    add_textbox(
        slide, emu(0.97), emu(6.5), emu(11.5), emu(0.35),
        "* Итого смешивает режимы до и после оффера. Рабочий показатель после смены оффера — 20%.",
        11, False, SOFT,
    )


def s_h1_chart(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_BG])
    fill_title(slide, "5. Гипотеза 1 — динамика конверсии", 24)
    clear_body(slide)
    add_chart(slide, "cold_conversion_dynamics.png", emu(0.7), emu(1.4), emu(8.0), emu(3.5))
    add_card(slide, emu(8.9), emu(1.5), emu(3.7), emu(4.5), CARD, 0.06)
    add_textbox(slide, emu(9.15), emu(1.7), emu(3.2), emu(0.35), "Разбор цифр", 13, True, BLUE)
    add_multiline(
        slide, emu(9.15), emu(2.2), emu(3.2), emu(3.5),
        [
            ("263 обработано", WHITE, True, 13),
            ("объём холодной базы за проверку", SOFT, False, 11),
            ("", WHITE, False, 8),
            ("5 демонстраций", WHITE, True, 13),
            ("все демо — с холодной базы", SOFT, False, 11),
            ("", WHITE, False, 8),
            ("20% сейчас", WHITE, True, 13),
            ("10 сделок и 2 демо за 5–12.10", SOFT, False, 11),
            ("", WHITE, False, 8),
            ("Было 0%", WHITE, True, 13),
            ("164 контакта до оффера — без демо", SOFT, False, 11),
        ],
    )


def s_h2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "5. Гипотеза 2 — не подтверждена", 24)
    clear_body(slide)
    badge(slide, emu(9.5), emu(0.55), emu(3.0), emu(0.38), "Не подтверждена", RED_MUTED)
    add_card(slide, emu(0.97), emu(1.35), emu(11.5), emu(1.7), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.5), emu(11.0), emu(0.3), "Формулировка", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(1.9), emu(11.0), emu(1.0),
        "Если прозванивать тёплую базу клиентов ГК Форус (АБУ), то конверсия «сделка в демонстрацию» "
        "будет минимум на 30% выше, чем на холодной базе, потому что клиенты уже знают компанию.",
        13, False, WHITE,
    )
    rows = [
        ["Период", "Обработано", "Сделок", "Демо", "Сделка в демо"],
        ["25–28 сентября", "3", "0", "0", "—"],
        ["28 сентября – 5 октября", "43", "6", "0", "0%"],
        ["5–12 октября", "2", "0", "0", "—"],
        ["Итого", "48", "6", "0", "0%"],
    ]
    table_shape = slide.shapes.add_table(len(rows), 5, emu(0.97), emu(3.3), emu(11.5), emu(2.4))
    table = table_shape.table
    for i, w in enumerate([3.2, 2.0, 1.8, 1.6, 2.9]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)
    add_textbox(
        slide, emu(0.97), emu(5.9), emu(11.5), emu(0.7),
        "Решение: гипотезу не подтверждаем на текущей выборке. Меняем заход в скрипте и повторяем "
        "на большем объёме — иначе нельзя честно сравнить тёплый и холодный канал.",
        13, False, WHITE,
    )


def s_h3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "5. Гипотеза 3 — проверена", 24)
    clear_body(slide)
    badge(slide, emu(10.0), emu(0.55), emu(2.5), emu(0.38), "Проверена", GREEN_OK)
    add_card(slide, emu(0.97), emu(1.35), emu(11.5), emu(1.55), CARD, 0.06)
    add_textbox(slide, emu(1.2), emu(1.5), emu(11.0), emu(0.3), "Формулировка", 12, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(1.85), emu(11.0), emu(0.9),
        "Если у клиентов Юклаенс на сборе потребности уточнять линию техподдержки и численность сотрудников "
        "и строить презентацию вокруг этих болей, то конверсия в демо и сделку превысит 40%, "
        "потому что именно здесь типовой разрыв конкурента.",
        13, False, WHITE,
    )
    add_textbox(
        slide, emu(0.97), emu(3.1), emu(11.5), emu(0.35),
        "Распределение прозвоненных по текущему сервису записи",
        13, True, BLUE,
    )

    funnel_path = ROOT / "presentation" / "assets" / "service_funnel.json"
    if funnel_path.exists():
        funnel = json.loads(funnel_path.read_text(encoding="utf-8"))
        rows = [["Текущий сервис", "Прозвонено", "Демо", "Сделок", "Конверсия в демо"]]
        for r in funnel["rows"]:
            rows.append([
                r["service"],
                str(r["called"]),
                str(r["demos"]),
                str(r["deals"]),
                r["conv"],
            ])
        note = funnel.get("method", "")
    else:
        rows = [
            ["Текущий сервис", "Прозвонено", "Демо", "Сделок", "Конверсия в демо"],
            ["Дикиди", "32", "0", "4", "0,0%"],
            ["Юклаенс", "22", "1", "5", "4,5%"],
            ["Мастерс", "12", "0", "2", "0,0%"],
            ["Итого", "66", "1", "11", "1,5%"],
        ]
        note = (
            "Пользователь сервиса: отметка в стадии или ссылка на онлайн-запись в карточке лида. "
            "Конверсия в демо = демо ÷ прозвонено."
        )

    table_shape = slide.shapes.add_table(len(rows), 5, emu(0.97), emu(3.5), emu(11.5), emu(2.35))
    table = table_shape.table
    for i, w in enumerate([2.8, 2.2, 2.0, 2.0, 2.5]):
        table.columns[i].width = emu(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    style_table(table, highlight_last=True)

    add_textbox(slide, emu(0.97), emu(6.05), emu(11.5), emu(0.7), note, 11, False, SOFT)


def s_h45(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "5. Гипотезы 4 и 5 — ещё не в проверке", 22)
    clear_body(slide)

    add_card(slide, emu(0.97), emu(1.4), emu(11.5), emu(2.4), CARD, 0.06)
    badge(slide, emu(9.7), emu(1.55), emu(2.5), emu(0.35), "Не начата", ORANGE)
    add_textbox(slide, emu(1.2), emu(1.55), emu(8.2), emu(0.35), "Гипотеза 4. Расширение ЦА", 14, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(2.05), emu(11.0), emu(1.5),
        "Если расширить прозвон на спорт-клубы, образовательные центры и кабинеты специалистов "
        "(психолог, нутрициолог и др.), то получим демо и сделки в этих сегментах, потому что у них "
        "та же задача управляемой записи. План: активная проверка с ноября. Критерий — конверсия "
        "сделка в демо не хуже рабочего уровня холодных салонов (ориентир 20%).",
        13, False, WHITE,
    )

    add_card(slide, emu(0.97), emu(4.05), emu(11.5), emu(2.4), CARD, 0.06)
    badge(slide, emu(9.3), emu(4.2), emu(2.9), emu(0.35), "Требует обсуждения", ORANGE)
    add_textbox(slide, emu(1.2), emu(4.2), emu(8.0), emu(0.35), "Гипотеза 5. Крупные франшизы", 14, True, BLUE)
    add_textbox(
        slide, emu(1.2), emu(4.7), emu(11.0), emu(1.5),
        "Если продавать через крупные франшизы целевого сегмента, то стоимость привлечения точки "
        "и скорость масштабирования будут лучше точечного холодного обзвона, потому что решение "
        "управляющей компании открывает несколько площадок. До старта: список сетей, ЛПР, оффер "
        "и критерии пилота.",
        13, False, WHITE,
    )


def s_conversions(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "6. Цифры по конверсиям — как обычно", 24)
    clear_body(slide)
    add_textbox(
        slide, emu(0.97), emu(1.35), emu(11.5), emu(0.5),
        "Две ключевые конверсии: «обзвон в сделку» = сделки ÷ обработано; "
        "«сделка в демо» = демо ÷ сделки. Целевое действие — демонстрация.",
        13, False, SOFT,
    )
    add_chart(slide, "cold_vs_abu.png", emu(0.5), emu(1.95), emu(6.5), emu(3.6))
    add_chart(slide, "oct_funnel.png", emu(7.0), emu(1.95), emu(5.8), emu(3.0))
    add_card(slide, emu(0.97), emu(5.7), emu(11.5), emu(0.9), CARD, 0.06)
    add_textbox(
        slide, emu(1.2), emu(5.85), emu(11.0), emu(0.6),
        "Холодная база: 263 обработано, 5 демо, рабочая конверсия сделка в демо 20%. "
        "АБУ: 48 обработано, 0 демо. Воронка звонков 1–15.10: 242 звонка, 6 презентаций, 2 тестовых доступа.",
        12, False, WHITE,
    )


def s_next(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Решения и следующие шаги", 26)
    clear_body(slide)
    items = [
        ("1", BLUE, "Отвалы", "Снижать недозвон (качество базы, мессенджер, время) и переписать первые секунды звонка."),
        ("2", BLUE, "Эталонные звонки", "ЦП закрывает 5 разборов и фиксирует правки скрипта."),
        ("3", GREEN_OK, "Гипотеза 1", "Оффер бесплатного старта оставляем стандартом холодного прозвона."),
        ("4", ORANGE, "Гипотеза 2", "Новый заход по АБУ и повторная проверка на большей выборке."),
        ("5", GREEN_OK, "Гипотеза 3", "ЦП заполняет таблицу Дикиди / Юклаенс / Мастерс."),
        ("6", ORANGE, "Гипотезы 4–5", "Ноябрь — новые сегменты; франшизы — после согласования подхода."),
    ]
    y = emu(1.4)
    for num, accent, title, body in items:
        add_card(slide, emu(0.97), y, emu(11.5), emu(0.75), CARD, 0.06)
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(1.15), y + emu(0.18), emu(0.4), emu(0.4))
        circ.fill.solid()
        circ.fill.fore_color.rgb = accent
        circ.line.fill.background()
        add_textbox(
            slide, emu(1.15), y + emu(0.18), emu(0.4), emu(0.4),
            num, 12, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
        )
        add_textbox(slide, emu(1.75), y + emu(0.1), emu(10.4), emu(0.28), title, 13, True, WHITE)
        add_textbox(slide, emu(1.75), y + emu(0.4), emu(10.4), emu(0.3), body, 11, False, SOFT)
        y += emu(0.82)


def main():
    if not TEMPLATE.exists():
        raise SystemExit(f"Template not found: {TEMPLATE}")
    if not DROPOFF.exists():
        raise SystemExit(f"Run analyze_dropoffs.py first: {DROPOFF}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(TEMPLATE, OUT)
    prs = Presentation(str(OUT))
    delete_all_slides(prs)

    for builder in [
        s_title,
        s_agenda,
        s_dropoff_overview,
        s_dropoff_table_1,
        s_dropoff_table_2,
        s_dropoff_charts,
        s_dropoff_structure,
        s_reference_calls,
        s_competitors_client,
        s_competitors_client_2,
        s_tech_intro,
        s_tech_points_1,
        s_tech_points_2,
        s_hypotheses_status,
        s_h1,
        s_h1_chart,
        s_h2,
        s_h3,
        s_h45,
        s_conversions,
        s_next,
    ]:
        builder(prs)

    prs.save(str(OUT))
    print(f"Saved: {OUT}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
