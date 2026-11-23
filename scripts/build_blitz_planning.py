#!/usr/bin/env python3
"""Планерка «Блиц-запись» на 2 часа.

Факты взяты из нового коммерческого предложения и таблицы звонков за ноябрь.
Формулировки для клиента на слайдах не заполнены: их пишут на встрече.
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
TEMPLATE = ROOT / "Презентация ГК Форус темный шаблон 16х9 (1).pptx"
OUT = ROOT / "presentation" / "Планерка_Блиц-запись_повестка.pptx"

BLUE = RGBColor(0x26, 0xA6, 0xE0)
CARD = RGBColor(0x3F, 0x3F, 0x3F)
CARD_LIGHT = RGBColor(0xF4, 0xF4, 0xF4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xC8, 0xC8, 0xC8)
NEAR_BLACK = RGBColor(0x1A, 0x1A, 0x1A)
DARK = RGBColor(0x2A, 0x2A, 0x2A)
MUTED = RGBColor(0xD2, 0xD2, 0xD2)
INK = RGBColor(0x2C, 0x2C, 0x2C)
RULE = RGBColor(0xC5, 0xC5, 0xC5)
TEAL_DEEP = RGBColor(0x14, 0x4A, 0x63)

FONT = "Verdana"
L_TITLE = 0
L_CONTENT = 3
L_LAST = 23


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
        body_pr.set("lIns", str(Emu(0)))
        body_pr.set("rIns", str(Emu(0)))
        body_pr.set("tIns", str(Emu(0)))
        body_pr.set("bIns", str(Emu(0)))


def add_lines(slide, left, top, width, height, lines, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    set_anchor(tf, anchor)
    for i, (text, size, bold, color, before) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(before)
        p.space_after = Pt(0)
        set_run(p.add_run(), text, size, bold, color)
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


def rule(slide, left, top, width):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, emu(0.015))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RULE
    bar.line.fill.background()
    return bar


def clear_body(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx in (13, 14, 1, 2, 10):
            if ph.has_text_frame:
                ph.text_frame.clear()


def fill_title(slide, text, size=24):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            set_run(p.add_run(), text, size, True, WHITE)
            return ph
    return add_lines(slide, emu(0.97), emu(0.42), emu(11.4), emu(0.7), [(text, size, True, WHITE, 0)])


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def content_slide(prs, title, size=24):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, title, size)
    clear_body(slide)
    return slide


def kicker(slide, text):
    add_lines(slide, emu(0.7), emu(1.38), emu(11.8), emu(0.28), [(text, 12, False, BLUE, 0)])


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------

def slide_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_TITLE])
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Планерка по сервису", 26, True, WHITE)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "«Блиц-запись»", 32, True, BLUE)
            p3 = tf.add_paragraph()
            p3.space_before = Pt(12)
            set_run(p3.add_run(), "Два часа. Повестка и места для записи.", 16, False, WHITE)
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(
                p.add_run(),
                "Центр продаж, производство и лаборатория.\n"
                "Формулировки для клиента пишем на встрече.",
                14, False, SOFT,
            )
    notes(slide, "Сразу сказать: готовых фраз для клиента на слайдах нет. Их записываем вместе.")
    return slide


def slide_questions(prs):
    slide = content_slide(prs, "Шесть вопросов встречи", 26)
    kicker(slide, "В конце по каждому вопросу должна появиться запись: что решили, кто делает, к какой дате")

    items = [
        ("1", "Что говорим клиенту", "Одна фраза в начале разговора и одно предложение, которое отправляют все менеджеры."),
        ("2", "Бесплатное подключение", "Кому его предлагаем, что входит и что клиент всё равно оплачивает."),
        ("3", "Сколько подключаем", "Сколько салонов в месяц успевают довести продажи и сколько успевает подключить производство."),
        ("4", "Что доделываем в сервисе", "Не больше шести задач на ближайший год. У каждой месяц и ответственный."),
        ("5", "Как работаем друг с другом", "Что команды просят друг у друга и за сколько отвечают."),
        ("6", "Точка и UDS", "Отдельный текст для каждого партнёра. В предложение для салона его не вставляем."),
    ]
    top = 1.74
    for num, title, body in items:
        y = emu(top)
        add_card(slide, emu(0.7), y, emu(11.9), emu(0.62), DARK, 0.1)
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(0.86), y + emu(0.11), emu(0.40), emu(0.40))
        badge.fill.solid()
        badge.fill.fore_color.rgb = BLUE
        badge.line.fill.background()
        add_lines(slide, emu(0.86), y + emu(0.16), emu(0.40), emu(0.30), [(num, 13, True, WHITE, 0)], align=PP_ALIGN.CENTER)
        add_lines(slide, emu(1.42), y + emu(0.06), emu(3.5), emu(0.50), [(title, 14, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(5.0), y + emu(0.06), emu(7.3), emu(0.50), [(body, 13, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE)
        top += 0.68
    notes(
        slide,
        "Если названия команд в зале другие, поправить в первые минуты. "
        "Лаборатория делает сервис. Производство подключает клиента. Центр продаж говорит с клиентом.",
    )
    return slide


def slide_agenda(prs):
    slide = content_slide(prs, "Как идём эти два часа", 26)
    kicker(slide, "Если вопрос не закрыли, называем человека и дату. Встречу не растягиваем")

    rows = [
        ("0:00", "5", "Как работаем", "Кто за что отвечает"),
        ("0:05", "8", "Что уже продаём и ноябрь", "Общие цифры, без спора о формулировках"),
        ("0:13", "18", "Фраза клиенту и бесплатное подключение", "Текст записываем на слайдах"),
        ("0:31", "16", "Одно предложение для менеджеров", "Заполняем пустые строки"),
        ("0:47", "16", "Сколько успеваем подключить", "Число на месяц по каждому типу салона"),
        ("1:03", "20", "Что доделываем в сервисе", "До шести задач и месяц каждой"),
        ("1:23", "16", "Как работаем друг с другом", "Просьбы команд и сроки ответа"),
        ("1:39", "12", "С кем нас сравнивают", "Три карточки, текст пишем сами"),
        ("1:51", "9", "Точка, UDS и фиксация", "Две фразы и список решений"),
    ]
    top = 1.70
    add_lines(slide, emu(0.86), emu(top), emu(0.9), emu(0.22), [("Старт", 11, True, BLUE, 0)])
    add_lines(slide, emu(1.8), emu(top), emu(0.6), emu(0.22), [("Мин", 11, True, BLUE, 0)])
    add_lines(slide, emu(2.5), emu(top), emu(5.2), emu(0.22), [("О чём говорим", 11, True, BLUE, 0)])
    add_lines(slide, emu(7.8), emu(top), emu(4.5), emu(0.22), [("Что должно остаться", 11, True, BLUE, 0)])
    top = 1.96
    for i, (clock, mins, title, out) in enumerate(rows):
        y = emu(top)
        add_card(slide, emu(0.7), y, emu(11.9), emu(0.38), CARD if i % 2 == 0 else DARK, 0.12)
        add_lines(slide, emu(0.86), y, emu(0.9), emu(0.38), [(clock, 12, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(1.8), y, emu(0.6), emu(0.38), [(mins, 12, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(2.5), y, emu(5.1), emu(0.38), [(title, 13, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(7.8), y, emu(4.6), emu(0.38), [(out, 12, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE)
        top += 0.41
    notes(slide, "Факты на следующих двух слайдах не обсуждаем долго. Они нужны, чтобы писать формулировки от одного и того же.")
    return slide


def slide_materials(prs):
    slide = content_slide(prs, "Что уже стоит в новых материалах", 24)
    kicker(slide, "Это не текст встречи. Это то, что уже можно показать клиенту")

    add_card(slide, emu(0.7), emu(1.76), emu(7.35), emu(2.35), CARD, 0.08)
    add_lines(
        slide, emu(0.9), emu(1.90), emu(6.95), emu(2.05),
        [
            ("Для кого собрано предложение", 14, True, BLUE, 0),
            ("Салоны красоты. Четыре формата: частный мастер, небольшая студия, салон, сеть салонов.", 13, False, WHITE, 8),
            ("В одном сервисе: запись, клиенты, сотрудники и зарплата, деньги, сводка по работе, уведомления, витрина услуг.", 13, False, WHITE, 6),
            ("Сервис российской компании «Форус», в реестре российского программного обеспечения. Запись 31611 от 30 октября 2025 года.", 13, False, WHITE, 6),
        ],
    )
    add_card(slide, emu(8.2), emu(1.76), emu(4.4), emu(2.35), TEAL_DEEP, 0.08)
    add_lines(
        slide, emu(8.4), emu(1.90), emu(4.05), emu(2.05),
        [
            ("Цена в материалах", 14, True, WHITE, 0),
            ("0 ₽ в месяц. Запись и календарь, до 5 сотрудников, 1 филиал.", 13, False, WHITE, 8),
            ("2 100 ₽. Набор «Управление».", 13, False, WHITE, 6),
            ("2 300 ₽. Набор «Премиум».", 13, False, WHITE, 4),
        ],
    )

    facts = [
        ("Переход", "Оставшийся оплаченный срок в другом сервисе компенсируем. Пример в материалах: осталось 6 месяцев, получаете 6 месяцев здесь без доплаты."),
        ("Подключение", "Помогаем настроить сервис, перенести базу, завести услуги и сотрудников, показать команде, как работать."),
        ("Демонстрация", "Бесплатная, онлайн или очно. Запись по ссылке из материалов."),
    ]
    for i, (title, body) in enumerate(facts):
        x = 0.7 + i * 4.05
        add_card(slide, emu(x), emu(4.26), emu(3.90), emu(1.55), DARK, 0.08)
        add_lines(slide, emu(x + 0.16), emu(4.38), emu(3.58), emu(1.32), [(title, 14, True, BLUE, 0), (body, 12, False, WHITE, 6)])
    notes(slide, "Не переписывать здесь цены. Если набор «Управление» и «Премиум» надо расшифровать по составу, это отдельная строка в предложении, не спор на этом слайде.")
    return slide


def slide_november(prs):
    slide = content_slide(prs, "Ноябрь. Холодные звонки по салонам", 24)
    kicker(slide, "Сумма дневных строк со 2 по 20 ноября 2026. 4 ноября выходной, 9 ноября строка пустая, 23 ноября без цифр")

    stats = [
        ("540", "звонков"),
        ("221", "дошли до того, кто решает"),
        ("35", "подтвердили интерес"),
        ("16", "ответили по делу"),
    ]
    for i, (num, label) in enumerate(stats):
        x = 0.7 + i * 3.05
        add_card(slide, emu(x), emu(1.78), emu(2.92), emu(1.28), CARD, 0.1)
        add_lines(slide, emu(x + 0.16), emu(1.90), emu(2.60), emu(1.05), [(num, 26, True, BLUE, 0), (label, 13, False, WHITE, 2)])

    more = [
        ("4", "встречи назначены"),
        ("1", "показали сервис"),
        ("3", "открыли тестовый доступ"),
    ]
    for i, (num, label) in enumerate(more):
        x = 0.7 + i * 3.05
        add_card(slide, emu(x), emu(3.20), emu(2.92), emu(1.05), DARK, 0.1)
        add_lines(
            slide, emu(x + 0.16), emu(3.32), emu(2.60), emu(0.82),
            [(num, 22, True, WHITE, 0), (label, 13, False, MUTED, 1)],
        )

    add_card(slide, emu(9.85), emu(3.20), emu(2.75), emu(1.05), TEAL_DEEP, 0.1)
    add_lines(
        slide, emu(10.0), emu(3.32), emu(2.48), emu(0.82),
        [("База", 12, True, BLUE, 0), ("Холодная база салонов красоты", 13, False, WHITE, 4)],
    )

    add_lines(
        slide, emu(0.7), emu(4.45), emu(11.6), emu(1.3),
        [
            ("Что написано в комментариях за эти дни", 14, True, BLUE, 0),
            ("Клиенты сравнивают сервис с DIKIDI и спрашивают про скидку и рассылки.", 14, False, WHITE, 6),
            ("Демонстрация дважды срывалась, показ переносили.", 14, False, WHITE, 4),
            ("16 ноября телефония не работала 12 часов.", 14, False, WHITE, 4),
        ],
    )
    notes(slide, "В исходной таблице столбцы называются «контакт с ЛПР», «квалификация», «обратная связь». На слайде те же числа обычными словами. 221, 35 и 16 не складывать между собой как воронку одного человека: это сумма отметок по дням.")
    return slide


def slide_phrase(prs):
    slide = content_slide(prs, "Одна фраза для клиента", 26)
    kicker(slide, "18 минут вместе со следующим слайдом. Готовую фразу не приносим, пишем здесь")

    add_card(slide, emu(0.7), emu(1.78), emu(11.9), emu(3.85), CARD_LIGHT, 0.06)
    add_lines(
        slide, emu(0.95), emu(1.95), emu(11.4), emu(0.7),
        [
            ("Фраза, с которой менеджер начинает разговор", 16, True, TEAL_DEEP, 0),
            ("Одна. Её говорят все. Второй фразы в этом разговоре нет.", 13, False, INK, 4),
        ],
    )
    y = emu(2.85)
    for _ in range(4):
        rule(slide, emu(0.95), y, emu(11.4))
        y += emu(0.55)
    notes(slide, "Прочитать вслух то, что записали. Если фразу нельзя сказать по телефону за двадцать секунд, она не годится.")
    return slide


def slide_free(prs):
    slide = content_slide(prs, "Бесплатное подключение", 26)
    kicker(slide, "Формулировку пишем на встрече. Слева только то, что уже напечатано в материалах")

    add_card(slide, emu(0.7), emu(1.76), emu(4.55), emu(3.95), CARD, 0.08)
    add_lines(
        slide, emu(0.9), emu(1.92), emu(4.2), emu(3.6),
        [
            ("Уже напечатано", 15, True, BLUE, 0),
            ("Можно начать с 0 рублей в месяц.", 13, False, WHITE, 10),
            ("Оплаченный срок в другом сервисе компенсируем.", 13, False, WHITE, 8),
            ("Демонстрация бесплатная, онлайн или очно.", 13, False, WHITE, 8),
            ("При переходе помогаем настроить, перенести базу и показать команде, как работать.", 13, False, WHITE, 8),
        ],
    )
    add_card(slide, emu(5.45), emu(1.76), emu(7.15), emu(3.95), CARD_LIGHT, 0.06)
    labels = [
        (1.92, "Кому предлагаем"),
        (2.70, "На какой срок"),
        (3.48, "Что делаем бесплатно"),
        (4.55, "Что клиент всё равно оплачивает"),
    ]
    for y, label in labels:
        add_lines(slide, emu(5.68), emu(y), emu(6.7), emu(0.28), [(label, 13, True, TEAL_DEEP, 0)])
        rule(slide, emu(5.68), emu(y + 0.40), emu(6.7))
    notes(slide, "Не собирать из левой колонки новую рекламную фразу. Левая колонка нужна, чтобы не обещать то, чего в материалах нет, и не забыть то, что там уже есть.")
    return slide


def slide_offer(prs):
    slide = content_slide(prs, "Одно предложение для всех менеджеров", 24)
    kicker(slide, "16 минут. Заполняем строки. Текст предложения после встречи собирает центр продаж")

    rows = [
        "Кому отправляем в первую очередь",
        "Первая фраза",
        "Какую цену называем",
        "Что обещаем сделать при подключении",
        "Чего не обещаем",
        "Какой один следующий шаг предлагаем",
    ]
    top = 1.74
    for label in rows:
        add_card(slide, emu(0.7), emu(top), emu(11.9), emu(0.56), CARD_LIGHT, 0.1)
        add_lines(slide, emu(0.88), emu(top + 0.04), emu(11.5), emu(0.22), [(label, 13, True, TEAL_DEEP, 0)])
        rule(slide, emu(0.88), emu(top + 0.34), emu(11.5))
        top += 0.60
    notes(slide, "Один файл. Текст для Точки и для UDS в этот файл не кладём. Цену не выдумываем: она уже есть на слайде с материалами, здесь записываем, какую из трёх называем первой.")
    return slide


def slide_capacity(prs):
    slide = content_slide(prs, "Сколько салонов успеваем подключить", 24)
    kicker(slide, "Берём меньшее из двух чисел: сколько доводит центр продаж и сколько берёт производство")

    add_card(slide, emu(0.7), emu(1.76), emu(11.9), emu(0.7), TEAL_DEEP, 0.08)
    add_lines(
        slide, emu(0.9), emu(1.86), emu(11.5), emu(0.52),
        [("Если для подключения нужна доработка сервиса, это задача лаборатории на следующий слайд, а не лишнее место в этом месяце.", 14, False, WHITE, 0)],
        anchor=MSO_ANCHOR.MIDDLE,
    )

    headers_y = 2.60
    add_lines(slide, emu(0.88), emu(headers_y), emu(4.2), emu(0.26), [("Кого подключаем", 12, True, BLUE, 0)])
    add_lines(slide, emu(5.3), emu(headers_y), emu(3.6), emu(0.26), [("Почему строка в списке", 12, True, BLUE, 0)])
    add_lines(slide, emu(9.2), emu(headers_y), emu(3.1), emu(0.26), [("Сколько в месяц", 12, True, BLUE, 0)])

    rows = [
        ("Частный мастер", "Есть в новых материалах"),
        ("Небольшая студия", "Есть в новых материалах"),
        ("Салон красоты", "По этой базе шли звонки в ноябре"),
        ("Сеть салонов", "Есть в новых материалах"),
        ("Точка и UDS", "Это партнёры, не салон"),
    ]
    y = 2.92
    for i, (name, why) in enumerate(rows):
        add_card(slide, emu(0.7), emu(y), emu(11.9), emu(0.46), CARD if i % 2 == 0 else DARK, 0.12)
        add_lines(slide, emu(0.88), emu(y), emu(4.2), emu(0.46), [(name, 13, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(5.3), emu(y), emu(3.7), emu(0.46), [(why, 12, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(9.2), emu(y), emu(3.1), emu(0.46), [("пишем на встрече", 13, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        y += 0.50
    notes(slide, "Нужны числа, не пожелания. Производство называет, сколько подключений тянет. Продажи называют, сколько доводят. Записываем меньшее. Кто останавливает новую обещанную дату, когда места кончились, называем по имени.")
    return slide


def slide_year(prs):
    slide = content_slide(prs, "Что доделываем в сервисе за год", 24)
    kicker(slide, "20 минут. Не больше шести задач. Седьмую берём только вместо одной из шести")

    add_card(slide, emu(0.7), emu(1.76), emu(5.35), emu(3.95), CARD, 0.08)
    add_lines(
        slide, emu(0.9), emu(1.90), emu(5.0), emu(3.65),
        [
            ("Уже показано клиенту", 15, True, BLUE, 0),
            ("Значит, это должно работать на демонстрации.", 12, False, SOFT, 4),
            ("Запись, календарь, загрузка сотрудников.", 13, False, WHITE, 8),
            ("База клиентов и история.", 13, False, WHITE, 5),
            ("Графики, показатели, расчёт зарплаты.", 13, False, WHITE, 5),
            ("Доходы, расходы и сводка по салону.", 13, False, WHITE, 5),
            ("Уведомления, Яндекс Карты, 2ГИС, витрина.", 13, False, WHITE, 5),
            ("Перенос данных из другого сервиса.", 13, False, WHITE, 5),
        ],
    )
    add_card(slide, emu(6.2), emu(1.76), emu(6.4), emu(3.70), CARD_LIGHT, 0.06)
    add_lines(slide, emu(6.4), emu(1.90), emu(6.0), emu(0.5), [("Шесть задач на год", 15, True, TEAL_DEEP, 0), ("Задача и месяц, к которому она нужна продажам", 12, False, INK, 2)])
    y = emu(2.65)
    for n in range(1, 7):
        add_lines(slide, emu(6.4), y - emu(0.22), emu(0.4), emu(0.24), [(str(n), 13, True, BLUE, 0)])
        rule(slide, emu(6.85), y, emu(5.45))
        y += emu(0.46)
    notes(slide, "Слева не план, а долг перед уже напечатанным предложением. Справа пишем только то, без чего менеджер не может говорить честно. Склад, касса и медицинские карты сами на слайд не ставим.")
    return slide


def slide_teams(prs):
    slide = content_slide(prs, "Что команды просят друг у друга", 24)
    kicker(slide, "16 минут на этот слайд и следующий. Подключение и обучение клиента делает производство")

    add_card(slide, emu(0.7), emu(1.76), emu(5.9), emu(3.95), CARD_LIGHT, 0.06)
    add_lines(
        slide, emu(0.9), emu(1.90), emu(5.5), emu(0.7),
        [("Центр продаж просит у лаборатории", 15, True, TEAL_DEEP, 0), ("Пишем конкретные просьбы, не пожелания.", 12, False, INK, 3)],
    )
    y = emu(2.85)
    for _ in range(5):
        rule(slide, emu(0.9), y, emu(5.5))
        y += emu(0.48)

    add_card(slide, emu(6.8), emu(1.76), emu(5.8), emu(3.95), CARD_LIGHT, 0.06)
    add_lines(
        slide, emu(7.0), emu(1.90), emu(5.4), emu(0.7),
        [("Лаборатория просит у центра продаж", 15, True, TEAL_DEEP, 0), ("Пишем конкретные просьбы, не пожелания.", 12, False, INK, 3)],
    )
    y = emu(2.85)
    for _ in range(5):
        rule(slide, emu(7.0), y, emu(5.4))
        y += emu(0.48)
    notes(slide, "Дать каждой стороне вычеркнуть лишнее, если строк не хватает. Производство подтверждает, что запуск клиента не ждут от лаборатории.")
    return slide


def slide_timing(prs):
    slide = content_slide(prs, "За сколько отвечаем друг другу", 24)
    kicker(slide, "Срок пишем в правой колонке. Новые строки не добавляем, пока не заполнены эти")

    rows = [
        "У подключённого клиента не работает запись",
        "Менеджер спрашивает, можно ли это обещать",
        "Клиента не получается подключить в срок",
        "Клиент отказался после разговора или демонстрации",
        "Вышло обновление сервиса",
    ]
    add_lines(slide, emu(0.88), emu(1.74), emu(7.4), emu(0.26), [("Что случилось", 12, True, BLUE, 0)])
    add_lines(slide, emu(8.5), emu(1.74), emu(3.8), emu(0.26), [("Кто и за какой срок", 12, True, BLUE, 0)])
    y = 2.04
    for i, event in enumerate(rows):
        add_card(slide, emu(0.7), emu(y), emu(11.9), emu(0.52), CARD if i % 2 == 0 else DARK, 0.12)
        add_lines(slide, emu(0.88), emu(y), emu(7.3), emu(0.52), [(event, 14, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(8.5), emu(y), emu(3.8), emu(0.52), [("пишем на встрече", 13, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        y += 0.56
    notes(slide, "Три общие цифры на месяц можно дописать под сроками, если останется время: сколько подключили против договорённого числа, сколько разговоров дошли до предложения, сколько обещаний не совпало с тем, что сервис умеет.")
    return slide


def slide_compare(prs):
    slide = content_slide(prs, "С кем нас сравнивают", 26)
    kicker(slide, "12 минут. В ноябре в комментариях клиенты сравнивали с DIKIDI. Остальных называем здесь")

    titles = ["Первый сервис", "Второй сервис", "Третий сервис"]
    for i, title in enumerate(titles):
        x = 0.7 + i * 4.05
        add_card(slide, emu(x), emu(1.76), emu(3.90), emu(3.68), CARD_LIGHT, 0.06)
        add_lines(slide, emu(x + 0.18), emu(1.90), emu(3.54), emu(0.3), [(title, 14, True, TEAL_DEEP, 0)])
        add_lines(slide, emu(x + 0.18), emu(2.28), emu(3.54), emu(0.24), [("Название", 12, False, INK, 0)])
        rule(slide, emu(x + 0.18), emu(2.62), emu(3.54))
        add_lines(slide, emu(x + 0.18), emu(2.82), emu(3.54), emu(0.24), [("Что говорим клиенту", 12, False, INK, 0)])
        rule(slide, emu(x + 0.18), emu(3.20), emu(3.54))
        rule(slide, emu(x + 0.18), emu(3.62), emu(3.54))
        add_lines(slide, emu(x + 0.18), emu(3.82), emu(3.54), emu(0.24), [("Чего не обещаем", 12, False, INK, 0)])
        rule(slide, emu(x + 0.18), emu(4.20), emu(3.54))
        rule(slide, emu(x + 0.18), emu(4.62), emu(3.54))
    notes(slide, "Первую карточку логично отдать DIKIDI, потому что это единственное имя из ноябрьских комментариев. Не подсказывать фразу. Если фразу нельзя произнести, карточка пустая.")
    return slide


def slide_partners(prs):
    slide = content_slide(prs, "Что говорим Точке и что говорим UDS", 24)
    kicker(slide, "Это разговор с партнёром, не абзац в предложении для салона. Текст пишем здесь")

    blocks = [
        (0.7, "Точка"),
        (6.8, "UDS"),
    ]
    labels = ["Кому внутри компании говорим", "Какая фраза", "О чём просим", "Чего в этом году не просим"]
    for x, title in blocks:
        add_card(slide, emu(x), emu(1.76), emu(5.85), emu(3.70), CARD_LIGHT, 0.06)
        add_lines(slide, emu(x + 0.2), emu(1.90), emu(5.45), emu(0.32), [(title, 16, True, TEAL_DEEP, 0)])
        y = 2.36
        for label in labels:
            add_lines(slide, emu(x + 0.2), emu(y), emu(5.45), emu(0.24), [(label, 12, True, INK, 0)])
            rule(slide, emu(x + 0.2), emu(y + 0.36), emu(5.45))
            y += 0.78
    notes(slide, "Если под UDS имеют в виду другую компанию, а не сервис лояльности, меняется только подпись адресата. Фразу всё равно пишут в зале.")
    return slide


def slide_decisions(prs):
    slide = content_slide(prs, "Что записываем в конце", 26)
    kicker(slide, "Пустая строка значит, что вопрос не закрыт. Тогда рядом появляются имя и дата")

    rows = [
        ("1", "Фраза клиенту", "Дословный текст"),
        ("2", "Бесплатное подключение", "Кому, срок, что бесплатно, что платно"),
        ("3", "Предложение", "Кто собирает один файл и к какой дате"),
        ("4", "Подключения", "Сколько в месяц и кто останавливает лишние обещания"),
        ("5", "Сервис на год", "Шесть задач и месяц каждой"),
        ("6", "Точка и UDS", "Две фразы. В предложение для салона их нет"),
    ]
    y = 1.74
    for num, title, body in rows:
        add_card(slide, emu(0.7), emu(y), emu(11.9), emu(0.60), DARK, 0.1)
        add_lines(slide, emu(0.88), emu(y), emu(0.45), emu(0.60), [(num, 16, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(1.45), emu(y), emu(3.3), emu(0.60), [(title, 14, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(4.8), emu(y), emu(7.5), emu(0.60), [(body, 14, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE)
        y += 0.66
    notes(slide, "Читать строку и записывать ответ дословно. «Примерно договорились» не считается.")
    return slide


def slide_after(prs):
    slide = content_slide(prs, "Что приносим через два рабочих дня", 24)
    kicker(slide, "Имена ставим в конце встречи")

    items = [
        ("01", "Центр продаж", "Один файл предложения по строкам, которые заполнили."),
        ("02", "Лаборатория", "Список на одну страницу: уже есть, будет к такой-то дате, этого не делаем."),
        ("03", "Производство и продажи", "Таблица: тип салона и сколько подключений в месяц."),
        ("04", "Отдельный лист", "Фразы для Точки и для UDS. В предложение для салона их не вставлять."),
    ]
    for i, (num, who, body) in enumerate(items):
        y = 1.78 + i * 1.00
        add_card(slide, emu(0.7), y, emu(11.9), emu(0.90), CARD, 0.08)
        add_lines(slide, emu(0.9), emu(y), emu(0.7), emu(0.90), [(num, 16, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(1.75), emu(y + 0.12), emu(10.5), emu(0.68), [(who, 15, True, WHITE, 0), (body, 14, False, MUTED, 2)])
    notes(slide, "Следующая короткая встреча через две недели смотрит только этот список. Новые темы туда не добавляем.")
    return slide


def slide_close(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_LAST])
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Сначала запись,", 30, True, WHITE)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "потом следующий вопрос", 28, True, BLUE)
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Через две недели проверяем только то,\nчто записали сегодня.", 16, False, WHITE)
        elif ph.placeholder_format.idx == 10:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Блиц-запись  ·  центр продаж, производство, лаборатория", 12, False, SOFT)
    notes(slide, "Новый вопрос из зала не открывать. Его дата уже есть, если вопрос не закрыли.")
    return slide


def build():
    if not TEMPLATE.exists():
        raise SystemExit(f"Template not found: {TEMPLATE}")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(TEMPLATE, OUT)
    prs = Presentation(str(OUT))
    delete_all_slides(prs)

    slide_title(prs)
    slide_questions(prs)
    slide_agenda(prs)
    slide_materials(prs)
    slide_november(prs)
    slide_phrase(prs)
    slide_free(prs)
    slide_offer(prs)
    slide_capacity(prs)
    slide_year(prs)
    slide_teams(prs)
    slide_timing(prs)
    slide_compare(prs)
    slide_partners(prs)
    slide_decisions(prs)
    slide_after(prs)
    slide_close(prs)

    prs.core_properties.title = "Планерка «Блиц-запись». Два часа"
    prs.core_properties.subject = "Повестка. Формулировки для клиента заполняются на встрече"
    prs.core_properties.category = "Внутренняя планерка"
    prs.save(str(OUT))
    print(f"Saved: {OUT}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    build()
