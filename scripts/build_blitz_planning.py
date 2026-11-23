#!/usr/bin/env python3
"""Рабочая презентация двухчасовой планерки по сервису «Блиц-запись».

Черновики на слайдах собраны из текущего КП, конкурентного анализа 2026
и выгрузки лидов. Это материал для решения в зале, не согласованная стратегия.
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
CARD_LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
GRAY = RGBColor(0x76, 0x76, 0x77)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xBF, 0xBF, 0xBF)
NEAR_BLACK = RGBColor(0x1A, 0x1A, 0x1A)
DARK = RGBColor(0x2A, 0x2A, 0x2A)
MUTED = RGBColor(0xD0, 0xD0, 0xD0)
TEAL_DEEP = RGBColor(0x14, 0x4A, 0x63)

FONT = "Verdana"

L_TITLE = 0
L_CONTENT = 3
L_BG = 6
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


def add_lines(
    slide,
    left,
    top,
    width,
    height,
    lines,
    anchor=MSO_ANCHOR.TOP,
    align=PP_ALIGN.LEFT,
):
    """lines: sequence of (text, size, bold, color, space_before_pt)."""
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
    return add_lines(
        slide, emu(0.97), emu(0.42), emu(11.4), emu(0.7),
        [(text, size, True, WHITE, 0)],
    )


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def content_slide(prs, title, size=24):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, title, size)
    clear_body(slide)
    return slide


def kicker(slide, text):
    add_lines(
        slide, emu(0.97), emu(1.42), emu(11.3), emu(0.28),
        [(text, 11, True, BLUE, 0)],
    )


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------

def slide_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_TITLE])
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Планерка по сервису", 26, True, WHITE)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "«Блиц-запись»", 32, True, BLUE)
            p3 = tf.add_paragraph()
            p3.space_before = Pt(14)
            set_run(p3.add_run(), "Повестка на 2 часа", 18, False, WHITE)
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            set_run(
                p.add_run(),
                "Лаборатория, производство и центр продаж.\n"
                "Шесть вопросов — шесть решений.",
                14, False, SOFT,
            )
    notes(
        slide,
        "Открыть так: сегодня не смотрим продукт и не защищаем план. "
        "Выходим с шестью записанными решениями. Черновики на слайдах можно отвергнуть.",
    )
    return slide


def slide_outcomes(prs):
    slide = content_slide(prs, "С чем выходим через 120 минут", 26)
    kicker(slide, "ВСТРЕЧА ЗАКРЫТА, ТОЛЬКО ЕСЛИ НАЗВАНЫ ВСЕ ШЕСТЬ РЕШЕНИЙ")

    items = [
        ("1", "УТП и единое КП", "Одна главная фраза, один сегмент квартала, каркас предложения, который говорят все."),
        ("2", "План подключений", "Квота запусков в месяц. Её задаёт меньшая мощность: центр продаж или производство."),
        ("3", "Техкарта на год", "Не больше шести пунктов. У каждого квартал, владелец и запрет обещать раньше срока."),
        ("4", "Ожидания команд", "Лаборатория и центр продаж: действия, срок реакции, три общих показателя."),
        ("5", "Отстройка", "Три карточки: против кого говорим в сделке и чего в этой сделке не обещаем."),
        ("6", "Точка и UDS", "Две разные фразы для двух каналов. Интеграцию в этот год — да или нет."),
    ]
    top = 1.70
    for num, title, body in items:
        y = emu(top)
        add_card(slide, emu(0.97), y, emu(11.4), emu(0.64), DARK, 0.1)
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(1.12), y + emu(0.12), emu(0.40), emu(0.40))
        badge.fill.solid()
        badge.fill.fore_color.rgb = BLUE
        badge.line.fill.background()
        add_lines(
            slide, emu(1.12), y + emu(0.16), emu(0.40), emu(0.30),
            [(num, 13, True, WHITE, 0)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_lines(
            slide, emu(1.7), y, emu(3.3), emu(0.64),
            [(title, 14, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE,
        )
        add_lines(
            slide, emu(5.05), y, emu(7.05), emu(0.64),
            [(body, 12, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE,
        )
        top += 0.68
    notes(
        slide,
        "Проговорить роли, если в зале спор, кто есть кто. "
        "Лаборатория — продукт и техкарта. Производство — запуски клиентов. "
        "Центр продаж — обещание клиенту и квота. Если роли называются иначе, поправить в первые пять минут.",
    )
    return slide


def slide_agenda(prs):
    slide = content_slide(prs, "Повестка на 120 минут", 26)
    kicker(slide, "ЧАСЫ ИДУТ С НАЧАЛА ВСТРЕЧИ. НЕЗАКРЫТЫЙ ВОПРОС ПОЛУЧАЕТ ВЛАДЕЛЬЦА И ДАТУ")

    rows = [
        ("0:00", "5", "Рамка", "Роли, правила, что считаем решением"),
        ("0:05", "22", "1. УТП и единое КП", "Фраза, сегмент квартала, каркас предложения"),
        ("0:27", "18", "2. План подключений", "Квота месяца и приоритет сегментов"),
        ("0:45", "22", "3. Техкарта на год", "До шести пунктов, квартал и владелец"),
        ("1:07", "18", "4. Ожидания команд", "Действия, реакция, три показателя"),
        ("1:25", "15", "5. Отстройка", "Три карточки против реальных конкурентов"),
        ("1:40", "15", "6. Точка и UDS", "Две фразы и ответ, есть ли стык в этом году"),
        ("1:55", "5", "Фиксация", "Лист решений вслух, имена, даты"),
    ]
    top = 1.72
    add_lines(slide, emu(0.97), emu(top), emu(1.15), emu(0.24), [("Старт", 11, True, BLUE, 0)])
    add_lines(slide, emu(2.15), emu(top), emu(0.7), emu(0.24), [("Мин", 11, True, BLUE, 0)])
    add_lines(slide, emu(2.95), emu(top), emu(4.2), emu(0.24), [("Вопрос", 11, True, BLUE, 0)])
    add_lines(slide, emu(7.3), emu(top), emu(4.6), emu(0.24), [("На выходе из блока", 11, True, BLUE, 0)])
    top += 0.28
    for i, (clock, mins, title, out) in enumerate(rows):
        y = emu(top)
        fill = CARD if i % 2 == 0 else DARK
        add_card(slide, emu(0.97), y, emu(11.4), emu(0.42), fill, 0.12)
        add_lines(slide, emu(1.12), y, emu(1.05), emu(0.42), [(clock, 13, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(2.15), y, emu(0.7), emu(0.42), [(mins, 13, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(2.95), y, emu(4.15), emu(0.42), [(title, 13, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(7.3), y, emu(4.85), emu(0.42), [(out, 12, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE)
        top += 0.46
    notes(
        slide,
        "Жёсткий тайминг. Если блок не закрыт, ведущий обрывает и пишет владельца и дату возврата. "
        "Вопросы 1–4 не отдаём в пользу 5–6: без них партнёрские фразы не на что опереть.",
    )
    return slide


def slide_facts(prs):
    slide = content_slide(prs, "Коммерческое предложение, как оно есть", 24)
    kicker(slide, "ИСТОЧНИК: КП «БЛИЦ-ЗАПИСЬ.РФ». ЭТО ТОЧКА СТАРТА, НЕ РЕШЕНИЕ ВОПРОСА 1")

    add_card(slide, emu(0.97), emu(1.78), emu(7.35), emu(2.55), CARD, 0.08)
    add_lines(
        slide, emu(1.18), emu(1.92), emu(6.95), emu(2.25),
        [
            ("Что обещаем сейчас", 14, True, BLUE, 0),
            ("Сервис записей, расписания, клиентской базы и загрузки сотрудников.", 13, False, WHITE, 8),
            ("Можно начать с внутреннего расписания и вторым шагом включить онлайн-запись, уведомления, виджет и аналитику.", 13, False, WHITE, 6),
            ("Запуск сопровождают эксперты. Платформа заявлена в реестре российского ПО Минцифры.", 13, False, WHITE, 6),
        ],
    )

    add_card(slide, emu(8.48), emu(1.78), emu(3.9), emu(2.55), TEAL_DEEP, 0.08)
    add_lines(
        slide, emu(8.68), emu(1.92), emu(3.52), emu(2.25),
        [
            ("Где расхождение", 14, True, BLUE, 0),
            ("В КП семь аудиторий сразу.", 13, False, WHITE, 8),
            ("В воронке — салоны.", 13, False, WHITE, 4),
            ("В конкурентном файле прямые конкуренты — автомойки, бьюти-системы записаны как косвенные.", 13, False, WHITE, 4),
        ],
    )

    prices = [
        ("2 года", "1 500 ₽", "вместо 2 500"),
        ("1 год", "1 900 ₽", "вместо 2 500"),
        ("6 месяцев", "2 300 ₽", "вместо 2 500"),
        ("3 месяца", "2 500 ₽", "без скидки"),
    ]
    for i, (term, price, note) in enumerate(prices):
        x = 0.97 + i * 2.95
        add_card(slide, emu(x), emu(4.50), emu(2.82), emu(1.22), DARK, 0.1)
        add_lines(
            slide, emu(x + 0.16), emu(4.60), emu(2.52), emu(1.02),
            [
                (term, 12, False, SOFT, 0),
                (price + " / мес", 16, True, WHITE, 2),
                (note, 11, False, BLUE, 1),
            ],
        )

    add_lines(
        slide, emu(0.97), emu(5.82), emu(9.3), emu(0.55),
        [(
            "Спецпредложение в КП: при переходе с другого сервиса увеличим срок подписки до 6 месяцев. "
            "Аудитории на листе: салоны, мойки, клиники, фитнес, обучение, частные специалисты.",
            12, False, SOFT, 0,
        )],
    )
    notes(
        slide,
        "Не обсуждать здесь правку формулировок. Задача слайда — общее фото. "
        "Развилку «салоны или мойки» перенести в вопрос 1: без неё техкарта и квота разъедутся.",
    )
    return slide


def slide_funnel(prs):
    slide = content_slide(prs, "Что уже показывает воронка", 26)
    kicker(slide, "ВЫГРУЗКА «ЛИДЫ БЛИЦ КОРШАК», 10 СЕНТЯБРЯ — 16 ОКТЯБРЯ 2026, 162 ЛИДА")

    stats = [
        ("162", "лида в выгрузке"),
        ("13", "дошли до стадии КП"),
        ("39", "недозвон больше 3 раз"),
        ("~20", "карточек с YCLIENTS"),
    ]
    for i, (num, label) in enumerate(stats):
        x = 0.97 + i * 2.95
        add_card(slide, emu(x), emu(1.82), emu(2.82), emu(1.35), CARD, 0.1)
        add_lines(
            slide, emu(x + 0.16), emu(1.94), emu(2.52), emu(1.12),
            [(num, 26, True, BLUE, 0), (label, 12, False, WHITE, 2)],
        )

    quotes = [
        ("Уже сидят в системе", "В карточках прямо названы YCLIENTS и бесплатный тариф DIKIDI. С ними спорим чаще, чем с «мойками» из конкурентного файла."),
        ("Онлайн-запись не нужна", "Есть отказы «записываем только после разговора» и «не нравится формат онлайн-записи». В КП такой сценарий уже разрешён."),
        ("До КП доходит мало", "13 из 162 — это стадия предложения, не оплата. Единое КП нужно не для красоты, а чтобы эти 13 звучали одинаково."),
    ]
    for i, (title, body) in enumerate(quotes):
        x = 0.97 + i * 3.93
        add_card(slide, emu(x), emu(3.35), emu(3.78), emu(2.35), DARK, 0.08)
        add_lines(
            slide, emu(x + 0.18), emu(3.50), emu(3.42), emu(2.05),
            [(title, 14, True, BLUE, 0), (body, 12, False, WHITE, 8)],
        )
    notes(
        slide,
        "Цифры не раздувать. 13 — это «лид сконвертирован (КП)», не продажи. "
        "«Около 20» — упоминания YCLIENTS / «ЮК» в стадии или комментарии, не отдельное поле CRM. "
        "DIKIDI в причинах встречается прямой фразой про бесплатный тариф.",
    )
    return slide


def slide_usp(prs):
    slide = content_slide(prs, "1. УТП: одна главная фраза", 26)
    kicker(slide, "22 МИНУТЫ. ВЫБИРАЕМ ГЛАВНУЮ И ПОДДЕРЖИВАЮЩУЮ. ТРЕТЬЮ МЕНЕДЖЕРАМ НЕ ДАЁМ")

    options = [
        (
            "А",
            "Внедрение, а не кабинет",
            "«Подключим запись под ключ и останемся рядом. Отечественное ПО из реестра Минцифры.»",
            "Бьёт в YCLIENTS и DIKIDI, которые продают продукт самообслуживания.",
        ),
        (
            "Б",
            "Одна цена, без конструктора",
            "«Фиксированная цена за рабочее место администратора: журнал, база и виджет без доплат за модули.»",
            "Бьёт в пересчёт YCLIENTS за пользователя, платные модули DIKIDI и пакеты СБИС.",
        ),
        (
            "В",
            "Любой бизнес по записи",
            "«Один сервис для салона, мойки, клиники, студии и частного специалиста.»",
            "Широко, но пока без глубины. В воронке салоны, в файле конкурентов — мойки. Так УТП не выбираем.",
        ),
    ]
    top = 1.74
    for letter, title, phrase, why in options:
        y = emu(top)
        add_card(slide, emu(0.97), y, emu(11.4), emu(1.18), CARD, 0.08)
        add_card(slide, emu(1.12), y + emu(0.30), emu(0.58), emu(0.58), BLUE, 0.2)
        add_lines(
            slide, emu(1.12), y + emu(0.42), emu(0.58), emu(0.36),
            [(letter, 16, True, WHITE, 0)], align=PP_ALIGN.CENTER,
        )
        add_lines(
            slide, emu(1.9), y + emu(0.08), emu(10.2), emu(1.02),
            [
                (title, 14, True, WHITE, 0),
                (phrase, 13, False, BLUE, 2),
                (why, 12, False, MUTED, 1),
            ],
        )
        top += 1.26

    notes(
        slide,
        "Голосование в два шага. Сначала сегмент квартала: салоны или мойки. "
        "Потом фраза. Если сегмент — салоны, главная скорее А, поддерживающая Б. "
        "Вариант В не произносим, пока нет отдельных слотов внедрения.",
    )
    return slide


def slide_offer(prs):
    slide = content_slide(prs, "1. Каркас единого предложения", 26)
    kicker(slide, "ТЕКСТ ПОТОМ ПИШЕТ ЦЕНТР ПРОДАЖ. ГРАНИЦУ ОБЕЩАНИЙ ПОДТВЕРЖДАЕТ ЛАБОРАТОРИЯ")

    points = [
        ("1", "Кому", "Первый абзац — один сегмент квартала. Остальные аудитории списком «также подходит», не на равных."),
        ("2", "Фраза", "Главное УТП дословно, как записали. Поддерживающее — одной строкой ниже, без третьей темы."),
        ("3", "Четыре результата", "Не восемь функций. Оставляем только то, что доказывает выбранную фразу."),
        ("4", "Цена", "Как в КП: 1 500 / 1 900 / 2 300 / 2 500 ₽ в месяц и бонус до 6 месяцев при переходе."),
        ("5", "Запуск", "Что входит во внедрение, сколько дней, кто со стороны производства ведёт клиента."),
        ("6", "Чего нет", "Склад, зарплата, касса, ЕГИСЗ, своя лояльность. Этого в речи менеджера нет."),
        ("7", "Шаг", "Один следующий шаг: демо, пилот или счёт. Не все три сразу."),
        ("8", "Носитель", "Один файл КП. Партнёрские фразы для Точки и UDS в него не кладём."),
    ]
    for i, (num, title, body) in enumerate(points):
        col, row = i % 2, i // 2
        x = 0.97 + col * 5.85
        y = 1.80 + row * 1.05
        add_card(slide, emu(x), emu(y), emu(5.70), emu(0.96), DARK if row % 2 == 0 else CARD, 0.1)
        add_lines(
            slide, emu(x + 0.16), emu(y + 0.10), emu(0.4), emu(0.76),
            [(num, 16, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE,
        )
        add_lines(
            slide, emu(x + 0.58), emu(y + 0.10), emu(4.95), emu(0.76),
            [(title, 13, True, WHITE, 0), (body, 11, False, MUTED, 2)],
            anchor=MSO_ANCHOR.MIDDLE,
        )
    notes(
        slide,
        "Пройти пункты быстро и спорить только о 1, 5 и 6. "
        "Цену не пересобирать на этой встрече, если нет отдельного решения о тарифе. "
        "Назвать дату, к которой центр продаж приносит текст.",
    )
    return slide


def slide_capacity(prs):
    slide = content_slide(prs, "2. План подключений равен мощности", 24)
    kicker(slide, "18 МИНУТ. КВОТА МЕСЯЦА = МЕНЬШЕЕ ИЗ ДВУХ: ЦЕНТР ПРОДАЖ И ПРОИЗВОДСТВО")

    add_card(slide, emu(0.97), emu(1.78), emu(11.4), emu(0.78), TEAL_DEEP, 0.08)
    add_lines(
        slide, emu(1.15), emu(1.88), emu(11.05), emu(0.60),
        [(
            "Лаборатория в квоту не входит, пока запуск не требует доработки. "
            "Доработка — это пункт техкарты, а не лишний слот внедрения.",
            13, False, WHITE, 0,
        )],
        anchor=MSO_ANCHOR.MIDDLE,
    )

    y = 2.68
    add_lines(slide, emu(1.12), emu(y), emu(3.4), emu(0.28), [("Сегмент или канал", 11, True, BLUE, 0)])
    add_lines(slide, emu(4.6), emu(y), emu(4.2), emu(0.28), [("Почему в списке", 11, True, BLUE, 0)])
    add_lines(slide, emu(9.15), emu(y), emu(3.0), emu(0.28), [("Запусков в месяц", 11, True, BLUE, 0)])

    rows = [
        ("Салоны и барбершопы", "Основная воронка сентября–октября"),
        ("Автомойки", "В конкурентном файле помечены как прямые"),
        ("Клиники, фитнес, обучение", "Есть в КП, в этой выгрузке почти нет"),
        ("Частные специалисты", "Другой цикл сделки, легко забить квоту"),
        ("Точка и UDS", "Канал. Не смешивать с прямыми, пока нет договорённости"),
    ]
    y = 3.02
    for i, (name, why) in enumerate(rows):
        fill = CARD if i % 2 == 0 else DARK
        add_card(slide, emu(0.97), emu(y), emu(11.4), emu(0.46), fill, 0.12)
        add_lines(slide, emu(1.12), emu(y), emu(3.4), emu(0.46), [(name, 12, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(4.6), emu(y), emu(4.4), emu(0.46), [(why, 11, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(9.15), emu(y), emu(3.0), emu(0.46), [("ставим в зале", 12, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        y += 0.50

    add_lines(
        slide, emu(0.97), emu(5.58), emu(7.3), emu(0.28),
        [("Сверх квоты срок не обещаем. Кто говорит «стоп» — называем на фиксации.", 12, False, SOFT, 0)],
    )
    notes(
        slide,
        "Не уходить в спор о годовом плане продаж. Нужны три числа: запуски салонов, запуски моек, "
        "всё остальное пауза или отдельная маленькая квота. Производство называет потолок, продажи — спрос. Берём меньшее.",
    )
    return slide


def slide_roadmap_rule(prs):
    slide = content_slide(prs, "3. Что имеет право попасть в техкарту", 24)
    kicker(slide, "22 МИНУТЫ НА ЭТОТ И СЛЕДУЮЩИЙ СЛАЙД. В ГОД БЕРЁМ НЕ БОЛЬШЕ ШЕСТИ ПУНКТОВ")

    add_card(slide, emu(0.97), emu(1.78), emu(5.7), emu(3.48), CARD, 0.08)
    add_lines(
        slide, emu(1.15), emu(1.92), emu(5.35), emu(3.40),
        [
            ("Уже обещано в КП", 15, True, BLUE, 0),
            ("Сначала сделать это честным на демо.", 12, False, SOFT, 4),
            ("Календарь без наложений: сотрудники, кабинеты, ресурсы.", 13, False, WHITE, 10),
            ("Онлайн-запись 24/7 и запрет двойной записи.", 13, False, WHITE, 6),
            ("Виджет на сайт, в соцсети, мессенджеры и карты.", 13, False, WHITE, 6),
            ("Уведомления о записи, переносе и отмене.", 13, False, WHITE, 6),
            ("Режим «только расписание», онлайн вторым шагом.", 13, False, WHITE, 6),
        ],
    )

    add_card(slide, emu(6.85), emu(1.78), emu(5.52), emu(3.48), DARK, 0.08)
    add_lines(
        slide, emu(7.03), emu(1.92), emu(5.18), emu(3.40),
        [
            ("Берём, только если бьёт в УТП", 15, True, BLUE, 0),
            ("Иначе это желание, не пункт года.", 12, False, SOFT, 4),
            ("Перенос базы с YCLIENTS и DIKIDI.", 13, False, WHITE, 10),
            ("Уведомления в мессенджер без отдельной платы.", 13, False, WHITE, 6),
            ("Одна отраслевая глубина: мойки или клиники.", 13, False, WHITE, 6),
            ("Стык с UDS или с Точкой. Один, не оба.", 13, False, WHITE, 6),
            ("Склад, зарплата, касса, ЕГИСЗ — не в этот год.", 13, False, WHITE, 6),
        ],
    )
    notes(
        slide,
        "Левая колонка — долг перед уже напечатанным КП. Правая — меню. "
        "Седьмой пункт можно взять только взамен одного из шести. Склад и ЕГИСЗ не торговать.",
    )
    return slide


def slide_roadmap_year(prs):
    slide = content_slide(prs, "3. Черновик карты на четыре квартала", 24)
    kicker(slide, "ЧЕРНОВИК НА СЛУЧАЙ, ЕСЛИ В ВОПРОСЕ 1 ПРИОРИТЕТ — САЛОНЫ. ИНАЧЕ МЕНЯЕТСЯ СОДЕРЖИМОЕ")

    quarters = [
        ("Кв. 1", "Честное демо", "Довести то, что уже написано в КП. Чеклист запуска одного салона для производства."),
        ("Кв. 2", "Переход", "Перенос с YCLIENTS и с бесплатного DIKIDI. Иначе бонус «до 6 месяцев» нечем подтвердить."),
        ("Кв. 3", "Канал связи", "Уведомления туда, где клиент уже есть. Свою лояльность не начинаем."),
        ("Кв. 4", "Один стык", "Либо UDS, либо Точка. Второй партнёр — следующий год, если шестой пункт уже занят."),
    ]
    for i, (q, title, body) in enumerate(quarters):
        x = 0.97 + i * 2.95
        add_card(slide, emu(x), emu(1.82), emu(2.82), emu(2.85), CARD if i % 2 == 0 else DARK, 0.08)
        add_lines(
            slide, emu(x + 0.14), emu(1.96), emu(2.54), emu(2.55),
            [
                (q, 12, True, BLUE, 0),
                (title, 16, True, WHITE, 6),
                (body, 12, False, MUTED, 8),
            ],
        )

    add_card(slide, emu(0.97), emu(4.82), emu(8.35), emu(1.15), TEAL_DEEP, 0.08)
    add_lines(
        slide, emu(1.15), emu(4.94), emu(8.0), emu(0.92),
        [
            ("Развилка, которую нельзя проскочить", 14, True, WHITE, 0),
            ("Если квартал отдан мойкам, карта 2–4 про мойку, не про перенос из YCLIENTS. Две карты сразу не ведём.", 13, False, WHITE, 4),
        ],
    )
    notes(
        slide,
        "Это не план лаборатории, это последовательность. В зале каждый квартал: пункт, владелец, "
        "что менеджер перестаёт обещать до готовности. Даты внутри квартала ставят производство и лаборатория после встречи, "
        "если на встрече успели только квартал.",
    )
    return slide


def slide_expectations(prs):
    slide = content_slide(prs, "4. Что команды ждут друг от друга", 24)
    kicker(slide, "18 МИНУТ НА ЭТОТ И СЛЕДУЮЩИЙ СЛАЙД. ЗАПУСК И ОБУЧЕНИЕ ОСТАЮТСЯ У ПРОИЗВОДСТВА")

    add_card(slide, emu(0.97), emu(1.78), emu(5.7), emu(3.55), CARD_LIGHT, 0.08)
    add_lines(
        slide, emu(1.15), emu(1.92), emu(5.35), emu(3.25),
        [
            ("Центр продаж ждёт от лаборатории", 14, True, TEAL_DEEP, 0),
            ("Демо, на котором можно показать КП без сюрпризов.", 13, False, NEAR_BLACK, 10),
            ("Карточку: есть / в техкарте / не будет.", 13, False, NEAR_BLACK, 6),
            ("Реакцию, если у подключённого клиента сбой.", 13, False, NEAR_BLACK, 6),
            ("Ответ в сделке, где просят нестандарт.", 13, False, NEAR_BLACK, 6),
            ("Пять строк менеджеру до каждой выкладки.", 13, False, NEAR_BLACK, 6),
        ],
    )
    add_card(slide, emu(6.85), emu(1.78), emu(5.52), emu(3.55), CARD, 0.08)
    add_lines(
        slide, emu(7.03), emu(1.92), emu(5.18), emu(3.25),
        [
            ("Лаборатория ждёт от центра продаж", 14, True, BLUE, 0),
            ("Лид с сегментом, текущей системой и числом людей.", 13, False, WHITE, 10),
            ("Ни одного обещания вне карточки продукта.", 13, False, WHITE, 6),
            ("Срок запуска не обещают сверх квоты вопроса 2.", 13, False, WHITE, 6),
            ("По каждому КП: взял, не взял, почему.", 13, False, WHITE, 6),
            ("Пилот доводят до запуска, а не бросают на счёте.", 13, False, WHITE, 6),
        ],
    )
    notes(
        slide,
        "Дать каждой стороне вычеркнуть один пункт и добавить один. Не раздувать список. "
        "Производство подтверждает, что запуск и обучение — их действие, и его не ждут от лаборатории.",
    )
    return slide


def slide_sla(prs):
    slide = content_slide(prs, "4. Реакция и три общих показателя", 24)
    kicker(slide, "СРОКИ НИЖЕ — ПРЕДЛОЖЕНИЕ. В ЗАЛЕ ИХ ПРИНИМАЮТ ИЛИ МЕНЯЮТ, НОВЫЕ СТРОКИ НЕ ДОБАВЛЯЮТ")

    rows = [
        ("Сбой у подключённого клиента", "Лаборатория", "4 рабочих часа, первый ответ"),
        ("«Можно это обещать?»", "Лаборатория", "1 рабочий день: да / нет / в техкарте"),
        ("Дефект блокирует запуск", "Лаборатория", "1 рабочий день: обход или дата"),
        ("Лид внутри квоты на запуск", "Центр продаж", "Квалификация заполнена до передачи"),
        ("Сделка проиграна", "Центр продаж", "Причина из короткого списка в тот же день"),
        ("Выкладка", "Лаборатория", "5 строк для менеджера до выкладки"),
    ]
    y = 1.78
    add_lines(slide, emu(1.12), emu(y), emu(4.6), emu(0.26), [("Событие", 11, True, BLUE, 0)])
    add_lines(slide, emu(5.9), emu(y), emu(2.3), emu(0.26), [("Кто", 11, True, BLUE, 0)])
    add_lines(slide, emu(8.2), emu(y), emu(4.0), emu(0.26), [("Срок, предложение", 11, True, BLUE, 0)])
    y = 2.08
    for i, (event, who, term) in enumerate(rows):
        fill = CARD if i % 2 == 0 else DARK
        add_card(slide, emu(0.97), emu(y), emu(11.4), emu(0.42), fill, 0.14)
        add_lines(slide, emu(1.12), emu(y), emu(4.7), emu(0.42), [(event, 12, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(5.9), emu(y), emu(2.2), emu(0.42), [(who, 12, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE)
        add_lines(slide, emu(8.2), emu(y), emu(4.0), emu(0.42), [(term, 12, False, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE)
        y += 0.46

    add_card(slide, emu(0.97), emu(4.92), emu(8.55), emu(1.22), TEAL_DEEP, 0.08)
    add_lines(
        slide, emu(1.14), emu(5.02), emu(8.22), emu(1.02),
        [
            ("Раз в месяц смотрим только это", 13, True, WHITE, 0),
            ("Запуски против квоты. Доля лидов с квалификацией. Доля КП без обещаний вне карточки.", 12, False, WHITE, 3),
            ("Контекст, не четвёртый KPI: отказы YCLIENTS, DIKIDI, «онлайн не нужна», недозвон.", 12, False, SOFT, 2),
        ],
    )
    notes(
        slide,
        "Если спор о часах, не торговаться дольше двух минут на строку. "
        "База конверсии для разговора: 13 из 162 дошли до КП за 10.09–16.10.2026. Оплаты в этой выгрузке нет.",
    )
    return slide


def slide_battle(prs):
    slide = content_slide(prs, "5. Отстройка: три карточки в разговор", 24)
    kicker(slide, "15 МИНУТ. ЭТО НЕ ПОВТОР УТП. ЭТО ФРАЗА ПРОТИВ КОНКРЕТНОГО ОТКАЗА")

    cards = [
        (
            "YCLIENTS",
            "Тяжёлая система, в карточках лидов всплывает чаще всего.",
            "«Не переводим вас на ещё одну ERP. Включаем запись и ведём запуск. Цена фиксированная, без пересчёта за пользователя.»",
            "Не обещаем склад, зарплату и кассу «как у них».",
        ),
        (
            "DIKIDI, бесплатно",
            "Причина «и так работает». В отказах назван прямо.",
            "«Журнал пусть живёт, пока хватает. Мы нужны, когда расписание путается и вам нужен человек на запуск. При переходе добавим срок до 6 месяцев.»",
            "Не говорим, что бесплатное — это плохо.",
        ),
        (
            "Без онлайн-записи",
            "«Записываем только после разговора».",
            "«Тогда ссылку клиенту не включаем. Блиц работает как расписание администратора. Онлайн добавите, когда сами решите.»",
            "Не спорим, что без онлайна сервис не нужен.",
        ),
    ]
    for i, (title, who, phrase, dont) in enumerate(cards):
        x = 0.70 + i * 4.15
        add_card(slide, emu(x), emu(1.78), emu(4.00), emu(3.62), CARD, 0.08)
        add_lines(
            slide, emu(x + 0.18), emu(1.90), emu(3.64), emu(3.35),
            [
                (title, 15, True, BLUE, 0),
                (who, 12, False, SOFT, 6),
                (phrase, 13, False, WHITE, 10),
                (dont, 12, True, BLUE, 10),
            ],
        )
    notes(
        slide,
        "Прочитать фразы вслух так, как их скажет менеджер. Если фраза не произносится, её нет. "
        "Автомойки в этот блок не добавлять, пока вопрос 2 не отдал им квоту: иначе получится четвёртая карточка без сделок.",
    )
    return slide


def slide_partners(prs):
    slide = content_slide(prs, "6. УТП для Точки и для UDS", 26)
    kicker(slide, "15 МИНУТ. ДВЕ ФРАЗЫ. ОБЕ — ПРЕДЛОЖЕНИЕ НА СОГЛАСОВАНИЕ, НЕ ТЕКСТ КП КЛИЕНТУ")

    add_card(slide, emu(0.70), emu(1.78), emu(6.05), emu(4.05), CARD_LIGHT, 0.08)
    add_lines(
        slide, emu(0.88), emu(1.92), emu(5.70), emu(3.75),
        [
            ("Точка", 16, True, TEAL_DEEP, 0),
            ("Адресат — партнёр банка, не салон.", 12, False, GRAY, 2),
            ("«Запись для клиентов Точки, которые работают по предварительной записи. Сервис в реестре российского ПО, фиксированная цена, подключение под ключ силами Форус. Своя команда внедрения банку не нужна.»", 13, False, NEAR_BLACK, 8),
            ("Просим рекомендацию в сервисах для бизнеса.", 12, True, TEAL_DEEP, 8),
            ("Не просим в этом году вход в интернет-банк.", 12, False, NEAR_BLACK, 3),
        ],
    )
    add_card(slide, emu(6.90), emu(1.78), emu(5.75), emu(4.05), CARD, 0.08)
    add_lines(
        slide, emu(7.08), emu(1.92), emu(5.40), emu(3.75),
        [
            ("UDS", 16, True, BLUE, 0),
            ("Адресат — партнёр сервиса лояльности.", 12, False, SOFT, 2),
            ("«UDS возвращает гостя после визита. Блиц-запись сажает его в расписание. Лояльность не строим — это ваш контур. Запись не собирайте — это наш. Рекомендуем друг друга.»", 13, False, WHITE, 8),
            ("Просим рекомендацию мерчанту без нормальной записи.", 12, True, BLUE, 8),
            ("Свой кешбэк в этом году не делаем.", 12, False, WHITE, 3),
        ],
    )
    notes(
        slide,
        "Если под UDS в зале имеют в виду конкретного клиента, а не сервис лояльности, адресата фразы меняют, "
        "развилку «не строим чужой контур» оставляют. Интеграцию в год берут только если в вопросе 3 на неё остался слот квартала 4. "
        "Иначе ответ «нет» в этом году — нормальное решение.",
    )
    return slide


def slide_decisions(prs):
    slide = content_slide(prs, "Лист решений, читаем вслух", 26)
    kicker(slide, "5 МИНУТ. ПУСТАЯ СТРОКА — ВОПРОС НЕ ЗАКРЫТ. У НЕЁ ПОЯВЛЯЕТСЯ ИМЯ И ДАТА")

    rows = [
        ("1", "УТП", "Главная фраза, поддерживающая, сегмент квартала."),
        ("2", "КП", "Каркас утверждён. Текст к дате __. Границу обещаний подтверждает лаборатория."),
        ("3", "Квота", "Сколько запусков в месяц и по каким сегментам. Кто говорит «стоп»."),
        ("4", "Техкарта", "До шести пунктов и квартал каждого. Что менеджер больше не обещает."),
        ("5", "Отстройка", "Три карточки приняты или переписаны одной фразой каждая."),
        ("6", "Каналы", "Фраза для Точки. Фраза для UDS. Стык в этом году: да или нет."),
    ]
    y = 1.74
    for num, title, body in rows:
        add_card(slide, emu(0.97), emu(y), emu(11.4), emu(0.60), DARK, 0.1)
        add_lines(
            slide, emu(1.12), emu(y), emu(0.45), emu(0.60),
            [(num, 16, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE,
        )
        add_lines(
            slide, emu(1.65), emu(y), emu(1.7), emu(0.60),
            [(title, 14, True, WHITE, 0)], anchor=MSO_ANCHOR.MIDDLE,
        )
        add_lines(
            slide, emu(3.45), emu(y), emu(8.5), emu(0.60),
            [(body, 13, False, MUTED, 0)], anchor=MSO_ANCHOR.MIDDLE,
        )
        y += 0.64
    notes(
        slide,
        "Ведущий читает строку и записывает ответ в протокол дословно. "
        "Не пересказывать «мы примерно договорились».",
    )
    return slide


def slide_after(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_BG])
    fill_title(slide, "Два рабочих дня после планерки", 26)
    clear_body(slide)
    kicker(slide, "ИМЕНА СТАВИМ НА ФИКСАЦИИ. ЗДЕСЬ — СОСТАВ РАБОТЫ, НЕ ПОРУЧЕНИЕ ВОЗДУХУ")

    items = [
        ("01", "Центр продаж", "Текст единого КП по каркасу. Одна версия, без личных вариаций."),
        ("02", "Лаборатория", "Карточка «есть / в техкарте / не будет» и техкарта на одну страницу."),
        ("03", "Производство и продажи", "Таблица квоты на квартал: сегмент, запуски в месяц, кто говорит «стоп»."),
        ("04", "Отдельный лист", "Фразы для Точки и UDS. В клиентское КП их не вставлять."),
    ]
    for i, (num, who, body) in enumerate(items):
        y = 1.82 + i * 1.05
        add_card(slide, emu(0.97), emu(y), emu(11.4), emu(0.95), CARD, 0.08)
        add_lines(
            slide, emu(1.15), emu(y), emu(0.7), emu(0.95),
            [(num, 16, True, BLUE, 0)], anchor=MSO_ANCHOR.MIDDLE,
        )
        add_lines(
            slide, emu(1.95), emu(y + 0.12), emu(10.1), emu(0.72),
            [(who, 14, True, WHITE, 0), (body, 13, False, MUTED, 2)],
        )
    notes(
        slide,
        "Назначить сверку через две недели. Повестка сверки — только лист решений, новые темы не добавлять.",
    )
    return slide


def slide_close(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_LAST])
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Решения", 32, True, WHITE)
            p2 = tf.add_paragraph()
            set_run(p2.add_run(), "важнее полноты", 32, True, BLUE)
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(
                p.add_run(),
                "Через две недели сверяем только лист решений.\nНовую повестку на ту сверку не приносим.",
                16, False, WHITE,
            )
        elif ph.placeholder_format.idx == 10:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Блиц-запись  ·  лаборатория, производство, центр продаж", 12, False, SOFT)
    notes(slide, "Закончить без нового вопроса из зала. Всё новое — в дату возврата.")
    return slide


def build():
    if not TEMPLATE.exists():
        raise SystemExit(f"Template not found: {TEMPLATE}")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(TEMPLATE, OUT)
    prs = Presentation(str(OUT))
    delete_all_slides(prs)

    slide_title(prs)
    slide_outcomes(prs)
    slide_agenda(prs)
    slide_facts(prs)
    slide_funnel(prs)
    slide_usp(prs)
    slide_offer(prs)
    slide_capacity(prs)
    slide_roadmap_rule(prs)
    slide_roadmap_year(prs)
    slide_expectations(prs)
    slide_sla(prs)
    slide_battle(prs)
    slide_partners(prs)
    slide_decisions(prs)
    slide_after(prs)
    slide_close(prs)

    prs.core_properties.title = "Планерка «Блиц-запись». Повестка на 2 часа"
    prs.core_properties.subject = "Лаборатория, производство и центр продаж"
    prs.core_properties.category = "Внутренняя планерка"

    prs.save(str(OUT))
    print(f"Saved: {OUT}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    build()
