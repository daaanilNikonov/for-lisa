#!/usr/bin/env python3
"""Жёлто-белая презентация Форус: статус работы с «1С:Кабинет сотрудника»."""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "presentation" / "Статус_Кабинет_сотрудника_август_2026.pptx"
BRAND = ROOT / "presentation" / "assets" / "brand"

YELLOW = RGBColor(0xFE, 0xCF, 0x68)
YELLOW_DEEP = RGBColor(0xE8, 0xB4, 0x3A)
CREAM = RGBColor(0xFF, 0xF8, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x1A, 0x1A, 0x1A)
SLATE = RGBColor(0x39, 0x5F, 0x75)
MUTED = RGBColor(0x5C, 0x5C, 0x5C)
SOFT_BG = RGBColor(0xFF, 0xFC, 0xF6)

FONT = "Verdana"
SW, SH = Inches(13.333), Inches(7.5)
TOTAL = 10


def emu(inches: float) -> int:
    return int(Inches(inches))


def set_run(run, text, size_pt, bold=False, color=INK, font_name=FONT):
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
    color=INK,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    set_anchor(tf, anchor)
    tf.paragraphs[0].alignment = align
    set_run(tf.paragraphs[0].add_run(), text, size_pt, bold, color)
    return box


def add_rect(slide, left, top, width, height, fill, corner=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner is not None else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    if corner is not None:
        try:
            shape.adjustments[0] = corner
        except Exception:
            pass
    return shape


def add_card(slide, left, top, width, height, fill=CREAM, corner=0.08):
    return add_rect(slide, left, top, width, height, fill, corner)


def add_picture(slide, name, left, top, width=None, height=None):
    path = BRAND / name
    if not path.exists():
        return None
    kwargs = {}
    if width is not None:
        kwargs["width"] = width
    if height is not None:
        kwargs["height"] = height
    return slide.shapes.add_picture(str(path), left, top, **kwargs)


def blank_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SW, SH, WHITE)
    return slide


def add_footer(slide, page: int):
    add_rect(slide, 0, emu(7.28), SW, emu(0.22), YELLOW)
    add_textbox(
        slide, emu(0.5), emu(7.08), emu(9.2), emu(0.22),
        "ГК Форус  ·  рабочая группа  ·  1С:Кабинет сотрудника",
        9, False, MUTED,
    )
    add_textbox(
        slide, emu(11.4), emu(7.08), emu(1.4), emu(0.22),
        f"{page} / {TOTAL}",
        9, False, MUTED, PP_ALIGN.RIGHT,
    )


def add_header(slide, kicker: str, title: str):
    add_picture(slide, "logo-forus.png", emu(0.5), emu(0.22), height=emu(0.38))
    add_picture(slide, "wave_tr.png", emu(10.55), emu(-0.05), width=emu(2.9))
    add_textbox(slide, emu(0.5), emu(0.72), emu(12.2), emu(0.28), kicker, 11, True, SLATE)
    add_textbox(slide, emu(0.5), emu(0.98), emu(12.2), emu(0.55), title, 24, True, INK)
    add_rect(slide, emu(0.5), emu(1.55), emu(1.35), emu(0.07), YELLOW)


def check_mark(slide, left, top, size=0.28):
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, emu(size), emu(size))
    circ.fill.solid()
    circ.fill.fore_color.rgb = YELLOW
    circ.line.fill.background()
    add_textbox(
        slide, left, top, emu(size), emu(size),
        "✓", 11, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
    )


def slide_title(prs):
    s = blank_slide(prs)
    add_rect(s, 0, 0, emu(0.22), SH, YELLOW)
    add_picture(s, "logo-forus.png", emu(0.7), emu(0.45), height=emu(0.55))
    add_picture(s, "wave_tr.png", emu(9.8), emu(-0.1), width=emu(3.7))
    add_textbox(
        s, emu(0.7), emu(1.7), emu(11.5), emu(0.35),
        "Рабочая группа  ·  август 2026", 14, True, SLATE,
    )
    add_textbox(
        s, emu(0.7), emu(2.15), emu(12.0), emu(1.5),
        "Как мы продаём\n«1С:Кабинет сотрудника»", 36, True, INK,
    )
    add_rect(s, emu(0.7), emu(3.85), emu(2.0), emu(0.1), YELLOW)
    add_textbox(
        s, emu(0.7), emu(4.2), emu(11.0), emu(1.15),
        "Что уже сделали по плану, какие идеи проверяем звонками\n"
        "и что из этого получается.",
        16, False, MUTED,
    )
    add_card(s, emu(0.7), emu(5.7), emu(3.5), emu(1.05), CREAM, 0.1)
    add_textbox(s, emu(0.9), emu(5.82), emu(3.2), emu(0.35), "Период звонков", 11, False, SLATE)
    add_textbox(s, emu(0.9), emu(6.12), emu(3.2), emu(0.45), "23.07 — 24.08.2026", 16, True, INK)
    add_card(s, emu(4.4), emu(5.7), emu(3.7), emu(1.05), CREAM, 0.1)
    add_textbox(s, emu(4.6), emu(5.82), emu(3.4), emu(0.35), "Кто звонит", 11, False, SLATE)
    add_textbox(s, emu(4.6), emu(6.12), emu(3.4), emu(0.45), "Юлиана · Соня · Данил", 16, True, INK)
    add_card(s, emu(8.3), emu(5.7), emu(4.3), emu(1.05), YELLOW, 0.1)
    add_textbox(s, emu(8.5), emu(5.82), emu(4.0), emu(0.35), "Звонков за период", 11, False, INK)
    add_textbox(s, emu(8.5), emu(6.12), emu(4.0), emu(0.45), "1 084", 16, True, INK)
    add_picture(s, "wave_bl.png", emu(-0.15), emu(6.55), width=emu(3.2))
    return s


def slide_agenda(prs):
    s = blank_slide(prs)
    add_header(s, "Содержание", "О чём этот отчёт")
    items = [
        ("01", "Проект", "Зачем мы запускаем Кабинет сотрудника и как устроена работа"),
        ("02", "Как работаем", "Четыре направления: свои клиенты, новые компании, семинары, два сервиса вместе"),
        ("03", "План проекта", "Все задачи из плана выполнены — таблица с отметками"),
        ("04", "Идеи, которые проверяем", "Четыре гипотезы простыми словами"),
        ("05", "Что показали звонки", "Сколько набрали, кто ответил, куда дошли"),
        ("06", "Почему не покупают", "Главные причины отказа после показа сервиса"),
    ]
    top = emu(1.85)
    for i, (num, title, body) in enumerate(items):
        y = top + emu(i * 0.8)
        add_card(s, emu(0.5), y, emu(12.3), emu(0.72), CREAM if i % 2 == 0 else SOFT_BG, 0.08)
        add_textbox(
            s, emu(0.7), y + emu(0.12), emu(0.8), emu(0.5), num, 20, True, YELLOW_DEEP,
            PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE,
        )
        add_textbox(
            s, emu(1.6), y + emu(0.08), emu(3.6), emu(0.55), title, 16, True, INK,
            PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE,
        )
        add_textbox(
            s, emu(5.3), y + emu(0.08), emu(7.2), emu(0.55), body, 13, False, MUTED,
            PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE,
        )
    add_footer(s, 2)
    return s


def slide_project(prs):
    s = blank_slide(prs)
    add_header(s, "1. О проекте", "Сервис, который убирает бумажную рутину у кадров и бухгалтерии")
    add_card(s, emu(0.5), emu(1.85), emu(7.6), emu(4.85), CREAM, 0.08)
    add_textbox(s, emu(0.75), emu(2.05), emu(7.1), emu(0.4), "Что это такое", 13, True, SLATE)
    add_textbox(
        s, emu(0.75), emu(2.45), emu(7.1), emu(1.7),
        "«1С:Кабинет сотрудника» — личный кабинет для сотрудников внутри привычной 1С. "
        "Через телефон или компьютер человек сам подаёт заявление на отпуск, "
        "смотрит расчётный листок, подписывает кадровые документы. "
        "Бухгалтеру не нужно собирать бумаги вручную. Компании не нужно менять свою программу учёта.",
        14, False, INK,
    )
    bullets = [
        ("Для клиента", "Меньше бумаги, быстрее заявления, документы с юридической силой, всё под рукой у сотрудника."),
        ("Для Форус", "Можно предложить сервис тем, кто уже с нами, и зайти в новые компании с понятной пользой."),
        ("Почему сейчас", "Кадровые документы уходят в электронный вид, но многие ещё боятся начинать. Мы показываем сервис вживую и помогаем подключить."),
    ]
    y = emu(4.2)
    for title, body in bullets:
        add_rect(s, emu(0.75), y + emu(0.08), emu(0.14), emu(0.14), YELLOW)
        add_textbox(s, emu(1.05), y, emu(6.8), emu(0.28), title, 13, True, INK)
        add_textbox(s, emu(1.05), y + emu(0.28), emu(6.8), emu(0.55), body, 12, False, MUTED)
        y += emu(0.78)

    add_card(s, emu(8.3), emu(1.85), emu(4.5), emu(4.85), CREAM, 0.08)
    add_rect(s, emu(8.3), emu(1.85), emu(4.5), emu(0.12), YELLOW)
    add_textbox(s, emu(8.55), emu(2.15), emu(4.05), emu(0.4), "Кто этим занимается", 13, True, SLATE)
    add_textbox(
        s, emu(8.55), emu(2.6), emu(4.05), emu(1.5),
        "Не просто отдел продаж. "
        "Три менеджера каждый день звонят, пробуют разные подходы "
        "и смотрят, что реально приводит человека на показ сервиса и к покупке.",
        13, False, INK,
    )
    add_textbox(s, emu(8.55), emu(4.25), emu(4.05), emu(0.35), "Юлиана Юнусова", 14, True, INK)
    add_textbox(s, emu(8.55), emu(4.6), emu(4.05), emu(0.35), "Соня Оглоблина", 14, True, INK)
    add_textbox(s, emu(8.55), emu(4.95), emu(4.05), emu(0.35), "Данил Кургузов", 14, True, INK)
    add_textbox(
        s, emu(8.55), emu(5.5), emu(4.05), emu(0.85),
        "За месяц на линии — 1 084 звонка.",
        13, False, MUTED,
    )
    add_footer(s, 3)
    return s


def slide_map(prs):
    s = blank_slide(prs)
    add_header(s, "2. Как устроена работа", "Четыре направления, по которым идём к клиенту")
    add_textbox(
        s, emu(0.5), emu(1.72), emu(12.3), emu(0.45),
        "Путь один и тот же: позвонили → человек ответил по делу → появился интерес → "
        "записались на показ → показали сервис → продали. Меняется только, кому звоним и что говорим.",
        13, False, MUTED,
    )
    contours = [
        ("1", "Свои клиенты", "Кто уже обслуживается у Форус",
         "Говорим: «Кабинет сотрудника уже входит в ваше сопровождение 1С». Показ без разговора про цену на входе."),
        ("2", "Новые компании", "Кого раньше не знали",
         "Звоним не всем подряд, а компаниям, где сервис правда нужен: торговля, производство, вахта, общепит и похожие."),
        ("3", "После семинаров", "Кто уже нас слышал",
         "Человек был на встрече или семинаре. Ещё раз объясняем пользу, отвечаем на вопросы и зовём на показ."),
        ("4", "Два сервиса вместе", "Документы и кадры пакетом",
         "Предлагаем не одну подписку, а связку: электронные документы для логистики и Кабинет сотрудника для кадров."),
    ]
    for i, (lit, title, bases, body) in enumerate(contours):
        left = emu(0.5) + emu(i * 3.2)
        add_card(s, left, emu(2.3), emu(3.05), emu(3.55), CREAM, 0.08)
        add_rect(s, left, emu(2.3), emu(3.05), emu(0.12), YELLOW)
        add_textbox(s, left + emu(0.18), emu(2.55), emu(2.7), emu(0.35), lit, 18, True, YELLOW_DEEP)
        add_textbox(s, left + emu(0.18), emu(2.95), emu(2.7), emu(0.7), title, 16, True, INK)
        add_textbox(s, left + emu(0.18), emu(3.6), emu(2.7), emu(0.55), bases, 12, True, SLATE)
        add_textbox(s, left + emu(0.18), emu(4.2), emu(2.7), emu(1.4), body, 12, False, MUTED)
    add_card(s, emu(0.5), emu(6.05), emu(12.3), emu(0.85), SOFT_BG, 0.08)
    add_textbox(
        s, emu(0.7), emu(6.18), emu(12.0), emu(0.6),
        "Каждый день у менеджера есть список дел и доска задач. "
        "После записи на показ действует порядок: письмо, напоминание, сам показ, звонок через три дня.",
        13, False, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE,
    )
    add_footer(s, 4)
    return s


def slide_plan(prs):
    s = blank_slide(prs)
    add_header(s, "3. План проекта", "Все задачи из плана выполнены")
    add_textbox(
        s, emu(0.5), emu(1.68), emu(12.3), emu(0.38),
        "Документ «План проекта Кабинет сотрудника»: исследование, материалы, обучение, запуск звонков и разбор результатов.",
        13, False, MUTED,
    )
    left_tasks = [
        "Изучили продукт и внутренние материалы",
        "Сравнили предложение с другими партнёрами 1С",
        "Посмотрели внешние кадровые сервисы",
        "Оценили своих клиентов и рынок в стране",
        "Собрали стратегию продвижения",
        "Поставили план продаж и финансовый план",
        "Написали сценарии разговора и шаблоны писем",
        "Собрали коммерческие предложения и листовки",
    ]
    right_tasks = [
        "Обучили менеджеров первому сценарию",
        "Подготовили страницу на сайте",
        "Начали звонить новым компаниям",
        "Начали звонить своим клиентам",
        "Настроили работу с входящими заявками",
        "Разобрали первую волну и зафиксировали выводы",
        "Подготовили второй сценарий, обучили, запустили звонки",
        "Собрали отчёт, пакет для других команд и защитили проект",
    ]
    headers = ["Готово", "Что сделали", "Готово", "Что сделали"]
    col_w = [1.15, 5.0, 1.15, 5.0]
    table_w = sum(col_w)
    start_x = emu(0.5)
    row_h = 0.52
    y0 = 2.12
    add_rect(s, start_x, emu(y0), emu(table_w), emu(0.42), INK)
    x = start_x
    for w, h in zip(col_w, headers):
        add_textbox(
            s, x, emu(y0), emu(w), emu(0.42), h, 12, True, WHITE,
            PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
        )
        x += emu(w)
    for r, (left, right) in enumerate(zip(left_tasks, right_tasks)):
        y = emu(y0 + 0.42 + r * row_h)
        bg = CREAM if r % 2 == 0 else WHITE
        add_rect(s, start_x, y, emu(table_w), emu(row_h), bg)
        check_mark(s, start_x + emu(0.42), y + emu(0.12), 0.28)
        add_textbox(
            s, start_x + emu(1.15), y, emu(5.0), emu(row_h), left, 12, False, INK,
            PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE,
        )
        check_mark(s, start_x + emu(6.57), y + emu(0.12), 0.28)
        add_textbox(
            s, start_x + emu(7.3), y, emu(5.0), emu(row_h), right, 12, False, INK,
            PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE,
        )
    add_footer(s, 5)
    return s


def slide_hypotheses(prs):
    s = blank_slide(prs)
    add_header(s, "4. Что мы проверяем", "Четыре идеи простыми словами")
    add_textbox(
        s, emu(0.5), emu(1.68), emu(12.3), emu(0.4),
        "Каждая идея читается так: если сделаем одно — человек чаще согласится на другое.",
        13, False, MUTED,
    )
    hyps = [
        (
            "1",
            "Свои клиенты",
            "Если действующему клиенту предложить показ Кабинета сотрудника словами "
            "«это уже входит в ваше сопровождение 1С», "
            "то на показ запишутся чаще, чем если просто звонить и предлагать «посмотреть сервис».",
        ),
        (
            "2",
            "Новые компании",
            "Если звонить не всем подряд, а только фирмам, чей бизнес подходит "
            "(торговля, производство, вахта, общепит и похожие), "
            "то из разговора чаще получается живой интерес и договор.",
        ),
        (
            "3",
            "Два сервиса сразу",
            "Если предлагать не один сервис, а сразу два — электронные документы для логистики "
            "и Кабинет сотрудника для кадров — то до договора доходим чаще, чем с одним сервисом.",
        ),
        (
            "4",
            "После семинара",
            "Если человеку после семинара ещё раз объяснить пользу и спокойно ответить на вопросы, "
            "то он согласится чаще, чем при самом первом звонке незнакомой компании.",
        ),
    ]
    for i, (num, title, body) in enumerate(hyps):
        y = emu(2.15) + emu(i * 1.15)
        add_card(s, emu(0.5), y, emu(12.3), emu(1.05), CREAM if i % 2 == 0 else SOFT_BG, 0.08)
        add_rect(s, emu(0.5), y, emu(0.7), emu(1.05), YELLOW)
        add_textbox(
            s, emu(0.5), y, emu(0.7), emu(1.05), num, 22, True, INK,
            PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
        )
        add_textbox(s, emu(1.4), y + emu(0.1), emu(11.1), emu(0.28), title, 14, True, SLATE)
        add_textbox(s, emu(1.4), y + emu(0.4), emu(11.1), emu(0.55), body, 13, False, INK)
    add_footer(s, 6)
    return s


def slide_funnel_totals(prs):
    s = blank_slide(prs)
    add_header(s, "Показатели по базам", "Воронка за 23.07–24.08.2026 · три менеджера")
    kpis = [
        ("1 084", "прозвонено", "100%"),
        ("483", "успешных", "44,6%"),
        ("70", "в сделку", "6,5%"),
        ("12", "записей", "1,1%"),
        ("11", "демо провели", "1,0%"),
        ("0", "продаж", "пока 0"),
    ]
    for i, (n, label, pct) in enumerate(kpis):
        left = emu(0.5) + emu(i * 2.12)
        fill = YELLOW if i == 0 else CREAM
        add_card(s, left, emu(1.85), emu(2.02), emu(1.85), fill, 0.1)
        add_textbox(
            s, left + emu(0.1), emu(2.0), emu(1.82), emu(0.7), n, 26, True, INK, PP_ALIGN.CENTER,
        )
        add_textbox(
            s, left + emu(0.1), emu(2.7), emu(1.82), emu(0.4), label, 12, True, SLATE, PP_ALIGN.CENTER,
        )
        add_textbox(
            s, left + emu(0.1), emu(3.1), emu(1.82), emu(0.35), pct, 12, False, MUTED, PP_ALIGN.CENTER,
        )

    add_textbox(s, emu(0.5), emu(3.9), emu(12.3), emu(0.35), "Кто набрал цифру", 14, True, INK)
    managers = [
        ("Юлиана Юнусова", "114", "60", "24", "9", "7"),
        ("Соня Оглоблина", "555", "238", "29", "2", "3"),
        ("Данил Кургузов", "415", "185", "17", "1", "1"),
    ]
    headers = ["Менеджер", "Прозвонено", "Успешных", "В сделку", "Записи", "Демо"]
    y = emu(4.35)
    add_rect(s, emu(0.5), y, emu(12.3), emu(0.42), INK)
    widths = [3.3, 1.8, 1.8, 1.8, 1.8, 1.8]
    x = emu(0.5)
    for w, h in zip(widths, headers):
        add_textbox(s, x, y, emu(w), emu(0.42), h, 12, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        x += emu(w)
    for i, row in enumerate(managers):
        y = emu(4.77) + emu(i * 0.55)
        bg = CREAM if i % 2 == 0 else SOFT_BG
        add_rect(s, emu(0.5), y, emu(12.3), emu(0.55), bg)
        x = emu(0.5)
        for j, (w, val) in enumerate(zip(widths, row)):
            align = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            pad = emu(0.2) if j == 0 else 0
            add_textbox(
                s, x + pad, y, emu(w) - pad, emu(0.55), val, 13, j == 0, INK,
                align, MSO_ANCHOR.MIDDLE,
            )
            x += emu(w)
    add_footer(s, 7)
    return s


def slide_funnel_bases(prs):
    s = blank_slide(prs)
    add_header(s, "Показатели по базам", "Откуда пришли звонки — и что из этого вышло")
    rows = [
        ["Источник", "Прозвон", "Успешн.", "Сделка", "Запись", "Демо", "Сделка %"],
        ["База ЗУП", "91", "46", "10", "0", "3", "11,0%"],
        ["База КП", "124", "60", "17", "10", "5", "13,7%"],
        ["База канбана", "12", "12", "0", "0", "0", "0%"],
        ["База вебинаров", "298", "132", "13", "1", "1", "4,4%"],
        ["Заявка с сайта", "4", "4", "4", "0", "0", "100%"],
        ["База вахта", "125", "34", "1", "0", "0", "0,8%"],
        ["База общепит", "266", "124", "13", "0", "0", "4,9%"],
        ["Холодка общепит", "31", "12", "3", "0", "0", "9,7%"],
        ["База производства", "123", "54", "8", "1", "2", "6,5%"],
        ["Новая КС / прочие", "10", "5", "1", "0", "0", "10%"],
        ["Итого", "1 084", "483", "70", "12", "11", "6,5%"],
    ]
    col_w = [2.7, 1.5, 1.5, 1.5, 1.5, 1.5, 1.6]
    table_w = sum(col_w)
    start_x = emu(0.7)
    row_h = 0.38
    y0 = 1.82
    for r, row in enumerate(rows):
        y = emu(y0 + r * row_h)
        if r == 0:
            bg, color, bold = INK, WHITE, True
        elif r == len(rows) - 1:
            bg, color, bold = YELLOW, INK, True
        else:
            bg = CREAM if r % 2 else WHITE
            color, bold = INK, False
        add_rect(s, start_x, y, emu(table_w), emu(row_h), bg)
        x = start_x
        for c, (w, val) in enumerate(zip(col_w, row)):
            align = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            pad = emu(0.12) if c == 0 else 0
            add_textbox(
                s, x + pad, y, emu(w) - pad, emu(row_h), val, 11, bold, color,
                align, MSO_ANCHOR.MIDDLE,
            )
            x += emu(w)
    add_textbox(
        s, emu(0.5), emu(6.55), emu(12.3), emu(0.45),
        "Узкое место воронки — не контакт, а запись и проведение демо. "
        "КП и ЗУП дают лучшую сделку; общепит и вебинары дают объём прозвона.",
        12, False, MUTED,
    )
    add_footer(s, 8)
    return s


def slide_refusals(prs):
    s = blank_slide(prs)
    add_header(s, "6. После показа сервиса", "Почему клиент чаще всего говорит «нет»")
    add_textbox(
        s, emu(0.5), emu(1.68), emu(12.3), emu(0.4),
        "За период провели 11 показов. Ниже — то, что слышим после встречи чаще всего.",
        13, False, MUTED,
    )
    reasons = [
        ("1", "Пока не обязаны — не будем",
         "«Когда государство заставит, тогда и подключим. Сейчас бумажный кадровый учёт нас устраивает»."),
        ("2", "Нет денег",
         "В бюджете этого года на сервис не заложено. «Дорого», «не сейчас», «в следующем году посмотрим»."),
        ("3", "Мало людей в офисе",
         "Два–пять человек сидят рядом с бухгалтером. Проще подойти лично, чем ставить кабинет."),
        ("4", "Телефоны не те",
         "На вахте и у возрастных сотрудников кнопочные телефоны. Приложением они не воспользуются."),
        ("5", "Не будут пользоваться",
         "Люди всё равно напишут бухгалтеру в чат. Приложение не приживётся — пользы не увидят."),
        ("6", "Нужно согласовать",
         "Понравилось, но решение за директором, юристом или тем, кто отвечает за безопасность данных."),
    ]
    for i, (num, title, body) in enumerate(reasons):
        col, row = i % 3, i // 3
        left = emu(0.5) + emu(col * 4.2)
        top = emu(2.2) + emu(row * 2.3)
        add_card(s, left, top, emu(4.0), emu(2.15), CREAM, 0.08)
        add_rect(s, left, top, emu(0.55), emu(2.15), YELLOW)
        add_textbox(
            s, left, top, emu(0.55), emu(2.15), num, 20, True, INK,
            PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE,
        )
        add_textbox(s, left + emu(0.7), top + emu(0.18), emu(3.1), emu(0.55), title, 14, True, INK)
        add_textbox(s, left + emu(0.7), top + emu(0.75), emu(3.1), emu(1.2), body, 12, False, MUTED)
    add_footer(s, 9)
    return s


def slide_next(prs):
    s = blank_slide(prs)
    add_header(s, "Дальше", "На чём держим внимание")
    steps = [
        ("1", "Больше записей на показ",
         "Сделок уже 70, а на показ записались 12 человек. После интереса сразу предлагаем конкретную дату встречи."),
        ("2", "Два сервиса — отдельная пометка",
         "В таблице звонков отдельно отмечать, когда предлагали логистику и Кабинет сотрудника вместе. Иначе не понять, сработала ли эта идея."),
        ("3", "Работать с отказами после показа",
         "Шесть причин повторяются. К каждой — короткий ответ менеджера, чтобы встреча не заканчивалась на «подумаем»."),
        ("4", "Сравнить тёплый и первый звонок",
         "Отдельно посмотреть: после семинара люди соглашаются чаще, чем при самом первом звонке незнакомой компании?"),
    ]
    for i, (n, title, body) in enumerate(steps):
        y = emu(1.85) + emu(i * 1.15)
        add_card(s, emu(0.5), y, emu(12.3), emu(1.05), CREAM, 0.08)
        add_rect(s, emu(0.5), y, emu(1.05), emu(1.05), YELLOW)
        add_textbox(s, emu(0.5), y, emu(1.05), emu(1.05), n, 24, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        add_textbox(s, emu(1.8), y + emu(0.15), emu(10.6), emu(0.35), title, 16, True, INK)
        add_textbox(s, emu(1.8), y + emu(0.52), emu(10.6), emu(0.42), body, 13, False, MUTED)
    add_footer(s, 10)
    return s


def main():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    slide_title(prs)
    slide_agenda(prs)
    slide_project(prs)
    slide_map(prs)
    slide_plan(prs)
    slide_hypotheses(prs)
    slide_funnel_totals(prs)
    slide_funnel_bases(prs)
    slide_refusals(prs)
    slide_next(prs)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print("saved", OUT, "slides", len(prs.slides), "bytes", OUT.stat().st_size)


if __name__ == "__main__":
    main()
