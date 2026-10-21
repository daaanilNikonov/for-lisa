#!/usr/bin/env python3
"""Схема продвижения «Блиц-запись» в Instagram — PDF для обсуждения с коллегами.

Возможности сервиса взяты из коммерческого предложения «Блиц-запись».
"""

from pathlib import Path

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, Color, white
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

ROOT = Path(__file__).resolve().parents[1]
OUT_PDF = ROOT / "presentation" / "Схема_продвижения_Блиц-запись_Instagram.pdf"
OUT_DIR = ROOT / "presentation"

PAGE_W = 1600
PAGE_H = 1000

INK = HexColor("#172026")
MUTED = HexColor("#5C6D7A")
SOFT = HexColor("#7E8E9B")
BLUE = HexColor("#1788C2")
BLUE_DARK = HexColor("#0E6E9E")
BLUE_DEEP = HexColor("#0C5A82")
BLUE_SOFT = HexColor("#E7F5FC")
BLUE_MID = HexColor("#3AA0D4")
LINE = HexColor("#D5DEE6")
CARD_LINE = HexColor("#E1E8EE")
BG = HexColor("#F3F6F8")
WHITE = HexColor("#FFFFFF")
AMBER = HexColor("#9A5B00")
AMBER_DARK = HexColor("#8A5600")
AMBER_BG = HexColor("#FFF6E8")
AMBER_LINE = HexColor("#F0D3A4")
GREEN = HexColor("#1B7A4A")
GREEN_BG = HexColor("#EEF8F2")
GREEN_LINE = HexColor("#C5E4D2")
SHADOW = HexColor("#E4EBF1")

FONT = "/usr/share/fonts/truetype/macos/Inter-Regular.ttf"
FONT_MED = "/usr/share/fonts/truetype/macos/Inter-Medium.ttf"
FONT_SEMI = "/usr/share/fonts/truetype/macos/Inter-SemiBold.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/macos/Inter-Bold.ttf"

pdfmetrics.registerFont(TTFont("Inter", FONT))
pdfmetrics.registerFont(TTFont("Inter-Med", FONT_MED))
pdfmetrics.registerFont(TTFont("Inter-Semi", FONT_SEMI))
pdfmetrics.registerFont(TTFont("Inter-Bold", FONT_BOLD))


def wrap(text, font, size, max_w):
    words = text.split()
    lines, cur = [], ""
    for word in words:
        trial = word if not cur else f"{cur} {word}"
        if stringWidth(trial, font, size) <= max_w:
            cur = trial
            continue
        if cur:
            lines.append(cur)
            cur = ""
        if stringWidth(word, font, size) <= max_w:
            cur = word
            continue
        buf = ""
        for ch in word:
            if stringWidth(buf + ch, font, size) <= max_w:
                buf += ch
            else:
                if buf:
                    lines.append(buf)
                buf = ch
        cur = buf
    if cur:
        lines.append(cur)
    return lines


def y_of(top, h=0):
    return PAGE_H - top - h


def draw_text(c, text, x, baseline, font, size, color, align="left"):
    c.setFillColor(color)
    c.setFont(font, size)
    if align == "center":
        c.drawCentredString(x, baseline, text)
    elif align == "right":
        c.drawRightString(x, baseline, text)
    else:
        c.drawString(x, baseline, text)


def paragraph(c, text, x, top, font, size, color, max_w, leading):
    lines = wrap(text, font, size, max_w)
    baseline = y_of(top, size)
    for line in lines:
        draw_text(c, line, x, baseline, font, size, color)
        baseline -= leading
    return len(lines) * leading


def round_card(c, x, top, w, h, fill, stroke=None, radius=14, sw=1.0, shadow=False):
    y = y_of(top, h)
    if shadow:
        c.setFillColor(SHADOW)
        c.roundRect(x, y - 3, w, h, radius, fill=1, stroke=0)
    c.setFillColor(fill)
    if stroke is None:
        c.setStrokeColor(fill)
        c.roundRect(x, y, w, h, radius, fill=1, stroke=0)
    else:
        c.setStrokeColor(stroke)
        c.setLineWidth(sw)
        c.roundRect(x, y, w, h, radius, fill=1, stroke=1)
    return y


def chevron(c, cx, cy, color):
    p = c.beginPath()
    p.moveTo(cx - 4, cy + 7)
    p.lineTo(cx + 5, cy)
    p.lineTo(cx - 4, cy - 7)
    p.close()
    c.setFillColor(color)
    c.drawPath(p, fill=1, stroke=0)


def footer(c, page, pages=2):
    c.setStrokeColor(LINE)
    c.setLineWidth(1)
    c.line(40, 32, PAGE_W - 40, 32)
    draw_text(
        c,
        "ГК Форус  ·  внутренняя схема, не материал для клиента",
        40,
        16,
        "Inter",
        10,
        SOFT,
    )
    draw_text(
        c,
        f"{page}  /  {pages}",
        PAGE_W - 40,
        16,
        "Inter-Med",
        10,
        SOFT,
        align="right",
    )


def header_block(c, eyebrow, title, subtitle):
    c.setFillColor(BLUE)
    c.rect(0, PAGE_H - 8, PAGE_W, 8, fill=1, stroke=0)
    top = 28
    draw_text(c, eyebrow, 40, y_of(top, 12), "Inter-Semi", 11, BLUE)
    top += 22
    draw_text(c, title, 40, y_of(top, 28), "Inter-Bold", 28, INK)
    top += 36
    draw_text(c, subtitle, 40, y_of(top, 15), "Inter", 13, MUTED)
    top += 24
    return top


def draw_page_1(c):
    c.setFillColor(BG)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    top = header_block(
        c,
        "ГК ФОРУС   ·   ДЛЯ ОБСУЖДЕНИЯ С КОЛЛЕГАМИ   ·   INSTAGRAM",
        "Схема продвижения «Блиц-запись»",
        "Персонаж-эксперт говорит о нехватке клиентов. Сервис вписываем в конец ролика — как разработку команды.",
    )

    # Idea
    idea_h = 132
    round_card(c, 40, top, PAGE_W - 80, idea_h, WHITE, LINE, 16, shadow=True)
    c.setFillColor(BLUE)
    c.rect(40, y_of(top, idea_h) + 16, 6, idea_h - 32, fill=1, stroke=0)

    draw_text(c, "МЫСЛЬ, КОТОРУЮ НЕСЁМ КОЛЛЕГАМ", 64, y_of(top + 18, 12), "Inter-Semi", 11, BLUE)
    paragraph(
        c,
        "Не рекламируем сервис в лоб. В кадре эксперт по работе с клиентами: он рассказывает, какие фишки есть у бизнеса услуг, и разбирает боль — «клиентов не хватает, как продвигаться?». «Блиц-запись» появляется как ответ на эту боль: сервис, который создала команда разработчиков.",
        64,
        top + 40,
        "Inter",
        13.5,
        INK,
        1040,
        18,
    )

    facts = [
        ("01", "Персонаж", "эксперт по клиентам, не продавец"),
        ("02", "Вход в ролик", "бизнесу не хватает клиентов"),
        ("03", "Где продукт", "последняя треть, одна фраза"),
    ]
    fx = 1148
    fy = top + 16
    for num, label, desc in facts:
        round_card(c, fx, fy, 412, 32, BLUE_SOFT, None, 8)
        draw_text(c, num, fx + 12, y_of(fy + 8, 13), "Inter-Bold", 12, BLUE)
        draw_text(c, label, fx + 42, y_of(fy + 8, 13), "Inter-Semi", 12, INK)
        draw_text(c, desc, fx + 168, y_of(fy + 9, 12), "Inter", 11, MUTED)
        fy += 36

    top += idea_h + 22

    draw_text(c, "ХОД ОДНОГО РОЛИКА", 40, y_of(top, 12), "Inter-Semi", 11, BLUE)
    draw_text(
        c,
        "Зритель приходит за советом. Название сервиса слышит один раз — когда уже понятно, зачем оно ему.",
        210,
        y_of(top, 12),
        "Inter",
        11.5,
        MUTED,
    )
    top += 22

    steps = [
        {
            "num": "01",
            "title": "Крючок",
            "head": BLUE,
            "body_bg": WHITE,
            "border": CARD_LINE,
            "text": "Первые секунды — боль, которую владелец узнаёт сразу. Названия сервиса ещё нет.",
            "quote": "Бизнес страдает от недостатка клиентов. Как можно продвигаться?",
        },
        {
            "num": "02",
            "title": "Переворот",
            "head": BLUE_MID,
            "body_bg": WHITE,
            "border": CARD_LINE,
            "text": "Дело часто не в том, что клиентов нет. Их теряют по дороге: не дозвонились, написали ночью, застряли в переписке.",
            "quote": "Реклама приводит человека. Запись его не удерживает.",
        },
        {
            "num": "03",
            "title": "Одна фишка",
            "head": BLUE_DARK,
            "body_bg": WHITE,
            "border": CARD_LINE,
            "text": "Эксперт даёт один приём для сферы услуг. В ролике один приём, без обзора всех функций.",
            "quote": "Вот приём, который возвращает тех, кто не дозвонился.",
        },
        {
            "num": "04",
            "title": "Вписываем сервис",
            "head": AMBER,
            "body_bg": AMBER_BG,
            "border": AMBER_LINE,
            "text": "Здесь «Блиц-запись» звучит впервые. Это инструмент, который создала команда разработчиков.",
            "quote": "Команда разработчиков создала сервис Блиц-запись.",
        },
        {
            "num": "05",
            "title": "Один призыв",
            "head": GREEN,
            "body_bg": GREEN_BG,
            "border": GREEN_LINE,
            "text": "Одно действие после ролика. Цена, тариф и слово «акция» остаются за кадром.",
            "quote": "Ссылка на запись — в шапке профиля.",
        },
    ]

    gap = 16
    card_w = (PAGE_W - 80 - gap * 4) / 5
    card_h = 252
    header_h = 40
    for i, step in enumerate(steps):
        x = 40 + i * (card_w + gap)
        round_card(c, x, top, card_w, card_h, step["body_bg"], step["border"], 14, shadow=True)
        # header bar
        c.setFillColor(step["head"])
        y_head = y_of(top, header_h)
        c.saveState()
        p = c.beginPath()
        p.roundRect(x, y_of(top, card_h), card_w, card_h, 14)
        c.clipPath(p, stroke=0, fill=0)
        c.rect(x, y_head, card_w, header_h, fill=1, stroke=0)
        c.restoreState()
        draw_text(
            c,
            f"{step['num']}   {step['title']}",
            x + 14,
            y_head + 13,
            "Inter-Semi",
            12,
            WHITE,
        )
        paragraph(
            c,
            step["text"],
            x + 14,
            top + header_h + 16,
            "Inter",
            12,
            INK,
            card_w - 28,
            16.5,
        )
        quote_h = 96
        quote_top = top + card_h - quote_h - 14
        qfill = WHITE if step["body_bg"] != WHITE else BG
        round_card(c, x + 12, quote_top, card_w - 24, quote_h, qfill, None, 10)
        draw_text(
            c,
            "ФРАЗА В КАДРЕ",
            x + 24,
            y_of(quote_top + 10, 10),
            "Inter-Semi",
            9,
            step["head"],
        )
        paragraph(
            c,
            step["quote"],
            x + 24,
            quote_top + 28,
            "Inter-Semi",
            12,
            INK,
            card_w - 48,
            16,
        )
        if i < 4:
            chevron(c, x + card_w + gap / 2, y_of(top, header_h / 2), step["head"])

    top += card_h + 22

    draw_text(c, "ИЗ ЧЕГО СОБИРАЕТСЯ СЕРИЯ", 40, y_of(top, 12), "Inter-Semi", 11, BLUE)
    draw_text(
        c,
        "Рабочее название серии: «Куда деваются клиенты». На встрече его можно сменить — логика ролика остаётся.",
        280,
        y_of(top, 12),
        "Inter",
        11.5,
        MUTED,
    )
    top += 22

    col_gap = 14
    col_w = (PAGE_W - 80 - col_gap * 2) / 3
    col_h = 292
    columns = [
        (
            "Персонаж",
            [
                ("Кто говорит", "Эксперт по работе с клиентами в бизнесе услуг. Говорит как практик, который видит, где салон, клиника или студия теряет запись."),
                ("Как звучит", "Коротко и на языке владельца: окна, администратор, ночные сообщения, клиент, который не дозвонился."),
                ("Граница роли", "Без чтения коммерческого предложения и без перечисления модулей. Практик, который разбирает чужую запись."),
                ("Кто в кадре", "Может быть сотрудник Форус. Роль важнее имени: в ролике он эксперт, не отдел продаж."),
            ],
        ),
        (
            "О чём серия",
            [
                ("Мало клиентов?", "Разбор, где заявка теряется ещё до рекламного бюджета."),
                ("Фишка для записи", "Один приём на 30–40 секунд: запись без администратора, кнопка в профиле, напоминание, общий календарь."),
                ("Ниша", "Салон, клиника, мойка, фитнес, учебный центр, частный специалист — отдельный ролик под узнаваемую ситуацию."),
                ("Вопрос из директа", "Короткий ответ на чужой вопрос и снова одна ссылка."),
            ],
        ),
        (
            "Как вписываем продукт",
            [
                ("Формула", "«Команда разработчиков создала сервис Блиц-запись». Так звучит авторство, а не витрина."),
                ("Когда", "Только последняя треть. На обложке и в крючке — боль, не название."),
                ("Сколько функций", "Одна на ролик. Календарь, виджет, база, уведомления и запись 24/7 не смешиваем."),
                ("Что остаётся после ролика", "Цена, срок, сравнение с другими сервисами и имя Форус — в шапке профиля и в личных сообщениях."),
            ],
        ),
    ]

    for i, (title, items) in enumerate(columns):
        x = 40 + i * (col_w + col_gap)
        round_card(c, x, top, col_w, col_h, WHITE, CARD_LINE, 14, shadow=True)
        draw_text(c, title, x + 18, y_of(top + 16, 16), "Inter-Bold", 16, INK)
        c.setStrokeColor(BLUE)
        c.setLineWidth(2)
        c.line(x + 18, y_of(top + 42), x + 58, y_of(top + 42))
        item_top = top + 54
        for label, text in items:
            draw_text(c, label, x + 18, y_of(item_top, 12), "Inter-Semi", 11.5, BLUE_DARK)
            used = paragraph(
                c,
                text,
                x + 18,
                item_top + 16,
                "Inter",
                11,
                INK,
                col_w - 36,
                14.5,
            )
            item_top += 16 + used + 8

    top += col_h + 16
    # chips
    draw_text(c, "КОМУ АДРЕСОВАНЫ РОЛИКИ", 40, y_of(top, 11), "Inter-Semi", 10, SOFT)
    labels = [
        "Салоны и барбершопы",
        "Клиники и стоматологии",
        "Автомойки",
        "Фитнес и студии",
        "Учебные центры",
        "Частные специалисты",
        "Любой бизнес по записи",
    ]
    cx = 230
    chip_h = 26
    for label in labels:
        tw = stringWidth(label, "Inter-Med", 11)
        w = tw + 22
        round_card(c, cx, top - 6, w, chip_h, WHITE, LINE, 13)
        draw_text(c, label, cx + 11, y_of(top - 6 + 6, 12), "Inter-Med", 11, INK)
        cx += w + 8

    draw_text(
        c,
        "На второй странице — три ролика, которые можно прочитать на встрече вслух, и вопросы, которые стоит закрыть до съёмки.",
        40,
        y_of(top + 32, 12),
        "Inter",
        12,
        MUTED,
    )

    footer(c, 1)
    print("page1 content ends at", top + 48)


def draw_page_2(c):
    c.setFillColor(BG)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    top = header_block(
        c,
        "ГК ФОРУС   ·   ПИЛОТ ДЛЯ СОГЛАСОВАНИЯ   ·   ТРИ РОЛИКА",
        "Как звучит ролик и куда ложится фишка",
        "Один ролик — одна ниша, одна потерянная запись, одна фраза про сервис. Ниже можно прочитать вслух на встрече.",
    )

    draw_text(c, "КАРКАС 40 СЕКУНД", 40, y_of(top, 12), "Inter-Semi", 11, BLUE)
    top += 20

    segments = [
        (3, "0–3 сек", "Крючок", BLUE),
        (9, "3–12 сек", "Ситуация", BLUE_MID),
        (13, "12–25 сек", "Одна фишка", BLUE_DARK),
        (10, "25–35 сек", "Блиц-запись", AMBER),
        (5, "35–40 сек", "Призыв", GREEN),
    ]
    inner = PAGE_W - 80
    seg_gap = 8
    usable = inner - seg_gap * (len(segments) - 1)
    bar_h = 46
    x = 40
    for dur, timing, label, color in segments:
        w = usable * dur / 40
        round_card(c, x, top, w, bar_h, color, None, 10)
        draw_text(c, timing, x + 12, y_of(top + 8, 11), "Inter", 10, HexColor("#EAF6FC") if color != AMBER else HexColor("#FFF4DE"))
        draw_text(c, label, x + 12, y_of(top + 24, 14), "Inter-Semi", 13, WHITE)
        x += w + seg_gap
    top += bar_h + 20

    draw_text(c, "ТРИ ПИЛОТНЫХ РОЛИКА", 40, y_of(top, 12), "Inter-Semi", 11, BLUE)
    draw_text(
        c,
        "Читаются сверху вниз как текст для съёмки. Жёлтая плашка — единственное место, где называется сервис.",
        230,
        y_of(top, 12),
        "Inter",
        11.5,
        MUTED,
    )
    top += 20

    scripts = [
        {
            "niche": "Салон и барбершоп",
            "hook": "Рекламу запустили, а записей больше не стало.",
            "situation": "Клиент пишет ночью. Утром администратор отвечает — человек уже записался туда, где в профиле есть кнопка.",
            "feature": "Онлайн-запись 24/7: клиент сам выбирает услугу, мастера и время. Без переписки и без ожидания утра.",
            "line": "Команда разработчиков создала сервис Блиц-запись. Ссылка — в шапке профиля.",
        },
        {
            "niche": "Клиника и стоматология",
            "hook": "Администратор не успевает брать трубку. Это тоже потерянные пациенты.",
            "situation": "Два звонка сразу. Один человек не дождался и записался в другую клинику.",
            "feature": "Виджет записи на сайте и в соцсетях. Пациент выбирает врача сам, все обращения падают в одно окно.",
            "line": "Команда разработчиков создала Блиц-запись: обращения и расписание в одном окне. Ссылка — в шапке профиля.",
        },
        {
            "niche": "Мойка, студия, специалист",
            "hook": "Спрос есть, а в расписании дыры.",
            "situation": "Запись живёт в переписке: два клиента на одно время, третий не пришёл — и это время никто не занял.",
            "feature": "Общий календарь по людям и местам плюс напоминание. Визит не теряется между чатами.",
            "line": "Команда разработчиков создала Блиц-запись: календарь, напоминания и база в одном сервисе. Ссылка — в шапке профиля.",
        },
    ]

    card_gap = 14
    card_w = (PAGE_W - 80 - card_gap * 2) / 3
    card_h = 318
    for i, sc in enumerate(scripts):
        x = 40 + i * (card_w + card_gap)
        round_card(c, x, top, card_w, card_h, WHITE, CARD_LINE, 14, shadow=True)
        c.setFillColor(BLUE if i < 2 else BLUE_DARK)
        y_head = y_of(top, 40)
        c.saveState()
        p = c.beginPath()
        p.roundRect(x, y_of(top, card_h), card_w, card_h, 14)
        c.clipPath(p, stroke=0, fill=0)
        c.rect(x, y_head, card_w, 40, fill=1, stroke=0)
        c.restoreState()
        draw_text(c, f"0{i+1}", x + 16, y_head + 13, "Inter-Bold", 14, WHITE)
        draw_text(c, sc["niche"], x + 48, y_head + 14, "Inter-Semi", 13, WHITE)

        blocks = [
            ("Крючок", sc["hook"], INK, None),
            ("Ситуация", sc["situation"], INK, None),
            ("Фишка", sc["feature"], INK, None),
        ]
        bt = top + 52
        for label, text, color, _ in blocks:
            draw_text(c, label.upper(), x + 16, y_of(bt, 10), "Inter-Semi", 9.5, SOFT)
            used = paragraph(c, text, x + 16, bt + 14, "Inter", 12, color, card_w - 32, 15.5)
            bt += 14 + used + 8

        line_h = 72
        line_top = top + card_h - line_h - 12
        round_card(c, x + 12, line_top, card_w - 24, line_h, AMBER_BG, AMBER_LINE, 10)
        draw_text(c, "В КОНЦЕ РОЛИКА", x + 24, y_of(line_top + 8, 10), "Inter-Semi", 9, AMBER_DARK)
        paragraph(c, sc["line"], x + 24, line_top + 24, "Inter-Semi", 12, INK, card_w - 48, 15.5)

    top += card_h + 20

    draw_text(c, "КАКАЯ БОЛЬ КАКОЙ ФИШКОЙ ЗАКРЫВАЕТСЯ", 40, y_of(top, 12), "Inter-Semi", 11, BLUE)
    draw_text(
        c,
        "В ролик берём одну строку. Остальное не перечисляем.",
        430,
        y_of(top, 12),
        "Inter",
        11.5,
        MUTED,
    )
    top += 20

    rows = [
        ("Клиентов мало, хотя реклама уже идёт", "Запись работает без администратора и ночью", "Онлайн-запись 24/7: услуга, специалист, время"),
        ("Написали ночью — ответили утром", "Кнопка «Записаться» прямо в профиле", "Виджет в соцсетях, на сайте, в мессенджерах и на картах"),
        ("Записи в переписке, накладки", "Одно расписание на всех", "Календарь по сотрудникам, кабинетам и ресурсам"),
        ("Клиент не пришёл, время пустое", "Напоминание и контроль отмен", "Уведомления о записи, изменении и отмене"),
        ("Непонятно, кто уже был", "История визитов под рукой", "База: контакты, посещения, комментарии"),
    ]
    headers = ("Боль, с которой начинается ролик", "Что говорит эксперт", "Где это в «Блиц-записи»")
    col_ws = [500, 500, PAGE_W - 80 - 1000]
    row_h = 36
    hx = 40
    round_card(c, 40, top, PAGE_W - 80, 28, BLUE_DARK, None, 8)
    hx = 40
    for head, w in zip(headers, col_ws):
        draw_text(c, head, hx + 12, y_of(top + 7, 12), "Inter-Semi", 11, WHITE)
        hx += w
    top += 32
    for r, row in enumerate(rows):
        fill = WHITE if r % 2 == 0 else HexColor("#F7FBFD")
        round_card(c, 40, top, PAGE_W - 80, row_h, fill, CARD_LINE, 8)
        hx = 40
        for text, w in zip(row, col_ws):
            draw_text(c, text, hx + 12, y_of(top + 10, 13), "Inter", 12, INK)
            hx += w
        top += row_h + 4

    top += 10
    draw_text(c, "ЧТО ЗАФИКСИРОВАТЬ НА ВСТРЕЧЕ", 40, y_of(top, 12), "Inter-Semi", 11, BLUE)
    top += 18

    decisions = [
        ("1", "Кто говорит", "Сотрудник в роли эксперта по клиентам. Не презентация продукта."),
        ("2", "Фраза авторства", "Оставляем «команда разработчиков создала» или меняем формулировку до съёмки."),
        ("3", "Куда ведём", "Шапка профиля, личные сообщения или пробный доступ. Один призыв на всю серию."),
        ("4", "Пилот", "Три ролика: салон, клиника, мойка или студия. Смотрим сохранения и ответы в директ."),
        ("5", "Чего нет в кадре", "Цены, акции, сравнение сервисов и «купите у Форус» — только после ролика."),
    ]
    d_gap = 10
    d_w = (PAGE_W - 80 - d_gap * 4) / 5
    d_h = 108
    x = 40
    for num, title, text in decisions:
        round_card(c, x, top, d_w, d_h, WHITE, CARD_LINE, 12, shadow=True)
        c.setFillColor(BLUE)
        c.circle(x + 22, y_of(top + 18, 0) - 2, 11, fill=1, stroke=0)
        draw_text(c, num, x + 22, y_of(top + 14, 11), "Inter-Bold", 11, WHITE, align="center")
        draw_text(c, title, x + 40, y_of(top + 12, 13), "Inter-Semi", 12, INK)
        paragraph(c, text, x + 14, top + 40, "Inter", 10.5, MUTED, d_w - 28, 13.5)
        x += d_w + d_gap

    footer(c, 2)
    print("page2 content ends at", top + d_h)


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT_PDF), pagesize=(PAGE_W, PAGE_H))
    c.setTitle("Схема продвижения «Блиц-запись» в Instagram")
    c.setAuthor("ГК Форус")
    c.setSubject("Контент-схема для обсуждения с коллегами")
    draw_page_1(c)
    c.showPage()
    draw_page_2(c)
    c.save()
    print(OUT_PDF)


if __name__ == "__main__":
    main()
