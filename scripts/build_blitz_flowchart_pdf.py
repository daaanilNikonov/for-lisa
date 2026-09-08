#!/usr/bin/env python3
"""PDF-блок-схема холодного звонка «Блиц-запись» — шпаргалка для МПП."""

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.colors import Color, white, HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
import os

OUT = "Блиц-запись_блок-схема_шпаргалка_МПП.pdf"
ASCII_OUT = "Blitz-zapis-flowchart-mpp.pdf"

NAVY = HexColor("#1B3A4B")
TEAL = HexColor("#2A9D8F")
ORANGE = HexColor("#E76F51")
SOFT = HexColor("#E8F5F3")
LIGHT = HexColor("#F4F7F8")
YELLOW = HexColor("#FFF3CD")
BORDER = HexColor("#D0D7DE")
DARK = HexColor("#243038")
GRAY = HexColor("#5A6570")

# Register a font that supports Cyrillic
FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
]
FONT_BOLD_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
]


def find_font(paths):
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def setup_fonts():
    reg = find_font(FONT_CANDIDATES)
    bold = find_font(FONT_BOLD_CANDIDATES)
    if not reg:
        raise SystemExit("No Cyrillic TTF font found")
    pdfmetrics.registerFont(TTFont("Body", reg))
    pdfmetrics.registerFont(TTFont("BodyBold", bold or reg))
    return "Body", "BodyBold"


def wrap(text, font, size, max_w):
    return simpleSplit(text, font, size, max_w)


def draw_round_rect(c, x, y, w, h, fill, stroke=BORDER, radius=6):
    c.setFillColor(fill)
    c.setStrokeColor(stroke)
    c.setLineWidth(1)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=1)


def draw_arrow_down(c, x, y, length=14):
    c.setStrokeColor(TEAL)
    c.setFillColor(TEAL)
    c.setLineWidth(1.5)
    c.line(x, y, x, y - length + 4)
    path = c.beginPath()
    path.moveTo(x, y - length)
    path.lineTo(x - 4, y - length + 7)
    path.lineTo(x + 4, y - length + 7)
    path.close()
    c.drawPath(path, fill=1, stroke=0)


def text_block(c, x, y_top, w, lines, font, size, color=DARK, leading=None):
    leading = leading or size + 3
    y = y_top
    c.setFillColor(color)
    c.setFont(font, size)
    for line in lines:
        c.drawString(x, y - size, line)
        y -= leading
    return y


def box_with_sections(c, x, y_top, w, title, sections, fonts, accent=TEAL):
    """
    sections: list of (label, text) — label is short header, text is body
    Returns bottom y of the box.
    """
    body, bold = fonts
    pad = 8
    title_size = 11
    label_size = 8.5
    body_size = 8.5
    max_w = w - 2 * pad

    # Measure
    title_lines = wrap(title, bold, title_size, max_w)
    content_h = 10 + len(title_lines) * (title_size + 3) + 6
    prepared = []
    for label, text in sections:
        lab = wrap(label, bold, label_size, max_w)
        body_lines = []
        for para in text.split("\n"):
            body_lines.extend(wrap(para, body, body_size, max_w) or [""])
        prepared.append((lab, body_lines))
        content_h += len(lab) * (label_size + 2) + 2
        content_h += len(body_lines) * (body_size + 2.5) + 8

    h = content_h + pad
    y_bottom = y_top - h

    # Card
    draw_round_rect(c, x, y_bottom, w, h, white, BORDER, 7)
    # accent bar
    c.setFillColor(accent)
    c.rect(x, y_bottom, 4, h, fill=1, stroke=0)

    y = y_top - pad
    # title
    c.setFillColor(NAVY)
    c.setFont(bold, title_size)
    for line in title_lines:
        c.drawString(x + pad + 2, y - title_size, line)
        y -= title_size + 3
    y -= 4

    for lab, body_lines in prepared:
        # label chip background
        chip_h = len(lab) * (label_size + 2) + 4
        c.setFillColor(SOFT)
        c.roundRect(x + pad, y - chip_h + 2, w - 2 * pad, chip_h, 3, fill=1, stroke=0)
        c.setFillColor(TEAL)
        c.setFont(bold, label_size)
        yy = y - 2
        for line in lab:
            c.drawString(x + pad + 4, yy - label_size, line)
            yy -= label_size + 2
        y = yy - 4

        c.setFillColor(DARK)
        c.setFont(body, body_size)
        for line in body_lines:
            c.drawString(x + pad + 2, y - body_size, line)
            y -= body_size + 2.5
        y -= 6

    return y_bottom


def gate_box(c, x, y_top, w, text, fonts):
    body, bold = fonts
    pad = 6
    size = 8.5
    lines = wrap(text, bold, size, w - 2 * pad)
    h = 10 + len(lines) * (size + 2.5) + pad
    y_bottom = y_top - h
    draw_round_rect(c, x, y_bottom, w, h, YELLOW, ORANGE, 5)
    c.setFillColor(ORANGE)
    c.setFont(bold, 7.5)
    c.drawString(x + pad, y_top - 10, "ПЕРЕХОД НА СЛЕДУЮЩИЙ ЭТАП")
    y = y_top - 14
    c.setFillColor(DARK)
    c.setFont(body, size)
    for line in lines:
        c.drawString(x + pad, y - size, line)
        y -= size + 2.5
    return y_bottom


def header(c, page_w, page_h, fonts, subtitle):
    body, bold = fonts
    c.setFillColor(NAVY)
    c.rect(0, page_h - 22 * mm, page_w, 22 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 14)
    c.drawString(14 * mm, page_h - 10 * mm, "Блиц-запись — блок-схема холодного звонка")
    c.setFont(body, 9)
    c.drawString(14 * mm, page_h - 16 * mm, subtitle)
    c.setFont(body, 8)
    c.drawRightString(page_w - 14 * mm, page_h - 10 * mm, "Шпаргалка для менеджера")
    c.drawRightString(page_w - 14 * mm, page_h - 16 * mm, "Цель звонка: запись на демонстрацию 15–20 минут")


def footer(c, page_w, page_num, total, fonts):
    body, bold = fonts
    c.setFillColor(GRAY)
    c.setFont(body, 8)
    c.drawCentredString(page_w / 2, 8 * mm, f"Страница {page_num} из {total}  ·  Группа компаний «Форус»  ·  «Блиц-запись»")


def build():
    fonts = setup_fonts()
    body, bold = fonts
    page_w, page_h = landscape(A4)

    # Content definition — human, no abbreviations (ЛПР kept as used)
    stages_page1 = [
        {
            "title": "Этап 1. Первичный контакт / приветствие и презентация компании",
            "sections": [
                (
                    "Что говорим",
                    "Добрый день, [Имя Отчество]! Меня зовут [имя], компания «Форус». Я попал(а) к владельцу или руководителю?\n"
                    "Если это ЛПР: есть две-три минуты? Хочу коротко рассказать, кто мы, и уточнить, как у вас устроена запись клиентов.\n"
                    "Презентация: наша компания занимается автоматизацией записи клиентов и предлагает для этого удобный сервис — «Блиц-запись». Мы из группы компаний «Форус», работаем с бизнесом в сфере услуг.\n"
                    "Переход к вопросам: расскажите, пожалуйста, как у вас сейчас записываются клиенты?",
                ),
                (
                    "Что выясняем",
                    "Это ЛПР или нужно перезвонить другому человеку?\nЕсть ли у собеседника две-три минуты сейчас?\nГотов ли человек коротко рассказать про текущую запись клиентов?",
                ),
            ],
            "gate": "Есть контакт с ЛПР (или назначен перезвон с конкретным временем), и собеседник согласился ответить на вопросы про запись.",
        },
        {
            "title": "Этап 2. Сбор потребности",
            "sections": [
                (
                    "Что спрашиваем",
                    "Как клиенты записываются сейчас: звонок, мессенджеры, соцсети или запись через интернет?\nКто ведёт расписание: владелец, администратор или мастер?\nЧем пользуются: тетрадь, таблица, «Юклиентс», «Дикиди» или другой сервис? Что нравится и что мешает?\nГде теряются клиенты: недозвоны, сообщения без ответа, двойные записи?\nЕсть ли напоминания о визите и насколько болезненны неявки?\nПишут ли вечером и в выходные, когда администратора нет?\nГде хранится база клиентов и история визитов?\nЧто больше всего отнимает время или нервы в записи и расписании?\nТему сейчас смотрят или «руки не доходят»?",
                ),
                (
                    "Какую информацию фиксируем",
                    "Канал записи · кто отвечает за расписание · текущий инструмент и боли по нему · потери заявок · неявки · запись вне рабочих часов · состояние базы · главная боль · готовность смотреть альтернативу.",
                ),
                (
                    "Обязательно перед следующим этапом",
                    "Повторяем услышанное своими словами: «Правильно понимаю: сейчас у вас [как записывают], больше всего мешает [боль], и важно [что хочет получить]. Верно?»",
                ),
            ],
            "gate": "Клиент подтвердил резюме потребности («да, верно»). Есть хотя бы одна понятная боль или задача, к которой можно привязать рассказ о сервисе.",
        },
    ]

    stages_page2 = [
        {
            "title": "Этап 3. Рассказ о «Блиц-записи» (только под боль клиента)",
            "sections": [
                (
                    "Правило",
                    "Называем два-три пункта из ответов клиента. Прайс не читаем. С текущим сервисом не спорим — предлагаем сравнить на демонстрации.",
                ),
                (
                    "Какой тезис берём",
                    "Теряют заявки / много звонков → клиент сам записывается через интернет, заявка сразу в календарь.\nПутаница в расписании → всё в одном месте, система не даёт задвоить запись.\nНеявки → автоматические напоминания, меньше простоев.\nБаза размазана → контакты и история визитов в одном сервисе.\nДорого / много лишнего → цена не от числа мастеров, перенос базы берём на себя, до шести месяцев в подарок при переходе.",
                ),
                (
                    "Пример внедрения (по необходимости)",
                    "Салон «Луна»: переход с «Юклиентс», экономия около трети в год, перенос базы под ключ.\nАвтомойка «Элис»: восемь постов, календарь под час пик, касса, шесть месяцев в подарок.",
                ),
            ],
            "gate": "Клиент услышал ценность «под свою боль» и не закрылся. Можно переходить к приглашению на демонстрацию (даже если остались сомнения — их закрываем вопросом про время).",
        },
        {
            "title": "Этап 4. Приглашение на демонстрацию",
            "sections": [
                (
                    "Что говорим",
                    "Давайте я вам просто покажу, как это выглядит. За 15–20 минут пройдёмся по календарю, записи и напоминаниям — уже на вашем формате. Когда удобнее: завтра до обеда или после? Или во вторник / в среду?",
                ),
                (
                    "Если «надо подумать»",
                    "Понимаю. Как раз поэтому и предлагаю короткую демонстрацию — посмотрите глазами, и уже после будет понятнее. Давайте поставим 15 минут на [день]?",
                ),
                (
                    "Что фиксируем при согласии",
                    "Дата и время · почта или мессенджер для ссылки · напоминание за день и за час.",
                ),
            ],
            "gate": "УСПЕХ ЗВОНКА: назначены дата и время демонстрации, есть канал для ссылки.\nЕсли демонстрацию сейчас не ставят — обязателен следующий шаг: коммерческое предложение на почту и конкретное время перезвона.",
        },
    ]

    stages_page3 = [
        {
            "title": "Если демонстрацию не ставят сразу",
            "sections": [
                (
                    "Что говорим",
                    "Хорошо, настаивать не буду. Зафиксирую, что для вас важно: [боль]. Могу скинуть короткое коммерческое предложение на почту и созвониться [день / время]. Какая почта удобнее?",
                ),
                (
                    "Что обязательно сделать",
                    "Письмо или материал · дата перезвона · комментарий в системе учёта. Звонок без следующего шага — потерянный контакт.",
                ),
            ],
            "gate": "Есть понятный следующий шаг с датой. Иначе этап не закрыт.",
        },
        {
            "title": "Частые возражения → куда возвращаем диалог",
            "sections": [
                (
                    "Некогда / всё устраивает / пришлите на почту",
                    "Короткая демонстрация в удобный слот · сравнение «как есть» и сервис · письмо + сразу время на просмотр.",
                ),
                (
                    "Уже есть другой сервис / дорого / маленький бизнес",
                    "Не спорим — зовём сравнить · цифры и цена не от числа мастеров · простой сценарий для небольшого бизнеса.",
                ),
                (
                    "Страх переноса / сложно внедрять / нет сайта",
                    "Перенос берём на себя · настраиваем сами · сайт не обязателен, достаточно ссылки в мессенджере и соцсетях.",
                ),
            ],
            "gate": "После отработки возражения снова предлагаем конкретное время демонстрации (выбор из двух вариантов).",
        },
        {
            "title": "Памятка на весь звонок",
            "sections": [
                (
                    "Как ведём разговор",
                    "Больше говорит клиент · после блока вопросов — пауза и слушаем · дольше 40–50 секунд подряд не говорим · тариф не продаём в холодном звонке · цену называем только если спросили.",
                ),
                (
                    "Цель",
                    "Не «продать подписку в звонке», а записать на демонстрацию сервиса на 15–20 минут.",
                ),
            ],
            "gate": None,
        },
    ]

    c = canvas.Canvas(OUT, pagesize=landscape(A4))
    pages = [
        ("Лист 1: контакт и сбор потребности", stages_page1),
        ("Лист 2: рассказ о сервисе и демонстрация", stages_page2),
        ("Лист 3: если не записали сразу · возражения · правила", stages_page3),
    ]
    total = len(pages)

    for page_idx, (subtitle, stages) in enumerate(pages, 1):
        header(c, page_w, page_h, fonts, subtitle)
        footer(c, page_w, page_idx, total, fonts)

        margin_x = 12 * mm
        top = page_h - 28 * mm
        gap = 6 * mm
        usable_w = page_w - 2 * margin_x

        # Two columns if 2 stages, or stack if 3
        if len(stages) == 2:
            col_w = (usable_w - gap) / 2
            positions = [
                (margin_x, top),
                (margin_x + col_w + gap, top),
            ]
            for stage, (x, y) in zip(stages, positions):
                yb = box_with_sections(
                    c, x, y, col_w, stage["title"], stage["sections"], fonts
                )
                if stage.get("gate"):
                    draw_arrow_down(c, x + col_w / 2, yb - 2, 12)
                    gate_box(c, x, yb - 16, col_w, stage["gate"], fonts)
        else:
            # 3 stages stacked / two on top one bottom spanning
            col_w = (usable_w - gap) / 2
            yb1 = box_with_sections(
                c, margin_x, top, col_w, stages[0]["title"], stages[0]["sections"], fonts
            )
            if stages[0].get("gate"):
                draw_arrow_down(c, margin_x + col_w / 2, yb1 - 2, 10)
                gate_box(c, margin_x, yb1 - 14, col_w, stages[0]["gate"], fonts)

            yb2 = box_with_sections(
                c,
                margin_x + col_w + gap,
                top,
                col_w,
                stages[1]["title"],
                stages[1]["sections"],
                fonts,
                accent=ORANGE,
            )
            if stages[1].get("gate"):
                draw_arrow_down(c, margin_x + col_w + gap + col_w / 2, yb2 - 2, 10)
                gate_box(
                    c,
                    margin_x + col_w + gap,
                    yb2 - 14,
                    col_w,
                    stages[1]["gate"],
                    fonts,
                )

            bottom_top = min(yb1, yb2) - 28 * mm
            if bottom_top > 40 * mm:
                yb3 = box_with_sections(
                    c,
                    margin_x,
                    bottom_top,
                    usable_w,
                    stages[2]["title"],
                    stages[2]["sections"],
                    fonts,
                    accent=NAVY,
                )

        # flow legend on page 1
        if page_idx == 1:
            c.setFillColor(GRAY)
            c.setFont(body, 7.5)
            c.drawString(
                margin_x,
                14 * mm,
                "Читать сверху вниз внутри карточки. Жёлтый блок — условие, без которого на следующий этап не переходим.",
            )

        c.showPage()

    c.save()
    # ascii copy
    import shutil

    shutil.copy(OUT, ASCII_OUT)
    print("Saved", OUT)
    print("Saved", ASCII_OUT)


if __name__ == "__main__":
    build()
