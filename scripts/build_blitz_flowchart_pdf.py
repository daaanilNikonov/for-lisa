#!/usr/bin/env python3
"""PDF-блок-схема холодного звонка «Блиц-запись» — только скрипт по этапам."""

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
import os
import shutil

OUT = "Блиц-запись_блок-схема_шпаргалка_МПП.pdf"
ASCII_OUT = "Blitz-zapis-flowchart-mpp.pdf"

NAVY = HexColor("#1B3A4B")
TEAL = HexColor("#2A9D8F")
ORANGE = HexColor("#E76F51")
BORDER = HexColor("#C5CDD4")
DARK = HexColor("#243038")
ARROW = HexColor("#2A9D8F")
GRAY = HexColor("#7A8792")

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
]
FONT_BOLD_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
]


def find_font(paths):
    for p in paths:
        if os.path.exists(p):
            return p
    raise SystemExit("No Cyrillic TTF font found")


def setup_fonts():
    pdfmetrics.registerFont(TTFont("Body", find_font(FONT_CANDIDATES)))
    pdfmetrics.registerFont(TTFont("BodyBold", find_font(FONT_BOLD_CANDIDATES)))
    return "Body", "BodyBold"


def wrap(text, font, size, max_w):
    return simpleSplit(text, font, size, max_w)


def draw_round_rect(c, x, y, w, h, fill, stroke=BORDER, radius=8):
    c.setFillColor(fill)
    c.setStrokeColor(stroke)
    c.setLineWidth(1.2)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=1)


def draw_arrow_right(c, x_from, x_to, y, label=None, fonts=None):
    c.setStrokeColor(ARROW)
    c.setFillColor(ARROW)
    c.setLineWidth(2.2)
    tip = x_to - 2
    c.line(x_from, y, tip - 9, y)
    path = c.beginPath()
    path.moveTo(tip, y)
    path.lineTo(tip - 10, y - 5)
    path.lineTo(tip - 10, y + 5)
    path.close()
    c.drawPath(path, fill=1, stroke=0)
    if label and fonts:
        _, bold = fonts
        c.setFillColor(TEAL)
        c.setFont(bold, 8)
        c.drawCentredString((x_from + tip) / 2, y + 7, label)


def draw_arrow_down(c, x, y_from, y_to, label=None, fonts=None):
    c.setStrokeColor(ARROW)
    c.setFillColor(ARROW)
    c.setLineWidth(2.2)
    tip = y_to + 2
    c.line(x, y_from, x, tip + 9)
    path = c.beginPath()
    path.moveTo(x, tip)
    path.lineTo(x - 5, tip + 10)
    path.lineTo(x + 5, tip + 10)
    path.close()
    c.drawPath(path, fill=1, stroke=0)
    if label and fonts:
        _, bold = fonts
        c.setFillColor(TEAL)
        c.setFont(bold, 8)
        c.drawCentredString(x + 28, (y_from + tip) / 2 - 2, label)


def stage_card(c, x, y_top, w, stage_num, title, script_lines, fonts, accent=TEAL):
    body, bold = fonts
    pad = 10
    title_size = 11
    body_size = 8.5
    max_w = w - 2 * pad

    if isinstance(stage_num, int):
        title_full = f"Этап {stage_num}. {title}"
        badge_text = str(stage_num)
    else:
        title_full = title
        badge_text = "↔"

    title_wrapped = wrap(title_full, bold, title_size, max_w - 30)

    content_blocks = []
    content_h = 0
    for block in script_lines:
        lines = []
        for para in block.split("\n"):
            lines.extend(wrap(para, body, body_size, max_w) or [""])
        content_blocks.append(lines)
        content_h += len(lines) * (body_size + 2.8) + 5

    header_h = 12 + len(title_wrapped) * (title_size + 2.5)
    h = header_h + content_h + pad
    y_bottom = y_top - h

    draw_round_rect(c, x, y_bottom, w, h, white, BORDER, 8)

    c.setFillColor(accent)
    c.roundRect(x, y_top - header_h, w, header_h, 8, fill=1, stroke=0)
    c.rect(x, y_top - header_h, w, 10, fill=1, stroke=0)

    badge_r = 9
    bx = x + pad + badge_r
    by = y_top - header_h / 2
    c.setFillColor(white)
    c.circle(bx, by, badge_r, fill=1, stroke=0)
    c.setFillColor(accent)
    c.setFont(bold, 10)
    c.drawCentredString(bx, by - 3.5, badge_text)

    c.setFillColor(white)
    c.setFont(bold, title_size)
    ty = y_top - 10
    for i, line in enumerate(title_wrapped):
        indent = 28 if i == 0 else 12
        c.drawString(x + pad + indent, ty - title_size, line)
        ty -= title_size + 2.5

    y = y_top - header_h - 7
    c.setFillColor(DARK)
    c.setFont(body, body_size)
    for lines in content_blocks:
        for line in lines:
            c.drawString(x + pad, y - body_size, line)
            y -= body_size + 2.8
        y -= 3

    return y_bottom, x + w / 2


def page_header(c, page_w, page_h, fonts):
    _, bold = fonts
    c.setFillColor(NAVY)
    c.rect(0, page_h - 14 * mm, page_w, 14 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 12)
    c.drawString(12 * mm, page_h - 9 * mm, "Блиц-запись — блок-схема холодного звонка")


def page_footer(c, page_w, n, total, fonts):
    body, _ = fonts
    c.setFillColor(GRAY)
    c.setFont(body, 8)
    c.drawCentredString(page_w / 2, 6 * mm, f"{n} / {total}")


def flow_strip(c, page_w, fonts, active=None):
    """Bottom strip: 1→2→3→4→5 with arrows."""
    _, bold = fonts
    labels = [
        (1, "Контакт"),
        (2, "Потребность"),
        (3, "Рассказ"),
        (4, "Демонстрация"),
        (5, "Если не записали"),
    ]
    margin = 18 * mm
    usable = page_w - 2 * margin
    y = 13 * mm
    step = usable / (len(labels) - 1)
    for i, (num, lab) in enumerate(labels):
        x = margin + i * step
        on = active is None or num in active
        c.setFillColor(TEAL if on else GRAY)
        c.circle(x, y + 9, 6, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont(bold, 8)
        c.drawCentredString(x, y + 6, str(num))
        c.setFillColor(NAVY if on else GRAY)
        c.setFont(bold, 7)
        c.drawCentredString(x, y - 2, lab)
        if i < len(labels) - 1:
            c.setStrokeColor(ARROW if on else GRAY)
            c.setFillColor(ARROW if on else GRAY)
            c.setLineWidth(1.6)
            c.line(x + 8, y + 9, x + step - 8, y + 9)
            path = c.beginPath()
            path.moveTo(x + step - 8, y + 9)
            path.lineTo(x + step - 13, y + 9 - 3)
            path.lineTo(x + step - 13, y + 9 + 3)
            path.close()
            c.drawPath(path, fill=1, stroke=0)


def build():
    fonts = setup_fonts()
    page_w, page_h = landscape(A4)
    c = canvas.Canvas(OUT, pagesize=landscape(A4))

    margin = 12 * mm
    top = page_h - 20 * mm
    usable = page_w - 2 * margin
    gap = 18 * mm
    col_w = (usable - gap) / 2
    left_x = margin
    right_x = margin + col_w + gap

    # ----- PAGE 1: 1 → 2 -----
    page_header(c, page_w, page_h, fonts)
    page_footer(c, page_w, 1, 3, fonts)
    flow_strip(c, page_w, fonts, active={1, 2})

    stage_card(
        c,
        left_x,
        top,
        col_w,
        1,
        "Первичный контакт / приветствие и презентация компании",
        [
            "Добрый день, [Имя Отчество]! Меня зовут [имя], компания «Форус». Я попал(а) к владельцу или руководителю [салона / автомойки / школы / студии]?",
            "Если не ЛПР: с кем лучше поговорить про запись клиентов и расписание? Как обратиться и когда перезвонить?",
            "Если ЛПР: скажите, у вас есть две-три минуты? Хочу коротко рассказать, кто мы, и уточнить, как у вас сейчас устроена запись клиентов.",
            "Наша компания занимается автоматизацией записи клиентов и предлагает для этого удобный сервис — «Блиц-запись». Мы из группы компаний «Форус», работаем с бизнесом в сфере услуг.",
            "Расскажите, пожалуйста, как у вас сейчас записываются клиенты?",
        ],
        fonts,
        TEAL,
    )

    stage_card(
        c,
        right_x,
        top,
        col_w,
        2,
        "Сбор потребности",
        [
            "Клиенты к вам чаще звонят, пишут в мессенджеры, приходят из соцсетей — или уже есть запись через интернет?",
            "Кто у вас этим занимается — вы сами, администратор или мастер?",
            "Чем пользуетесь для записи — тетрадью, таблицей, «Юклиентс», «Дикиди» или чем-то ещё? Что нравится, а что раздражает?",
            "Бывает, что в час пик клиент не дозванивается, пишет в мессенджер — и запись пропадает? Как часто так бывает?",
            "Вы напоминаете клиентам о визите? Насколько болезненны неявки?",
            "Пишут или звонят вечером и в выходные, когда администратора нет?",
            "Где хранится база клиентов и история визитов?",
            "Что в записи и расписании сейчас больше всего отнимает время или нервы?",
            "Эту тему сейчас смотрите, или руки не доходят?",
            "Правильно понимаю: сейчас у вас [как записывают], больше всего мешает [боль], и важно [что хочет получить]. Верно?",
        ],
        fonts,
        TEAL,
    )

    draw_arrow_right(
        c,
        left_x + col_w + 2,
        right_x - 2,
        top - 45 * mm,
        "этап 1 → этап 2",
        fonts,
    )
    c.showPage()

    # ----- PAGE 2: 3 → 4 -----
    page_header(c, page_w, page_h, fonts)
    page_footer(c, page_w, 2, 3, fonts)
    flow_strip(c, page_w, fonts, active={3, 4})

    stage_card(
        c,
        left_x,
        top,
        col_w,
        3,
        "Рассказ о «Блиц-записи»",
        [
            "Берём два-три пункта из ответов клиента. Прайс не читаем. С текущим сервисом не спорим.",
            "Теряют заявки / много звонков: клиент сам записывается через интернет — с сайта, из соцсетей, с карт. Заявка сразу в календарь администратора.",
            "Путаница в расписании: всё в одном месте по сотрудникам, кабинетам, постам. Система не даёт задвоить запись.",
            "Неявки: сервис сам напоминает о визите — смс или мессенджер. Простоев становится меньше.",
            "База размазана: контакты, история визитов, комментарии в одном месте.",
            "Дорого / много лишнего: цена не растёт от числа мастеров. При переходе — до шести месяцев в подарок, базу переносим сами.",
            "Если нужен общий рассказ: «Блиц-запись» — запись через интернет, расписание, база и напоминания в одном окне. Работает с компьютера, планшета и телефона. Мы помогаем подключить и настроить.",
            "Примеры при необходимости: салон «Луна» — ушли с «Юклиентс», экономия около трети в год. Автомойка «Элис» — восемь постов, календарь под час пик, касса, шесть месяцев в подарок.",
        ],
        fonts,
        TEAL,
    )

    stage_card(
        c,
        right_x,
        top,
        col_w,
        4,
        "Приглашение на демонстрацию",
        [
            "Давайте я вам просто покажу, как это выглядит. За 15–20 минут пройдёмся по календарю, записи и напоминаниям — уже на вашем формате. Когда удобнее: завтра до обеда или после? Или во вторник / в среду?",
            "Если «надо подумать»: понимаю. Как раз поэтому и предлагаю короткую демонстрацию — посмотрите глазами, и уже после будет понятнее. Давайте поставим 15 минут на [день]?",
            "Хорошо. Куда удобнее отправить ссылку — на почту или в мессенджер? Подтверждаем: [дата], [время], демонстрация «Блиц-запись», минут 15–20. Я за день и за час напомню. Хорошего дня, [Имя Отчество]!",
        ],
        fonts,
        ORANGE,
    )

    draw_arrow_right(
        c,
        left_x + col_w + 2,
        right_x - 2,
        top - 45 * mm,
        "этап 3 → этап 4",
        fonts,
    )
    c.showPage()

    # ----- PAGE 3: 5 + objections → back to 4 -----
    page_header(c, page_w, page_h, fonts)
    page_footer(c, page_w, 3, 3, fonts)
    flow_strip(c, page_w, fonts, active={4, 5})

    yb5, _ = stage_card(
        c,
        margin,
        top,
        usable,
        5,
        "Если демонстрацию сейчас не ставят",
        [
            "Хорошо, настаивать не буду. Зафиксирую, что для вас важно: [боль]. Могу скинуть короткое коммерческое предложение на почту и созвониться [день / время] — сравним с тем, как у вас сейчас. Какая почта удобнее?",
        ],
        fonts,
        NAVY,
    )

    draw_arrow_down(
        c,
        page_w / 2,
        yb5 - 2,
        yb5 - 11 * mm,
        "этап 5 → возражения / снова этап 4",
        fonts,
    )

    stage_card(
        c,
        margin,
        yb5 - 13 * mm,
        usable,
        "↔",
        "Отработка возражений → снова этап 4 (демонстрация)",
        [
            "Некогда: предлагаю 15 минут демонстрации в удобное время. Когда спокойнее — утром или вечером?",
            "Всё устраивает: часто смотрят, чтобы администратор меньше сидел на телефоне и записи вечером не терялись. Давайте за 15 минут сравните с тем, как у вас сейчас.",
            "Уже есть другой сервис: не спорим. К нам приходят из‑за цены «за мастера» и тяжёлого интерфейса. Могу коротко показать разницу на демонстрации.",
            "Дорого: ориентир от 1 500 рублей в месяц при оплате на два года, на год — около 1 900; от числа мастеров цена не растёт. На демонстрации прикинем под вас.",
            "Маленький бизнес: как раз для небольшого — запись, календарь, база, напоминания без лишнего. На демонстрации покажем простой сценарий.",
            "Клиенты всё равно звонят: часть будет звонить — это нормально. Сервис забирает тех, кто готов записаться сам. Администратор тоже может записывать, как раньше.",
            "Боимся переносить базу: перенос берём на себя. Сначала смотрите в тестовом доступе.",
            "Сложно внедрять: мы сами подключаем и настраиваем услуги, расписание, сотрудников, ссылки на запись.",
            "Пришлите на почту: направлю. И давайте сразу поставим 15 минут — так материал не потеряется. Когда удобно?",
            "Надо посоветоваться: пригласим партнёра или администратора сразу на демонстрацию. Когда вам двоим удобно?",
            "Тетрадь / таблица: обычно ломается в час пик. «Блиц» — тот же журнал в телефоне и на компьютере. Посмотрите 15 минут.",
            "Нет сайта: сайт не обязателен. Ссылку ставят в мессенджер, соцсети, на карты.",
            "После любого возражения: когда удобнее на демонстрацию — завтра до обеда или после?",
        ],
        fonts,
        ORANGE,
    )

    c.showPage()
    c.save()
    shutil.copy(OUT, ASCII_OUT)
    print("Saved", OUT)
    print("Saved", ASCII_OUT)


if __name__ == "__main__":
    build()
