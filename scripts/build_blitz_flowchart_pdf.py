#!/usr/bin/env python3
"""
Вертикальная блок-схема холодного звонка «Блиц-запись»
Новый заход: бесплатный Старт + самостоятельная онлайн-запись.
Формат: Н → уровни со стрелками → К, вопросы под каждым уровнем.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
import shutil

OUT = "Блиц-запись_блок-схема.pdf"
ASCII_OUT = "Blitz-zapis-flowchart.pdf"

NAVY = HexColor("#1B3A4B")
TEAL = HexColor("#2A9D8F")
ORANGE = HexColor("#E76F51")
BLUE = HexColor("#3D7EA6")
LIGHT = HexColor("#F0F4F5")
YELLOW = HexColor("#FFF6E8")
BORDER = HexColor("#B8C2CA")
DARK = HexColor("#243038")
GRAY = HexColor("#6B7680")

FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"


def setup_fonts():
    pdfmetrics.registerFont(TTFont("Body", FONT_REG))
    pdfmetrics.registerFont(TTFont("BodyBold", FONT_BOLD))
    return "Body", "BodyBold"


def wrap(text, font, size, max_w):
    return simpleSplit(text, font, size, max_w)


def draw_circle(c, cx, cy, r, letter, fonts, fill=TEAL):
    _, bold = fonts
    c.setFillColor(fill)
    c.circle(cx, cy, r, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 13)
    c.drawCentredString(cx, cy - 4.5, letter)


def draw_arrow_v(c, x, y_from, y_to):
    c.setStrokeColor(TEAL)
    c.setFillColor(TEAL)
    c.setLineWidth(1.8)
    tip = y_to + 1
    c.line(x, y_from, x, tip + 8)
    path = c.beginPath()
    path.moveTo(x, tip)
    path.lineTo(x - 4.5, tip + 8)
    path.lineTo(x + 4.5, tip + 8)
    path.close()
    c.drawPath(path, fill=1, stroke=0)


def draw_level(c, x, y_top, w, title, fonts, accent=TEAL):
    _, bold = fonts
    pad = 10
    size = 12
    lines = wrap(title, bold, size, w - 2 * pad)
    h = 14 + len(lines) * (size + 3)
    y_bottom = y_top - h
    c.setFillColor(accent)
    c.roundRect(x, y_bottom, w, h, 6, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, size)
    ty = y_top - 10
    for line in lines:
        c.drawCentredString(x + w / 2, ty - size, line)
        ty -= size + 3
    return y_bottom


def draw_questions(c, x, y_top, w, items, fonts, caption="Что говорим / спрашиваем"):
    body, bold = fonts
    pad = 8
    cap_size = 8
    body_size = 8.2
    max_w = w - 2 * pad - 8

    prepared = []
    content_h = 8 + cap_size + 4
    for item in items:
        lines = wrap("• " + item, body, body_size, max_w)
        prepared.append(lines)
        content_h += len(lines) * (body_size + 2.3) + 3
    h = content_h + pad
    y_bottom = y_top - h

    c.setFillColor(LIGHT)
    c.setStrokeColor(BORDER)
    c.setLineWidth(0.8)
    c.roundRect(x, y_bottom, w, h, 5, fill=1, stroke=1)

    c.setFillColor(TEAL)
    c.setFont(bold, cap_size)
    c.drawString(x + pad, y_top - cap_size - 4, caption)

    y = y_top - cap_size - 10
    c.setFillColor(DARK)
    c.setFont(body, body_size)
    for lines in prepared:
        for line in lines:
            c.drawString(x + pad, y - body_size, line)
            y -= body_size + 2.3
        y -= 2
    return y_bottom


def draw_branch_table(c, x, y_top, w, rows, fonts, title="Если клиент говорит → что говорим → куда дальше"):
    body, bold = fonts
    pad = 6
    title_size = 8
    cell = 7.2
    col_w = [(w - 2 * pad) * 0.28, (w - 2 * pad) * 0.44, (w - 2 * pad) * 0.28]

    prepared = []
    content_h = 8 + title_size + 6
    headers = ["Клиент говорит", "Что говорим", "Дальше"]
    header_lines = [wrap(h, bold, cell, col_w[i] - 4) for i, h in enumerate(headers)]
    header_h = max(len(hl) for hl in header_lines) * (cell + 2) + 4
    content_h += header_h + 2

    for a, b, d in rows:
        cells = [
            wrap(a, body, cell, col_w[0] - 4),
            wrap(b, body, cell, col_w[1] - 4),
            wrap(d, body, cell, col_w[2] - 4),
        ]
        rh = max(len(x) for x in cells) * (cell + 2) + 4
        prepared.append((cells, rh))
        content_h += rh

    h = content_h + pad
    y_bottom = y_top - h
    c.setFillColor(YELLOW)
    c.setStrokeColor(BORDER)
    c.setLineWidth(0.8)
    c.roundRect(x, y_bottom, w, h, 5, fill=1, stroke=1)

    c.setFillColor(ORANGE)
    c.setFont(bold, title_size)
    c.drawString(x + pad, y_top - title_size - 3, title)

    y = y_top - title_size - 8
    c.setFillColor(HexColor("#F3D9C0"))
    c.rect(x + pad, y - header_h, w - 2 * pad, header_h, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont(bold, cell)
    cx = x + pad
    for i, hl in enumerate(header_lines):
        yy = y - 3
        for line in hl:
            c.drawString(cx + 2, yy - cell, line)
            yy -= cell + 2
        cx += col_w[i]
    y -= header_h

    for idx, (cells, rh) in enumerate(prepared):
        c.setFillColor(white if idx % 2 == 0 else HexColor("#FFFBF3"))
        c.rect(x + pad, y - rh, w - 2 * pad, rh, fill=1, stroke=0)
        c.setStrokeColor(BORDER)
        c.setLineWidth(0.4)
        c.line(x + pad, y - rh, x + w - pad, y - rh)
        c.setFillColor(DARK)
        c.setFont(body, cell)
        cx = x + pad
        for i, cell_lines in enumerate(cells):
            yy = y - 3
            for line in cell_lines:
                c.drawString(cx + 2, yy - cell, line)
                yy -= cell + 2
            cx += col_w[i]
        y -= rh
    return y_bottom


def page_head(c, page_w, page_h, fonts, n, total, cont=False):
    body, bold = fonts
    c.setFillColor(NAVY)
    c.rect(0, page_h - 16 * mm, page_w, 16 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(bold, 12)
    c.drawCentredString(page_w / 2, page_h - 7.5 * mm, "Блиц-запись — блок-схема холодного звонка")
    c.setFont(body, 8)
    c.drawCentredString(page_w / 2, page_h - 12.5 * mm, "бесплатный Старт · самостоятельная онлайн-запись · демонстрация")
    c.setFillColor(GRAY)
    c.drawCentredString(page_w / 2, 7 * mm, f"{n} / {total}")
    if cont:
        c.setFillColor(TEAL)
        c.setFont(bold, 8)
        c.drawCentredString(page_w / 2, page_h - 20 * mm, "↓ продолжение схемы")


def build():
    fonts = setup_fonts()
    body, bold = fonts
    page_w, page_h = A4
    c = canvas.Canvas(OUT, pagesize=A4)

    margin = 16 * mm
    box_w = page_w - 2 * margin
    cx = page_w / 2
    total = 3

    # ========== PAGE 1 ==========
    page_head(c, page_w, page_h, fonts, 1, total)
    y = page_h - 24 * mm

    c.setFillColor(NAVY)
    c.setFont(bold, 11)
    c.drawCentredString(cx, y, "Холодный звонок по «Блиц-записи»")
    y -= 9 * mm

    draw_circle(c, cx, y - 7, 8, "Н", fonts, TEAL)
    y -= 7 + 8
    draw_arrow_v(c, cx, y, y - 7 * mm)
    y -= 7 * mm

    # Level 1
    y = draw_level(
        c,
        margin,
        y,
        box_w,
        "1. Первичный контакт / приветствие и презентация компании",
        fonts,
        TEAL,
    )
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions(
        c,
        margin + 5 * mm,
        y,
        box_w - 10 * mm,
        [
            "Добрый день! Меня зовут [имя], я из компании «Форус». Мы предлагаем вам посотрудничать с нами и бесплатно установить наше новое приложение для самостоятельной онлайн-записи клиентов. Оно интегрируется куда угодно: 2ГИС, ваш сайт или социальные сети. С кем могу обсудить вопрос сотрудничества?",
            "Если не ЛПР: с кем лучше обсудить запись клиентов и расписание? Как обратиться и когда перезвонить?",
            "Если ЛПР: у вас есть две-три минуты? Хочу коротко показать идею и уточнить, как у вас сейчас устроена запись.",
        ],
        fonts,
    )
    draw_arrow_v(c, cx, y, y - 7 * mm)
    y -= 7 * mm

    # Level 2
    y = draw_level(c, margin, y, box_w, "2. Короткий рассказ про сервис", fonts, TEAL)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions(
        c,
        margin + 5 * mm,
        y,
        box_w - 10 * mm,
        [
            "«Блиц-запись» — сервис для самостоятельной онлайн-записи. Клиент сам выбирает услугу и время с сайта, из 2ГИС, Яндекс Карт или соцсетей.",
            "Можно начать бесплатно на тарифе «Старт»: календарь, онлайн-запись, прайс, до пяти сотрудников — без абонентской платы.",
            "Если нужно больше — подключаете только нужные инструменты или берёте полный тариф.",
            "Переход: расскажите, пожалуйста, как у вас сейчас записываются клиенты?",
        ],
        fonts,
    )
    draw_arrow_v(c, cx, y, y - 7 * mm)
    y -= 7 * mm

    # Level 3
    y = draw_level(c, margin, y, box_w, "3. Сбор потребности", fonts, BLUE)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions(
        c,
        margin + 5 * mm,
        y,
        box_w - 10 * mm,
        [
            "Клиенты чаще звонят, пишут в мессенджеры, приходят из соцсетей — или уже есть онлайн-запись?",
            "Кто этим занимается — вы, администратор или мастер?",
            "Чем пользуетесь: тетрадь, таблица, «Юклиентс», «Дикиди»? Что нравится и что мешает?",
            "Бывает, что в час пик не дозваниваются и запись пропадает?",
            "Могут ли клиенты записаться сами с сайта, 2ГИС, соцсетей — или всё через администратора?",
            "Напоминаете о визите? Насколько болезненны неявки?",
            "Что в записи и расписании больше всего отнимает время или нервы?",
            "Повтор: сейчас у вас [как записывают], мешает [боль], важно [цель]. Верно?",
        ],
        fonts,
        "Вопросы",
    )

    c.setFillColor(TEAL)
    c.setFont(bold, 8)
    c.drawCentredString(cx, 13 * mm, "↓ продолжение на следующей странице")
    c.showPage()

    # ========== PAGE 2 ==========
    page_head(c, page_w, page_h, fonts, 2, total, cont=True)
    y = page_h - 26 * mm
    draw_circle(c, cx, y - 5, 5, "↓", fonts, GRAY)
    y -= 10
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    # Level 4
    y = draw_level(
        c, margin, y, box_w, "4. Рассказ о сервисе под боль клиента", fonts, TEAL
    )
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions(
        c,
        margin + 5 * mm,
        y,
        box_w - 10 * mm,
        [
            "Теряют заявки / много звонков → клиент сам записывается с сайта, 2ГИС, карт и соцсетей; заявка сразу в календарь.",
            "Не готовы платить сразу → тариф «Старт» 0 ₽/мес: календарь, онлайн-запись, прайс, до 5 сотрудников, 1 филиал.",
            "Путаница в расписании → всё в одном месте, нельзя задвоить запись; на «Управлении» — графики и загрузка.",
            "Нужна база и возврат клиентов → на «Старте» база есть; история, выгрузка и лояльность — в модулях или в «Управлении».",
            "Нужен запуск под ключ и бренд → «Премиум»: менеджер, брендирование, витрина, каналы привлечения, аналитика из 2ГИС и соцсетей.",
        ],
        fonts,
        "Какой тезис берём",
    )
    draw_arrow_v(c, cx, y, y - 7 * mm)
    y -= 7 * mm

    # Level 5 tariffs mini
    y = draw_level(c, margin, y, box_w, "5. Тарифы (коротко, если спросили)", fonts, ORANGE)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions(
        c,
        margin + 5 * mm,
        y,
        box_w - 10 * mm,
        [
            "Формула КП: начните бесплатно, подключайте только нужные инструменты или получите все возможности в одном тарифе.",
            "«Старт» — 0 ₽/мес.",
            "Модули к «Старту» — аналитика, график, зарплата, полная база, Telegram-бот, лояльность.",
            "«Управление» — 1 900 ₽/мес при оплате за год (22 800 ₽/год) или 2 100 ₽ помесячно. Экономия за год — 2 400 ₽.",
            "«Премиум» — 2 067 ₽/мес при оплате за год (24 800 ₽/год) или 2 300 ₽ помесячно.",
        ],
        fonts,
        "Ориентиры по цене",
    )
    draw_arrow_v(c, cx, y, y - 7 * mm)
    y -= 7 * mm

    # Level 6 demo
    y = draw_level(c, margin, y, box_w, "6. Приглашение на демонстрацию", fonts, ORANGE)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions(
        c,
        margin + 5 * mm,
        y,
        box_w - 10 * mm,
        [
            "Давайте я просто покажу, как это выглядит. За 15–20 минут пройдёмся по онлайн-записи, календарю и бесплатному старту. Когда удобнее: завтра до обеда или после? Или во вторник / в среду?",
            "Если «надо подумать»: понимаю. Поэтому и предлагаю короткую демонстрацию — особенно бесплатный «Старт». Давайте 15 минут на [день]?",
            "Фиксация: куда отправить ссылку — почта или мессенджер? [дата], [время], демонстрация «Блиц-запись». Напомню за день и за час.",
        ],
        fonts,
    )

    c.setFillColor(TEAL)
    c.setFont(bold, 8)
    c.drawCentredString(cx, 13 * mm, "↓ возражения и завершение — на следующей странице")
    c.showPage()

    # ========== PAGE 3 ==========
    page_head(c, page_w, page_h, fonts, 3, total, cont=True)
    y = page_h - 26 * mm
    draw_circle(c, cx, y - 5, 5, "↓", fonts, GRAY)
    y -= 10
    draw_arrow_v(c, cx, y, y - 6 * mm)
    y -= 6 * mm

    y = draw_level(
        c, margin, y, box_w, "7. Если демонстрацию сейчас не ставят", fonts, NAVY
    )
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions(
        c,
        margin + 5 * mm,
        y,
        box_w - 10 * mm,
        [
            "Хорошо, настаивать не буду. Зафиксирую, что важно: [боль]. Могу скинуть описание тарифов — там есть бесплатный «Старт» — и созвониться [день / время]. Какая почта удобнее?",
        ],
        fonts,
    )
    draw_arrow_v(c, cx, y, y - 7 * mm)
    y -= 7 * mm

    y = draw_level(c, margin, y, box_w, "8. Отработка возражений → снова к демонстрации", fonts, ORANGE)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_branch_table(
        c,
        margin + 3 * mm,
        y,
        box_w - 6 * mm,
        [
            (
                "«Некогда»",
                "15 минут демонстрации в удобное время. Утро или вечер?",
                "→ слот на демо",
            ),
            (
                "«Всё устраивает»",
                "Можно начать бесплатно и сравнить с тем, как сейчас принимают запись.",
                "→ демо",
            ),
            (
                "«Уже есть сервис»",
                "Не спорим. У нас можно с 0 ₽ и только нужные модули. Покажу разницу.",
                "→ демо",
            ),
            (
                "«Дорого / платно»",
                "«Старт» — 0 ₽/мес. Полный тариф — от 1 900 ₽/мес при оплате за год.",
                "→ демо / КП",
            ),
            (
                "«Нет сайта»",
                "Сайт не обязателен: 2ГИС, карты, соцсети, ссылка в мессенджере.",
                "→ демо",
            ),
            (
                "«Пришлите на почту»",
                "Отправлю. Что больнее — звонки, неявки или цена? И сразу 15 минут на просмотр.",
                "→ письмо + слот",
            ),
            (
                "«Надо посоветоваться»",
                "Пригласим партнёра/администратора сразу на демонстрацию.",
                "→ демо вдвоём",
            ),
        ],
        fonts,
    )
    draw_arrow_v(c, cx, y, y - 8 * mm)
    y -= 8 * mm

    y = draw_level(c, margin, y, box_w, "9. Фиксация результата звонка", fonts, NAVY)
    draw_arrow_v(c, cx, y, y - 4 * mm)
    y -= 4 * mm
    y = draw_questions(
        c,
        margin + 5 * mm,
        y,
        box_w - 10 * mm,
        [
            "Есть следующий шаг: демонстрация / перезвон / коммерческое предложение?",
            "Зафиксированы почта или мессенджер, роль ЛПР, дата?",
            "Комментарий сохранён в CRM?",
        ],
        fonts,
        "Что проверить",
    )
    draw_arrow_v(c, cx, y, y - 10 * mm)
    y -= 10 * mm

    draw_circle(c, cx, y - 8, 8, "К", fonts, NAVY)
    y -= 8 + 12
    c.setFillColor(GRAY)
    c.setFont(body, 8)
    c.drawCentredString(cx, y, "конец звонка")

    y -= 12 * mm
    c.setFillColor(NAVY)
    c.setFont(bold, 8)
    c.drawCentredString(cx, y, "Цепочка уровней:")
    y -= 5 * mm
    c.setFillColor(DARK)
    c.setFont(body, 7.2)
    chain = "Н → 1 Контакт → 2 Рассказ → 3 Потребность → 4 Презентация → 5 Тарифы → 6 Демонстрация → 7/8 Возражения → 9 Фиксация → К"
    for line in wrap(chain, body, 7.2, box_w - 8 * mm):
        c.drawCentredString(cx, y, line)
        y -= 10

    c.showPage()
    c.save()
    shutil.copy(OUT, ASCII_OUT)
    print("Saved", OUT)
    print("Saved", ASCII_OUT)


if __name__ == "__main__":
    build()
