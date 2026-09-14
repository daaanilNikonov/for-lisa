#!/usr/bin/env python3
"""Generate a client-facing pricing PDF for 1С:Кабинет сотрудника (КЭДО)."""

from __future__ import annotations

from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Flowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "presentation" / "Стоимость_пакетов_Кабинет_сотрудника.pdf"

# ГК Форус brand + package accents from product leaflet
BLUE = HexColor("#26A6E0")
BLUE_DARK = HexColor("#1B7FAF")
NEAR_BLACK = HexColor("#1A1A1A")
GRAY = HexColor("#767677")
SOFT = HexColor("#F2F2F2")
LINE = HexColor("#D9D9D9")
YELLOW = HexColor("#E8B84A")
YELLOW_BG = HexColor("#FFF8E8")
YELLOW_HEAD = HexColor("#F0C75A")
CYAN = HexColor("#5BB8C9")
CYAN_BG = HexColor("#EEF8FA")
CYAN_HEAD = HexColor("#7BC8D6")
LICENSE_HEAD = HexColor("#4A4A4A")
ROW_ALT = HexColor("#F7F7F7")

FONT_REG = "NotoSans"
FONT_BOLD = "NotoSans-Bold"

pdfmetrics.registerFont(TTFont(FONT_REG, "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont(FONT_BOLD, "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf"))


def fmt(n: int) -> str:
    return f"{n:,}".replace(",", " ") + " ₽"


LICENSE = [
    (10, 3360),
    (25, 8400),
    (50, 16800),
    (75, 25200),
    (100, 33600),
    (200, 62400),
    (300, 96000),
    (400, 124800),
    (500, 144000),
]

# (start_admin, start_full, support_admin, support_full)
PACKAGES = {
    10: (7000, 15000, 30000, 60000),
    25: (7000, 15000, 30000, 60000),
    50: (7000, 15000, 30000, 60000),
    75: (7000, 15000, 30000, 60000),
    100: (7000, 15000, 30000, 60000),
    200: (7000, 15000, 30000, 60000),
    300: (7500, 17500, 35000, 75000),
    400: (7500, 17500, 35000, 75000),
    500: (8000, 20000, 40000, 90000),
}


class AccentBar(Flowable):
    def __init__(self, width, height=3, color=BLUE):
        super().__init__()
        self.width = width
        self.height = height
        self.color = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.roundRect(0, 0, self.width, self.height, 1.5, fill=1, stroke=0)


def styles():
    return {
        "title": ParagraphStyle(
            "title",
            fontName=FONT_BOLD,
            fontSize=22,
            textColor=NEAR_BLACK,
            leading=26,
            alignment=TA_LEFT,
            spaceAfter=2,
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            fontName=FONT_REG,
            fontSize=10,
            textColor=GRAY,
            leading=14,
            spaceAfter=6,
        ),
        "formula": ParagraphStyle(
            "formula",
            fontName=FONT_BOLD,
            fontSize=11,
            textColor=BLUE_DARK,
            leading=15,
            alignment=TA_CENTER,
        ),
        "h2": ParagraphStyle(
            "h2",
            fontName=FONT_BOLD,
            fontSize=11,
            textColor=NEAR_BLACK,
            leading=15,
            spaceBefore=0,
            spaceAfter=0,
        ),
        "how": ParagraphStyle(
            "how",
            fontName=FONT_REG,
            fontSize=9,
            textColor=GRAY,
            leading=12,
        ),
        "note": ParagraphStyle(
            "note",
            fontName=FONT_REG,
            fontSize=8,
            textColor=NEAR_BLACK,
            leading=11,
        ),
        "note_muted": ParagraphStyle(
            "note_muted",
            fontName=FONT_REG,
            fontSize=7.5,
            textColor=GRAY,
            leading=10,
        ),
        "cell": ParagraphStyle(
            "cell",
            fontName=FONT_REG,
            fontSize=8.5,
            textColor=NEAR_BLACK,
            leading=11,
            alignment=TA_CENTER,
        ),
        "cell_bold": ParagraphStyle(
            "cell_bold",
            fontName=FONT_BOLD,
            fontSize=9,
            textColor=NEAR_BLACK,
            leading=11,
            alignment=TA_CENTER,
        ),
        "th": ParagraphStyle(
            "th",
            fontName=FONT_BOLD,
            fontSize=8,
            textColor=white,
            leading=10,
            alignment=TA_CENTER,
        ),
        "th_dark": ParagraphStyle(
            "th_dark",
            fontName=FONT_BOLD,
            fontSize=7.5,
            textColor=NEAR_BLACK,
            leading=10,
            alignment=TA_CENTER,
        ),
        "example": ParagraphStyle(
            "example",
            fontName=FONT_REG,
            fontSize=9,
            textColor=NEAR_BLACK,
            leading=13,
            alignment=TA_LEFT,
        ),
        "footer": ParagraphStyle(
            "footer",
            fontName=FONT_REG,
            fontSize=7.5,
            textColor=GRAY,
            leading=10,
            alignment=TA_CENTER,
        ),
    }


def header_block(s, content_width):
    elems = [
        Paragraph("1С:Кабинет сотрудника · КЭДО", s["subtitle"]),
        Paragraph("Стоимость пакетов", s["title"]),
        AccentBar(content_width, 3.5, BLUE),
        Spacer(1, 8),
    ]
    formula = Table(
        [
            [
                Paragraph(
                    "Итого для клиента = лицензия на сервис (за год) + выбранный пакет запуска КЭДО",
                    s["formula"],
                )
            ]
        ],
        colWidths=[content_width],
    )
    formula.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), HexColor("#E8F6FC")),
                ("BOX", (0, 0), (-1, -1), 1.2, BLUE),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    how = Paragraph(
        "Как выбрать: 1) найдите строку по числу сотрудников → 2) возьмите лицензию → "
        "3) добавьте пакет «Старт» или «Старт + сопровождение» и формат внедрения.",
        s["how"],
    )
    elems += [formula, Spacer(1, 6), how, Spacer(1, 8)]
    return elems


def unified_table(s, content_width):
    """One row = one scenario: employees | license | Start×2 | Support×2."""
    col_emps = 20 * mm
    col_lic = 32 * mm
    rest = content_width - col_emps - col_lic
    col_pkg = rest / 4

    th = s["th"]
    th_d = s["th_dark"]
    empty = Paragraph("", th)

    top = [
        Paragraph("Кол-во<br/>сотрудников", th),
        Paragraph("Лицензия<br/>на сервис / год", th),
        Paragraph("Старт КЭДО", th_d),
        empty,
        Paragraph("Старт + сопровождение КЭДО", th_d),
        empty,
    ]
    sub = [
        empty,
        empty,
        Paragraph("С вашим<br/>администратором", th_d),
        Paragraph("С нашим полным<br/>внедрением", th_d),
        Paragraph("С вашим<br/>администратором", th_d),
        Paragraph("С нашим полным<br/>внедрением", th_d),
    ]

    data = [top, sub]
    for count, price in LICENSE:
        sa, sf, sua, suf = PACKAGES[count]
        data.append(
            [
                Paragraph(str(count), s["cell_bold"]),
                Paragraph(fmt(price), s["cell_bold"]),
                Paragraph(fmt(sa), s["cell"]),
                Paragraph(fmt(sf), s["cell"]),
                Paragraph(fmt(sua), s["cell"]),
                Paragraph(fmt(suf), s["cell"]),
            ]
        )

    t = Table(
        data,
        colWidths=[col_emps, col_lic, col_pkg, col_pkg, col_pkg, col_pkg],
        repeatRows=2,
    )

    style_cmds = [
        ("SPAN", (2, 0), (3, 0)),
        ("SPAN", (4, 0), (5, 0)),
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (1, 0), (1, 1)),
        ("BACKGROUND", (0, 0), (1, 1), LICENSE_HEAD),
        ("BACKGROUND", (2, 0), (3, 0), YELLOW_HEAD),
        ("BACKGROUND", (2, 1), (3, 1), YELLOW),
        ("BACKGROUND", (4, 0), (5, 0), CYAN_HEAD),
        ("BACKGROUND", (4, 1), (5, 1), CYAN),
        ("GRID", (0, 0), (-1, -1), 0.45, LINE),
        ("BOX", (0, 0), (-1, -1), 1, HexColor("#BDBDBD")),
        ("LINEBEFORE", (2, 0), (2, -1), 1.2, YELLOW),
        ("LINEBEFORE", (4, 0), (4, -1), 1.2, CYAN),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, 1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, 1), 6),
        ("TOPPADDING", (0, 2), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 2), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (2, 2), (3, -1), YELLOW_BG),
        ("BACKGROUND", (4, 2), (5, -1), CYAN_BG),
    ]

    for i in range(2, len(data)):
        if (i - 2) % 2 == 1:
            style_cmds.append(("BACKGROUND", (0, i), (1, i), ROW_ALT))

    t.setStyle(TableStyle(style_cmds))
    return t


def explanations(s, content_width):
    gap = 4 * mm
    half = (content_width - gap) / 2
    left = Paragraph(
        "<b>С вашим администратором</b><br/>"
        "Клиент выделяет администратора 1С: настройка сервиса и ЭП, "
        "базовая поддержка сотрудников. Подрядчик помогает на старте.",
        s["note"],
    )
    right = Paragraph(
        "<b>С нашим полным внедрением</b><br/>"
        "Подрядчик выполняет все работы по запуску. "
        "Выделять администратора со стороны клиента не требуется.",
        s["note"],
    )
    box_style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, -1), SOFT),
            ("BOX", (0, 0), (-1, -1), 0.8, LINE),
            ("LINEBEFORE", (0, 0), (0, 0), 3, BLUE),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ]
    )
    left_box = Table([[left]], colWidths=[half])
    left_box.setStyle(box_style)
    right_box = Table([[right]], colWidths=[half])
    right_box.setStyle(box_style)
    row = Table([[left_box, "", right_box]], colWidths=[half, gap, half])
    row.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return row


def example_block(s, content_width):
    text = (
        "<b>Пример расчёта:</b> 100 сотрудников + «Старт КЭДО» с вашим администратором&nbsp;— "
        f"{fmt(33600)} (лицензия) + {fmt(7000)} (пакет) = "
        f"<font color='#1B7FAF'><b>{fmt(40600)}</b></font>"
    )
    t = Table([[Paragraph(text, s["example"])]], colWidths=[content_width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), HexColor("#E8F6FC")),
                ("BOX", (0, 0), (-1, -1), 1, BLUE),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
            ]
        )
    )
    return t


def build():
    page = landscape(A4)
    margin = 14 * mm
    content_width = page[0] - 2 * margin
    s = styles()

    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=page,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=11 * mm,
        bottomMargin=9 * mm,
        title="Стоимость пакетов — 1С:Кабинет сотрудника",
        author="ГК Форус",
    )

    story = []
    story += header_block(s, content_width)
    story.append(unified_table(s, content_width))
    story.append(Spacer(1, 5))
    story.append(
        Paragraph(
            "Лицензии без НДС (ПО включено в реестр российского ПО). "
            "Цены — типовой функционал облачного решения для одного юридического лица.",
            s["note_muted"],
        )
    )
    story.append(Spacer(1, 8))
    story.append(
        Paragraph(
            "Формат внедрения (одинаково для обоих пакетов):",
            s["how"],
        )
    )
    story.append(Spacer(1, 4))
    story.append(explanations(s, content_width))
    story.append(Spacer(1, 8))
    story.append(example_block(s, content_width))
    story.append(Spacer(1, 6))
    story.append(
        Paragraph(
            "ГК Форус · 1С:Кабинет сотрудника · актуальные цены уточняйте у менеджера",
            s["footer"],
        )
    )

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.build(story)
    print(f"Written: {OUT}")


if __name__ == "__main__":
    build()
