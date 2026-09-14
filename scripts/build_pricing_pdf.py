#!/usr/bin/env python3
"""
Visualize the proposed pricing-block scheme for 1С:Кабинет сотрудника (КЭДО).

Scheme:
  Formula → ШАГ 1 (лицензии) → ШАГ 2 (пакеты × формат внедрения) → пояснения → пример

Prices from the client screenshot.
"""

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
STEP_BG = HexColor("#E8F6FC")

FONT_REG = "NotoSans"
FONT_BOLD = "NotoSans-Bold"
pdfmetrics.registerFont(TTFont(FONT_REG, "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont(FONT_BOLD, "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf"))

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

# Screenshot prices collapsed where identical
PACKAGE_TIERS = [
    # label, start_admin, start_full, support_admin, support_full
    ("до 200", 7000, 15000, 30000, 60000),
    ("300–400", 7500, 17500, 35000, 75000),
    ("500", 8000, 20000, 40000, 90000),
]


def fmt(n: int) -> str:
    return f"{n:,}".replace(",", " ") + " ₽"


class AccentBar(Flowable):
    def __init__(self, width, height=3.2, color=BLUE):
        super().__init__()
        self.width = width
        self.height = height
        self.color = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.roundRect(0, 0, self.width, self.height, 1.4, fill=1, stroke=0)


class StepBadge(Flowable):
    def __init__(self, text: str):
        super().__init__()
        self.text = text
        self.width = 48
        self.height = 16

    def draw(self):
        c = self.canv
        c.setFillColor(BLUE)
        c.roundRect(0, 0, self.width, self.height, 8, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont(FONT_BOLD, 7.5)
        c.drawCentredString(self.width / 2, 4.5, self.text)


def make_styles():
    return {
        "eyebrow": ParagraphStyle(
            "eyebrow", fontName=FONT_REG, fontSize=9, textColor=GRAY, leading=11, spaceAfter=1
        ),
        "title": ParagraphStyle(
            "title", fontName=FONT_BOLD, fontSize=18, textColor=NEAR_BLACK, leading=22, spaceAfter=1
        ),
        "formula": ParagraphStyle(
            "formula",
            fontName=FONT_BOLD,
            fontSize=10.5,
            textColor=BLUE_DARK,
            leading=13,
            alignment=TA_CENTER,
        ),
        "step_title": ParagraphStyle(
            "step_title", fontName=FONT_BOLD, fontSize=10.5, textColor=NEAR_BLACK, leading=13
        ),
        "step_sub": ParagraphStyle(
            "step_sub", fontName=FONT_REG, fontSize=8, textColor=GRAY, leading=10
        ),
        "th": ParagraphStyle(
            "th", fontName=FONT_BOLD, fontSize=7.5, textColor=white, leading=9, alignment=TA_CENTER
        ),
        "th_dark": ParagraphStyle(
            "th_dark",
            fontName=FONT_BOLD,
            fontSize=7.5,
            textColor=NEAR_BLACK,
            leading=9,
            alignment=TA_CENTER,
        ),
        "cell": ParagraphStyle(
            "cell",
            fontName=FONT_REG,
            fontSize=8,
            textColor=NEAR_BLACK,
            leading=10,
            alignment=TA_CENTER,
        ),
        "cell_b": ParagraphStyle(
            "cell_b",
            fontName=FONT_BOLD,
            fontSize=8.5,
            textColor=NEAR_BLACK,
            leading=10,
            alignment=TA_CENTER,
        ),
        "pkg": ParagraphStyle(
            "pkg",
            fontName=FONT_BOLD,
            fontSize=8.5,
            textColor=NEAR_BLACK,
            leading=11,
            alignment=TA_LEFT,
        ),
        "note": ParagraphStyle(
            "note", fontName=FONT_REG, fontSize=7.5, textColor=NEAR_BLACK, leading=10
        ),
        "muted": ParagraphStyle(
            "muted", fontName=FONT_REG, fontSize=7, textColor=GRAY, leading=9
        ),
        "example": ParagraphStyle(
            "example", fontName=FONT_REG, fontSize=8.5, textColor=NEAR_BLACK, leading=11
        ),
        "footer": ParagraphStyle(
            "footer",
            fontName=FONT_REG,
            fontSize=7,
            textColor=GRAY,
            leading=9,
            alignment=TA_CENTER,
        ),
    }


def step_header(s, badge: str, title: str, subtitle: str, width: float):
    inner = Table(
        [
            [
                StepBadge(badge),
                [Paragraph(title, s["step_title"]), Paragraph(subtitle, s["step_sub"])],
            ]
        ],
        colWidths=[54, width - 70],
    )
    inner.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ("LEFTPADDING", (1, 0), (1, 0), 8),
            ]
        )
    )
    wrap = Table([[inner]], colWidths=[width])
    wrap.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), STEP_BG),
                ("BOX", (0, 0), (-1, -1), 0.8, BLUE),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return wrap


def license_table(s, width: float):
    n = len(LICENSE)
    label_w = 30 * mm
    col_w = (width - label_w) / n
    head = [Paragraph("Сотрудники", s["th"])] + [Paragraph(str(c), s["th"]) for c, _ in LICENSE]
    row = [Paragraph("Стоимость за год", s["th_dark"])] + [
        Paragraph(fmt(p), s["cell_b"]) for _, p in LICENSE
    ]
    t = Table([head, row], colWidths=[label_w] + [col_w] * n)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), LICENSE_HEAD),
                ("BACKGROUND", (0, 1), (0, 1), SOFT),
                ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                ("BOX", (0, 0), (-1, -1), 1, HexColor("#BDBDBD")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 1),
                ("RIGHTPADDING", (0, 0), (-1, -1), 1),
            ]
        )
    )
    return t


def package_matrix(s, width: float):
    """Packages as columns (transposed): Старт | Старт + сопровождение."""
    format_w = 48 * mm
    tier_w = 24 * mm
    half = (width - format_w - tier_w) / 2

    data = [
        [
            Paragraph("Формат внедрения", s["th"]),
            Paragraph("Сотрудники", s["th"]),
            Paragraph("Старт КЭДО", s["th_dark"]),
            Paragraph("Старт + сопровождение КЭДО", s["th_dark"]),
        ]
    ]

    # Rows: your admin × tiers, then full implementation × tiers
    for i, (label, sa, _sf, sua, _suf) in enumerate(PACKAGE_TIERS):
        data.append(
            [
                Paragraph("С вашим администратором", s["pkg"]) if i == 0 else Paragraph("", s["pkg"]),
                Paragraph(label, s["cell_b"]),
                Paragraph(fmt(sa), s["cell"]),
                Paragraph(fmt(sua), s["cell"]),
            ]
        )
    for i, (label, _sa, sf, _sua, suf) in enumerate(PACKAGE_TIERS):
        data.append(
            [
                Paragraph("С нашим полным внедрением", s["pkg"]) if i == 0 else Paragraph("", s["pkg"]),
                Paragraph(label, s["cell_b"]),
                Paragraph(fmt(sf), s["cell"]),
                Paragraph(fmt(suf), s["cell"]),
            ]
        )

    t = Table(data, colWidths=[format_w, tier_w, half, half])
    t.setStyle(
        TableStyle(
            [
                ("SPAN", (0, 1), (0, 3)),
                ("SPAN", (0, 4), (0, 6)),
                ("BACKGROUND", (0, 0), (1, 0), LICENSE_HEAD),
                ("BACKGROUND", (2, 0), (2, 0), YELLOW_HEAD),
                ("BACKGROUND", (3, 0), (3, 0), CYAN_HEAD),
                ("BACKGROUND", (0, 1), (1, 3), SOFT),
                ("BACKGROUND", (0, 4), (1, 6), SOFT),
                ("BACKGROUND", (2, 1), (2, 6), YELLOW_BG),
                ("BACKGROUND", (3, 1), (3, 6), CYAN_BG),
                ("BACKGROUND", (2, 2), (2, 2), white),
                ("BACKGROUND", (3, 2), (3, 2), white),
                ("BACKGROUND", (2, 5), (2, 5), white),
                ("BACKGROUND", (3, 5), (3, 5), white),
                ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                ("BOX", (0, 0), (-1, -1), 1, HexColor("#BDBDBD")),
                ("LINEABOVE", (0, 4), (-1, 4), 1.0, LINE),
                ("LINEBEFORE", (2, 0), (2, -1), 1.1, YELLOW),
                ("LINEBEFORE", (3, 0), (3, -1), 1.1, CYAN),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return t


def explanations(s, width: float):
    gap = 3 * mm
    half = (width - gap) / 2
    left = Paragraph(
        "<b>С вашим администратором</b><br/>"
        "Клиент выделяет администратора 1С: настройка сервиса и ЭП, "
        "базовая поддержка сотрудников. Подрядчик помогает на старте.",
        s["note"],
    )
    right = Paragraph(
        "<b>С нашим полным внедрением</b><br/>"
        "Подрядчик выполняет все работы по запуску. "
        "Свой администратор со стороны клиента не нужен.",
        s["note"],
    )
    box = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, -1), SOFT),
            ("BOX", (0, 0), (-1, -1), 0.7, LINE),
            ("LINEBEFORE", (0, 0), (0, 0), 2.5, BLUE),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ]
    )
    a = Table([[left]], colWidths=[half])
    a.setStyle(box)
    b = Table([[right]], colWidths=[half])
    b.setStyle(box)
    row = Table([[a, "", b]], colWidths=[half, gap, half])
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


def example_block(s, width: float):
    text = (
        "<b>Пример:</b> 100 сотрудников + «Старт КЭДО» с вашим администратором = "
        f"{fmt(33600)} (лицензия) + {fmt(7000)} (пакет) = "
        f"<font color='#1B7FAF'><b>{fmt(40600)}</b></font>"
    )
    t = Table([[Paragraph(text, s["example"])]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), STEP_BG),
                ("BOX", (0, 0), (-1, -1), 1, BLUE),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return t


def build():
    page = landscape(A4)
    margin = 11 * mm
    width = page[0] - 2 * margin
    s = make_styles()

    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=page,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=8 * mm,
        bottomMargin=7 * mm,
        title="Стоимость пакетов — схема блока цен",
        author="ГК Форус",
    )

    story = [
        Paragraph("1С:Кабинет сотрудника · КЭДО", s["eyebrow"]),
        Paragraph("Стоимость пакетов", s["title"]),
        AccentBar(width),
        Spacer(1, 5),
    ]

    formula = Table(
        [
            [
                Paragraph(
                    "Итого = лицензия на сервис (за год) + выбранный пакет запуска КЭДО",
                    s["formula"],
                )
            ]
        ],
        colWidths=[width],
    )
    formula.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), STEP_BG),
                ("BOX", (0, 0), (-1, -1), 1.2, BLUE),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    story += [formula, Spacer(1, 6)]

    story += [
        step_header(
            s,
            "ШАГ 1",
            "Лицензия на сервис",
            "Выберите стоимость по числу сотрудников (за год, без НДС)",
            width,
        ),
        Spacer(1, 4),
        license_table(s, width),
        Spacer(1, 2),
        Paragraph(
            "Лицензии без НДС (ПО в реестре российского ПО). Типовой облачный функционал, одно юридическое лицо.",
            s["muted"],
        ),
        Spacer(1, 6),
        step_header(
            s,
            "ШАГ 2",
            "Пакет запуска КЭДО",
            "Выберите пакет и формат внедрения. Цены зависят от числа сотрудников.",
            width,
        ),
        Spacer(1, 4),
        package_matrix(s, width),
        Spacer(1, 5),
        explanations(s, width),
        Spacer(1, 5),
        example_block(s, width),
        Spacer(1, 4),
        Paragraph(
            "ГК Форус · 1С:Кабинет сотрудника · актуальные цены уточняйте у менеджера",
            s["footer"],
        ),
    ]

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.build(story)
    print(f"Written: {OUT}")


if __name__ == "__main__":
    build()
