#!/usr/bin/env python3
"""Отчёт по «1С:Кабинет сотрудника» за сентябрь 2026 (жёлто-белый стиль Форус)."""

from __future__ import annotations

import re
import shutil
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "КС аналитика (4).xlsx"
OUT_DIR = ROOT / "presentation"
OUT = OUT_DIR / "Отчёт_Кабинет_сотрудника_сентябрь_2026.pptx"
OUT_ASCII = OUT_DIR / "KS_Report_September_2026.pptx"
BRAND = OUT_DIR / "assets" / "brand"

YELLOW = RGBColor(0xFE, 0xCF, 0x68)
YELLOW_DEEP = RGBColor(0xE8, 0xB4, 0x3A)
CREAM = RGBColor(0xFF, 0xF8, 0xE8)
CREAM2 = RGBColor(0xFF, 0xF1, 0xCC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x1A, 0x1A, 0x1A)
SLATE = RGBColor(0x39, 0x5F, 0x75)
MUTED = RGBColor(0x5C, 0x5C, 0x5C)
SOFT_BG = RGBColor(0xFF, 0xFC, 0xF6)
GREEN = RGBColor(0x2E, 0x7D, 0x4F)

FONT = "Verdana"
SW, SH = Inches(13.333), Inches(7.5)
SEP_START, SEP_END = date(2026, 9, 1), date(2026, 9, 30)
TOTAL = 8

# Фактические KPI менеджеров за сентябрь (от заказчика)
MGR = {
    "Оглоблина Софья": {
        "book": 6,
        "demo": 5,
        "invoice": 0,
        "pilot": 0,
        "sale": 1,
    },
    "Кургузов Данил": {
        "book": 3,
        "demo": 3,
        "invoice": 0,
        "pilot": 0,
        "sale": 0,
    },
}


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


def add_multitext(slide, left, top, width, height, paragraphs, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    set_anchor(tf, anchor)
    for i, item in enumerate(paragraphs):
        text, size, bold, color = item[:4]
        space_before = item[4] if len(item) > 4 else 0
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(space_before)
        set_run(p.add_run(), text, size, bold, color)
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
    # Use blank layout if available, else last layout
    layout = prs.slide_layouts[-1]
    s = prs.slides.add_slide(layout)
    # paint white background
    add_rect(s, 0, 0, SW, SH, WHITE)
    return s


def add_footer(slide, page: int):
    add_rect(slide, 0, emu(7.28), SW, emu(0.22), YELLOW)
    add_textbox(
        slide,
        emu(0.5),
        emu(7.08),
        emu(9.2),
        emu(0.22),
        "ГК Форус  ·  рабочая группа  ·  1С:Кабинет сотрудника",
        9,
        False,
        MUTED,
    )
    add_textbox(
        slide,
        emu(11.4),
        emu(7.08),
        emu(1.4),
        emu(0.22),
        f"{page} / {TOTAL}",
        9,
        False,
        MUTED,
        PP_ALIGN.RIGHT,
    )


def add_header(slide, kicker: str, title: str):
    add_picture(slide, "logo-forus.png", emu(0.5), emu(0.22), height=emu(0.38))
    add_picture(slide, "wave_tr.png", emu(10.55), emu(-0.05), width=emu(2.9))
    add_textbox(slide, emu(0.5), emu(0.72), emu(12.2), emu(0.28), kicker, 11, True, SLATE)
    add_textbox(slide, emu(0.5), emu(0.98), emu(12.2), emu(0.55), title, 22, True, INK)
    add_rect(slide, emu(0.5), emu(1.52), emu(1.35), emu(0.07), YELLOW)


def parse_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
            try:
                return datetime.strptime(v.strip()[:10], fmt).date()
            except ValueError:
                pass
    return None


def num(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def norm_base(s):
    if not s:
        return "Без базы"
    raw = str(s).strip()
    key = re.sub(r"\s+", " ", raw.lower()).strip()
    rules = [
        (r"зуп\s*проф|база\s*зуп\s*проф", "ЗУП ПРОФ"),
        (r"зуп\s*корп", "ЗУП КОРП"),
        (r"зуп\s*базов", "ЗУП базовые"),
        (r"база\s*зуп|^зуп$", "База ЗУП"),
        (r"холодк[аи].*транспорт|база\s*транспорт|^транспорт", "Холодка транспорт"),
        (r"канбан", "База канбана"),
        (r"отвал", "База отвалов / отвал КС"),
        (r"вебинар", "База вебинаров"),
        (r"общепит", "База общепит"),
        (r"производ", "База производство"),
        (r"^пилот", "Пилот"),
        (r"^сделк", "Сделки (донабор)"),
        (r"упал", "Упавшие"),
        (r"база\s*кп|^кп$", "База КП"),
        (r"вахт", "База вахта"),
    ]
    for pat, name in rules:
        if re.search(pat, key):
            return name
    return raw.strip().title()


def pct(a, b):
    if not b:
        return "—"
    return f"{a / b * 100:.1f}%".replace(".", ",")


def fmt_int(n):
    return f"{int(round(n)):,}".replace(",", " ")


def load_september():
    if not XLSX.exists():
        raise SystemExit(f"Excel not found: {XLSX}")
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    sheets = {"КС Соня": "Оглоблина Софья", "КС Данил": "Кургузов Данил"}
    by_base = defaultdict(
        lambda: {
            "calls": 0.0,
            "ok": 0.0,
            "deal": 0.0,
            "book": 0.0,
            "demo": 0.0,
            "trial": 0.0,
            "sale": 0.0,
            "mgrs": set(),
        }
    )
    by_mgr_calls = defaultdict(float)
    for sheet, mgr in sheets.items():
        ws = wb[sheet]
        for r in range(2, (ws.max_row or 1) + 1):
            d = parse_date(ws.cell(r, 1).value)
            if not d or not (SEP_START <= d <= SEP_END):
                continue
            base = norm_base(ws.cell(r, 3).value)
            row = {
                "calls": num(ws.cell(r, 4).value),
                "ok": num(ws.cell(r, 5).value),
                "deal": num(ws.cell(r, 6).value),
                "book": num(ws.cell(r, 7).value),
                "demo": num(ws.cell(r, 8).value),
                "trial": num(ws.cell(r, 9).value),
                "sale": num(ws.cell(r, 10).value),
            }
            for k, v in row.items():
                by_base[base][k] += v
            by_base[base]["mgrs"].add(mgr.split()[0])
            by_mgr_calls[mgr] += row["calls"]
    # drop empty
    by_base = {k: v for k, v in by_base.items() if v["calls"] > 0 or v["demo"] > 0 or v["book"] > 0}
    return by_base, by_mgr_calls


def slide_title(prs, totals):
    s = blank_slide(prs)
    add_rect(s, 0, 0, emu(0.22), SH, YELLOW)
    add_picture(s, "logo-forus.png", emu(0.7), emu(0.45), height=emu(0.55))
    add_picture(s, "wave_tr.png", emu(9.8), emu(-0.1), width=emu(3.7))
    add_textbox(
        s,
        emu(0.7),
        emu(1.7),
        emu(11.5),
        emu(0.35),
        "Рабочая группа  ·  сентябрь 2026",
        14,
        True,
        SLATE,
    )
    add_textbox(
        s,
        emu(0.7),
        emu(2.15),
        emu(12.0),
        emu(1.5),
        "Отчёт по проработке\n«1С:Кабинет сотрудника»",
        34,
        True,
        INK,
    )
    add_rect(s, emu(0.7), emu(3.85), emu(2.0), emu(0.1), YELLOW)
    add_textbox(
        s,
        emu(0.7),
        emu(4.15),
        emu(11.0),
        emu(0.9),
        "Базы, конверсии проработки и результаты менеджеров.\n"
        "Кейсы продажи и пилота + выводы: что двигает клиента к покупке.",
        15,
        False,
        MUTED,
    )
    cards = [
        ("Период", "01.09 — 29.09.2026", CREAM),
        ("Менеджеры", "Софья · Данил", CREAM),
        ("Звонков", fmt_int(totals["calls"]), YELLOW),
        ("Продаж", str(totals["sale"]), YELLOW),
    ]
    for i, (k, v, fill) in enumerate(cards):
        left = emu(0.7) + emu(i * 3.1)
        add_card(s, left, emu(5.5), emu(2.95), emu(1.15), fill, 0.1)
        add_textbox(s, left + emu(0.18), emu(5.62), emu(2.6), emu(0.3), k, 11, False, SLATE if fill == CREAM else INK)
        add_textbox(s, left + emu(0.18), emu(5.95), emu(2.6), emu(0.5), v, 18, True, INK)
    add_picture(s, "wave_bl.png", emu(-0.15), emu(6.55), width=emu(3.2))
    return s


def slide_summary(prs, totals):
    s = blank_slide(prs)
    add_header(s, "1. Итоги месяца", "Сентябрь в цифрах")
    kpis = [
        ("Звонков", fmt_int(totals["calls"]), "прозвон по базам"),
        ("Успешных", fmt_int(totals["ok"]), f"дозвон/диалог · {pct(totals['ok'], totals['calls'])}"),
        ("В сделку", fmt_int(totals["deal"]), f"интерес · {pct(totals['deal'], totals['calls'])}"),
        ("Записей", str(totals["book"]), "на демо / подключение"),
        ("Демо", str(totals["demo"]), "фактически проведено"),
        ("Продаж", str(totals["sale"]), "закрыто в сентябре"),
    ]
    for i, (title, value, sub) in enumerate(kpis):
        col, row = i % 3, i // 3
        left = emu(0.5) + emu(col * 4.2)
        top = emu(1.85) + emu(row * 2.15)
        add_card(s, left, top, emu(4.0), emu(1.95), CREAM, 0.1)
        add_rect(s, left, top, emu(4.0), emu(0.1), YELLOW)
        add_textbox(s, left + emu(0.25), top + emu(0.3), emu(3.5), emu(0.35), title, 13, True, SLATE)
        add_textbox(s, left + emu(0.25), top + emu(0.7), emu(3.5), emu(0.6), value, 32, True, INK)
        add_textbox(s, left + emu(0.25), top + emu(1.4), emu(3.5), emu(0.35), sub, 12, False, MUTED)
    add_footer(s, 2)
    return s


def slide_bases(prs, bases):
    s = blank_slide(prs)
    add_header(s, "2. Базы", "Что прорабатывали и какие конверсии")
    add_textbox(
        s,
        emu(0.5),
        emu(1.65),
        emu(12.3),
        emu(0.35),
        "Данные из «КС аналитика (4).xlsx» за сентябрь. Конверсии: успешный / звонок → сделка / звонок → запись / сделка → демо / запись.",
        12,
        False,
        MUTED,
    )
    # table header
    headers = ["База", "Звонки", "Успешн.", "Сделка", "Запись", "Демо", "Усп/зв", "Сд/зв", "Зап/сд", "Демо/зап"]
    widths = [2.8, 0.95, 0.95, 0.9, 0.9, 0.8, 0.95, 0.9, 0.95, 1.05]
    left0, top0, row_h = emu(0.5), emu(2.1), emu(0.38)
    x = left0
    for h, w in zip(headers, widths):
        add_rect(s, x, top0, emu(w), row_h, YELLOW)
        add_textbox(s, x, top0, emu(w), row_h, h, 10, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        x += emu(w)

    rows = sorted(bases.items(), key=lambda kv: -kv[1]["calls"])
    # limit to bases with meaningful volume
    rows = [(b, v) for b, v in rows if v["calls"] >= 10 or v["demo"] > 0 or v["book"] > 0][:11]
    for i, (base, v) in enumerate(rows):
        y = top0 + row_h * (i + 1)
        fill = CREAM if i % 2 == 0 else SOFT_BG
        vals = [
            base,
            fmt_int(v["calls"]),
            fmt_int(v["ok"]),
            fmt_int(v["deal"]),
            fmt_int(v["book"]),
            fmt_int(v["demo"]),
            pct(v["ok"], v["calls"]),
            pct(v["deal"], v["calls"]),
            pct(v["book"], v["deal"]),
            pct(v["demo"], v["book"]),
        ]
        x = left0
        for j, (val, w) in enumerate(zip(vals, widths)):
            add_rect(s, x, y, emu(w), row_h, fill)
            align = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            pad = emu(0.08) if j == 0 else 0
            add_textbox(
                s,
                x + pad,
                y,
                emu(w) - pad,
                row_h,
                val,
                10 if j else 11,
                j == 0,
                INK,
                align,
                MSO_ANCHOR.MIDDLE,
            )
            x += emu(w)

    add_textbox(
        s,
        emu(0.5),
        emu(6.55),
        emu(12.3),
        emu(0.4),
        "Лидеры по объёму: холодка транспорт, ЗУП ПРОФ, ЗУП базовые. По воронке к демо лучше отработали ЗУП ПРОФ и канбан/донабор сделок.",
        12,
        False,
        SLATE,
    )
    add_footer(s, 3)
    return s


def slide_base_insights(prs, bases):
    s = blank_slide(prs)
    add_header(s, "3. Фокус по базам", "Где сентябрь дал движение к демо и продаже")
    cards = [
        (
            "ЗУП ПРОФ",
            bases.get("ЗУП ПРОФ", {}),
            "Клиенты ЗУП, ранее сегментированные по ИТС. Отсюда вышла продажа ООО «СМ».",
            YELLOW,
        ),
        (
            "Холодка транспорт",
            bases.get("Холодка транспорт", {}),
            "Самый большой объём прозвона. Много диалогов, узкое горлышко в запись на демо.",
            CREAM,
        ),
        (
            "ЗУП базовые",
            bases.get("ЗУП базовые", {}),
            "Средний объём, есть записи и демо — база продолжает давать встречи.",
            CREAM,
        ),
        (
            "Канбан + сделки",
            {
                "calls": bases.get("База канбана", {}).get("calls", 0)
                + bases.get("Сделки (донабор)", {}).get("calls", 0),
                "ok": bases.get("База канбана", {}).get("ok", 0)
                + bases.get("Сделки (донабор)", {}).get("ok", 0),
                "deal": bases.get("База канбана", {}).get("deal", 0)
                + bases.get("Сделки (донабор)", {}).get("deal", 0),
                "book": bases.get("База канбана", {}).get("book", 0)
                + bases.get("Сделки (донабор)", {}).get("book", 0),
                "demo": bases.get("База канбана", {}).get("demo", 0)
                + bases.get("Сделки (донабор)", {}).get("demo", 0),
            },
            "Донабор и прогрев уже тёплых контактов: демо без большого прозвона «вхолостую».",
            CREAM,
        ),
    ]
    for i, (title, v, note, fill) in enumerate(cards):
        col, row = i % 2, i // 2
        left = emu(0.5) + emu(col * 6.4)
        top = emu(1.8) + emu(row * 2.5)
        add_card(s, left, top, emu(6.15), emu(2.3), fill, 0.1)
        add_textbox(s, left + emu(0.25), top + emu(0.2), emu(5.7), emu(0.4), title, 16, True, INK)
        line = (
            f"Звонки {fmt_int(v.get('calls', 0))}  ·  успешн. {fmt_int(v.get('ok', 0))}  ·  "
            f"сделки {fmt_int(v.get('deal', 0))}  ·  записи {fmt_int(v.get('book', 0))}  ·  "
            f"демо {fmt_int(v.get('demo', 0))}"
        )
        add_textbox(s, left + emu(0.25), top + emu(0.7), emu(5.7), emu(0.4), line, 12, True, SLATE)
        add_textbox(
            s,
            left + emu(0.25),
            top + emu(1.15),
            emu(5.7),
            emu(0.35),
            f"Конверсия в сделку: {pct(v.get('deal', 0), v.get('calls', 0))}   ·   "
            f"запись из сделки: {pct(v.get('book', 0), v.get('deal', 0))}",
            12,
            False,
            MUTED,
        )
        add_textbox(s, left + emu(0.25), top + emu(1.55), emu(5.7), emu(0.55), note, 12, False, INK)
    add_footer(s, 4)
    return s


def slide_managers(prs, mgr_calls):
    s = blank_slide(prs)
    add_header(s, "4. Менеджеры", "Эффективность и результативность за сентябрь")
    add_textbox(
        s,
        emu(0.5),
        emu(1.65),
        emu(12.3),
        emu(0.35),
        "Записи / проведено / счета / пилоты / продажи — фактические показатели сентября. Звонки — из таблицы аналитики.",
        12,
        False,
        MUTED,
    )
    people = [
        (
            "Оглоблина Софья",
            MGR["Оглоблина Софья"],
            mgr_calls.get("Оглоблина Софья", 0),
            "1 продажа в сентябре — ООО «СМ» (10 кабинетов)",
        ),
        (
            "Кургузов Данил",
            MGR["Кургузов Данил"],
            mgr_calls.get("Кургузов Данил", 0),
            "Стабильный набор демо; пилотов и продаж в сентябре пока нет",
        ),
    ]
    for i, (name, m, calls, note) in enumerate(people):
        left = emu(0.5) + emu(i * 6.4)
        add_card(s, left, emu(2.15), emu(6.15), emu(4.55), CREAM, 0.1)
        add_rect(s, left, emu(2.15), emu(6.15), emu(0.12), YELLOW)
        add_textbox(s, left + emu(0.3), emu(2.45), emu(5.5), emu(0.45), name, 18, True, INK)
        metrics = [
            ("Звонков (таблица)", fmt_int(calls)),
            ("Записей на демо", str(m["book"])),
            ("Проведено демо", str(m["demo"])),
            ("Выставлено счетов", str(m["invoice"])),
            ("Пилот (3 мес.)", str(m["pilot"])),
            ("Продаж", str(m["sale"])),
        ]
        for j, (lab, val) in enumerate(metrics):
            y = emu(3.05) + emu(j * 0.42)
            add_textbox(s, left + emu(0.35), y, emu(3.6), emu(0.38), lab, 13, False, MUTED, anchor=MSO_ANCHOR.MIDDLE)
            add_textbox(s, left + emu(4.0), y, emu(1.8), emu(0.38), val, 16, True, INK, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
        add_textbox(s, left + emu(0.3), emu(5.7), emu(5.5), emu(0.7), note, 12, False, SLATE)
    add_footer(s, 5)
    return s


def slide_sale_case(prs):
    s = blank_slide(prs)
    add_header(s, "5. Кейс продажи", "ООО «СМ» — 10 кабинетов после демо")
    # meta strip
    add_card(s, emu(0.5), emu(1.7), emu(12.3), emu(0.7), YELLOW, 0.08)
    add_textbox(
        s,
        emu(0.7),
        emu(1.78),
        emu(12.0),
        emu(0.55),
        "Клиент: ООО «СМ»  ·  ИНН 3801102232  ·  Источник: база ЗУП ПРОФ (сегмент по ИТС)  ·  Менеджер: Оглоблина Софья",
        13,
        True,
        INK,
        PP_ALIGN.LEFT,
        MSO_ANCHOR.MIDDLE,
    )

    blocks = [
        (
            "Как вышли",
            "Заход к бухгалтеру по ранее купленному ЗУП: уточнили, как идёт работа с программой, "
            "напомнили про неизрасходованные услуги при покупке. Бухгалтер о них не знала. "
            "Рассказали про бесплатную демонстрацию Кабинета сотрудника — КЭДО внутри 1С, меньше бумаги. "
            "Тема заинтересовала, но занимается кадровая служба; попросили оставить контакты "
            "(были сложности с телефонией). Отправили письмо — бухгалтер передала кадровикам, "
            "и уже кадры сами вышли на связь.",
        ),
        (
            "Работа с кадрами",
            "Кадровики написали, что хотят подключить КЭДО. Сначала перепутали с демо-периодом "
            "Смартвея и просили «попробовать 7 дней». Уточнили: доступна бесплатная демонстрация. "
            "Клиент согласился на демо.",
        ),
        (
            "Демо и решение",
            "На встрече — две сотрудницы: молодая специалистка (основной контакт) и более старшая ЛПР. "
            "Показали возможности сервиса и сценарии КЭДО. По итогам ЛПР сразу: «Давай купим 10 кабинетов». "
            "Решение подтверждено на демонстрации. Запись демо сохранена. В сентябре — оплата.",
        ),
    ]
    for i, (title, body) in enumerate(blocks):
        top = emu(2.55) + emu(i * 1.4)
        add_card(s, emu(0.5), top, emu(12.3), emu(1.28), CREAM if i % 2 == 0 else SOFT_BG, 0.08)
        add_rect(s, emu(0.5), top, emu(0.12), emu(1.28), YELLOW)
        add_textbox(s, emu(0.85), top + emu(0.12), emu(11.7), emu(0.3), title, 14, True, SLATE)
        add_textbox(s, emu(0.85), top + emu(0.42), emu(11.7), emu(0.75), body, 12, False, INK)
    add_footer(s, 6)
    return s


def slide_pilot_case(prs):
    s = blank_slide(prs)
    add_header(s, "6. Кейс пилота", "ООО «ЦЗ» (Центр зрения) — согласие на пилот")
    add_card(s, emu(0.5), emu(1.7), emu(12.3), emu(0.7), YELLOW, 0.08)
    add_textbox(
        s,
        emu(0.7),
        emu(1.78),
        emu(12.0),
        emu(0.55),
        "Клиент: ООО «ЦЗ» · Центр зрения · Долгая проработка · Есть сотрудники · Интерес к ЭДО",
        13,
        True,
        INK,
        PP_ALIGN.LEFT,
        MSO_ANCHOR.MIDDLE,
    )
    blocks = [
        (
            "Контекст",
            "Клиента обрабатывали долго. Уже был пробный период, но попользоваться сервисом "
            "по-настоящему не успели — поэтому интерес не «закрылся», а остался открытым.",
        ),
        (
            "Что цепляет",
            "В компании есть сотрудники, которым нужен кадровый контур. Нравится идея "
            "электронного документооборота: меньше бумаги, понятные процессы, автоматизация.",
        ),
        (
            "Почему пилот",
            "На пилот согласились именно потому, что на демо / в пробном периоде не успели "
            "«прожить» сервис руками. Сервис интересен — нужен спокойный срок, чтобы проверить "
            "в работе и принять решение о покупке.",
        ),
    ]
    for i, (title, body) in enumerate(blocks):
        top = emu(2.55) + emu(i * 1.35)
        add_card(s, emu(0.5), top, emu(12.3), emu(1.22), CREAM if i % 2 == 0 else SOFT_BG, 0.08)
        add_rect(s, emu(0.5), top, emu(0.12), emu(1.22), YELLOW)
        add_textbox(s, emu(0.85), top + emu(0.12), emu(11.7), emu(0.28), title, 14, True, SLATE)
        add_textbox(s, emu(0.85), top + emu(0.42), emu(11.7), emu(0.65), body, 13, False, INK)
    add_footer(s, 7)
    return s


def slide_insights(prs):
    s = blank_slide(prs)
    add_header(s, "7. Выводы", "Что двигает клиента купить Кабинет сотрудника")
    add_textbox(
        s,
        emu(0.5),
        emu(1.62),
        emu(12.3),
        emu(0.4),
        "По кейсам ООО «СМ» (продажа) и ООО «ЦЗ» (пилот). Общее: интерес к автоматизации и электронным процессам.",
        13,
        False,
        MUTED,
    )
    drivers = [
        (
            "01",
            "Уже «болит» бумага / кадры",
            "Есть сотрудники и живой кадровый контур. Клиент чувствует рутину заявлений, "
            "подписей, согласований — и ищет, как упростить.",
        ),
        (
            "02",
            "Тяга к ЭДО и автоматизации",
            "Не спорят с идеей «всё в электронном виде». Готовы слушать сервис, который "
            "встраивается в 1С и не ломает привычный учёт.",
        ),
        (
            "03",
            "Понять ценность на себе",
            "Решение зреет после живого контакта: демо у «СМ», пилот у «ЦЗ». "
            "Если не успели попробовать — оставляют дверь открытой через пилот.",
        ),
        (
            "04",
            "Правильный ЛПР внутри",
            "У «СМ» бухгалтер открыла дверь, купили кадры/ЛПР на демо. "
            "Нужно рано выяснять: кто реально решает по КЭДО и кадрам.",
        ),
        (
            "05",
            "Повод связаться без «холода»",
            "Напоминание про неизрасходованные услуги ЗУП, прошлое касание, долгая "
            "проработка — снижают барьер и дают право на разговор про Кабинет.",
        ),
        (
            "06",
            "Низкий риск следующего шага",
            "Бесплатное демо или пилот на 3 месяца — безопасный способ «пожить» в сервисе "
            "до оплаты. Это и есть мост от интереса к решению.",
        ),
    ]
    for i, (num, title, body) in enumerate(drivers):
        col, row = i % 3, i // 3
        left = emu(0.5) + emu(col * 4.2)
        top = emu(2.1) + emu(row * 2.35)
        add_card(s, left, top, emu(4.0), emu(2.2), CREAM if (col + row) % 2 == 0 else SOFT_BG, 0.1)
        add_textbox(s, left + emu(0.2), top + emu(0.15), emu(3.6), emu(0.35), num, 16, True, YELLOW_DEEP)
        add_textbox(s, left + emu(0.2), top + emu(0.5), emu(3.6), emu(0.55), title, 14, True, INK)
        add_textbox(s, left + emu(0.2), top + emu(1.1), emu(3.6), emu(0.9), body, 12, False, MUTED)
    add_footer(s, 8)
    return s


def delete_all_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        r_id = sld_id.get(qn("r:id"))
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def main():
    bases, mgr_calls = load_september()
    totals = {
        "calls": sum(v["calls"] for v in bases.values()),
        "ok": sum(v["ok"] for v in bases.values()),
        "deal": sum(v["deal"] for v in bases.values()),
        "book": MGR["Оглоблина Софья"]["book"] + MGR["Кургузов Данил"]["book"],
        "demo": MGR["Оглоблина Софья"]["demo"] + MGR["Кургузов Данил"]["demo"],
        "sale": MGR["Оглоблина Софья"]["sale"] + MGR["Кургузов Данил"]["sale"],
        "pilot": MGR["Оглоблина Софья"]["pilot"] + MGR["Кургузов Данил"]["pilot"],
    }

    # Start from empty presentation 16:9
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH
    # remove default blank if any by building on fresh
    delete_all_slides(prs)

    slide_title(prs, totals)
    slide_summary(prs, totals)
    slide_bases(prs, bases)
    slide_base_insights(prs, bases)
    slide_managers(prs, mgr_calls)
    slide_sale_case(prs)
    slide_pilot_case(prs)
    slide_insights(prs)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    shutil.copy2(OUT, OUT_ASCII)
    print(f"Saved: {OUT}")
    print(f"Copy:  {OUT_ASCII}")
    print(f"Slides: {len(prs.slides)}")
    print(
        f"Totals: calls={totals['calls']:.0f} ok={totals['ok']:.0f} deal={totals['deal']:.0f} "
        f"book={totals['book']} demo={totals['demo']} sale={totals['sale']}"
    )
    print(f"Bases: {len(bases)}")


if __name__ == "__main__":
    main()
