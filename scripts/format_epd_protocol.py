#!/usr/bin/env python3
"""Insert EPD plan block into original protocol, keeping native formatting."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
SRC_NAME = "Протокол встреч по проекту ЭПД  ЦП-ЦКС-ЦСВ.xlsx"
OUT_PATH = ROOT / SRC_NAME

# Original palette from the file
GRAY_HDR = "ECECEC"
GREEN_HDR = "8BC34A"
GREEN_SOFT = "E2EFD9"
WHITE = "FFFFFF"
YELLOW = "FFF2CC"
ORANGE_SOFT = "FCE4D6"

thin = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

font_section = Font(name="Calibri", bold=True, size=13)
font_hdr = Font(name="Calibri", bold=True, size=11)
font_cell = Font(name="Calibri", size=12)
font_small = Font(name="Calibri", size=10, italic=True, color="666666")
font_metric = Font(name="Calibri", bold=True, size=14)

fill_gray = PatternFill("solid", fgColor=GRAY_HDR)
fill_green = PatternFill("solid", fgColor=GREEN_HDR)
fill_green_soft = PatternFill("solid", fgColor=GREEN_SOFT)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_yellow = PatternFill("solid", fgColor=YELLOW)
fill_orange_soft = PatternFill("solid", fgColor=ORANGE_SOFT)

align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_lt = Alignment(horizontal="left", vertical="top", wrap_text=True)


def fix_workbook(src: Path, dst: Path) -> None:
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                text = data.decode("utf-8")
                text = text.replace('style="solid"', 'style="thin"')
                text = re.sub(r'\s*style="none"', "", text)
                text = re.sub(r'rgb="none"', 'rgb="FF000000"', text)
                data = text.encode("utf-8")
            zout.writestr(item, data)


def style_cell(cell, *, value=None, font=None, fill=None, alignment=None, border=thin, num=None):
    if value is not None:
        cell.value = value
    if font is not None:
        cell.font = font
    if fill is not None:
        cell.fill = fill
    if alignment is not None:
        cell.alignment = alignment
    if border is not None:
        cell.border = border
    if num is not None:
        cell.number_format = num


def paint_range(ws, r1, r2, c1, c2, **kwargs):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            style_cell(ws.cell(r, c), **kwargs)


def unmerge_overlapping(ws, min_row, max_row, min_col=1, max_col=20):
    for m in list(ws.merged_cells.ranges):
        if m.max_row < min_row or m.min_row > max_row:
            continue
        if m.max_col < min_col or m.min_col > max_col:
            continue
        try:
            ws.unmerge_cells(str(m))
        except Exception:
            pass


def clear_cells(ws, r1, r2, c1=1, c2=7):
    from openpyxl.cell.cell import MergedCell

    unmerge_overlapping(ws, r1, r2, c1, c2)
    for row in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row, c)
            if isinstance(cell, MergedCell):
                continue
            cell.value = None
            cell.fill = PatternFill()
            cell.border = Border()
            cell.font = Font()
            cell.alignment = Alignment()


def merge_header(ws, range_str, text, fill, font=font_section):
    ws.merge_cells(range_str)
    top = range_str.split(":")[0]
    style_cell(ws[top], value=text, font=font, fill=fill, alignment=align_c)
    start, end = range_str.split(":")
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    sc, sr = coordinate_from_string(start)
    ec, er = coordinate_from_string(end)
    for r in range(sr, er + 1):
        for c in range(column_index_from_string(sc), column_index_from_string(ec) + 1):
            cell = ws.cell(r, c)
            cell.fill = fill
            cell.border = thin
            if cell.coordinate != top:
                cell.font = font


def insert_plan_block(ws, start_row: int = 3) -> int:
    """Insert plan + period tables at start_row. Returns number of rows inserted."""
    n = 14
    ws.insert_rows(start_row, amount=n)
    r = start_row

    # Ensure side columns wide enough
    ws.column_dimensions["I"].width = max(ws.column_dimensions["I"].width or 12, 14)
    ws.column_dimensions["J"].width = max(ws.column_dimensions["J"].width or 12, 12)
    ws.column_dimensions["K"].width = max(ws.column_dimensions["K"].width or 12, 28)

    # Section title
    merge_header(ws, f"A{r}:G{r}", "План по ЭПД", fill_gray)
    ws.row_dimensions[r].height = 22

    # Side section title (same row)
    merge_header(ws, f"I{r}:K{r}", "Соотношение к прошлому периоду", fill_gray, font=font_hdr)

    # Summary labels
    r = start_row + 1
    labels = [
        (f"A{r}:B{r}", "План"),
        (f"C{r}:D{r}", "Текущее значение"),
        (f"E{r}:F{r}", "Итого до плана"),
        (f"G{r}", "% выполнения плана"),
    ]
    for rng, text in labels:
        if ":" in rng:
            ws.merge_cells(rng)
            top = rng.split(":")[0]
        else:
            top = rng
        style_cell(ws[top], value=text, font=font_hdr, fill=fill_green, alignment=align_c)
        if ":" in rng:
            from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
            sc, sr = coordinate_from_string(rng.split(":")[0])
            ec, er = coordinate_from_string(rng.split(":")[1])
            for c in range(column_index_from_string(sc), column_index_from_string(ec) + 1):
                style_cell(ws.cell(sr, c), font=font_hdr, fill=fill_green, alignment=align_c)

    for col, text in [("I", "Дата"), ("J", "Клиенты"), ("K", "Изменение")]:
        style_cell(ws[f"{col}{r}"], value=text, font=font_hdr, fill=fill_green, alignment=align_c)

    # Summary values — fact cells will be at start_row+5 .. +7 (E column) → absolute later
    # We'll set formulas after we know absolute rows
    r = start_row + 2
    plan_row = r
    # Placeholders; formulas patched below with absolute refs
    ws.merge_cells(f"A{r}:B{r}")
    style_cell(ws[f"A{r}"], value=500, font=font_metric, fill=fill_white, alignment=align_c, num="0")
    style_cell(ws[f"B{r}"], fill=fill_white)

    ws.merge_cells(f"C{r}:D{r}")
    # current = sum of facts — set after dept rows known
    style_cell(ws[f"C{r}"], value=None, font=font_metric, fill=fill_green_soft, alignment=align_c, num="0")
    style_cell(ws[f"D{r}"], fill=fill_green_soft)

    ws.merge_cells(f"E{r}:F{r}")
    style_cell(ws[f"E{r}"], value=None, font=font_metric, fill=fill_white, alignment=align_c, num="0")
    style_cell(ws[f"F{r}"], fill=fill_white)

    style_cell(ws[f"G{r}"], value=None, font=font_metric, fill=fill_yellow, alignment=align_c, num="0.0%")

    ws.row_dimensions[r].height = 28

    # Period first rows (aligned with summary)
    style_cell(ws[f"I{r}"], value="27.08.2026", font=font_cell, fill=fill_white, alignment=align_c)
    style_cell(ws[f"J{r}"], value=None, font=font_cell, fill=fill_green_soft, alignment=align_c, num="0")  # =current
    style_cell(ws[f"K{r}"], value="—", font=font_cell, fill=fill_white, alignment=align_c)

    # Hint
    r = start_row + 3
    ws.merge_cells(f"A{r}:G{r}")
    style_cell(
        ws[f"A{r}"],
        value="«Текущее значение» = сумма ячеек «Факт» по отделам (ниже). Жёлтые ячейки — для ввода.",
        font=font_small,
        fill=fill_gray,
        alignment=align_l,
        border=thin,
    )
    for c in range(2, 8):
        style_cell(ws.cell(r, c), fill=fill_gray)

    # Period example rows
    style_cell(ws[f"I{r}"], value="03.09.2026", font=font_cell, fill=fill_white, alignment=align_c)
    style_cell(ws[f"J{r}"], value=200, font=font_cell, fill=fill_yellow, alignment=align_c, num="0")
    style_cell(ws[f"K{r}"], value=None, font=font_cell, fill=fill_white, alignment=align_l)

    r = start_row + 4
    merge_header(ws, f"A{r}:G{r}", "План на МПП в месяц — по отделам", fill_gray, font=font_hdr)

    style_cell(ws[f"I{r}"], value="10.09.2026", font=font_cell, fill=fill_white, alignment=align_c)
    style_cell(ws[f"J{r}"], value=250, font=font_cell, fill=fill_yellow, alignment=align_c, num="0")
    style_cell(ws[f"K{r}"], value=None, font=font_cell, fill=fill_white, alignment=align_l)

    # Department headers
    r = start_row + 5
    dept_hdr_row = r
    ws.merge_cells(f"A{r}:B{r}")
    style_cell(ws[f"A{r}"], value="Отдел / группа", font=font_hdr, fill=fill_green, alignment=align_c)
    style_cell(ws[f"B{r}"], fill=fill_green)
    ws.merge_cells(f"C{r}:D{r}")
    style_cell(ws[f"C{r}"], value="План на МПП в мес", font=font_hdr, fill=fill_green, alignment=align_c)
    style_cell(ws[f"D{r}"], fill=fill_green)
    style_cell(ws[f"E{r}"], value="Факт", font=font_hdr, fill=fill_green, alignment=align_c)
    ws.merge_cells(f"F{r}:G{r}")
    style_cell(ws[f"F{r}"], value="% выполнения плана", font=font_hdr, fill=fill_green, alignment=align_c)
    style_cell(ws[f"G{r}"], fill=fill_green)

    # Extra empty period rows with formulas
    for pr in range(start_row + 5, start_row + 10):
        for col in "IJK":
            style_cell(ws[f"{col}{pr}"], font=font_cell, fill=fill_yellow if col == "J" else fill_white, alignment=align_c)
        ws[f"J{pr}"].number_format = "0"

    departments = [
        ("Группа «Новые деньги»", 75, 90),
        ("Группа продуктового запуска", 25, 40),
        ("ЦКС", None, 20),
    ]
    fact_rows = []
    for i, (name, plan, fact) in enumerate(departments):
        row = start_row + 6 + i
        fact_rows.append(row)
        ws.merge_cells(f"A{row}:B{row}")
        style_cell(ws[f"A{row}"], value=name, font=font_cell, fill=fill_white, alignment=align_l)
        style_cell(ws[f"B{row}"], fill=fill_white)
        ws.merge_cells(f"C{row}:D{row}")
        style_cell(
            ws[f"C{row}"],
            value=plan,
            font=font_cell,
            fill=fill_white if plan is not None else fill_gray,
            alignment=align_c,
            num="0",
        )
        style_cell(ws[f"D{row}"], fill=fill_white if plan is not None else fill_gray)
        style_cell(ws[f"E{row}"], value=fact, font=font_cell, fill=fill_yellow, alignment=align_c, num="0")
        ws.merge_cells(f"F{row}:G{row}")
        style_cell(
            ws[f"F{row}"],
            value=f'=IF(OR(C{row}="",C{row}=0),"—",E{row}/C{row})',
            font=font_cell,
            fill=fill_green_soft,
            alignment=align_c,
            num="0.0%",
        )
        style_cell(ws[f"G{row}"], fill=fill_green_soft)
        ws.row_dimensions[row].height = 20

    # Spacer row
    spacer = start_row + 9
    ws.row_dimensions[spacer].height = 10

    # Wire summary formulas
    facts = "+".join(f"E{x}" for x in fact_rows)
    style_cell(ws[f"C{plan_row}"], value=f"={facts}", font=font_metric, fill=fill_green_soft, alignment=align_c, num="0")
    style_cell(ws[f"E{plan_row}"], value=f"=A{plan_row}-C{plan_row}", font=font_metric, fill=fill_white, alignment=align_c, num="0")
    style_cell(
        ws[f"G{plan_row}"],
        value=f'=IF(A{plan_row}=0,"—",C{plan_row}/A{plan_row})',
        font=font_metric,
        fill=fill_yellow,
        alignment=align_c,
        num="0.0%",
    )
    style_cell(ws[f"J{plan_row}"], value=f"=C{plan_row}", font=font_cell, fill=fill_green_soft, alignment=align_c, num="0")

    # Period change formulas for rows plan_row+1 .. plan_row+7 (relative)
    # Row start_row+3 (03.09), start_row+4 (10.09), then empty rows start_row+5..
    period_value_rows = list(range(plan_row, start_row + 10))  # includes baseline
    for i, prow in enumerate(period_value_rows):
        if i == 0:
            ws[f"K{prow}"].value = "—"
            continue
        prev = period_value_rows[i - 1]
        ws[f"K{prow}"].value = (
            f'=IF(OR(J{prev}="",J{prev}=0,J{prow}=""),"—",'
            f'IF(J{prow}>=J{prev},"увеличилось на "&TEXT((J{prow}-J{prev})/J{prev},"0.0%"),'
            f'"уменьшилось на "&TEXT((J{prev}-J{prow})/J{prev},"0.0%")))'
        )
        ws[f"K{prow}"].alignment = align_l
        ws[f"K{prow}"].font = font_cell
        ws[f"K{prow}"].border = thin

    # Note under period
    note_r = start_row + 10
    ws.merge_cells(f"I{note_r}:K{note_r}")
    style_cell(
        ws[f"I{note_r}"],
        value="Строка 27.08 подтягивает текущее значение. Новые даты — в жёлтые ячейки «Клиенты».",
        font=font_small,
        fill=fill_gray,
        alignment=align_l,
    )
    for c in range(10, 12):
        style_cell(ws.cell(note_r, c), fill=fill_gray)

    # Empty spacer before original content
    # rows start_row .. start_row+13 used (14 rows): 3..16 if start=3
    # start_row+0 title
    # +1 labels
    # +2 values
    # +3 hint
    # +4 dept section title
    # +5 dept headers
    # +6..+8 depts
    # +9 spacer
    # +10 period note (only I-K) / left empty
    # +11,+12,+13 spare → keep empty so original content has breathing room
    return n


def find_row(ws, text_prefix: str, max_row: int = 200) -> int | None:
    for r in range(1, max_row + 1):
        v = ws.cell(r, 1).value
        if v and str(v).strip().startswith(text_prefix):
            return r
    return None


def expand_decisions(ws):
    """Add extracted tasks into «2. Решения текущей встречи», original style."""
    header = find_row(ws, "2. Решения текущей встречи")
    if header is None:
        return

    # Table header is next row
    th = header + 1
    # Existing first data row
    first_data = header + 2

    # Collect existing tasks in B column until empty stretch / next section
    next_section = find_row(ws, "3. Обсуждаемые темы")
    existing = []
    r = first_data
    while r < (next_section or 999):
        b = ws.cell(r, 2).value
        if b:
            existing.append(str(b).strip())
        r += 1
        # stop if we've gone past used area of decisions (before section 3)
        if next_section and r >= next_section:
            break

    new_tasks = [
        ("Разобраться с показателями охвата сервисом наших клиентов 1С-ЭПД из таблицы А. Дворяк",
         "Корнева, Степанова", "до след. встречи", "в работе",
         "9000 ед. в охвате — уточнить: подключение или предоплаченные пакеты? Цифры расходятся с Калугой."),
        ("Уточнить у Суворовой статус Элиттрейд; предложить акцию 3 мес. Доки для захода к контрагентам",
         "Блохина", "до след. встречи", "в работе",
         "Крупные производители задают правила → возможность выхода на контрагентов."),
        ("Подключить Антона Гуркова к процессу изучения привязки клиента (Доки.Логистика / ЦКС)",
         "Гурков А.", "до след. встречи", "взял на себя",
         "Задача О. Ремез — прописать подробнее после пересмотра встречи."),
        ("Выяснить у Калуги: можно ли отследить активность клиента в сервисе Доки",
         "Дивакова М.", "до след. встречи", "в работе", ""),
        ("Уточнить у Никитченко: считается ли Доки в охват сервисами ЭПД",
         "Дворяк", "до след. встречи", "в работе", "Саша спросит."),
        ("Подготовить со СМАК материалы для прогрева клиента на период демо-доступа",
         "СМАК", "к 01.10", "в работе",
         "С 01.10 демо = 14 дней; до 01.10 — 3 месяца бесплатно."),
        ("Подготовить калькулятор пополнения титулов для менеджеров ЦКС",
         "Вика / Маша", "до след. встречи", "в работе",
         "Обсуждали накануне лектория с 1С."),
        ("Подготовить текст рассылки «проверьте контрагента и впишите в тандем» — сроки и текст",
         "Блохина", "до след. встречи", "в работе", ""),
        ("Обзвонить крупных клиентов (от 2000 титулов / генераторы трафика) и предложить ЭПД ТАНДЕМ",
         "ЦП / ЦКС", "сентябрь", "в работе",
         "Начать с крупных. Корп. сегмент продавал ЭПД даже в малом объёме."),
        ("Обзвон подключённых клиентов: оценка удовлетворённости + проекты ЦАС (Доки)",
         "Дворяк / Лазарчук", "сентябрь", "в работе",
         "Задача по клиентам ЦАС (Доки)."),
        ("Взять неуспешные сделки ЭПД по корп. сегменту и предложить Доки в 1-ю неделю сентября",
         "Юлиана", "1-я нед. сентября", "в работе",
         "Клиенты сбытового офиса уходят на СБИС — разобрать причины."),
        ("Проверить в базе подключение клиентов к каналу ЭПД (МАКС / ТГ / ВК) как альт. канал",
         "ЦП / ЦКС", "до след. встречи", "в работе",
         "Связать со стратегией и целями."),
        ("Зафиксировать план 500 клиентов до конца 2026: 50 крупных / 300 средних / 150 мелких",
         "Все участники", "принято", "принято",
         "Приоритет — количество клиентов, не сумма выручки на старте."),
    ]

    # Replace decision rows with structured task list (includes original + extracted)
    filtered = new_tasks

    # How many empty rows available before section 3?
    available_start = first_data
    end = (next_section or available_start + 20) - 1

    needed = len(filtered)
    existing_slots = end - available_start + 1
    if existing_slots < needed:
        ws.insert_rows(next_section, amount=needed - existing_slots)
        next_section = find_row(ws, "3. Обсуждаемые темы")
        end = next_section - 1

    clear_cells(ws, available_start, end, 1, 7)

    # Rebuild table header (may have been damaged by row shifts/merges)
    clear_cells(ws, th, th, 1, 7)
    headers = ["№", "Задача", "Лидер", "Срок", "Статус*", "Комментарии. Выводы", ""]
    for c, text in enumerate(headers, 1):
        style_cell(ws.cell(th, c), value=text or None, font=font_hdr, fill=fill_green, alignment=align_c, num="General")
    ws.merge_cells(start_row=th, start_column=6, end_row=th, end_column=7)

    for i, (task, leader, deadline, status, comment) in enumerate(filtered):
        row = available_start + i
        # Reset formats inherited from original date columns
        for c in range(1, 8):
            cell = ws.cell(row, c)
            cell.number_format = "General"
        style_cell(ws.cell(row, 1), value=i + 1, font=font_cell, fill=fill_white, alignment=align_c, num="0")
        style_cell(ws.cell(row, 2), value=task, font=font_cell, fill=fill_white, alignment=align_l, num="General")
        style_cell(ws.cell(row, 3), value=leader, font=font_cell, fill=fill_white, alignment=align_c, num="General")
        style_cell(ws.cell(row, 4), value=deadline, font=font_cell, fill=fill_white, alignment=align_c, num="@")
        st_fill = fill_green_soft if status in ("принято", "выполнено", "взял на себя") else fill_orange_soft
        style_cell(ws.cell(row, 5), value=status, font=font_cell, fill=st_fill, alignment=align_c, num="General")
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=7)
        style_cell(ws.cell(row, 6), value=comment, font=font_cell, fill=fill_white, alignment=align_lt, num="General")
        style_cell(ws.cell(row, 7), fill=fill_white)
        ws.row_dimensions[row].height = 36


def format_discussion_blocks(ws):
    """Group discussion topics into blocks, keep original green/gray style."""
    header = find_row(ws, "3. Обсуждаемые темы")
    if header is None:
        return

    th = header + 1  # column headers
    last = header + 1
    for r in range(header + 2, ws.max_row + 5):
        vals = []
        for c in (1, 2, 4):
            cell = ws.cell(r, c)
            from openpyxl.cell.cell import MergedCell
            if not isinstance(cell, MergedCell) and cell.value:
                vals.append(cell.value)
        if vals:
            last = r

    clear_cells(ws, th, max(last + 5, th + 40), 1, 9)

    # Column headers — original style
    ws.merge_cells(start_row=th, start_column=1, end_row=th, end_column=3)
    style_cell(ws.cell(th, 1), value="Обсуждаемая тема (проблема)", font=font_section, fill=fill_green, alignment=align_c)
    for c in (2, 3):
        style_cell(ws.cell(th, c), fill=fill_green)
    ws.merge_cells(start_row=th, start_column=4, end_row=th, end_column=7)
    style_cell(ws.cell(th, 4), value="Выводы и решения", font=font_section, fill=fill_green, alignment=align_c)
    for c in (5, 6, 7):
        style_cell(ws.cell(th, c), fill=fill_green)

    blocks = [
        ("Блок: Акция и ЦСВ", [
            ("Обсудить с ЦСВ трудозатраты на акцию по ЭПД.",
             "Оксана: уточнить адресата вопроса. Акцию продлили до 30.09."),
        ]),
        ("Блок: Каналы МАКС / ТГ / ВК", [
            ("Канал ЭПД в МАКС, ТГ, ВК. Можем ли прорабатывать этих клиентов силами ЦП и ЦКС? Связать со стратегией; альтернативный канал привлечения.",
             "Проверить из базы: подключён клиент к каналу или нет. Идея — просить менеджеров подключать клиентов (вопрос мотивации клиента открыт)."),
        ]),
        ("Блок: Стратегия и сегментация", [
            ("Заход в Элиттрейд (таблица А. Дворяк) — выход к контрагентам. Предложение Гуркова: сегментировать охваченных на крупных / средних / мелких.",
             "Производители (крупные) задают правила для перевозчиков → те для получателей. Лиза уточнит у Суворовой статус Элиттрейд; можно предложить акцию 3 мес. Доки."),
            ("К признаку «крупный» добавить: является ли клиент генератором трафика (привлечение контрагентов).",
             "По соотношению базы (крупные / средние / мелкие) пришли к соглашению по процентам."),
            ("Приоритет — количество привлечённых клиентов, а не сумма заработка. Выстроить цепочку «закрепления» и почкования с одного пакета на новых клиентов.",
             "ПЛАН: 500 клиентов до конца 2026 → 50 крупных + 300 средних + 150 мелких."),
        ]),
        ("Блок: Доки.Логистика и охват", [
            ("Задача по Доки.Логистика (Оксана Ремез) — прописать подробнее после пересмотра встречи.",
             "Подключить Антона Гуркова к изучению привязки клиента. Антон взял на себя."),
            ("Можем ли отследить активность клиента в сервисе Доки? (вопрос Калуге)",
             "Маша Дивакова выяснит."),
            ("Считается ли Доки в охват сервисами ЭПД?",
             "Вопрос Никитченко — Саша спросит. Цифры расходятся с Калугой; проблемы с закреплением."),
            ("Все проданные ЭПД: центр компетенции изучает — всё ли продано, подключились ли.",
             "Трафик не видим → не квалифицируем активность пользования пакетом. 1С обещают данные в сентябре."),
        ]),
        ("Блок: Демо и прогрев (СМАК)", [
            ("Со СМАК подготовить материалы для прогрева клиента во время демо-доступа.",
             "С 01.10 — 14 дней демо; до 01.10 — 3 месяца бесплатно."),
        ]),
        ("Блок: Сопровождение МПП / ЦКС", [
            ("План для МПП: платные / бесплатные клиенты; «выхаживать» может другой человек (идея — Антон Гурков).",
             "Если есть ИТС — менеджер ЦКС смотрит динамику расхода ЭПД/Доки. Трафика пока нет; можно писать в Калугу с перечнем ИНН. Риск: кто выхаживает — текущие менеджеры или новый человек?"),
            ("Клиенту продают минимальный пакет («пока не понятно»). Кто ведёт дальше при продлении? Получают ли ЦКС сигнал о дозакупке титулов?",
             "По логике — ЦКС. Нужен калькулятор — Вика с Машей обсуждали накануне лектория с 1С."),
            ("Ожидание: ЦКС категоризирует клиента по кол-ву титулов / УАТ и своевременно пополняет баланс.",
             "Зафиксировано как рабочий процесс."),
        ]),
        ("Блок: Тандем и обзвоны", [
            ("Рассылка по клиентам: «проверьте своего контрагента и впишите в тандем».",
             "Лиза — когда и какой текст."),
            ("Начать с крупных клиентов: прозвонить и предложить подключить контрагента в ЭПД ТАНДЕМ.",
             "Крупные = от 2000 титулов. Корп. клиенты продавали ЭПД даже в малом объёме."),
            ("Обзвон подключённых клиентов: оценка удовлетворённости + проекты ЦАС (Доки).",
             "Задача на Сашу Дворяк / Лазарчук — клиенты ЦАС."),
        ]),
        ("Блок: Корп. сегмент / сбытовой офис", [
            ("Почему клиенты сбытового офиса не переходят в ЭПД / уходят на СБИС?",
             "Взять неуспешные сделки ЭПД по корп. сегменту и предложить Доки в 1-ю неделю сентября (Юлиана)."),
            ("Юля: предлагать всем клиентам бесплатно Доки (дублирующий сервис) на 3 мес.; KPI менеджерам на сентябрь.",
             "Зафиксировано как предложение."),
        ]),
    ]

    row = th + 1
    for block_name, items in blocks:
        # Block banner — same gray as section headers
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        style_cell(ws.cell(row, 1), value=block_name, font=font_hdr, fill=fill_gray, alignment=align_l)
        for c in range(2, 8):
            style_cell(ws.cell(row, c), fill=fill_gray)
        ws.row_dimensions[row].height = 18
        row += 1

        for topic, conclusion in items:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
            style_cell(ws.cell(row, 1), value=topic, font=font_cell, fill=fill_white, alignment=align_lt)
            for c in (2, 3):
                style_cell(ws.cell(row, c), fill=fill_white)
            ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=7)
            # Highlight conclusions that contain tasks
            concl_fill = fill_orange_soft if any(
                k in conclusion.lower() for k in ("взял", "выяснит", "спросит", "лиза", "юлиана", "план:", "задача")
            ) else fill_white
            style_cell(ws.cell(row, 4), value=conclusion, font=font_cell, fill=concl_fill, alignment=align_lt)
            for c in (5, 6, 7):
                style_cell(ws.cell(row, c), fill=concl_fill)
            ws.row_dimensions[row].height = 48
            row += 1

    # Legend
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    style_cell(
        ws.cell(row, 1),
        value="Подсветка выводов: строки с поручениями/задачами выделены мягким оранжевым. Поставленные задачи продублированы в блоке «2. Решения текущей встречи».",
        font=font_small,
        fill=fill_gray,
        alignment=align_l,
    )
    for c in range(2, 8):
        style_cell(ws.cell(row, c), fill=fill_gray)


def ensure_title(ws):
    # Keep A1:G1 if present; set title if empty
    if ws["A1"].value in (None, ""):
        # may be merged
        style_cell(
            ws["A1"],
            value="ВСТРЕЧА ЦП-ЦКС-ЦСВ-СМАК",
            font=font_section,
            fill=fill_gray,
            alignment=align_c,
        )
        for c in range(2, 8):
            style_cell(ws.cell(1, c), fill=fill_gray)


def main():
    # Always start from origin/main version to preserve original formatting
    import subprocess
    src_bytes = subprocess.check_output(
        ["git", "show", f"origin/main:{SRC_NAME}"],
        cwd=ROOT,
    )
    raw = Path("/tmp/protocol_from_main.xlsx")
    raw.write_bytes(src_bytes)
    fixed = Path("/tmp/protocol_work.xlsx")
    fix_workbook(raw, fixed)

    wb = load_workbook(fixed)
    # Drop template sheet if present from earlier attempts (won't be in main)
    if "Шаблон (новая встреча)" in wb.sheetnames:
        del wb["Шаблон (новая встреча)"]

    ws = wb["27.08.2026"]
    ensure_title(ws)

    # Insert plan above advertising campaign (original row 3)
    insert_plan_block(ws, start_row=3)

    # Expand decisions + format discussion (rows already shifted)
    expand_decisions(ws)
    format_discussion_blocks(ws)

    wb.save(OUT_PATH)
    print(f"Saved {OUT_PATH}")

    # Sanity
    wb2 = load_workbook(OUT_PATH)
    ws2 = wb2["27.08.2026"]
    print("A3", ws2["A3"].value)
    print("A17", ws2["A17"].value)  # should be ad campaign after +14
    # find sections
    for prefix in ["План по ЭПД", "Показатели рекламной", "1. Мониторинг", "2. Решения", "3. Обсуждаемые"]:
        row = None
        for r in range(1, 120):
            v = ws2.cell(r, 1).value
            if v and str(v).startswith(prefix[:12] if len(prefix) > 12 else prefix):
                # looser
                pass
            if v and prefix.split()[0] in str(v):
                print(f"found '{prefix}' ~ at row {r}: {v}")
                break
    print("C5 formula", ws2["C5"].value)
    print("E9/E10/E11 facts", ws2["E9"].value, ws2["E10"].value, ws2["E11"].value)


if __name__ == "__main__":
    main()
