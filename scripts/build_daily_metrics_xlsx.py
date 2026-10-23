#!/usr/bin/env python3
"""Сборка Excel: ежедневные метрики воронки (месяц нарастающим итогом)."""

from __future__ import annotations

from calendar import monthrange
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUTPUT = Path(__file__).resolve().parents[1] / "Ежедневные_метрики_звонки_и_воронка.xlsx"

BASES = ("холодная база салонов красоты", "база АБУ")

WEEKDAY_RU = {
    0: "пн",
    1: "вт",
    2: "ср",
    3: "чт",
    4: "пт",
}

HEADERS = [
    "Дата",
    "День",
    "База",
    "1. Число звонков\n(месяц нарастающим)",
    "2. Получен номер ЛПР /\nСостоялся контакт с ЛПР",
    "3. Квалификация пройдена /\nВыявлен интерес / Отправлено КП",
    "4. Получена ОС от ЛПР",
    "5. Назначена встреча\nна текущий месяц",
    "6. Проведено\nпрезентаций",
    "7. Тестовый доступ\nподключен",
    "8. Онбординг пройден партнером /\nСписок услуг, прайс, вводные",
    "9. Личный кабинет настроен\nи подготовлен к работе",
    "10. Ссылка на запись\nопубликована",
    "Формат\n2/3/4/5/6/7/8/9/10/11",
]

METRIC_FILLS = [
    PatternFill("solid", fgColor="1F4E79"),
    PatternFill("solid", fgColor="2E75B6"),
    PatternFill("solid", fgColor="2E75B6"),
    PatternFill("solid", fgColor="2E75B6"),
    PatternFill("solid", fgColor="548235"),
    PatternFill("solid", fgColor="548235"),
    PatternFill("solid", fgColor="C65911"),
    PatternFill("solid", fgColor="C65911"),
    PatternFill("solid", fgColor="7030A0"),
    PatternFill("solid", fgColor="7030A0"),
]


def weekdays_in_month(year: int, month: int) -> list[date]:
    """Пн–пт с 1-го числа месяца."""
    result = []
    for day in range(1, monthrange(year, month)[1] + 1):
        d = date(year, month, day)
        if d.weekday() < 5:  # 0=пн … 4=пт
            result.append(d)
    return result


def slash_format_formula(row: int, first_col: str = "D", last_col: str = "M") -> str:
    """Автосборка D/E/F/... как 2/3/4/5/6/7/8/9/10/11."""
    letters = [get_column_letter(c) for c in range(4, 14)]  # D..M
    parts = [f'TEXT(N({letter}{row}),"0")' for letter in letters]
    joined = '&"/"&'.join(parts)
    return f'=IF(COUNTA({first_col}{row}:{last_col}{row})=0,"",{joined})'


def build(year: int = 2026, month: int = 10) -> Path:
    wb = Workbook()
    month_names = {
        1: "январь",
        2: "февраль",
        3: "март",
        4: "апрель",
        5: "май",
        6: "июнь",
        7: "июль",
        8: "август",
        9: "сентябрь",
        10: "октябрь",
        11: "ноябрь",
        12: "декабрь",
    }
    sheet_name = f"{month_names[month]} {year}"
    ws = wb.active
    ws.title = sheet_name

    # Служебный лист со списком баз (для надёжного выпадающего списка)
    lists = wb.create_sheet("_списки")
    lists["A1"] = "База"
    for i, base in enumerate(BASES, 2):
        lists.cell(i, 1, base)
    lists.sheet_state = "hidden"
    wb.defined_names.add(DefinedName(name="СписокБаз", attr_text=f"_списки!$A$2:$A${1 + len(BASES)}"))

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    title_font = Font(name="Calibri", bold=True, size=16, color="1F4E79")
    hint_font = Font(name="Calibri", italic=True, size=10, color="666666")
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    input_fill = PatternFill("solid", fgColor="FFF2CC")
    alt_fill = PatternFill("solid", fgColor="F2F2F2")
    red_fill = PatternFill("solid", fgColor="FF6B6B")
    red_font = Font(name="Calibri", bold=True, color="9C0006")
    ok_fill = PatternFill("solid", fgColor="C6EFCE")
    summary_fill = PatternFill("solid", fgColor="D6EAF8")
    locked_fill = PatternFill("solid", fgColor="E7E6E6")

    ws.merge_cells("A1:N1")
    ws["A1"] = f"Ежедневные метрики продаж — {sheet_name} (месяц нарастающим итогом)"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(vertical="center")

    ws.merge_cells("A2:N2")
    ws["A2"] = (
        "Только рабочие дни (пн–пт). Сб и вс не заполняются. "
        "В колонке «База» выберите значение из раскрывающегося списка. "
        "Метрики — нарастающий итог с 1-го числа. "
        "Колонка «Формат» собирается автоматически: звонки/ЛПР/квал/ОС/встреча/през/тест/онборд/ЛК/ссылка."
    )
    ws["A2"].font = hint_font
    ws["A2"].alignment = left_align

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 40
    ws.row_dimensions[3].height = 8
    ws.row_dimensions[4].height = 58

    for col, text in enumerate(HEADERS, 1):
        cell = ws.cell(4, col, text)
        cell.font = header_font
        cell.alignment = center
        cell.border = thin
        if col <= 3:
            cell.fill = header_fill
        elif 4 <= col <= 13:
            cell.fill = METRIC_FILLS[col - 4]
        else:
            cell.fill = PatternFill("solid", fgColor="833C0C")

    widths = {
        "A": 12,
        "B": 6,
        "C": 32,
        "D": 14,
        "E": 16,
        "F": 18,
        "G": 14,
        "H": 14,
        "I": 12,
        "J": 12,
        "K": 18,
        "L": 16,
        "M": 14,
        "N": 28,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Выпадающий список базы (ссылка на именованный диапазон)
    dv = DataValidation(
        type="list",
        formula1="=СписокБаз",
        allow_blank=True,
        showDropDown=False,  # False = показывать стрелку списка
        showErrorMessage=True,
        showInputMessage=True,
    )
    dv.error = "Выберите базу из раскрывающегося списка"
    dv.errorTitle = "База"
    dv.prompt = "Нажмите стрелку и выберите базу"
    dv.promptTitle = "Выбор базы"
    ws.add_data_validation(dv)

    days = weekdays_in_month(year, month)
    start_row = 5
    for i, d in enumerate(days):
        row = start_row + i

        date_cell = ws.cell(row, 1, d)
        date_cell.number_format = "DD.MM.YYYY"
        date_cell.alignment = center
        date_cell.border = thin
        if i % 2 == 1:
            date_cell.fill = alt_fill

        day_cell = ws.cell(row, 2, WEEKDAY_RU[d.weekday()])
        day_cell.alignment = center
        day_cell.border = thin
        if i % 2 == 1:
            day_cell.fill = alt_fill

        # База — пустая, только выбор из списка
        base_cell = ws.cell(row, 3, None)
        base_cell.alignment = center
        base_cell.border = thin
        base_cell.fill = input_fill

        for col in range(4, 14):
            cell = ws.cell(row, col)
            cell.border = thin
            cell.alignment = center
            cell.fill = input_fill
            cell.number_format = "0"

        fmt_cell = ws.cell(row, 14, slash_format_formula(row))
        fmt_cell.border = thin
        fmt_cell.alignment = center
        fmt_cell.fill = summary_fill
        fmt_cell.font = Font(name="Calibri", bold=True, size=10)

    end_row = start_row + len(days) - 1
    dv.add(f"C{start_row}:C{end_row}")

    ws.conditional_formatting.add(
        f"D{start_row}:D{end_row}",
        CellIsRule(operator="lessThan", formula=["20"], fill=red_fill, font=red_font),
    )
    ws.conditional_formatting.add(
        f"D{start_row}:D{end_row}",
        CellIsRule(operator="greaterThanOrEqual", formula=["20"], fill=ok_fill),
    )

    # Демо на ближайший будний день к 13.10.2026 (вт 13.10)
    demo_date = date(year, month, 13)
    if demo_date.weekday() < 5 and demo_date in days:
        demo_row = start_row + days.index(demo_date)
        ws.cell(demo_row, 3, BASES[0])
        for col, val in enumerate([22, 4, 2, 1, 1, 0, 0, 0, 0, 0], 4):
            ws.cell(demo_row, col, val)

    # Итоги
    sum_row = end_row + 2
    ws.merge_cells(f"A{sum_row}:C{sum_row}")
    ws.cell(sum_row, 1, "ИТОГО по базе (макс. нарастающий за месяц)").font = Font(
        bold=True, size=12, color="1F4E79"
    )

    ws.cell(sum_row + 1, 1, "База").font = header_font
    ws.cell(sum_row + 1, 1).fill = header_fill
    ws.cell(sum_row + 1, 1).border = thin
    for col in (2, 3):
        ws.cell(sum_row + 1, col).fill = header_fill
        ws.cell(sum_row + 1, col).border = thin

    for col, text in enumerate(HEADERS[3:13], 4):
        cell = ws.cell(sum_row + 1, col, text.split("\n")[0])
        cell.font = Font(bold=True, color="FFFFFF", size=9)
        cell.fill = METRIC_FILLS[col - 4]
        cell.alignment = center
        cell.border = thin

    for i, base in enumerate(BASES):
        r = sum_row + 2 + i
        ws.merge_cells(f"A{r}:C{r}")
        ws.cell(r, 1, base).border = thin
        ws.cell(r, 1).alignment = left_align
        ws.cell(r, 1).fill = PatternFill("solid", fgColor="E2EFDA" if i == 0 else "FCE4D6")
        for col in range(4, 14):
            letter = get_column_letter(col)
            formula = (
                f"=IFERROR(MAXIFS({letter}${start_row}:{letter}${end_row},"
                f"$C${start_row}:$C${end_row},A{r}),0)"
            )
            cell = ws.cell(r, col, formula)
            cell.border = thin
            cell.alignment = center
            cell.number_format = "0"

    legend_row = sum_row + 5
    ws.cell(legend_row, 1, "Легенда:").font = Font(bold=True)
    ws.cell(legend_row + 1, 1, "Жёлтые ячейки").fill = input_fill
    ws.cell(legend_row + 1, 2, "— ввод: база (список) и метрики нарастающим итогом")
    ws.cell(legend_row + 2, 1, "Красный").fill = red_fill
    ws.cell(legend_row + 2, 2, "— число звонков меньше 20")
    ws.cell(legend_row + 3, 1, "Зелёный").fill = ok_fill
    ws.cell(legend_row + 3, 2, "— число звонков 20 и более")
    ws.cell(legend_row + 4, 1, "Формат").fill = summary_fill
    ws.cell(legend_row + 4, 2, "— считается автоматически через /")
    ws.cell(legend_row + 5, 1, "Сб/вс").fill = locked_fill
    ws.cell(legend_row + 5, 2, "— в таблице нет: базы в выходные не обрабатывают")

    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:N{end_row}"
    ws.print_title_rows = "4:4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    wb.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Сборка ежедневных метрик Excel")
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--month", type=int, default=10)
    args = parser.parse_args()
    path = build(args.year, args.month)
    print(f"Saved {path}")
