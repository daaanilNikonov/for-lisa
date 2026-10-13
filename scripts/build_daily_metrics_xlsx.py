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
from openpyxl.worksheet.datavalidation import DataValidation

OUTPUT = Path(__file__).resolve().parents[1] / "Ежедневные_метрики_звонки_и_воронка.xlsx"

BASES = ("холодная база салонов красоты", "база АБУ")

HEADERS = [
    "Дата",
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

METRIC_HELPS = [
    (1, "Число звонков", "Все исходящие звонки по выбранной базе, нарастающим с начала месяца. Норма: ≥ 20 в день."),
    (2, "Получен номер ЛПР / Состоялся контакт с ЛПР", "Дозвон до лица, принимающего решение, или получение его контакта."),
    (3, "Квалификация / Интерес / КП", "Пройдена квалификация, выявлен интерес или отправлено коммерческое предложение."),
    (4, "Получена ОС от ЛПР", "Обратная связь от ЛПР по КП / предложению."),
    (5, "Назначена встреча на текущий месяц", "Встреча/демо назначена именно на текущий календарный месяц."),
    (6, "Проведено презентаций", "Фактически проведённые презентации / демо."),
    (7, "Тестовый доступ подключен", "Партнёру/клиенту выдан тестовый доступ."),
    (8, "Онбординг / услуги и прайс", "Партнёр прошёл онбординг; получены список услуг, прайс и вводные."),
    (9, "Личный кабинет настроен", "ЛК настроен и готов к работе."),
    (10, "Ссылка на запись опубликована", "Ссылка на запись опубликована (для Блиц-запись и аналогов)."),
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


def build(year: int = 2026, month: int = 10) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ежедневные метрики"

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

    ws.merge_cells("A1:M1")
    ws["A1"] = "Ежедневные метрики продаж (месяц нарастающим итогом)"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(vertical="center")

    ws.merge_cells("A2:M2")
    ws["A2"] = (
        "Заполняйте нарастающий итог с начала месяца. База — из списка. "
        "Число звонков < 20 подсвечивается красным. "
        "Колонка «Формат» собирает метрики как 2/3/4/5/6/7/8/9/10/11 автоматически."
    )
    ws["A2"].font = hint_font
    ws["A2"].alignment = left_align

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 8
    ws.row_dimensions[4].height = 55

    for col, text in enumerate(HEADERS, 1):
        cell = ws.cell(4, col, text)
        cell.font = header_font
        cell.alignment = center
        cell.border = thin
        if col <= 2:
            cell.fill = header_fill
        elif 3 <= col <= 12:
            cell.fill = METRIC_FILLS[col - 3]
        else:
            cell.fill = PatternFill("solid", fgColor="833C0C")

    widths = {
        "A": 12,
        "B": 28,
        "C": 14,
        "D": 16,
        "E": 18,
        "F": 14,
        "G": 14,
        "H": 12,
        "I": 12,
        "J": 18,
        "K": 16,
        "L": 14,
        "M": 28,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    dv = DataValidation(
        type="list",
        formula1='"холодная база салонов красоты,база АБУ"',
        allow_blank=True,
    )
    dv.error = "Выберите базу из списка"
    dv.errorTitle = "База"
    dv.prompt = "Выберите базу"
    dv.promptTitle = "База"
    ws.add_data_validation(dv)

    ws.freeze_panes = "C5"

    days_in_month = monthrange(year, month)[1]
    start_row = 5
    row = start_row
    for day in range(1, days_in_month + 1):
        for base in BASES:
            d = date(year, month, day)
            ws.cell(row, 1, d).number_format = "DD.MM.YYYY"
            ws.cell(row, 1).alignment = center
            ws.cell(row, 1).border = thin

            ws.cell(row, 2, base).alignment = center
            ws.cell(row, 2).border = thin
            ws.cell(row, 2).fill = input_fill

            for col in range(3, 13):
                cell = ws.cell(row, col)
                cell.border = thin
                cell.alignment = center
                cell.fill = input_fill
                cell.number_format = "0"

            formula = (
                f'=IF(COUNTA(C{row}:L{row})=0,"",'
                f'TEXT(IF(C{row}="",0,C{row}),"0")&"/"&'
                f'TEXT(IF(D{row}="",0,D{row}),"0")&"/"&'
                f'TEXT(IF(E{row}="",0,E{row}),"0")&"/"&'
                f'TEXT(IF(F{row}="",0,F{row}),"0")&"/"&'
                f'TEXT(IF(G{row}="",0,G{row}),"0")&"/"&'
                f'TEXT(IF(H{row}="",0,H{row}),"0")&"/"&'
                f'TEXT(IF(I{row}="",0,I{row}),"0")&"/"&'
                f'TEXT(IF(J{row}="",0,J{row}),"0")&"/"&'
                f'TEXT(IF(K{row}="",0,K{row}),"0")&"/"&'
                f'TEXT(IF(L{row}="",0,L{row}),"0"))'
            )
            fmt_cell = ws.cell(row, 13, formula)
            fmt_cell.border = thin
            fmt_cell.alignment = center
            fmt_cell.fill = summary_fill
            fmt_cell.font = Font(name="Calibri", bold=True, size=10)

            if day % 2 == 0:
                ws.cell(row, 1).fill = alt_fill
            row += 1

    end_row = row - 1
    dv.add(f"B{start_row}:B{end_row}")

    ws.conditional_formatting.add(
        f"C{start_row}:C{end_row}",
        CellIsRule(operator="lessThan", formula=["20"], fill=red_fill, font=red_font),
    )
    ws.conditional_formatting.add(
        f"C{start_row}:C{end_row}",
        CellIsRule(operator="greaterThanOrEqual", formula=["20"], fill=ok_fill),
    )

    # Demo: 13.10 — норма и ниже нормы
    demo_ok = start_row + (12 * 2)  # 13-е число, первая база
    demo_low = demo_ok + 1
    for col, val in enumerate([22, 4, 2, 1, 1, 0, 0, 0, 0, 0], 3):
        ws.cell(demo_ok, col, val)
    for col, val in enumerate([12, 1, 0, 0, 0, 0, 0, 0, 0, 0], 3):
        ws.cell(demo_low, col, val)

    sum_row = end_row + 2
    ws.merge_cells(f"A{sum_row}:B{sum_row}")
    ws.cell(sum_row, 1, "ИТОГО по базе (макс. нарастающий за месяц)").font = Font(
        bold=True, size=12, color="1F4E79"
    )

    ws.cell(sum_row + 1, 1, "База").font = header_font
    ws.cell(sum_row + 1, 1).fill = header_fill
    ws.cell(sum_row + 1, 1).border = thin
    ws.cell(sum_row + 1, 2).fill = header_fill
    ws.cell(sum_row + 1, 2).border = thin

    for col, text in enumerate(HEADERS[2:12], 3):
        cell = ws.cell(sum_row + 1, col, text.split("\n")[0])
        cell.font = Font(bold=True, color="FFFFFF", size=9)
        cell.fill = METRIC_FILLS[col - 3]
        cell.alignment = center
        cell.border = thin

    for i, base in enumerate(BASES):
        r = sum_row + 2 + i
        ws.cell(r, 1, base).border = thin
        ws.merge_cells(f"A{r}:B{r}")
        ws.cell(r, 1).alignment = left_align
        ws.cell(r, 1).fill = PatternFill("solid", fgColor="E2EFDA" if i == 0 else "FCE4D6")
        for col in range(3, 13):
            letter = get_column_letter(col)
            formula = (
                f"=IFERROR(MAXIFS({letter}${start_row}:{letter}${end_row},"
                f"$B${start_row}:$B${end_row},A{r}),0)"
            )
            cell = ws.cell(r, col, formula)
            cell.border = thin
            cell.alignment = center
            cell.number_format = "0"

    legend_row = sum_row + 5
    ws.cell(legend_row, 1, "Легенда:").font = Font(bold=True)
    ws.cell(legend_row + 1, 1, "Жёлтые ячейки").fill = input_fill
    ws.cell(legend_row + 1, 2, "— поля для ввода (нарастающий итог с 1-го числа месяца)")
    ws.cell(legend_row + 2, 1, "Красный").fill = red_fill
    ws.cell(legend_row + 2, 2, "— число звонков меньше 20")
    ws.cell(legend_row + 3, 1, "Зелёный").fill = ok_fill
    ws.cell(legend_row + 3, 2, "— число звонков 20 и более")
    ws.cell(legend_row + 4, 1, "Формат").fill = summary_fill
    ws.cell(legend_row + 4, 2, "— автосборка: звонки/ЛПР/квал/ОС/встреча/през/тест/онборд/ЛК/ссылка")

    ws2 = wb.create_sheet("Справочник метрик")
    ws2["A1"] = "№"
    ws2["B1"] = "Метрика"
    ws2["C1"] = "Пояснение"
    for col in range(1, 4):
        ws2.cell(1, col).font = header_font
        ws2.cell(1, col).fill = header_fill
        ws2.cell(1, col).border = thin

    for i, (num, name, help_) in enumerate(METRIC_HELPS, 2):
        ws2.cell(i, 1, num).border = thin
        ws2.cell(i, 1).alignment = center
        ws2.cell(i, 2, name).border = thin
        ws2.cell(i, 3, help_).border = thin
        ws2.cell(i, 3).alignment = left_align

    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 42
    ws2.column_dimensions["C"].width = 80

    ws2["A13"] = "Базы для выбора:"
    ws2["A13"].font = Font(bold=True)
    ws2["A14"] = BASES[0]
    ws2["A15"] = BASES[1]

    ws2["A17"] = "Как заполнять:"
    ws2["A17"].font = Font(bold=True)
    ws2["A18"] = "1. Найдите строку с нужной датой и базой."
    ws2["A19"] = "2. Внесите нарастающие итоги с 1-го числа месяца (не за день, а кумулятивно)."
    ws2["A20"] = "3. Колонка «Формат» сама соберёт строку вида 12/3/1/0/1/0/0/0/0/0 для отчёта."
    ws2["A21"] = "4. Если по одной из баз в день не работали — строку можно оставить пустой."
    ws2["A22"] = "5. Чтобы сменить месяц: python3 scripts/build_daily_metrics_xlsx.py --year YYYY --month M"

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
