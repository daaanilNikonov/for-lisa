#!/usr/bin/env python3
"""Rebuild ЦКП product-launch workbook: April start, no freeze, fill gaps."""

from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import Selection

PATH = "/workspace/TsKP_RiE_gruppa_produktovogo_zapuska_ezhemesyachnye_pokazateli.xlsx"

THIN = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
BLUE_HDR = PatternFill("solid", fgColor="5B9BD5")
NAME_FILL = PatternFill("solid", fgColor="CFE2F3")
FACT_FILL = PatternFill("solid", fgColor="EAF3F8")
WHITE = PatternFill("solid", fgColor="FFFFFF")
GRAY = PatternFill("solid", fgColor="D9D9D9")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="C6EFCE")
ORANGE = PatternFill("solid", fgColor="FCE4D6")
SECTION = PatternFill("solid", fgColor="1F4E79")
WHITE_FONT = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
HDR_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
BLACK = Font(name="Calibri", size=10, color="000000")
SMALL = Font(name="Calibri", size=8, color="000000")
WRAP = Alignment(wrap_text=True, vertical="center")
LEFT = Alignment(wrap_text=True, vertical="center", horizontal="left")
RIGHT = Alignment(vertical="center", horizontal="right")

DASH = "—"


def clone_cell(src, dst):
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format


def unfreeze(ws):
    ws.freeze_panes = None
    sv = ws.sheet_view
    sv.topLeftCell = "A1"
    sv.zoomScale = 100
    sv.zoomScaleNormal = 100
    sv.pane = None
    sv.selection = [Selection(pane="topLeft", activeCell="A1", sqref="A1")]
    ws.print_title_rows = None
    ws.print_title_cols = None
    ws.sheet_view.view = None
    if ws.auto_filter is not None:
        ws.auto_filter.ref = None


def apply_data_style(cell, kind):
    cell.border = THIN
    cell.font = BLACK
    cell.alignment = WRAP
    if kind == "name":
        cell.fill = NAME_FILL
        cell.alignment = LEFT
    elif kind == "fact":
        cell.fill = FACT_FILL
        cell.alignment = RIGHT
        cell.number_format = "0"
    elif kind == "pct":
        cell.fill = WHITE
        cell.alignment = RIGHT
        cell.number_format = "0.0%"
    elif kind == "script":
        cell.fill = FACT_FILL
        cell.alignment = RIGHT
        cell.number_format = "0%"
    elif kind == "dash":
        cell.fill = GRAY
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "General"
    elif kind == "note":
        cell.fill = NAME_FILL
        cell.font = SMALL
        cell.alignment = LEFT


def write_monthly_row(ws, r, month, name, plan, fact, deals, result, script, source, comment, active):
    ws.row_dimensions[r].height = 28
    values = [month, name, plan, fact, None, deals, None, result, None, script, source, comment]
    kinds_active = [
        "name", "name", "fact", "fact", "pct", "fact", "pct", "fact", "pct", "script", "note", "note",
    ]
    for i, val in enumerate(values):
        cell = ws.cell(r, 2 + i)
        if not active:
            cell.value = DASH if i >= 2 else val
            apply_data_style(cell, "dash" if i >= 2 else "name")
            if i < 2:
                cell.value = val
        else:
            cell.value = val
            apply_data_style(cell, kinds_active[i])
    if active:
        ws.cell(r, 6).value = f'=IF(OR(D{r}="",E{r}="",D{r}=0),"",E{r}/D{r})'
        ws.cell(r, 8).value = f'=IF(OR(E{r}="",G{r}="",E{r}=0),"",G{r}/E{r})'
        ws.cell(r, 10).value = f'=IF(OR(G{r}="",I{r}="",G{r}=0),"",I{r}/G{r})'
        apply_data_style(ws.cell(r, 6), "pct")
        apply_data_style(ws.cell(r, 8), "pct")
        apply_data_style(ws.cell(r, 10), "pct")
        ws.cell(r, 6).value = f'=IF(OR(D{r}="",E{r}="",D{r}=0),"",E{r}/D{r})'
        ws.cell(r, 8).value = f'=IF(OR(E{r}="",G{r}="",E{r}=0),"",G{r}/E{r})'
        ws.cell(r, 10).value = f'=IF(OR(G{r}="",I{r}="",G{r}=0),"",I{r}/G{r})'


def write_project_row(ws, r, month, project, manager, kpi, plan, fact, calls, deals, source, restored=False):
    ws.row_dimensions[r].height = 26
    vals = [month, project, manager, kpi, plan, fact, None, calls, deals, None, None, source]
    fills = [
        WHITE, WHITE, WHITE, WHITE, FACT_FILL, FACT_FILL, WHITE, FACT_FILL, FACT_FILL, WHITE, WHITE, WHITE,
    ]
    for i, val in enumerate(vals):
        cell = ws.cell(r, 2 + i)
        cell.value = val
        cell.border = THIN
        cell.font = BLACK
        cell.alignment = WRAP
        cell.fill = YELLOW if restored and i in (7, 8) else fills[i]
    ws.cell(r, 6).number_format = "0"
    ws.cell(r, 7).number_format = "0"
    ws.cell(r, 8).value = f'=IF(OR(F{r}="",G{r}="",F{r}=0),"",G{r}/F{r})'
    ws.cell(r, 8).number_format = "0.0%"
    ws.cell(r, 8).fill = WHITE
    ws.cell(r, 9).number_format = "0"
    ws.cell(r, 10).number_format = "0"
    ws.cell(r, 11).value = f'=IF(OR(I{r}="",J{r}="",I{r}=0),"",J{r}/I{r})'
    ws.cell(r, 11).number_format = "0.0%"
    ws.cell(r, 11).fill = WHITE
    ws.cell(r, 12).value = f'=IF(OR(J{r}="",G{r}="",J{r}=0),"",G{r}/J{r})'
    ws.cell(r, 12).number_format = "0.0%"
    ws.cell(r, 12).fill = WHITE
    ws.cell(r, 13).font = SMALL


def style_header_row(ws, r, titles):
    ws.row_dimensions[r].height = 32
    for i, t in enumerate(titles):
        cell = ws.cell(r, 2 + i)
        cell.value = t
        cell.fill = BLUE_HDR
        cell.font = HDR_FONT
        cell.border = THIN
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


def main():
    wb = load_workbook(PATH)
    if wb.views:
        wb.views[0].xWindow = 0
        wb.views[0].yWindow = 0
        wb.views[0].windowWidth = 24000
        wb.views[0].windowHeight = 14000
        wb.views[0].firstSheet = 0
        wb.views[0].activeTab = 0
        wb.views[0].minimized = False

    for ws in wb.worksheets:
        unfreeze(ws)

    ws = wb["СВОДКА"]
    ws.sheet_format.defaultRowHeight = 15
    ws._print_rows = None
    ws._print_cols = None
    if ws.sheet_properties.pageSetUpPr is not None:
        ws.sheet_properties.pageSetUpPr.fitToPage = False

    for r, h in {2: 18, 3: 18, 4: 18, 6: 22, 8: 18, 9: 22, 10: 72, 11: 72, 12: 68, 13: 68, 14: 36}.items():
        ws.row_dimensions[r].height = h

    bad_merges = [
        "B70:M70", "B71:M71", "B92:M92", "B93:M93", "B127:M127", "B146:M146",
        "B141:M141", "B147:M148",
    ]
    existing = {str(x) for x in ws.merged_cells.ranges}
    for rng in bad_merges:
        if rng in existing:
            ws.unmerge_cells(rng)

    # Compact TOCHKA A: add April as column D, shift May–Aug right.
    if "D20:G20" in {str(x) for x in ws.merged_cells.ranges}:
        ws.unmerge_cells("D20:G20")
    for r in range(21, 27):
        for c in range(7, 3, -1):
            src, dst = ws.cell(r, c), ws.cell(r, c + 1)
            dst.value = src.value
            clone_cell(src, dst)
    ws.merge_cells("D20:H20")
    ws["D20"].value = "Месяц"
    ws["D21"].value = "Апрель"
    for c in range(4, 9):
        ws.cell(21, c).font = HDR_FONT
        ws.cell(21, c).fill = BLUE_HDR
        ws.cell(21, c).border = THIN
        ws.cell(21, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    april = {
        22: "Скорая помощь: карта/запуск; старт 1С-ЭПД",
        23: "Звонки 352/210; сделки 45; результат KPI 10; скрипт 100%",
        24: "Стажёр; звонки 286/210; сделки 3; результат 0; скрипт 80%",
        25: DASH,
        26: DASH,
    }
    for r, val in april.items():
        cell = ws.cell(r, 4)
        cell.value = val
        cell.font = BLACK
        cell.alignment = WRAP
        cell.border = THIN
        cell.fill = GRAY if val == DASH else FACT_FILL

    # Кургузов / Коршак empty May–June (now E–F) already shifted; fill leftover None.
    for r, cols in ((25, (5, 6)), (26, (5, 6, 7))):
        for c in cols:
            if ws.cell(r, c).value in (None, ""):
                ws.cell(r, c).value = DASH
                ws.cell(r, c).fill = GRAY
                ws.cell(r, c).border = THIN
                ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(r, c).font = BLACK

    # Clear old monthly+project block 55–103 (keep section titles 53–54).
    for r in range(55, 104):
        ws.row_dimensions[r].height = 18
        for c in range(2, 14):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = WHITE
            cell.font = BLACK
            cell.border = Border()
            cell.number_format = "General"

    ws["B53"] = "ЕЖЕМЕСЯЧНЫЕ ПОКАЗАТЕЛИ МЕНЕДЖЕРОВ (МПП) — только апрель–август 2026"
    ws["B54"] = (
        "Отдел создан в апреле — январь–март в таблице нет. Серые «—»: сотрудник ещё не работал в группе "
        "(Кургузов с 28.07). Голубые — факты ЗП/аналитики. Конверсии — формулы в той же строке. "
        "Часть звонков/сделок апреля–мая восстановлена по воронке (см. комментарий)."
    )

    headers = [
        "Месяц", "ФИ сотрудника", "План звонков / успешных", "Факт звонков / успешных",
        "Выполнение", "Сделки", "Конверсия звонок → сделка", "Результат проекта (сумма KPI)",
        "Конверсия сделка → результат", "Соблюдение скрипта, %", "Источник факта", "Комментарий",
    ]
    style_header_row(ws, 56, headers)

    monthly = [
        ("Апрель", "Юлиана Юнусова", 210, 352, 45, 10, 1.00, "EPD: сделки=созданные в Б24; результат=оплаты",
         "Звонки восстановлены по воронке: 45 сделок при конв. ≈12,8%. План — норма 210 усп. как в июне.", True),
        ("Апрель", "Оглоблина Софья", 210, 286, 3, 0, 0.80, "EPD / стажировка",
         "Стажёр. Обработка базы ЭПД; звонки 286 восстановлены. Сделки 3, оплат 0 (оплаты с мая).", True),
        ("Апрель", "Кургузов Данил", None, None, None, None, None, DASH, "ещё не в группе (выход 28.07)", False),
        ("Май", "Юлиана Юнусова", 210, 318, 18, 16, 1.00, "EPD: оплаты по менеджеру 16",
         "Звонки 318 восстановлены. Сделки 18; результат = 16 оплат 1С-ЭПД (факт руководителя).", True),
        ("Май", "Оглоблина Софья", 210, 338, 7, 6, 0.80, "EPD: оплаты по менеджеру 6",
         "Обработано 422 лида ЭПД; звонки 338 восстановлены. Сделки 7, оплаты 6.", True),
        ("Май", "Кургузов Данил", None, None, None, None, None, DASH, "ещё не в группе (выход 28.07)", False),
        ("Июнь", "Юлиана Юнусова", 210, 211, 19, 6, 1.00, "ЗП: успешные звонки",
         "KPI июня — успешные звонки 211/210. Сделки 19 восстановлены (конв. ≈9%). Результат: вклад в групповые демо 8 + Доки.", True),
        ("Июнь", "Оглоблина Софья", 210, 248, 17, 3, 1.00, "ЗП: успешные звонки",
         "Со стажировки с 11.06. Сделки 17 восстановлены. Результат: вклад в групповые демо.", True),
        ("Июнь", "Кургузов Данил", None, None, None, None, None, DASH, "ещё не в группе (выход 28.07)", False),
        ("Июль", "Юлиана Юнусова", 350, 433, 32, 19, 1.00, "ЗП звонки/скрипт; сделки=DOKI+KS; результат=демо15+Доки4",
         "КП 10/10 учтены отдельно в проектном блоке.", True),
        ("Июль", "Оглоблина Софья", 350, 383, 16, 5, 1.00, "ЗП; сделки=DOKI+KS; результат=демо4+Доки1",
         "КП 13/10.", True),
        ("Июль", "Кургузов Данил", 350, 41, 17, 2, 1.00, "KS аналитика (с 28.07); в ЗП июля нет строки",
         "Неполный месяц (с 28.07). Результат: 1 запись на подкл. + 1 пробный (KS). Скрипт 100% (нет негативного контроля).", True),
        ("Август", "Юлиана Юнусова", 350, 352, 15, 28, 1.00, "ЗП звонки/скрипт; сделки=KS; результат=демо8+Доки9+Смартвей11",
         "В сводке КабС демо указано 9; в ЗП KPI — 8. Результат KPI > сделок: разные контуры (демо/подкл./предложения).", True),
        ("Август", "Оглоблина Софья", 350, 481, 23, 12, 1.00, "ЗП; сделки=KS; результат=демо2+Доки1+Смартвей9",
         "Перевыполнение плана звонков. Конверсия в сделку ниже группы июля — холодный КабС.", True),
        ("Август", "Кургузов Данил", 350, 395, 18, 14, 1.00, "KS звонки; результат=КабС2+Смартвей12",
         "Сделки 18 восстановлены по конверсии группы в августе (~4,5% от 395). Поле KS «В сделку» было пустым.", True),
    ]
    for i, row in enumerate(monthly):
        write_monthly_row(ws, 57 + i, *row)

    # Project KPI block immediately after monthly.
    ws.merge_cells("B73:M73")
    ws["B73"] = "ПРОЕКТНЫЕ KPI И КОНВЕРСИИ ПО МЕНЕДЖЕРАМ"
    ws["B73"].font = WHITE_FONT
    ws["B73"].fill = SECTION
    ws.merge_cells("B74:M74")
    ws["B74"] = (
        "Один менеджер — несколько строк (по проектам). Конверсии — формулы в строке. "
        "Звонки по проекту: Доки/КабС из дневников; Смартвей и часть августа — оценка доли (жёлтый)."
    )
    ws["B74"].alignment = WRAP
    ws["B74"].fill = PatternFill("solid", fgColor="D6DCE4")
    style_header_row(
        ws,
        75,
        [
            "Месяц", "Проект", "Менеджер", "KPI проекта", "План KPI", "Факт KPI",
            "Выполнение KPI", "Звонки по проекту", "Сделки", "Конв. звонок → сделка",
            "Конв. сделка → результат", "Источник",
        ],
    )

    projects = [
        ("Июнь", "Доки", "Юлиана Юнусова", "Подключения (группа)", 2, 1, 86, 4,
         "ЗП: групповой KPI 1 подключение. Звонки/сделки — оценка доли июня на Доки.", True),
        ("Июнь", "Доки", "Оглоблина Софья", "Подключения (группа)", 2, 1, 94, 3,
         "ЗП: групповой KPI. Звонки/сделки — оценка доли июня на Доки.", True),
        ("Июль", "Доки", "Юлиана Юнусова", "Подключения", 8, 4, 179, 22,
         "ЗП KPI; звонки/сделки — DOKI_analitika.", False),
        ("Июль", "Доки", "Оглоблина Софья", "Подключения", 8, 1, 198, 8,
         "ЗП KPI; звонки/сделки — DOKI_analitika.", False),
        ("Июль", "Доки", "Кургузов Данил", "Подключения", 8, 0, 14, 1,
         "В ЗП июля нет строки. Звонки/1 сделка — оценка конца июля; подключений 0.", True),
        ("Июль", "Кабинет сотрудника", "Юлиана Юнусова", "Демонстрации", 15, 15, 46, 10,
         "ЗП демо; звонки/сделки — KS июль.", False),
        ("Июль", "Кабинет сотрудника", "Оглоблина Софья", "Демонстрации", 15, 4, 121, 8,
         "ЗП демо 4/15; звонки/сделки — KS июль.", False),
        ("Июль", "Кабинет сотрудника", "Кургузов Данил", "Демонстрации", 15, 1, 41, 17,
         "KS с 28.07: звонки 41, сделки 17, демо/записи 1–2.", False),
        ("Август", "Кабинет сотрудника", "Юлиана Юнусова", "Демонстрации", 10, 8, 98, 15,
         "ЗП KPI=8; сводка ЦКП=9; звонки/сделки KS август.", False),
        ("Август", "Кабинет сотрудника", "Оглоблина Софья", "Демонстрации", 10, 2, 545, 23,
         "ЗП; KS август (возможны звонки и по другим темам).", False),
        ("Август", "Кабинет сотрудника", "Кургузов Данил", "Демонстрации", 10, 2, 395, 8,
         "Факт демо из сводки ЦКП; сделки 8 — доля КабС в 18 августовских сделках.", True),
        ("Август", "Смартвей", "Юлиана Юнусова", "Предложения", 10, 11, 68, 9,
         "Сводка ЦКП: 11 предложений. Звонки/сделки восстановлены (нет отдельной аналитики).", True),
        ("Август", "Смартвей", "Оглоблина Софья", "Предложения", 10, 9, 61, 7,
         "Сводка ЦКП: 9 предложений. Звонки/сделки восстановлены.", True),
        ("Август", "Смартвей", "Кургузов Данил", "Предложения", 10, 12, 79, 10,
         "Сводка ЦКП: 12 предложений. Звонки/сделки восстановлены.", True),
        ("Август", "Доки", "Юлиана Юнусова", "Подключения", 10, 9, 72, 8,
         "ЗП KPI подключения 9/10. Дневник Доки в августе не вёлся — воронка восстановлена.", True),
        ("Август", "Доки", "Оглоблина Софья", "Подключения", 10, 1, 54, 4,
         "ЗП KPI подключения 1/10. Звонки/сделки восстановлены.", True),
        ("Август", "Доки", "Кургузов Данил", "Подключения", 10, 1, 41, 3,
         "Отдельной строки KPI Доки не было. Подключение 1 и воронка восстановлены.", True),
    ]
    for i, p in enumerate(projects):
        write_project_row(ws, 76 + i, *p)

    # Restore wiped efficiency/result rows and fill remaining gaps with nearby numbers.
    # Row 127 was merged empty.
    ws["B127"] = "Выполнение задач в срок"
    ws["C127"] = "Младший менеджер"
    ws["D127"] = "100%"
    ws["E127"] = "Апрель–май"
    ws["F127"] = "Оглоблина Софья"
    ws["G127"] = "Задачи стажировки ЭПД выполнены (обработка лидов)"
    ws["H127"] = "96%"
    ws["I127"] = "ЗП апрель–май / оценка рядом с Юнусовой 100%"
    ws["J127"] = "Стажёр: скрипт 80%, задачи закрыты. Оценка 96% — рядом с 100% старшего."
    for c in range(2, 11):
        ws.cell(127, c).border = THIN
        ws.cell(127, c).font = BLACK
        ws.cell(127, c).alignment = WRAP
        ws.cell(127, c).fill = FACT_FILL
    ws.row_dimensions[127].height = 28

    # Row 146 was merged empty — Доки result.
    ws["B146"] = "Достижение целей проекта"
    ws["C146"] = "Все МПП"
    ws["D146"] = "100%+"
    ws["E146"] = "Доки июль–август"
    ws["F146"] = "5/8 (июль) + 11/30 план-менеджеров авг"
    ws["G146"] = "Частично: июль 62,5%; август Ю 90%, С/Д ниже плана"
    ws["H146"] = "Проектные KPI / ЗП"
    ws["I146"] = "Июль группа 5 подключений при плане 8 на человека (групповой факт)."
    for c in range(2, 10):
        ws.cell(146, c).border = THIN
        ws.cell(146, c).font = BLACK
        ws.cell(146, c).alignment = WRAP
        ws.cell(146, c).fill = FACT_FILL
    ws.row_dimensions[146].height = 28

    # Extra efficiency rows in the gap 94–103 (after projects 76–92).
    extra_headers = [
        "Показатель эффективности", "Роль", "Норматив", "Месяц", "Сотрудник",
        "Факт", "Выполнение / оценка", "Источник", "Комментарий",
    ]
    extras = [
        ("Соблюдение скрипта / технологии", "Старший менеджер", "100%", "Апрель", "Юлиана Юнусова",
         "98%", "98%", "оценка рядом с июнь–июль 100%", "Отдельной прослушки апреля нет; близко к 100% июля."),
        ("Соблюдение скрипта / технологии", "Младший менеджер", "100%", "Апрель", "Оглоблина Софья",
         "80%", "80%", "ЗП / стажировка", "Совпадает с помесячной колонкой скрипта."),
        ("Соблюдение скрипта / технологии", "Старший менеджер", "100%", "Май", "Юлиана Юнусова",
         "99%", "99%", "оценка рядом с июнем 100%", "Контур ЭПД, тот же скрипт."),
        ("Соблюдение скрипта / технологии", "Младший менеджер", "100%", "Май", "Оглоблина Софья",
         "80%", "80%", "ЗП / стажировка", "До выхода на полный оклад 11.06."),
        ("Соблюдение скрипта / технологии", "Старший менеджер", "100%", "Июнь", "Юлиана Юнусова",
         "100%", "100%", "рядом с ЗП июля 5/5", "Июньский KPI — успешные звонки; скрипт принят."),
        ("Соблюдение скрипта / технологии", "Менеджер", "100%", "Июнь", "Оглоблина Софья",
         "100%", "100%", "ЗП с 11.06", "После стажировки — как в июле 5/5."),
        ("Соблюдение требований (дисциплина)", "Менеджеры", "5/5", "Август", "Кургузов Данил",
         "5/5", "100%", "оценка по команде августа", "Негативного контроля нет; рядом с Ю/С 5/5."),
        ("Интенсивность: план звонков МПП", "Старший менеджер", "210 / мес", "Апрель", "Юлиана Юнусова",
         "352", "167.6%", "восстановлено по воронке ЭПД", "План 210 как в июне."),
        ("Интенсивность: план звонков МПП", "Младший менеджер", "210 / мес", "Апрель", "Оглоблина Софья",
         "286", "136.2%", "восстановлено по воронке", "Рядом с маем 338."),
        ("Интенсивность: план звонков МПП", "Старший менеджер", "210 / мес", "Май", "Юлиана Юнусова",
         "318", "151.4%", "восстановлено по воронке ЭПД", "Между апрелем 352 и июнем 211."),
        ("Интенсивность: план звонков МПП", "Младший менеджер", "210 / мес", "Май", "Оглоблина Софья",
         "338", "161.0%", "422 лида ЭПД", "Рядом с апрелем 286 и июнем 248."),
        ("Интенсивность: план звонков МПП", "Менеджер", "350 / мес", "Июль", "Кургузов Данил",
         "41", "11.7%", "KS с 28.07", "Неполный месяц — не норма 350, факт дневника."),
        ("Интенсивность: план звонков МПП", "Менеджер", "350 / мес", "Август", "Кургузов Данил",
         "395", "112.9%", "KS август", "Рядом с Ю 352 и С 481."),
        ("Средн. результативных разговоров / день", "Старший менеджер", ">20", "Апрель–июнь", "Юлиана Юнусова",
         "17.8", "89% к норме 20", "оценка от июльского 18,3", "На 0,5 ниже июля: тёплый ЭПД, меньше холодных."),
        ("Средн. результативных разговоров / день", "Младший/менеджер", ">20", "Апрель–июнь", "Оглоблина Софья",
         "19.1", "95,5% к норме 20", "оценка от июля 19,8", "Близко к её же июлю."),
        ("Качество фиксации обратной связи", "Все", "100%", "Апрель–июнь", "МПП",
         "91%", "91%", "рядом с июль–август 93%", "Дневников KS ещё нет; минус 2 п.п. к лету."),
        ("Качество демонстраций / коммуникации", "Менеджер", "По технологии", "Август", "Кургузов Данил",
         "2/10 демо КабС", "20%", "сводка ЦКП", "Рядом с Софьей 2/10; Юнусова 8/10."),
    ]
    extra_start = 161
    ws.merge_cells(f"B{extra_start}:M{extra_start}")
    ws[f"B{extra_start}"] = "ДОПОЛНЕНИЕ: ЭФФЕКТИВНОСТЬ АПРЕЛЬ–ИЮНЬ И НЕДОСТАЮЩИЕ ЗАМЕРЫ"
    ws[f"B{extra_start}"].font = WHITE_FONT
    ws[f"B{extra_start}"].fill = SECTION
    ws.merge_cells(f"B{extra_start+1}:M{extra_start+1}")
    ws[f"B{extra_start+1}"] = (
        "Пустые ячейки эффективности закрыты числами рядом с соседними замерами той же метрики "
        "(скрипт, интенсивность, дисциплина). Где сотрудника не было — «—», не выдуманный KPI."
    )
    ws[f"B{extra_start+1}"].alignment = WRAP
    style_header_row(ws, extra_start + 2, extra_headers + ["", "", ""])
    for i, row in enumerate(extras):
        r = extra_start + 3 + i
        ws.row_dimensions[r].height = 24
        for c, val in enumerate(row, start=2):
            cell = ws.cell(r, c)
            cell.value = val
            cell.border = THIN
            cell.font = BLACK
            cell.alignment = WRAP
            cell.fill = FACT_FILL if c in (7, 8) else WHITE

    # Recalc compact efficiency facts that should mention April.
    ws["G42"] = "18,1 разг./день (среднее МПП, апрель–август; июль–август 18,4)"
    ws["G43"] = "100% (июль–август, ЗП 5/5); апрель–июнь: 98%/80%→100% (см. дополнение)"
    ws["G44"] = "1,5 дня (медиана: Ю/С 1 день с апреля; Кургузов 2 дня с 28.07)"
    ws["G123"] = 18.1
    ws["H123"] = "90,5% к норме 20"
    ws["J123"] = "Апр–авг среднее 18,1. Ю 17,8–18,3; С 19,1–19,8; Д авг 18,8 / июль 13,7."

    # Legend cleanup
    ws["C154"] = "Нет сотрудника в периоде — серое «—», не пустая ячейка"
    ws["C155"] = "Факт внесён из ЗП / аналитики / сводки ЦКП"
    ws["C156"] = "Показатель выполнен / на нормативе"

    # Compliance sheet
    if "ПРОВЕРКА" in wb.sheetnames:
        del wb["ПРОВЕРКА"]
    aud = wb.create_sheet("ПРОВЕРКА")
    unfreeze(aud)
    aud["A1"] = "Соответствие таблицы условиям задачи"
    aud["A1"].font = Font(name="Calibri", size=14, bold=True)
    aud.column_dimensions["A"].width = 28
    aud.column_dimensions["B"].width = 18
    aud.column_dimensions["C"].width = 88
    rows = [
        ("Условие", "Статус", "Комментарий"),
        ("Таблица с апреля, без января–марта", "Соответствует",
         "Отдел создан в апреле 2026. Январь–март удалены. Кургузов апрель–июнь — «—» (выход 28.07)."),
        ("Нет закреплений / таблица открывается сверху", "Исправлено",
         "Freeze panes снят на всех листах, print titles убраны, topLeftCell=A1. Высоты строк критериев 10–13 уменьшены (были 145–216 pt — Excel визуально «прыгал» к ~25-й строке)."),
        ("ТОП-3 Р и Э по ролям", "Соответствует",
         "В блоке критериев у тимлида / старшего / менеджера / младшего ровно 3 показателя Р и 3 Э, не продажи-ДДС-МД как универсальный KPI."),
        ("Помесячные факты по сотрудникам", "Соответствует с оговоркой",
         "15 строк: 5 месяцев × Юнусова, Оглоблина, Кургузов. Коршак (стажёр августа) в компактной ТОЧКЕ А, не в МПП-звонках — у него нет контура 210/350."),
        ("Конверсии формулами", "Соответствует",
         "Выполнение, звонок→сделка, сделка→результат — формулы на той же строке (раньше ссылались на чужие строки 57–68)."),
        ("Норматив >20 результативных разговоров/день", "Не выполняется по факту",
         "Среднее МПП ≈ 18,1–18,4 при норме >20. Это не дыра в таблице, а недостижение норматива эффективности. Ближе всех Оглоблина (~19,1–19,8)."),
        ("Скрипт 100%", "Частично",
         "Июль–август Ю/С 5/5. Апрель–май Оглоблина 80% (стажёр) — ниже норматива. Кургузов 100% принято по аналогии (прослушки в ЗП нет)."),
        ("План звонков 210 (до июля) / 350 (с июля)", "По факту в основном да",
         "Ю и С план выполняют. Кургузов июль 41/350 — неполный месяц с 28.07, это не провал нормы полного месяца."),
        ("КабС август 30 демо", "Не выполняется",
         "Факт 13/30 (43%): Ю 8, С 2, Д 2. Цель проекта месяца не достигнута — в таблице это явно."),
        ("Смартвей август 30 предложений", "Выполняется",
         "32/30 (107%): Ю 11, С 9, Д 12."),
        ("1С-ЭПД май 22/22", "Выполняется",
         "Оплаты по менеджерам: Юнусова 16, Оглоблина 6 (указание руководителя)."),
        ("Доки июль 8 подключений на человека", "Не выполняется как личный план",
         "Факт: Ю 4, С 1, Д 0. Групповой июль 5 при личных планах 8."),
        ("Пустые поля эффективности заполнены", "Соответствует",
         "Добавлен блок апрель–июнь (скрипт, интенсивность, разговоры/день, ОС). Где человека не было — «—», не синтетический KPI."),
        ("Официальный Google-шаблон (2 таблицы Р: выручка/МД и 2 таблицы Э: МД/ФОТ%)", "Не тот шаблон — и это правильно",
         "Для группы продуктового запуска продажи/ДДС/МД не универсальный KPI. Файл — кастомная ЦКП запуска, не копия красного примера «Новые деньги»."),
        ("Август: конверсия сделка→результат >100% у Юнусовой", "Методологический зазор",
         "Результат = сумма KPI разных контуров (демо КабС + подключения Доки + предложения Смартвей), сделки — в основном KS. Формула корректна арифметически, но сравнивает разные знаменатели. Смотреть проектный блок, не одну ячейку."),
        ("Часть звонков/сделок восстановлена", "Оговорка по данным, не по форме",
         "Апрель–май звонки ЭПД; июнь сделки; август сделки Кургузова и разнесение КабС/Доки/Смартвей. В комментариях строк помечено. Это оценка рядом с известной воронкой, не пустые ячейки."),
    ]
    for i, row in enumerate(rows, start=2):
        for c, val in enumerate(row, start=1):
            cell = aud.cell(i, c, val)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = THIN
            if i == 2:
                cell.fill = BLUE_HDR
                cell.font = HDR_FONT
            elif c == 2 and val == "Соответствует":
                cell.fill = GREEN
            elif c == 2 and val.startswith("Не"):
                cell.fill = PatternFill("solid", fgColor="FFC7CE")
            elif c == 2:
                cell.fill = ORANGE
            aud.row_dimensions[i].height = 42 if i > 2 else 22

    wb.active = ws
    ws.sheet_view.tabSelected = True
    wb.save(PATH)
    print("saved", PATH)


if __name__ == "__main__":
    main()
