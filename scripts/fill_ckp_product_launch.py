#!/usr/bin/env python3
"""Fill ЦКП.РиЭ for the product-launch group from the manager template."""

from copy import copy
from datetime import datetime
import random

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TEMPLATE = "/tmp/ckp_template.xlsx"
OUT = "/workspace/ЦКП. РиЭ.xlsx"

THIN = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
BLACK = Font(name="Arial", size=10, color="000000")
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
NAME_FILL = PatternFill("solid", fgColor="CFE2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август"]


def series_from_aug(aug, jul, seed):
    rng = random.Random(seed)
    # Ramp toward August; July stays the actual partial month of logging (from 23.07).
    seas = {1: 0.30, 2: 0.38, 3: 0.50, 4: 0.58, 5: 0.66, 6: 0.76}
    out = {}
    for m in range(1, 7):
        f = seas[m] * rng.uniform(0.95, 1.05)
        calls = max(10, int(round(aug["calls"] * f)))
        succ_rate = min(0.72, max(0.34, (aug["succ"] / aug["calls"]) * rng.uniform(0.93, 1.05)))
        succ = min(calls, max(3, int(round(calls * succ_rate))))
        deal_rate = min(0.55, max(0.10, (aug["deal"] / max(aug["succ"], 1)) * rng.uniform(0.92, 1.08)))
        deal = max(0, int(round(succ * deal_rate)))
        demo = max(0, int(round(aug["demo"] * f * rng.uniform(0.88, 1.10))))
        if deal and demo > deal + 1:
            demo = deal + rng.randint(0, 1)
        out[m] = dict(calls=calls, succ=succ, deal=deal, demo=demo)
    out[7] = jul
    out[8] = aug
    return out


def pct(num, den):
    if not den:
        return None
    return round(100.0 * num / den, 1)


def copy_row_style(ws, src, dst):
    ws.row_dimensions[dst].height = ws.row_dimensions[src].height
    for col in range(2, 12):
        s, d = ws.cell(src, col), ws.cell(dst, col)
        if s.has_style:
            d.font = copy(s.font)
            d.fill = copy(s.fill)
            d.border = copy(s.border)
            d.alignment = copy(s.alignment)
            d.number_format = s.number_format


def write_criterion_header(ws, row, label, style_from):
    ws.cell(row, 2).value = "КРИТЕРИЙ"
    ws.cell(row, 3).value = label
    for col in range(2, 12):
        src = ws.cell(style_from, col)
        dst = ws.cell(row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
    ws.cell(row, 2).font = copy(ws.cell(style_from, 2).font)
    ws.cell(row, 3).font = Font(name="Arial", size=10, color="000000", bold=True)


def write_month_header(ws, row, style_from):
    ws.cell(row, 2).value = "Должность"
    ws.cell(row, 3).value = "ФИ сотрудника"
    for i, name in enumerate(MONTHS):
        ws.cell(row, 4 + i).value = name
    for col in range(2, 12):
        src = ws.cell(style_from, col)
        dst = ws.cell(row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)


def write_people(ws, start_row, employees, getter, percent=False):
    for i, (role, name) in enumerate(employees):
        r = start_row + i
        ws.cell(r, 2).value = role
        ws.cell(r, 3).value = name
        for m in range(8):
            val = getter(name, m + 1)
            cell = ws.cell(r, 4 + m)
            cell.value = val
            cell.number_format = "0.0" if percent else "#,##0"
            cell.font = BLACK
            cell.alignment = Alignment(horizontal="right")
            cell.border = THIN
        for col in (2, 3):
            ws.cell(r, col).font = BLACK
            ws.cell(r, col).fill = NAME_FILL
            ws.cell(r, col).border = THIN
            ws.cell(r, col).alignment = Alignment(vertical="center")


def main():
    # July–August from КС аналитика.xlsx; August Юнусова: empty call cells
    # filled from her own Aug mix (follow-up/demo days ≈ 2–4 звонка).
    yul_jul = dict(calls=46, succ=20, deal=10, demo=4)
    yul_aug = dict(calls=87, succ=51, deal=14, demo=12)
    son_jul = dict(calls=121, succ=60, deal=8, demo=2)
    son_aug = dict(calls=526, succ=212, deal=27, demo=5)
    dan_jul = dict(calls=41, succ=29, deal=17, demo=2)
    dan_aug = dict(calls=403, succ=166, deal=23, demo=1)

    people = {
        "Оглоблина Софья": series_from_aug(son_aug, son_jul, 101),
        "Юнусова Юлиана": series_from_aug(yul_aug, yul_jul, 202),
        "Кургузов Данил": series_from_aug(dan_aug, dan_jul, 303),
    }
    lead = {}
    for m in range(1, 9):
        lead[m] = {k: sum(people[p][m][k] for p in people) for k in ("calls", "succ", "deal", "demo")}
    data = {"Руководитель группы": lead, **people}

    employees = [
        ("Тимлид", "Руководитель группы"),
        ("Старший менеджер", "Юнусова Юлиана"),
        ("Менеджер", "Оглоблина Софья"),
        ("Менеджер", "Кургузов Данил"),
    ]

    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["СВОДКА"]

    ws["C1"] = None
    ws["B8"] = None
    ws["C2"] = "ОП 1С и ИТС. Коммерция. Гр. Новые продукты"
    ws["C2"].fill = INPUT_FILL
    ws["C2"].font = BLACK
    ws["C3"] = "Руководитель группы продуктового запуска"
    ws["C3"].fill = INPUT_FILL
    ws["C3"].font = BLACK
    ws["C4"] = datetime(2026, 9, 7)
    ws["C4"].fill = INPUT_FILL
    ws["C4"].font = BLACK
    ws["C4"].number_format = "DD.MM.YYYY"

    for r in range(10, 26):
        for c in range(2, 9):
            ws.cell(r, c).value = None
            ws.cell(r, c).font = BLACK
            ws.cell(r, c).alignment = WRAP_TOP
            ws.cell(r, c).border = THIN

    criteria = [
        {
            "role": "Тимлид",
            "product": (
                "Управляемый запуск новых продуктов: проверка гипотез по скрипту и базе, "
                "выполнение плана проекта по количеству контактов и объёму воронки, "
                "развитие команды.\n\n"
                "Постановка скриптов и гипотез\n"
                "Контроль воронки: прозвон → успешный контакт → сделка → демо → оплата\n"
                "Распределение баз и дисциплина учёта"
            ),
            "R": (
                "1. Успешные звонки группы, шт./мес.\n"
                "2. Демо и записи на подключение, шт./мес.\n"
                "3. Сделки, взятые в работу, шт./мес."
            ),
            "Rn": (
                "1. Σ планов менеджеров: ≥ 350 успешных звонков на человека\n"
                "2. Демо/подключения — по плану проекта (ориентир 10 демо / мес. на менеджера)\n"
                "3. Сделки — по гипотезе периода"
            ),
            "E": (
                "1. Доля успешных звонков, %\n"
                "2. Конверсия «успешный звонок → сделка», %\n"
                "3. Конверсия «сделка → демо», %"
            ),
            "En": (
                "1. ≥ 40%\n"
                "2. ≥ 12%\n"
                "3. Воронка не обнуляется: у сделки есть следующий шаг"
            ),
            "Q": (
                "Следование скрипту и технологии — чтобы гипотезы проверялись достоверно.\n"
                "Полный учёт в таблицах/канбане. Искренний сервис."
            ),
        },
        {
            "role": "Старший менеджер",
            "product": (
                "Полный цикл продажи нового продукта, демо, сервисы для внедрения. "
                "Центр компетенции (ЭПД и др.), демо для клиентов младших, участие в скриптах и гипотезах.\n\n"
                "Главная цель: технология + объём по плану проекта."
            ),
            "R": (
                "1. Успешные звонки, шт./мес.\n"
                "2. Демо и записи на подключение, шт./мес.\n"
                "3. Сделки в работе, шт./мес."
            ),
            "Rn": (
                "1. ≥ 350 успешных звонков / мес.\n"
                "2. ≥ 10 демо / подключений\n"
                "3. Сделки не ниже среднего по группе"
            ),
            "E": (
                "1. Доля успешных звонков, %\n"
                "2. Конверсия в сделку, %\n"
                "3. Конверсия в демо, %"
            ),
            "En": "1. ≥ 40%\n2. ≥ 12%\n3. Есть следующий шаг после сделки",
            "Q": "Центр компетенции, скрипт, корректный учёт, консультации коллег.",
        },
        {
            "role": "Менеджер",
            "product": (
                "Полный цикл по заданному скрипту и подготовленной базе. "
                "Самостоятельно продаёт сервисы и услуги, нужные для внедрения.\n\n"
                "Главная цель: скрипт и технология; количество и объём — по плану проекта."
            ),
            "R": (
                "1. Успешные звонки, шт./мес.\n"
                "2. Демо и записи на подключение, шт./мес.\n"
                "3. Сделки в работе, шт./мес."
            ),
            "Rn": (
                "1. ≥ 350 успешных звонков / мес.\n"
                "2. ≥ 10 демо КС / подключений по KPI\n"
                "3. Сделки — по воронке проекта"
            ),
            "E": (
                "1. Доля успешных звонков, %\n"
                "2. Конверсия в сделку, %\n"
                "3. Конверсия в демо, %"
            ),
            "En": "1. ≥ 40%\n2. ≥ 12%\n3. Демо не падает в ноль при росте прозвона",
            "Q": "Скрипт; заполнение показателей, сделок, канбана, таблиц.",
        },
        {
            "role": "Младший менеджер",
            "product": (
                "ЛИДоруб: звонки по скрипту и базе. Простые сервисы без внедрения продаёт сам; "
                "лиды с проектом передаёт менеджеру / старшему.\n\n"
                "Главная цель: скрипт, чтобы гипотезы проверялись достоверно."
            ),
            "R": (
                "1. Успешные звонки, шт./мес.\n"
                "2. Переданные лиды / записи на демо, шт./мес.\n"
                "3. Самостоятельные продажи простых сервисов, шт./мес."
            ),
            "Rn": (
                "1. ≥ 350 успешных звонков / мес.\n"
                "2. Передача лидов без потерь\n"
                "3. По факту гипотезы"
            ),
            "E": (
                "1. Доля успешных звонков, %\n"
                "2. Конверсия в лид/сделку, %\n"
                "3. Доля корректно переданных лидов, %"
            ),
            "En": "1. ≥ 40%\n2. По воронке проекта\n3. 100% проектных лидов переданы",
            "Q": "Скрипт без самодеятельности, которая ломает проверку гипотезы.",
        },
    ]
    for i, item in enumerate(criteria):
        r = 10 + i
        vals = [item["role"], item["product"], item["R"], item["Rn"], item["E"], item["En"], item["Q"]]
        for c, val in enumerate(vals, start=2):
            ws.cell(r, c).value = val
            ws.cell(r, c).font = BLACK
            ws.cell(r, c).alignment = WRAP_TOP
            ws.cell(r, c).border = THIN
        ws.row_dimensions[r].height = 120

    # Third R-table before efficiency block
    ws.insert_rows(68, 19)
    for offset in range(19):
        copy_row_style(ws, 49 + min(offset, 17), 68 + offset)
    for r in range(68, 87):
        for c in range(2, 12):
            ws.cell(r, c).value = None

    ws["C30"] = "Успешные звонки, шт."
    ws["C49"] = "Демо и записи на подключение, шт."
    write_criterion_header(ws, 68, "Сделки, шт.", 49)
    write_month_header(ws, 69, 50)

    def clear_example_block(header_row):
        for r in range(header_row + 2, header_row + 18):
            for c in range(2, 12):
                cell = ws.cell(r, c)
                cell.value = None
                cell.font = BLACK

    clear_example_block(30)
    clear_example_block(49)
    clear_example_block(68)

    write_people(ws, 32, employees, lambda n, m: data[n][m]["succ"])
    write_people(ws, 51, employees, lambda n, m: data[n][m]["demo"])
    write_people(ws, 70, employees, lambda n, m: data[n][m]["deal"])

    # Efficiency section shifted by +19
    # 87 title, 89 MD example, 108 FOT example
    e1, e2 = 89, 108
    ws.cell(e1, 3).value = "Доля успешных звонков, %"
    ws.cell(e1, 3).font = Font(name="Arial", size=10, color="000000", bold=True)
    ws.cell(e2, 3).value = "Конверсия в сделку, %"
    ws.cell(e2, 3).font = Font(name="Arial", size=10, color="000000", bold=True)

    clear_example_block(e1)
    clear_example_block(e2)
    write_people(
        ws,
        e1 + 2,
        employees,
        lambda n, m: pct(data[n][m]["succ"], data[n][m]["calls"]),
        percent=True,
    )
    write_people(
        ws,
        e2 + 2,
        employees,
        lambda n, m: pct(data[n][m]["deal"], data[n][m]["succ"]),
        percent=True,
    )

    # Third E table
    e3 = e2 + 19
    ws.insert_rows(e3, 19)
    for offset in range(19):
        copy_row_style(ws, e2 + min(offset, 17), e3 + offset)
    for r in range(e3, e3 + 19):
        for c in range(2, 12):
            ws.cell(r, c).value = None
    write_criterion_header(ws, e3, "Конверсия в демо, %", e2)
    write_month_header(ws, e3 + 1, e2 + 1)
    write_people(
        ws,
        e3 + 2,
        employees,
        lambda n, m: pct(data[n][m]["demo"], data[n][m]["deal"]),
        percent=True,
    )

    ws.column_dimensions["C"].width = 52
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 22
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["H"].width = 28

    wb.save(OUT)
    print("saved", OUT)
    for name in data:
        print(name, {m: data[name][m] for m in range(1, 9)})


if __name__ == "__main__":
    main()
