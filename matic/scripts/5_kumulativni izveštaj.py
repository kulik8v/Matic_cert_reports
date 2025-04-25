#!/usr/bin/env python3
# coding: utf-8
"""
create_kumulativni_izveštaj.py

Скрипт создаёт кумулятивный отчёт на основе файлов из <root>/Output/Izvedeno:
1. Находит корень проекта (родитель папки scripts).
2. Копирует шаблон izvedeno_template.xlsx в <root>/Output/Izvedeno и повторно сохраняет кумулятивный отчёт с датой.
3. Суммирует ячейки листов 'K_00_REKAP', 'K_03_AB radovi', 'K_04_Armiracki' из всех файлов (кроме текущего отчёта).
4. Записывает итоговые суммы в новый отчёт и переносит его в папку kumulativni izveštaj.
"""
import sys
import shutil
from pathlib import Path
from datetime import datetime

import pythoncom
import win32com.client as win32
from openpyxl import load_workbook


def to_number(raw):
    if raw is None:
        return 0.0
    if isinstance(raw, (int, float)):
        return float(raw)
    if isinstance(raw, str):
        s = raw.strip().replace('\u00A0', '').replace(' ', '')
        if '.' in s and ',' in s:
            if s.rfind(',') > s.rfind('.'):
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '')
        else:
            if ',' in s and '.' not in s:
                s = s.replace(',', '.')
            else:
                s = s.replace(',', '')
        try:
            return float(s)
        except ValueError:
            return 0.0
    return 0.0


def create_kumulativni_izveštaj():
    # Определяем корень проекта
    scripts_dir = Path(__file__).resolve().parent
    root = scripts_dir.parent

    templates_dir = root / "templates"
    template = templates_dir / "izvedeno_template.xlsx"
    output_dir = root / "Output" / "Izvedeno"
    kum_dir = output_dir / "kumulativni izveštaj"
    kum_dir.mkdir(parents=True, exist_ok=True)

    # Проверка шаблона
    if not template.is_file():
        print(f"Ошибка: не найден шаблон: {template}", file=sys.stderr)
        sys.exit(1)

    # Имя итогового файла
    today = datetime.now().strftime("%d.%m.%Y")
    report_name = f"kumulativni izveštaj_{today}.xlsx"
    temp_report = output_dir / report_name

    shutil.copy2(template, temp_report)

    # Словари для суммирования
    cells_rekap = ["F5","F6","F7","F8","F9","F11","F12","F13","D17","F17","F18","F20"]
    totals_rekap = dict.fromkeys(cells_rekap, 0.0)

    cells_ab = [
        "D13","D14","D15","D22","D30","D31","D32","D40","D41","D42",
        "D43","D50","D58","D66","D67","D68","D76","D83","D90","D98",
        "D106","D107","D108","D116","D117","D118","D125","D132","D139",
        "D147","D148","D154","D161","D169","D170"
    ]
    totals_ab = dict.fromkeys(cells_ab, 0.0)

    cells_arm = ["D13","D20","D27"]
    totals_arm = dict.fromkeys(cells_arm, 0.0)

    # COM Excel
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # Проходим по всем файлам кроме текущего отчёта и kum_dir
    for fn in output_dir.glob("*.xlsx"):
        if fn.name == report_name:
            continue
        wb = excel.Workbooks.Open(str(fn), ReadOnly=True)
        try:
            ws = wb.Worksheets("K_00_REKAP")
            for cell in cells_rekap:
                totals_rekap[cell] += to_number(ws.Range(cell).Value)
        except Exception:
            print(f"Предупреждение: нет листа K_00_REKAP в {fn.name}", file=sys.stderr)
        try:
            ws = wb.Worksheets("K_03_AB radovi")
            for cell in cells_ab:
                totals_ab[cell] += to_number(ws.Range(cell).Value)
        except Exception:
            print(f"Предупреждение: нет листа K_03_AB radovi в {fn.name}", file=sys.stderr)
        try:
            ws = wb.Worksheets("K_04_Armiracki")
            for cell in cells_arm:
                totals_arm[cell] += to_number(ws.Range(cell).Value)
        except Exception:
            print(f"Предупреждение: нет листа K_04_Armiracki в {fn.name}", file=sys.stderr)
        wb.Close(False)

    excel.Quit()
    pythoncom.CoUninitialize()

    # Запись итогов
    wb_rep = load_workbook(temp_report)
    for cell, total in totals_rekap.items():
        wb_rep["K_00_REKAP"][cell].value = total
    for cell, total in totals_ab.items():
        wb_rep["K_03_AB radovi"][cell].value = total
    for cell, total in totals_arm.items():
        wb_rep["K_04_Armiracki"][cell].value = total
    wb_rep.save(temp_report)

    # Перемещение в kum_dir
    final_path = kum_dir / report_name
    shutil.move(str(temp_report), str(final_path))
    print(f"Кумулятивный отчёт создан: {final_path}")


if __name__ == "__main__":
    create_kumulativni_izveštaj()