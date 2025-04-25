#!/usr/bin/env python3
"""
Автоматически заполняет отчёты Situacija по журналу:
  - Читает journal.xlsx из templates
  - Вычисляет кумулятивные суммы
  - Находит или копирует шаблон situacija_template.xlsx
  - Заменяет теги в Excel через COM
Скрипт располагается в папке scripts внутри корня проекта.
"""
import sys
import shutil
from pathlib import Path
import pandas as pd
import win32com.client as win32

def fill_situacija_reports_com():
    # Определяем корень проекта из расположения скрипта
    scripts_dir = Path(__file__).resolve().parent
    root = scripts_dir.parent

    templates_dir = root / "templates"
    out_sit = root / "Output" / "Situacija"
    journal_fp = templates_dir / "journal.xlsx"
    template_fp = templates_dir / "situacija_template.xlsx"

    # Проверяем доступность шаблонов и папок
    if not journal_fp.exists():
        print(f"❌ Файл журнала не найден: {journal_fp}", file=sys.stderr)
        sys.exit(1)
    if not template_fp.exists():
        print(f"❌ Шаблон situacija не найден: {template_fp}", file=sys.stderr)
        sys.exit(1)

    out_sit.mkdir(parents=True, exist_ok=True)

    # Читаем журнал и готовим данные
    df = pd.read_excel(journal_fp)
    df["CertDate_dt"] = pd.to_datetime(
        df["Certificate Date"], dayfirst=True, format="%d.%m.%Y"
    )
    df = df.sort_values("CertDate_dt").reset_index(drop=True)

    # Предыдущие и кумулятивные суммы
    df["previous_total"] = df["Total Amount"].cumsum().shift(1).fillna(0)
    df["all_total_on_date"] = df["Total Amount"].cumsum()
    df["Total Amount Din"] = pd.to_numeric(
        df["Total Amount Din"], errors="coerce"
    ).fillna(0)
    df["previous_total_din"] = df["Total Amount Din"].cumsum().shift(1).fillna(0)

    # Маппинг тегов на колонки и набор числовых тегов
    tag_map = {
        "[date]":              "Certificate Date",
        "[number]":            "Certificate Number",
        "[current_total]":     "Total Amount",
        "[previous_total]":    "previous_total",
        "[all_total_on date]": "all_total_on_date",
        "[total_amount_din]":  "previous_total_din",
    }
    numeric_tags = {
        "[current_total]",
        "[previous_total]",
        "[all_total_on date]",
        "[total_amount_din]",
    }

    # Инициализируем COM Excel
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    for _, row in df.iterrows():
        invoice_name = row.get("Invoice")
        dest_fp = out_sit / invoice_name

        # Копируем шаблон, если отчёт ещё не создан
        if not dest_fp.exists():
            shutil.copy(template_fp, dest_fp)

        wb = excel.Workbooks.Open(str(dest_fp))
        wb.Worksheets.Select()

        # Заменяем теги в выделении
        for tag, col in tag_map.items():
            val = row.get(col)
            if pd.isna(val):
                repl_str = ""
            else:
                repl_str = (
                    str(val).replace(".", ",")
                    if tag in numeric_tags
                    else str(val)
                )
            excel.Selection.Replace(
                What=str(tag),
                Replacement=repl_str,
                LookAt=2,        # xlPart
                SearchOrder=1,   # xlByRows
                MatchCase=False
            )

        wb.Save()
        wb.Close(False)

    excel.Quit()
    print(f"✅ All situacija reports updated in {out_sit}")

if __name__ == "__main__":
    fill_situacija_reports_com()
