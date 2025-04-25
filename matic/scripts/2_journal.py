#!/usr/bin/env python3
"""
Обрабатывает новые файлы из папки Input:
  - Парсит имя файла для получения номера и даты сертификата
  - Суммирует суммы по столбцу "Amount in certificate with VAT"
  - Копирует шаблоны Situacija и Izvedeno, заполняет их (TODO)
  - Ведёт журнал в templates/journal.xlsx
Скрипт находится в папке scripts внутри корня проекта.
"""
import sys
from pathlib import Path
import pandas as pd
import shutil


def get_processed_files(journal_path: Path):
    cols = [
        "Source File",
        "Certificate Number",
        "Certificate Date",
        "Total Amount",
        "Advance Rate",
        "Total Rate",
        "Total Amount Din",
        "Invoice",
        "Certificate"
    ]
    if journal_path.exists():
        df = pd.read_excel(journal_path)
        missing = [c for c in cols if c not in df.columns]
        if missing:
            print(f"⚠ Warning: journal exists but missing columns: {missing}. Reinitializing journal.")
            df = pd.DataFrame(columns=cols)
    else:
        df = pd.DataFrame(columns=cols)
    processed = set(df["Source File"].astype(str).tolist())
    return processed, df


def parse_filename(fn: str):
    stem = Path(fn).stem
    if "_Progress_certificate_" not in stem:
        raise ValueError(f"Unexpected filename format: {fn}")
    _, rest = stem.split("_Progress_certificate_", 1)
    parts = rest.split("_", 1)
    return parts[0], parts[1]


def extract_total_amount(input_path: Path):
    df = pd.read_excel(input_path, sheet_name="Completion certificate")
    mask = df["Name"].astype(str).str.strip().isin(["Radovi", "RADOVI PO PONUDI"])
    if not mask.any():
        raise ValueError(f"Rows 'Radovi' or 'RADOVI PO PONUDI' not found in {input_path.name}")
    total = 0.0
    for raw in df.loc[mask, "Amount in certificate with VAT"]:
        s = str(raw).replace(" ", "").replace(",", "")
        try:
            val = float(s)
        except ValueError:
            val = float(s.replace(",", "."))
        total += val
    return total


def get_next_index(folder: Path, prefix: str):
    nums = []
    for f in folder.glob(f"{prefix}_*.xlsx"):
        parts = f.stem.split("_")
        if parts[0] == prefix and parts[1].isdigit():
            nums.append(int(parts[1]))
    return max(nums) + 1 if nums else 1


def main():
    # Определяем корень проекта из расположения скрипта
    scripts_dir = Path(__file__).resolve().parent
    root = scripts_dir.parent

    input_dir  = root / "Input"
    out_sit    = root / "Output" / "Situacija"
    out_izv    = root / "Output" / "Izvedeno"
    templates  = root / "templates"
    journal_fp = templates / "journal.xlsx"

    # Создаём папки, если нужно
    for d in (input_dir, out_sit, out_izv, templates):
        d.mkdir(parents=True, exist_ok=True)

    print(f"Working root:      {root}")
    print(f"Input folder:      {input_dir}")
    print(f"Situacija output:  {out_sit}")
    print(f"Izvedeno output:   {out_izv}")
    print(f"Templates folder:  {templates}\n")

    processed, journal_df = get_processed_files(journal_fp)
    print(f"Already processed: {len(processed)} files in journal")

    all_files = sorted(input_dir.glob("*.xlsx"))
    new_files = [f for f in all_files if f.name not in processed]
    print(f"Found {len(all_files)} files, {len(new_files)} new\n")

    next_idx = get_next_index(out_sit, "situacija")

    for src in new_files:
        fn = src.name
        print(f"-> Processing {fn} ...", end=" ")
        try:
            cert_num, cert_date = parse_filename(fn)
            total_eur = extract_total_amount(src)

            # TODO: расчёт курсов и суммы в дин.
            advance_rate = None
            total_rate   = None
            total_din    = None

            idx = next_idx
            next_idx += 1

            invoice_fn = f"situacija_{idx}_{cert_date}.xlsx"
            cert_fn    = f"izvedeno_{idx}_{cert_date}.xlsx"

            # TODO: копировать и заполнять:
            # shutil.copy(templates/"situacija_template.xlsx", out_sit/invoice_fn)
            # shutil.copy(templates/"izvedeno_template.xlsx", out_izv/cert_fn)
            # затем openpyxl для вставки данных

            journal_df.loc[len(journal_df)] = {
                "Source File":        fn,
                "Certificate Number": cert_num,
                "Certificate Date":   cert_date,
                "Total Amount":       total_eur,
                "Advance Rate":       advance_rate,
                "Total Rate":         total_rate,
                "Total Amount Din":   total_din,
                "Invoice":            invoice_fn,
                "Certificate":        cert_fn
            }
            print("OK")
        except Exception as e:
            print(f"ERROR: {e}")

    journal_df.to_excel(journal_fp, index=False)
    print(f"\nJournal saved to {journal_fp}")
    print("=== Done ===")

if __name__ == "__main__":
    main()
