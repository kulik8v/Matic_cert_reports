#!/usr/bin/env python3
"""
main_runner.py

Универсальный скрипт-обёртка (упакованный в exe), который последовательно импортирует и выполняет
визначенные функции из скриптов в папке scripts без рекурсивного запуска exe самого себя.

Скрипты должны лежать рядом с exe в скрытой папке scripts:
  scripts/
    1_journal_update.py        (функция update_journal_total_amount_din)
    2_journal.py               (функция main)
    3_situacija.py             (функция fill_situacija_reports_com)
    4_izvedeno.py              (функция main)
    5_kumulativni izveštaj.py  (функция create_kumulativni_izveštaj)
"""
import sys
import traceback
import importlib.util
from pathlib import Path

def load_and_run(script_path: Path, func_name: str):
    print(f"\n=== Running: {script_path.name} ===")
    if not script_path.exists():
        print(f"⚠ Skipping missing script: {script_path.name}")
        return
    try:
        spec = importlib.util.spec_from_file_location(script_path.stem, str(script_path))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        func = getattr(module, func_name, None)
        if not func:
            raise AttributeError(f"Function '{func_name}' not found in {script_path.name}")
        func()
        print(f"✅ {script_path.name} finished successfully.")
    except Exception as e:
        print(f"❌ Error in {script_path.name}: {e}")
        traceback.print_exc()
        input("Press Enter to exit...")
        sys.exit(1)


def main():
    # Определяем корень проекта: папка, где лежит exe (или скрипт в режиме разработки)
    if getattr(sys, 'frozen', False):
        root = Path(sys.executable).resolve().parent
    else:
        root = Path(__file__).resolve().parent

    scripts_dir = root / 'scripts'
    if not scripts_dir.exists():
        print(f"❌ Папка scripts не найдена: {scripts_dir}")
        sys.exit(1)

    # Явный порядок и соответствие функций
    to_run = [
        ('1_journal_update.py', 'update_journal_total_amount_din'),
        ('2_journal.py', 'main'),
        ('3_situacija.py', 'fill_situacija_reports_com'),
        ('4_izvedeno.py', 'main'),
        ('5_kumulativni izveštaj.py', 'create_kumulativni_izveštaj')
    ]
    print("Будет запущено в порядке:")
    for name, func in to_run:
        print(f"  - {name} -> {func}()")

    # Последовательно загружаем и выполняем
    for name, func in to_run:
        script_path = scripts_dir / name
        load_and_run(script_path, func)

    print("\n🎉 Все скрипты успешно выполнены.")
    input("Press Enter to exit...")

if __name__ == '__main__':
    main()
