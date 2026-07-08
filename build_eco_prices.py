import pandas as pd
import json
import math
from pathlib import Path
from datetime import datetime

BASE_DIR = Path.home()
SOURCE_DIR = BASE_DIR / "source"
OUTPUT_FILE = BASE_DIR / "output" / "eco_prices.json"

excel_files = list(SOURCE_DIR.glob("*.xlsx"))
if not excel_files:
    raise Exception("В папке source нет Excel-файлов")
EXCEL_FILE = excel_files[0]
print(f"Используем файл: {EXCEL_FILE.name}")


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_price(value):
    if pd.isna(value):
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, str):
        v = value.strip().lower()
        if v in ["бесплатно", "free", "0", "—", "-", "цена (руб.)"]:
            return None
    try:
        return float(value)
    except:
        return None


def is_empty_row(row):
    return all(pd.isna(v) or str(v).strip() == "" for v in row[:4])

def is_header_row(row):
    return clean_text(row.iloc[2]).strip() == "Наименование услуги"

def is_total_row(row):
    return clean_text(row.iloc[2]).strip().rstrip() == "Стоимость программы"

def is_note_row(row):
    col0 = clean_text(row.iloc[0])
    col1 = clean_text(row.iloc[1])
    col2 = clean_text(row.iloc[2])
    return len(col0) > 40 and not col1 and not col2

def is_program_header(row):
    code = clean_text(row.iloc[0])
    name = clean_text(row.iloc[2])
    price = clean_price(row.iloc[3])
    # medical_code может быть пустым (как у программы 2.3.2)
    return bool(code) and bool(name) and price is None


def build_eco_programs():
    print("Читаем Excel (лист 'Программы ЭКО')...")
    df = pd.read_excel(EXCEL_FILE, sheet_name="Программы ЭКО", header=None)
    print(f"Найдено строк: {len(df)}")

    programs = []
    current_program = None
    last_program = None  # сохраняем ссылку после завершения, для примечания

    for idx, row in df.iterrows():
        if is_empty_row(row):
            continue
        if is_header_row(row):
            continue

        # Примечание — идёт ПОСЛЕ завершения программы (current_program уже None)
        if is_note_row(row):
            target = current_program or last_program
            if target is not None:
                target["note"] = clean_text(row.iloc[0])
            continue

        # Строка "Стоимость программы"
        if is_total_row(row):
            if current_program:
                current_program["price_full"] = clean_price(row.iloc[3])
                current_program["price_discount"] = clean_price(row.iloc[4]) if len(row) > 4 else None
                programs.append(current_program)
                print(f"  Программа: {current_program['name']} | {current_program['price_full']}")
                last_program = current_program
                current_program = None
            continue

        # Заголовок новой программы
        if is_program_header(row):
            current_program = {
                "internal_code": clean_text(row.iloc[0]),
                "medical_code": clean_text(row.iloc[1]),
                "name": clean_text(row.iloc[2]),
                "price_full": None,
                "price_discount": None,
                "note": None,
                "services": []
            }
            print(f"PROGRAM: {current_program['name']}")
            continue

        # Строка услуги внутри программы
        if current_program is not None:
            name = clean_text(row.iloc[2])
            if not name:
                continue
            price = clean_price(row.iloc[3])
            current_program["services"].append({
                "internal_code": clean_text(row.iloc[0]),
                "medical_code": clean_text(row.iloc[1]),
                "name": name,
                "price": price if price is not None else None
            })

    if current_program:
        programs.append(current_program)
        print(f"WARNING: программа без итоговой строки: {current_program['name']}")

    return programs


def save_json(programs):
    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    result = {
        "generated_at": datetime.now().isoformat(),
        "programs": programs
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"Сохранено: {OUTPUT_FILE}")


if __name__ == "__main__":
    programs = build_eco_programs()
    print(f"\nПрограмм найдено: {len(programs)}")
    save_json(programs)
    print("Готово.")
