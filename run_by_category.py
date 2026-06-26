"""
Запуск парсера для каждой категории отдельно (параллельно).
Сохраняет JSON и XLSX для каждой категории в отдельный файл.
"""
import os
import json
from concurrent.futures import ThreadPoolExecutor, as_completed

from parse_technobuyer import parse_urls_from_file, generate_yandex_kit_xlsx

CATEGORIES = {
    "iphone": "urls_iphone.txt",
    "macbook": "urls_macbook.txt",
    "ipad": "urls_ipad.txt",
    "airpods": "urls_airpods.txt",
    "apple_watch": "urls_apple_watch.txt",
}

OUT_DIR = os.path.dirname(os.path.abspath(__file__))


def _process_category(name, filepath):
    print(f"\n>>> Старт категории: {name}")
    variants = parse_urls_from_file(filepath, max_workers=8)
    print(f"[{name}] Найдено вариаций: {len(variants)}")

    xlsx_path = os.path.join(OUT_DIR, f"yandex-kit-import-{name}.xlsx")
    n = generate_yandex_kit_xlsx(variants, xlsx_path)
    print(f"[{name}] XLSX: {xlsx_path} ({n} строк)")

    json_path = os.path.join(OUT_DIR, f"parsed-data-{name}.json")
    serializable = []
    for v in variants:
        serializable.append({
            "sku": v["sku"],
            "name": v["name"],
            "price": v["price"],
            "old_price": v["old_price"],
            "stock": v["stock"],
            "url": v["url"],
            "features": v["features"],
            "images": v.get("images", []),
        })
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"variants": serializable}, f, ensure_ascii=False, indent=2)
    print(f"[{name}] JSON: {json_path}")
    return name, len(variants)


if __name__ == "__main__":
    print(f"Запуск {len(CATEGORIES)} категорий параллельно...")
    with ThreadPoolExecutor(max_workers=len(CATEGORIES)) as executor:
        futures = {
            executor.submit(_process_category, name, filepath): name
            for name, filepath in CATEGORIES.items()
        }
        for future in as_completed(futures):
            name, count = future.result()
            print(f"  [OK] {name}: {count} вариаций")

    print(f"\n{'=' * 60}")
    print("Готово! Все категории обработаны.")
