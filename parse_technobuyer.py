"""
Парсер товаров techno-buyer.ru → Яндекс Кит (XLSX) + Яндекс Маркет (YML)
Использование:
  python parse_technobuyer.py "URL-товара"         — автосбор вариантов с одного URL
  python parse_technobuyer.py --file urls.txt      — парсинг списка URL из файла
"""
import sys, re, json, os, html as html_mod
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.parse import urljoin
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom.minidom import parseString

import openpyxl

BASE = "https://techno-buyer.ru"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}


import socket

socket.setdefaulttimeout(30)

def fetch(url):
    req = Request(url, headers=HEADERS)
    with urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8", errors="replace")


def parse_product_page(url):
    text = fetch(url)

    price_match = re.search(r'<meta\s+itemprop="price"\s+content="([^"]+)"', text)
    price = float(price_match.group(1)) if price_match else 0

    cp_match = re.search(r'"compare_price":"([^"]+)"', text)
    old_price = float(cp_match.group(1)) if cp_match else 0

    sku_match = re.search(r'<meta\s+itemprop="sku"\s+content="([^"]+)"', text)
    sku = sku_match.group(1) if sku_match else ""

    pid_match = re.search(r'"product_id":"(\d+)"', text)
    product_id = pid_match.group(1) if pid_match else ""

    stock_match = re.search(r'data-sku-count="(\d+)"', text)
    stock = int(stock_match.group(1)) if stock_match else 0

    name_match = re.search(r'<h1[^>]*>([^<]+)</h1>', text)
    name = html_mod.unescape(name_match.group(1).strip()) if name_match else ""

    features = {}
    for m in re.finditer(
        r'<div class="features-two__name"><span>([^<]+)</span></div>\s*<div class="features-two__value">(.*?)</div>',
        text, re.DOTALL
    ):
        key = html_mod.unescape(m.group(1).strip())
        val = re.sub(r'<[^>]+>', '', m.group(2)).strip()
        val = html_mod.unescape(val)
        features[key] = val

    images = []
    seen_img_urls = set()
    for m in re.finditer(r'href="(/wa-data/public/shop/products/.*?\.(?:jpg|png))"', text):
        img_url = urljoin(BASE, m.group(1))
        if img_url in seen_img_urls:
            continue
        seen_img_urls.add(img_url)
        images.append(img_url)

    desc_match = re.search(r'<meta\s+name="description"\s+content="([^"]+)"', text)
    desc = html_mod.unescape(desc_match.group(1)) if desc_match else ""

    brand_match = re.search(r'<meta\s+itemprop="name"\s+content="([^"]+)"', text)
    brand = brand_match.group(1) if brand_match else "Apple"

    base_slug = re.search(r'/([^/]+)/?(?:index\.php)?$', url)
    base_pattern = ""
    if base_slug:
        parts = base_slug.group(1).split("-")
        base_pattern = "-".join(parts[:3]) + "-" if len(parts) >= 3 else ""

    variant_urls = set()
    for m in re.finditer(r'<a\s+[^>]*class="[^"]*\bproduct-group__item\b[^"]*"[^>]*href="([^"]+)"', text):
        href = m.group(1)
        if not href:
            continue
        if base_pattern and base_pattern not in href:
            continue
        if "pro-max" in href.lower():
            continue
        variant_urls.add(urljoin(BASE, href))

    return {
        "url": url,
        "name": name,
        "price": price,
        "old_price": old_price,
        "sku": sku,
        "product_id": product_id,
        "stock": stock,
        "brand": brand,
        "features": features,
        "images": images,
        "description": desc,
        "variant_urls": variant_urls,
    }


def parse_product(main_url):
    print(f"Парсинг главной страницы...")
    main = parse_product_page(main_url)
    variants = [main]
    parsed_urls = {main_url}

    for vurl in sorted(main["variant_urls"]):
        if vurl in parsed_urls:
            continue
        parsed_urls.add(vurl)
        try:
            print(f"  Парсинг варианта: {vurl}")
            v = parse_product_page(vurl)
            variants.append(v)
        except Exception as e:
            print(f"  [!] Ошибка: {e}")

    return variants


def parse_urls_from_file(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        urls = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]

    if not urls:
        raise SystemExit(f"Файл {filepath} пуст или не содержит URL.")

    print(f"Загружено URL: {len(urls)}")
    variants = []
    parsed_urls = set()

    for i, url in enumerate(urls, 1):
        try:
            print(f"\n[{i}/{len(urls)}] Парсинг: {url}")
            main = parse_product_page(url)
            if main["url"] not in parsed_urls:
                parsed_urls.add(main["url"])
                variants.append(main)

            for vurl in sorted(main["variant_urls"]):
                if vurl in parsed_urls:
                    continue
                parsed_urls.add(vurl)
                try:
                    print(f"  Вариант: {vurl}")
                    v = parse_product_page(vurl)
                    variants.append(v)
                except Exception as e:
                    print(f"  [!] Ошибка варианта: {e}")
        except Exception as e:
            print(f"  [!] Ошибка: {e}")

    return variants


# ====================================================================
# YANDEX KIT XLSX GENERATION
# ====================================================================

def _storage_label(val):
    m = re.search(r'(\d+)', val)
    if not m:
        return val
    gb = int(m.group(1))
    if gb == 1024:
        return "1Tb"
    return f"{gb}Gb"


def _color_label(val):
    return val.strip().capitalize()


def _build_product_name(variant, features, series):
    color = features.get("Цвет", "")
    memory = features.get("Встроенная память", "")
    conn = features.get("Связь", "")
    base = series
    if memory:
        base += " " + _storage_label(memory)
    if conn:
        base += ", " + conn
    base += ", без RuStore"
    if color:
        base += " (" + _color_label(color) + ")"
    return base


def generate_yandex_kit_xlsx(variants, output_path):

    if not variants:
        return 0

    # Шаблон Яндекс Кит — столбцы строго по порядку из официального шаблона
    fieldnames = [
        "KIT ID*", "Название товара*", "Артикул", "Описание товара", "Штрихкод",
        "Статус", "Цена до скидки, руб.", "Цена со скидкой, руб.", "Цена по акции, руб.",
        "Минимальная цена, руб.", "Приоритет в каталоге", "Ставка НДС", "Маркируемый товар",
        "Объединять по (Group ID)", "Группирующие характеристики", "Разделять на карточки по",
        "Бренд",
        "Категория 1-го уровня*", "Категория 2-го уровня", "Категория 3-го уровня",
        "Склад: Склад №1",
        "Внешний ID: 1C", "Внешний ID: МойСклад", "Внешний ID: Wildberries",
        "Внешний ID: Ozon", "Внешний ID: YML", "Внешний ID: Яндекс.Маркет",
        "Изображения и видео", "Ссылки на файлы", "Бейджи",
        "Количество упаковок", "Высота упаковки, см", "Ширина упаковки, см",
        "Длина упаковки, см", "Вес с упаковкой, г",
        # Характеристики
        "Цвет", "Объём встроенной памяти", "Тип связи",
        "Серия", "Процессор", "Диагональ",
        "Основная камера", "Фронтальная камера",
        "ОС", "Защита от воды", "Разъём", "Год модели",
    ]

    group_id = str(int(datetime.now().timestamp() * 1000))
    series = variants[0]["features"].get("Серия", "iPhone 17 Pro")

    wb = openpyxl.Workbook()
    ws = wb.active

    # header row
    for ci, h in enumerate(fieldnames, 1):
        ws.cell(row=1, column=ci, value=h)

    for ri, v in enumerate(variants, 2):
        f = v["features"]
        weight_raw = f.get("Вес", "").replace(" г", "").strip()
        row = [
            "",                             # KIT ID* — пусто для новых
            _build_product_name(v, f, series),  # Название товара*
            v["sku"],                       # Артикул
            v["description"][:5000] if v["description"] else "",  # Описание товара
            "",                             # Штрихкод
            "Опубликован",                  # Статус
            int(v["old_price"]),            # Цена до скидки, руб.
            int(v["price"]),                # Цена со скидкой, руб.
            "",                             # Цена по акции, руб.
            "",                             # Минимальная цена, руб.
            "",                             # Приоритет в каталоге
            22,                             # Ставка НДС
            0,                              # Маркируемый товар
            group_id,                       # Объединять по (Group ID)
            "Цвет; Объём встроенной памяти; Тип связи",  # Группирующие характеристики
            "",                             # Разделять на карточки по
            v["brand"],                     # Бренд
            "Электроника",                  # Категория 1-го уровня*
            "Смартфоны",                    # Категория 2-го уровня
            "Apple iPhone",                 # Категория 3-го уровня
            v["stock"],                     # Склад: Склад №1
            "", "", "",                     # Внешние ID (1C, МойСклад, WB)
            "", "", "",                     # Внешние ID (Ozon, YML, Яндекс.Маркет)
            " ".join(v.get("images", [])[:10]),  # Изображения и видео
            "",                             # Ссылки на файлы
            "",                             # Бейджи
            "",                             # Количество упаковок
            "", "", "",                     # Габариты (высота, ширина, длина)
            weight_raw,                     # Вес с упаковкой, г
            # Характеристики
            f.get("Цвет", ""),
            f.get("Встроенная память", ""),
            f.get("Связь", ""),
            f.get("Серия", ""),
            f.get("Процессор", ""),
            f.get("Диагональ", ""),
            f.get("Разрешение камеры", ""),
            f.get("Разрешение фронтальной камеры, Мп", ""),
            f.get("Операционная система", ""),
            f.get("Защита от воды", ""),
            f.get("Разъём", ""),
            f.get("Год (модель представлена)", ""),
        ]
        for ci, val in enumerate(row, 1):
            ws.cell(row=ri, column=ci, value=val)

    wb.save(output_path)
    return len(variants)


# ====================================================================
# YANDEX MARKET YML GENERATION
# ====================================================================

def _feature_val(features, *keys):
    for k in keys:
        v = features.get(k, "")
        if v:
            return v
    return ""


def generate_yandex_market_yml(variants, output_path):

    if not variants:
        return 0

    main = variants[0]
    features = main["features"]
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+03:00")

    yml = Element("yml_catalog", {"date": now})
    shop = SubElement(yml, "shop")

    SubElement(shop, "name").text = "Techno Buyer"
    SubElement(shop, "company").text = "Techno Buyer"
    SubElement(shop, "url").text = "https://techno-buyer.ru"
    SubElement(shop, "email").text = "info@techno-buyer.ru"
    SubElement(shop, "phone").text = "+7 (909) 964-77-74"

    currencies = SubElement(shop, "currencies")
    SubElement(currencies, "currency", {"id": "RUB", "rate": "1"})

    categories = SubElement(shop, "categories")
    SubElement(categories, "category", {"id": "1"}).text = "Смартфоны"
    SubElement(categories, "category", {"id": "2", "parentId": "1"}).text = "Apple iPhone"
    SubElement(categories, "category", {"id": "3", "parentId": "2"}).text = f"Apple {features.get('Серия', 'iPhone')}"

    group_id = str(int(datetime.now().timestamp() * 1000))
    offers = SubElement(shop, "offers")

    for v in variants:
        f = v["features"]
        offer = SubElement(offers, "offer", {
            "id": v["sku"],
            "available": "true",
            "group_id": group_id,
        })

        SubElement(offer, "url").text = v["url"]
        SubElement(offer, "price").text = str(int(v["price"]))
        SubElement(offer, "oldprice").text = str(int(v["old_price"])) if v["old_price"] else str(int(v["price"]))
        SubElement(offer, "currencyId").text = "RUB"
        SubElement(offer, "categoryId").text = "3"

        for img_url in v["images"][:10]:
            SubElement(offer, "picture").text = img_url

        name = _feature_val(f, "Серия")
        SubElement(offer, "name").text = name if name else v["name"]

        SubElement(offer, "vendor").text = v["brand"]
        SubElement(offer, "vendorCode").text = v["sku"]

        color = f.get("Цвет", "")
        processor = f.get("Процессор", "")
        display = f.get("Диагональ", "")
        display_tech = f.get("Технологии дисплея", "")
        camera = f.get("Разрешение камеры", "")
        desc = f"{name} {_storage_label(f.get('Встроенная память', ''))}, {f.get('Связь', '')}, без RuStore"
        if color:
            desc += f" ({_color_label(color)})"
        desc += "."
        if processor:
            desc += f" Процессор {processor},"
        if display:
            desc += f" дисплей {display}"
        if display_tech:
            desc += f" ({display_tech}),"
        if camera:
            desc += f" камера {camera},"
        desc += " защита IP68."
        SubElement(offer, "description").text = desc

        SubElement(offer, "barcode")
        SubElement(offer, "weight").text = "0.206"

        params = [
            ("Бренд", v["brand"]),
            ("Серия", f.get("Серия", "")),
            ("Модель", f.get("Серия", "")),
            ("Цвет", f.get("Цвет", "")),
            ("Объём встроенной памяти", f.get("Встроенная память", "")),
            ("Тип связи", f.get("Связь", "")),
            ("Диагональ экрана", f.get("Диагональ", "").replace(" дюйм", '"')),
            ("Процессор", f.get("Процессор", "")),
            ("Основная камера", f.get("Разрешение камеры", "")),
            ("Фронтальная камера", f.get("Разрешение фронтальной камеры, Мп", "")),
            ("Защита от воды", f.get("Защита от воды", "")),
            ("ОС", f.get("Операционная система", "")),
            ("Разъём", f.get("Разъём", "")),
            ("Вес", f.get("Вес", "")),
            ("Год модели", f.get("Год (модель представлена)", "")),
        ]
        for pname, pval in params:
            if pval:
                SubElement(offer, "param", {"name": pname}).text = pval

    raw = tostring(yml, encoding="unicode")
    pretty = parseString(raw).toprettyxml(indent="  ", encoding="UTF-8")
    with open(output_path, "wb") as f:
        f.write(pretty)

    return len(variants)


# ====================================================================
# MAIN
# ====================================================================

def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--file":
        variants = parse_urls_from_file(sys.argv[2])
    elif len(sys.argv) > 1 and not sys.argv[1].startswith("--"):
        variants = parse_product(sys.argv[1])
    else:
        url = "https://techno-buyer.ru/iphone-17-pro-1tb-bez-rustore-oranzhevyy-2/"
        variants = parse_product(url)

    print(f"\n{'=' * 60}")
    print(f"Найдено вариаций: {len(variants)}")
    print(f"{'=' * 60}")

    for v in variants:
        f = v["features"]
        color = f.get("Цвет", "?")
        memory = f.get("Встроенная память", "?")
        conn = f.get("Связь", "?")
        print(f"  {v['sku']:12s} | {color:12s} | {memory:8s} | {conn:14s} | {int(v['price']):>8,} RUB | stock: {v['stock']}".replace(",", " "))

    # XLSX для Яндекс Кит
    xlsx_path = "yandex-kit-import.xlsx"
    n = generate_yandex_kit_xlsx(variants, xlsx_path)
    print(f"\nXLSX для Яндекс Кит сохранён: {xlsx_path} ({n} строк)")

    # YML для Яндекс Маркет
    yml_path = "yandex-market-import.yml"
    n2 = generate_yandex_market_yml(variants, yml_path)
    print(f"YML для Яндекс Маркет сохранён: {yml_path} ({n2} офферов)")

    # JSON для справки
    json_path = "parsed-data.json"
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
    print(f"JSON сохранён: {json_path}")

    print("\nГотово!")
    print("  Загрузка в Яндекс Кит:   Товары -> Добавить -> из Excel-файла -> yandex-kit-import.xlsx")
    print("  Загрузка в Яндекс Маркет: yandex-market-import.yml")


if __name__ == "__main__":
    main()
