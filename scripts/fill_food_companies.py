from __future__ import annotations

import os
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import backend.app as app


ROWS = [
    {
        "companyName": 'АО "МАКФА"',
        "inn": "7438015885",
        "companyType": "Производитель конечного товара",
        "contactName": "Коммерческий отдел",
        "contactPosition": "Представитель компании",
        "productName": "Макаронные изделия MAKFA",
        "productCategory": "Макаронные изделия",
        "productDescription": "Макаронные изделия, мука, крупы и зерновые хлопья под брендом MAKFA.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "опытный",
        "website": "https://www.makfa.ru/",
        "country": "Россия",
        "city": "Челябинск",
        "preferredNetworks": "Федеральные сети, Региональные сети, Маркетплейсы",
        "priceCategory": "Средний",
    },
    {
        "companyName": 'ПАО "Группа Черкизово"',
        "inn": "7718560636",
        "companyType": "Производитель конечного товара",
        "contactName": "Михайлов Сергей Игоревич",
        "contactPosition": "Генеральный директор",
        "productName": "Мясная продукция Черкизово",
        "productCategory": "Мясо и мясопродукты",
        "productDescription": "Мясо птицы, свинина, колбасы и мясопереработанная продукция.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "опытный",
        "website": "https://cherkizovo-group.com/",
        "country": "Россия",
        "city": "Кашира",
        "preferredNetworks": "Федеральные сети, Региональные сети, Дискаунтеры",
        "priceCategory": "Средний",
    },
    {
        "companyName": 'АО "Русская Рыбная Компания"',
        "inn": "7701174512",
        "companyType": "Дистрибьютор",
        "contactName": "Дангауэр Дмитрий Сергеевич",
        "contactPosition": "Генеральный директор",
        "productName": "Рыба и морепродукты РРК",
        "productCategory": "Рыба и морепродукты",
        "productDescription": "Охлажденная и свежемороженая рыба, морепродукты, продукты рыбной переработки.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "опытный",
        "website": "https://rusfishcom.ru/",
        "country": "Россия",
        "city": "Москва",
        "preferredNetworks": "Федеральные сети, Региональные сети, HoReCa",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'АО "Сады Придонья"',
        "inn": "3403014273",
        "companyType": "Производитель конечного товара",
        "contactName": 'ООО "УК "Сады Придонья"',
        "contactPosition": "Управляющая компания",
        "productName": "Соки Сады Придонья",
        "productCategory": "Соки и напитки",
        "productDescription": "Соки, сокосодержащие напитки, детское питание и растительные продукты.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "опытный",
        "website": "https://pridonie.ru/",
        "country": "Россия",
        "city": "Сады Придонья",
        "preferredNetworks": "Федеральные сети, Региональные сети, Дискаунтеры",
        "priceCategory": "Средний",
    },
    {
        "companyName": 'ООО "Ресурс"',
        "inn": "7440007056",
        "companyType": "Производитель конечного товара",
        "contactName": "Зяблин Виталий Владимирович",
        "contactPosition": "Директор",
        "productName": "Крупы Увелка",
        "productCategory": "Крупы",
        "productDescription": "Крупы, хлопья и зерновые продукты легкого приготовления под брендом Увелка.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "опытный",
        "website": "https://www.uvelka.ru/",
        "country": "Россия",
        "city": "Увельский",
        "preferredNetworks": "Федеральные сети, Региональные сети, Маркетплейсы",
        "priceCategory": "Средний",
    },
    {
        "companyName": 'ООО "Любятово"',
        "inn": "3661048688",
        "companyType": "Производитель конечного товара",
        "contactName": "Коммерческий отдел",
        "contactPosition": "Представитель компании",
        "productName": "Печенье Любятово",
        "productCategory": "Печенье и готовые завтраки",
        "productDescription": "Печенье, крекеры, хлебцы и готовые завтраки под брендом Любятово.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "опытный",
        "website": "https://www.lubyatovo.ru/",
        "country": "Россия",
        "city": "Вязьма",
        "preferredNetworks": "Федеральные сети, Региональные сети, Дискаунтеры",
        "priceCategory": "Средний",
    },
    {
        "companyName": 'ООО "Кондитерская фабрика "Победа"',
        "inn": "7729390753",
        "companyType": "Производитель конечного товара",
        "contactName": "Коммерческий отдел",
        "contactPosition": "Представитель компании",
        "productName": "Шоколад Победа вкуса",
        "productCategory": "Кондитерские изделия",
        "productDescription": "Шоколад, конфеты, мармелад и сахаристые кондитерские изделия.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://pobedavkusa.ru/",
        "country": "Россия",
        "city": "Москва",
        "preferredNetworks": "Федеральные сети, Региональные сети, Маркетплейсы",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'ООО "ЭФКО Пищевые Ингредиенты"',
        "inn": "3662065051",
        "companyType": "Производитель конечного товара",
        "contactName": "Самченко Константин Владимирович",
        "contactPosition": "Генеральный директор",
        "productName": "Пищевые ингредиенты ЭФКО",
        "productCategory": "Масложировая продукция и ингредиенты",
        "productDescription": "Пищевые ингредиенты, масложировая продукция и продукты для пищевой промышленности.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "опытный",
        "website": "https://www.efko.ru/",
        "country": "Россия",
        "city": "Алексеевка",
        "preferredNetworks": "Федеральные сети, HoReCa, Специализированные сети",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'ООО "Русагро-Сахар"',
        "inn": "7728307368",
        "companyType": "Производитель конечного товара",
        "contactName": "Попова Елена Михайловна",
        "contactPosition": "Генеральный директор",
        "productName": "Сахар Русагро",
        "productCategory": "Сахар",
        "productDescription": "Сахарный песок и сахарная продукция группы Русагро.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "опытный",
        "website": "https://www.rusagrogroup.ru/",
        "country": "Россия",
        "city": "Тамбов",
        "preferredNetworks": "Федеральные сети, Региональные сети, Дискаунтеры",
        "priceCategory": "Эконом",
    },
    {
        "companyName": 'ООО "Томское молоко"',
        "inn": "7014055531",
        "companyType": "Производитель конечного товара",
        "contactName": "Безбородов А. Н.",
        "contactPosition": "Генеральный директор",
        "productName": "Молочная продукция Томское молоко",
        "productCategory": "Молочная продукция",
        "productDescription": "Молоко, сливки, масло и прочая молочная продукция.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://tomskoemoloko.ru/",
        "country": "Россия",
        "city": "Нелюбино",
        "preferredNetworks": "Региональные сети, Специализированные сети",
        "priceCategory": "Средний",
    },
]


def main() -> None:
    names, parts = app._xlsx_parts()
    shared = app._shared_strings(parts)
    sheet_path = app._sheet_target(parts, "Карточка клиента")
    root = ET.fromstring(parts[sheet_path])
    ns = f"{{{app.MAIN_NS}}}"
    sheet_data = root.find(f"{ns}sheetData")
    if sheet_data is None:
        sheet_data = ET.SubElement(root, f"{ns}sheetData")

    app._ensure_headers(root, sheet_data, shared)
    header_row = sheet_data.find(f"{ns}row[@r='1']")
    headers_by_col = app._row_values(header_row, shared)
    col_by_header = {header.strip(): col for col, header in headers_by_col.items()}
    rows_by_inn = {row["inn"]: row for row in ROWS}
    updated = []

    for excel_row in sheet_data.findall(f"{ns}row"):
        if excel_row.attrib.get("r") == "1":
            continue
        values = app._row_values(excel_row, shared)
        inn_col = col_by_header.get("ИНН")
        inn = values.get(inn_col, "").strip() if inn_col else ""
        data = rows_by_inn.pop(inn, None)
        if data is None:
            continue

        for field, header in app.FIELD_TO_HEADER.items():
            col = col_by_header[header.strip()]
            ref = app._cell_ref(col, int(excel_row.attrib["r"]))
            for cell in list(excel_row.findall(f"{ns}c")):
                if cell.attrib.get("r") == ref:
                    excel_row.remove(cell)
            excel_row.append(app._inline_cell(col, int(excel_row.attrib["r"]), data.get(field, "")))
        app._sort_cells(excel_row)
        updated.append((excel_row.attrib["r"], data["companyName"]))

    for data in rows_by_inn.values():
        row_number = app._max_used_row(sheet_data, shared) + 1
        excel_row = ET.Element(f"{ns}row", {"r": str(row_number)})
        row_values = [""] * len(app.HEADERS)
        row_values[0] = str(row_number - 1)
        header_indexes = {header: index for index, header in enumerate(app.HEADERS)}
        for field, header in app.FIELD_TO_HEADER.items():
            row_values[header_indexes[header]] = data.get(field, "")
        for index, value in enumerate(row_values, start=1):
            excel_row.append(app._inline_cell(app._number_to_col(index), row_number, value))
        sheet_data.append(excel_row)
        updated.append((str(row_number), data["companyName"]))

    app._sort_rows(sheet_data)
    dimension = root.find(f"{ns}dimension")
    if dimension is not None:
        dimension.set("ref", f"A1:S{max(app._max_used_row(sheet_data, shared), 1)}")

    parts[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=app.ROOT)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
            for name in names:
                target.writestr(name, parts[name])
        temp_path.replace(app.DATA_FILE)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise

    print(f"updated {len(updated)}")
    for row_number, company in updated:
        print(row_number, company)


if __name__ == "__main__":
    main()
