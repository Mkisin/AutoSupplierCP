from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


OUTPUT = Path("Каталог фотографий сетей.xlsx")
PHOTO_DIR = Path("Фотки переговоров")

HEADERS = [
    "Файл",
    "Сеть",
    "Тип сети",
    "Подходит для групп товаров",
    "Ценовой сегмент",
    "Приоритет",
    "Уровень универсальности",
    "Комментарий",
]

ROWS = [
    ["Фотки переговоров/Азбука.jpg", "Азбука вкуса", "Федеральная; премиальная food-сеть", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food", "Средний+; Премиум", "1", "2 - близкий по каналу/сегменту", "Хорошо подходит для premium food и качественных локальных производителей"],
    ["Фотки переговоров/ВкусВилл.jpg", "ВкусВилл", "Федеральная; специализированная food-сеть", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food", "Средний+; Премиум", "1", "2 - близкий по каналу/сегменту", "Подходит для натуральных, локальных, ЗОЖ и фермерских продуктов"],
    ["Фотки переговоров/Магнит.jpg", "Магнит", "Федеральная сеть", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food; Косметика; Товары для животных", "Эконом; Средний", "1", "3 - универсальный сетевой пример", "Универсальная федеральная сеть для массового рынка"],
    ["Фотки переговоров/Пятерочка.jpg", "Пятерочка", "Федеральная сеть", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food", "Эконом; Средний", "1", "3 - универсальный сетевой пример", "Массовый федеральный food-ритейл"],
    ["Фотки переговоров/Окей.jpg", "ОКЕЙ", "Федеральная сеть; гипермаркет", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food; Товары для животных", "Средний; Средний+", "2", "3 - универсальный сетевой пример", "Хорошо подходит для широкого ассортимента и производителей с большим SKU"],
    ["Фотки переговоров/Ашан.jpg", "Ашан", "Федеральная сеть; гипермаркет", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food; Товары для животных", "Эконом; Средний", "2", "3 - универсальный сетевой пример", "Подходит для массовых категорий и производителей с объемом"],
    ["Фотки переговоров/Metro.jpg", "METRO", "Федеральная сеть; cash&carry; HoReCa", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food", "Средний; Средний+", "2", "2 - близкий по каналу/сегменту", "Подходит для food-производителей и HoReCa-ориентированных продуктов"],
    ["Фотки переговоров/Магнолия.jpg", "Магнолия", "Региональная сеть", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food", "Средний; Средний+", "2", "2 - близкий по каналу/сегменту", "Хороший пример региональной сети для Москвы"],
    ["Фотки переговоров/Квартал.jpg", "Квартал", "Региональная сеть", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food", "Средний", "3", "2 - близкий по каналу/сегменту", "Региональный food-ритейл"],
    ["Фотки переговоров/Вкустер.jpg", "Вкустер", "Региональная food-сеть", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food", "Средний; Средний+", "3", "2 - близкий по каналу/сегменту", "Локальная/региональная food-сеть"],
    ["Фотки переговоров/DNS.jpg", "DNS", "Федеральная non-food сеть", "Стройтовары; DIY; Non-food", "Средний; Средний+", "3", "2 - близкий по каналу/сегменту", "Использовать для non-food, техники, электроники и смежных категорий"],
    ["Фотки переговоров/Familia.jpg", "Familia", "Федеральная off-price сеть", "Одежда; Текстиль; Non-food; Товары для дома", "Эконом; Средний", "3", "2 - близкий по каналу/сегменту", "Подходит для fashion, home goods и non-food категорий"],
    ["Фотки переговоров/Сбермаркет.jpg", "СберМаркет", "Маркетплейс; e-grocery", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Food; Косметика; Товары для животных", "Эконом; Средний; Средний+", "3", "3 - универсальный сетевой пример", "Онлайн-канал и доставка, универсальный пример для разных категорий"],
    ["Фотки переговоров/Газпром.jpg", "Газпром", "АЗС; convenience retail", "Крупы; Мясные деликатесы; Овощная консервация; Food; Товары импульсного спроса", "Средний; Средний+", "4", "3 - универсальный сетевой пример", "Использовать для товаров у дома, convenience и импульсного спроса"],
    ["Фотки переговоров/Татнефть.jpg", "Татнефть", "АЗС; convenience retail", "Крупы; Мясные деликатесы; Овощная консервация; Food; Товары импульсного спроса", "Средний", "4", "3 - универсальный сетевой пример", "Региональный/АЗС-канал, запасной вариант для convenience retail"],
    ["Фотки переговоров/Seller.jpg", "Универсальное фото переговоров", "Универсальное фото", "Крупы; Мясные деликатесы; Овощная консервация; Молочная продукция Сыры; Косметика; Товары для животных; Стройтовары; Food; Non-food", "Эконом; Средний; Средний+; Премиум", "9", "4 - универсальный fallback", "Использовать, если нет подходящего фото конкретной сети"],
]


def col_name(index: int) -> str:
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def build_sheet(rows: list[list[str]]) -> str:
    sheet_rows = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            style = ' s="1"' if row_index == 1 else ""
            cells.append(
                f'<c r="{col_name(col_index)}{row_index}" t="inlineStr"{style}>'
                f"<is><t>{escape(str(value))}</t></is></c>"
            )
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    cols = "".join(
        [
            '<col min="1" max="1" width="36" customWidth="1"/>',
            '<col min="2" max="2" width="24" customWidth="1"/>',
            '<col min="3" max="3" width="28" customWidth="1"/>',
            '<col min="4" max="4" width="56" customWidth="1"/>',
            '<col min="5" max="5" width="26" customWidth="1"/>',
            '<col min="6" max="6" width="12" customWidth="1"/>',
            '<col min="7" max="7" width="30" customWidth="1"/>',
            '<col min="8" max="8" width="58" customWidth="1"/>',
        ]
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:H{len(rows)}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>{cols}</cols>
  <sheetData>{"".join(sheet_rows)}</sheetData>
  <autoFilter ref="A1:H{len(rows)}"/>
</worksheet>'''


def main() -> None:
    existing = {str(path).replace("\\", "/") for path in PHOTO_DIR.glob("*") if path.is_file()}
    rows = [row for row in ROWS if row[0] in existing]
    all_rows = [HEADERS] + rows

    files = {
        "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>''',
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>''',
        "xl/workbook.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Фотографии сетей" sheetId="1" r:id="rId1"/></sheets>
</workbook>''',
        "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>''',
        "xl/worksheets/sheet1.xml": build_sheet(all_rows),
        "xl/styles.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
</styleSheet>''',
        "docProps/core.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Каталог фотографий сетей</dc:title><dc:creator>AutoSupplierCP</dc:creator>
</cp:coreProperties>''',
        "docProps/app.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>AutoSupplierCP</Application></Properties>''',
    }

    with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)
    print(f"created {OUTPUT} with {len(rows)} photo rows")


if __name__ == "__main__":
    main()
