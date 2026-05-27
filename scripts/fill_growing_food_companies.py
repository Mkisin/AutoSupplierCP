from __future__ import annotations

from pathlib import Path

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import backend.app as app


ROWS = [
    {
        "companyName": 'ООО "БиоФудЛаб"',
        "inn": "7708753075",
        "companyType": "Производитель конечного товара",
        "contactName": "Шифрина Елена Владимировна",
        "contactPosition": "Генеральный директор",
        "productName": "Полезные снеки Bite",
        "productCategory": "Полезные снеки",
        "productDescription": "Фруктово-ореховые батончики, снеки и продукты здорового питания под брендом Bite.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://bite.ru/",
        "country": "Россия",
        "city": "Москва",
        "preferredNetworks": "Федеральные сети, Маркетплейсы, Специализированные сети",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'ООО "Иван-Поле"',
        "inn": "5024153076",
        "companyType": "Производитель конечного товара",
        "contactName": "Василевская Анна Валерьевна",
        "contactPosition": "Генеральный директор",
        "productName": "Растительные продукты Иван-поле",
        "productCategory": "Растительное питание",
        "productDescription": "Растительные напитки, продукты из сырья, выращенного в экологически чистых районах Калужской области.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://ivan-pole.ru/",
        "country": "Россия",
        "city": "Красногорск",
        "preferredNetworks": "Федеральные сети, Специализированные сети, Маркетплейсы",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'ООО "ВАСТЭКО ПАРТНЕРС"',
        "inn": "5259034951",
        "companyType": "Производитель конечного товара",
        "contactName": "Андреев Дмитрий Николаевич",
        "contactPosition": "Исполнительный директор",
        "productName": "Продукты здорового питания ВАСТЭКО",
        "productCategory": "Здоровое питание",
        "productDescription": "Топленое масло ГХИ, натуральные продукты и товары для здорового питания.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://vasteco.ru/",
        "country": "Россия",
        "city": "Нижний Новгород",
        "preferredNetworks": "Специализированные сети, Маркетплейсы, Региональные сети",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'ООО "ГРИНВАЙЗ"',
        "inn": "4011031618",
        "companyType": "Производитель конечного товара",
        "contactName": "Коммерческий отдел",
        "contactPosition": "Представитель компании",
        "productName": "Растительные альтернативы Greenwise",
        "productCategory": "Растительное мясо",
        "productDescription": "Растительные альтернативы мясным, рыбным и молочным продуктам.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://greenwise.ru/",
        "country": "Россия",
        "city": "Малоярославец",
        "preferredNetworks": "Федеральные сети, Специализированные сети, HoReCa",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'ООО "Котлетарь"',
        "inn": "4401049402",
        "companyType": "Производитель конечного товара",
        "contactName": "Александров Кирилл Сергеевич",
        "contactPosition": "Генеральный директор",
        "productName": "Полуфабрикаты Котлетарь",
        "productCategory": "Мясные полуфабрикаты",
        "productDescription": "Мясные и мясосодержащие полуфабрикаты, котлеты, пельмени и замороженная продукция.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "http://www.kotletar.ru/",
        "country": "Россия",
        "city": "Кострома",
        "preferredNetworks": "Региональные сети, Федеральные сети, Дискаунтеры",
        "priceCategory": "Средний",
    },
    {
        "companyName": 'ООО "ЭкоНива-Продукты Питания"',
        "inn": "3602011911",
        "companyType": "Производитель конечного товара",
        "contactName": "Дюрр Штефан Маттиас",
        "contactPosition": "Генеральный директор",
        "productName": "Молочная продукция ЭкоНива",
        "productCategory": "Молочная продукция",
        "productDescription": "Молоко и молочные продукты от собственных ферм ЭкоНивы.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://www.ekoniva-moloko.com/",
        "country": "Россия",
        "city": "Бобров",
        "preferredNetworks": "Федеральные сети, Региональные сети, Дискаунтеры",
        "priceCategory": "Средний",
    },
    {
        "companyName": 'ООО "Молочная Культура"',
        "inn": "7813409970",
        "companyType": "Производитель конечного товара",
        "contactName": "Бородин Алексей Андреевич",
        "contactPosition": "Генеральный директор",
        "productName": "Молочная продукция Молочная Культура",
        "productCategory": "Молочная продукция",
        "productDescription": "Молоко и кисломолочные продукты бренда Молочная Культура из Ленинградской области.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://dairyculture.ru/",
        "country": "Россия",
        "city": "Сельцо",
        "preferredNetworks": "Федеральные сети, Специализированные сети, Региональные сети",
        "priceCategory": "Премиум",
    },
    {
        "companyName": 'ООО "Чистая Линия"',
        "inn": "5008060096",
        "companyType": "Производитель конечного товара",
        "contactName": "Коммерческий отдел",
        "contactPosition": "Представитель компании",
        "productName": "Мороженое Чистая Линия",
        "productCategory": "Мороженое",
        "productDescription": "Мороженое и молочная продукция бренда Чистая Линия.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://icecream-chl.ru/",
        "country": "Россия",
        "city": "Долгопрудный",
        "preferredNetworks": "Федеральные сети, Региональные сети, Маркетплейсы",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'ООО "Варина Мама"',
        "inn": "3662191440",
        "companyType": "Производитель конечного товара",
        "contactName": "Кузьмин Евгений Викторович",
        "contactPosition": "Генеральный директор",
        "productName": "Кондитерские изделия Варина мама",
        "productCategory": "Кондитерские изделия",
        "productDescription": "Семейная кондитерская, хлебобулочные и мучные кондитерские изделия.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://varinamama.ru/",
        "country": "Россия",
        "city": "Воронеж",
        "preferredNetworks": "Региональные сети, Специализированные сети, Маркетплейсы",
        "priceCategory": "Средний+",
    },
    {
        "companyName": 'ООО "Агропродмаш"',
        "inn": "3662302640",
        "companyType": "Производитель конечного товара",
        "contactName": "Мерный Евгений Александрович",
        "contactPosition": "Генеральный директор",
        "productName": "Макаронные изделия Агропродмаш",
        "productCategory": "Макаронные изделия",
        "productDescription": "Макаронные изделия и мука из собственного зерна.",
        "industry": "Food",
        "activity": "Продукты питания\\напитки (фуд)",
        "networkCategories": "растущий",
        "website": "https://www.agropm.ru/",
        "country": "Россия",
        "city": "Воронеж",
        "preferredNetworks": "Региональные сети, Дискаунтеры, Федеральные сети",
        "priceCategory": "Средний",
    },
]


def main() -> None:
    existing_inns = {card.get("ИНН", "").strip() for card in app._read_cards()}
    added = []
    skipped = []

    for row in ROWS:
        if row["inn"] in existing_inns:
            skipped.append(row["companyName"])
            continue
        added.append((app._append_card(row, ""), row["companyName"]))
        existing_inns.add(row["inn"])

    print(f"added {len(added)}")
    for excel_row, name in added:
        print(excel_row, name)
    if skipped:
        print("skipped", len(skipped), "; ".join(skipped))
    print("total", len(app._read_cards()))


if __name__ == "__main__":
    main()
