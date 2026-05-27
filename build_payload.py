import json
import re
import statistics
import sys
import zipfile
import xml.etree.ElementTree as ET
from argparse import ArgumentParser

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def cell_col(ref):
    return re.sub(r"\d+", "", ref or "")


def col_to_index(col):
    result = 0
    for char in col:
        result = result * 26 + ord(char.upper()) - 64
    return result - 1


def read_xlsx(path):
    with zipfile.ZipFile(path) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", NS):
                shared.append("".join(t.text or "" for t in item.findall(".//a:t", NS)))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall("rel:Relationship", NS)}

        result = {}
        for sheet in workbook.findall(".//a:sheet", NS):
            title = sheet.attrib["name"]
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            sheet_path = "xl/" + relmap[rid].lstrip("/")
            root = ET.fromstring(archive.read(sheet_path))
            rows = []

            for row in root.findall(".//a:row", NS):
                values = []
                for cell in row.findall("a:c", NS):
                    ref = cell.attrib.get("r", "")
                    index = col_to_index(cell_col(ref))
                    while len(values) <= index:
                        values.append("")
                    cell_type = cell.attrib.get("t")
                    value_node = cell.find("a:v", NS)
                    value = "" if value_node is None else value_node.text or ""
                    if cell_type == "s" and value:
                        value = shared[int(value)]
                    elif cell_type == "inlineStr":
                        value = "".join(t.text or "" for t in cell.findall(".//a:t", NS))
                    values[index] = value.strip() if isinstance(value, str) else value
                if any(str(value).strip() for value in values):
                    rows.append(values)
            result[title] = rows
        return result


def to_num(value):
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return None


def normalize_name(value):
    value = str(value).upper()
    value = re.sub(r"[^A-ZА-Я0-9]+", " ", value)
    return " ".join(value.split())


parser = ArgumentParser(description="Build presentation payload from project data.")
parser.add_argument("company", nargs="?", default="ГОСУДАРЕВ СТАНДАРТ")
args = parser.parse_args()

workbook = read_xlsx("Данные для разработки.xlsx")
target_company = args.company

headers = workbook["Карточка клиента"][0]
records = []
for row in workbook["Карточка клиента"][1:]:
    row = row + [""] * (len(headers) - len(row))
    record = dict(zip(headers, row))
    if record.get("Название компании"):
        records.append(record)

target_normalized = normalize_name(target_company)
client_candidates = [
    record for record in records
    if record.get("Название компании", "").upper() == target_company.upper()
    or normalize_name(record.get("Название компании", "")) == target_normalized
    or (
        len(target_normalized) >= 4
        and target_normalized in normalize_name(record.get("Название компании", ""))
    )
]
if not client_candidates:
    raise SystemExit(f"Company not found in Карточка клиента: {target_company}")


def contact_priority(record):
    position = record.get("Должность контакта", "").lower()
    if "коммерческий" in position:
        return 0
    if "генеральный" in position:
        return 1
    if "директор" in position:
        return 2
    return 3


client = sorted(client_candidates, key=contact_priority)[0]

company = client.get("Название компании", "").strip()
contact = client.get("ФИО контакта", "").strip()
position = client.get("Должность контакта", "").strip()
category = client.get("Категория товара", "").strip() or "Молочная продукция Сыры"
product_name = client.get("Название товара", "").strip()
product_photo = client.get("Фото товара", "").strip()
company_type = client.get("Тип компании", "").strip()
price_category = client.get("Ценовая категория", "").strip()

stats_rows = workbook["Статистика по группам"][1:]
stats = []
for row in stats_rows:
    row = row + [""] * 10
    if row[0]:
        stats.append({
            "Группа": row[0],
            "Переговоров": row[1],
            "Интерес": row[3],
            "Чемпионы": row[9],
        })

category_lower = category.lower()
stats_source = "exact"
stat = next(
    (
        record for record in stats
        if record.get("Группа", "").lower() in category_lower
        or category_lower in record.get("Группа", "").lower()
    ),
    None,
)

if not stat:
    category_words = {
        word for word in re.split(r"\W+", category_lower)
        if len(word) >= 5
    }
    scored = []
    for record in stats:
        group_words = {
            word for word in re.split(r"\W+", record.get("Группа", "").lower())
            if len(word) >= 5
        }
        score = len(category_words & group_words)
        if score:
            scored.append((score, record))
    if scored:
        scored.sort(key=lambda item: item[0], reverse=True)
        stat = scored[0][1]
        stats_source = "word_match"

if not stat and ("мясо" in category_lower or "мясопродукт" in category_lower):
    stat = next((record for record in stats if record.get("Группа") == "Мясные деликатесы"), None)
    if stat:
        stats_source = "mapped_from_data_group"

if not stat:
    stats_source = "average_food_benchmark"
    stat = {
        "Переговоров": round(statistics.mean([
            to_num(record.get("Переговоров")) for record in stats
            if to_num(record.get("Переговоров")) is not None
        ])),
        "Интерес": round(statistics.mean([
            to_num(record.get("Интерес")) for record in stats
            if to_num(record.get("Интерес")) is not None
        ]), 2),
        "Чемпионы": round(statistics.mean([
            to_num(record.get("Чемпионы")) for record in stats
            if to_num(record.get("Чемпионы")) is not None
        ])),
    }

def words(value):
    return {
        word
        for word in re.split(r"\W+", str(value).lower())
        if len(word) >= 4
    }


def category_terms(value):
    terms = words(value)
    text = str(value or "").lower()
    if "мяс" in text:
        terms.update({"мясо", "мясные", "мясопродукты", "деликатесы"})
    if "молоч" in text or "сыр" in text:
        terms.update({"молочная", "молоко", "сыр", "сыры", "food"})
    if "круп" in text:
        terms.update({"крупы", "бакалея", "food"})
    if "овощ" in text or "консерв" in text:
        terms.update({"овощная", "консервация", "консервы", "food"})
    if "космет" in text:
        terms.update({"косметика", "уходовая", "beauty"})
    if "живот" in text or "зоотовар" in text:
        terms.update({"зоотовары", "животных", "pet"})
    if "строй" in text or "diy" in text:
        terms.update({"стройтовары", "строительные", "diy"})
    return terms


def split_tags(value):
    return {
        item.strip().lower()
        for item in re.split(r"[;,]", str(value or ""))
        if item.strip()
    }


def universality_rank(value):
    text = str(value or "").strip().lower()
    match = re.match(r"(\d+)", text)
    return int(match.group(1)) if match else 9


def review_candidates(category_value, company_type_value="", price_category_value=""):
    reviews = read_xlsx("\u041e\u0442\u0437\u044b\u0432\u044b.xlsx")
    rows = next(iter(reviews.values()), [])
    if len(rows) <= 1:
        return []

    headers = [str(value).strip() for value in rows[0]]
    category_words = category_terms(category_value)
    company_type_words = words(company_type_value)
    price_tags = split_tags(price_category_value)
    candidates = []

    for index, row in enumerate(rows[1:], start=2):
        row = row + [""] * (len(headers) - len(row))
        record = dict(zip(headers, row))
        review_group = record.get("\u0413\u0440\u0443\u043f\u043f\u0430 \u0442\u043e\u0432\u0430\u0440\u043e\u0432", row[0] if len(row) > 0 else "")
        review_company = record.get("\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u0438", row[1] if len(row) > 1 else "")
        review_person = record.get("\u0424\u0418\u041e", row[2] if len(row) > 2 else "")
        review_text = record.get("\u041e\u0442\u0437\u044b\u0432", row[3] if len(row) > 3 else "")
        review_company_type = record.get("\u0422\u0438\u043f \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u0438", "")
        review_networks = record.get("\u0421\u0435\u0442\u0438 / \u043a\u0430\u043d\u0430\u043b", "")
        review_price_segment = record.get("\u0426\u0435\u043d\u043e\u0432\u043e\u0439 \u0441\u0435\u0433\u043c\u0435\u043d\u0442", "")
        review_universality = record.get("\u0423\u0440\u043e\u0432\u0435\u043d\u044c \u0443\u043d\u0438\u0432\u0435\u0440\u0441\u0430\u043b\u044c\u043d\u043e\u0441\u0442\u0438", "")
        review_keywords = record.get("\u041a\u043b\u044e\u0447\u0435\u0432\u044b\u0435 \u0441\u043b\u043e\u0432\u0430 \u043f\u043e\u0434\u0431\u043e\u0440\u0430", "")
        if not review_company and not review_text:
            continue

        score = 0
        reasons = []
        category_matched = False
        review_group_lower = str(review_group).strip().lower()
        category_lower_value = str(category_value).strip().lower()
        if review_group_lower and (
            review_group_lower == category_lower_value
            or review_group_lower in category_lower_value
            or category_lower_value in review_group_lower
        ):
            score += 100
            category_matched = True
            reasons.append("\u0442\u043e\u0447\u043d\u0430\u044f/\u0447\u0430\u0441\u0442\u0438\u0447\u043d\u0430\u044f \u0442\u043e\u0432\u0430\u0440\u043d\u0430\u044f \u0433\u0440\u0443\u043f\u043f\u0430")

        searchable_category = " ".join([str(review_group), str(review_keywords)])
        category_overlap = category_words & words(searchable_category)
        if category_overlap:
            score += 15 * len(category_overlap)
            category_matched = True
            reasons.append("\u0441\u043e\u0432\u043f\u0430\u043b\u0438 \u0441\u043b\u043e\u0432\u0430 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438: " + ", ".join(sorted(category_overlap)))

        type_overlap = company_type_words & words(review_company_type)
        if type_overlap:
            score += 20
            reasons.append("\u0441\u043e\u0432\u043f\u0430\u043b \u0442\u0438\u043f \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u0438")

        review_price_tags = split_tags(review_price_segment)
        if price_tags and review_price_tags and price_tags & review_price_tags:
            score += 15
            reasons.append("\u0441\u043e\u0432\u043f\u0430\u043b \u0446\u0435\u043d\u043e\u0432\u043e\u0439 \u0441\u0435\u0433\u043c\u0435\u043d\u0442")

        rank = universality_rank(review_universality)
        if rank == 1 and category_matched:
            score += 25
            reasons.append("\u043e\u0442\u0437\u044b\u0432 \u0442\u043e\u0447\u043d\u044b\u0439 \u043f\u043e \u0433\u0440\u0443\u043f\u043f\u0435")
        elif rank == 2 and category_matched:
            score += 10
            reasons.append("\u043e\u0442\u0437\u044b\u0432 \u0431\u043b\u0438\u0437\u043a\u0438\u0439 \u043f\u043e \u043a\u0430\u043d\u0430\u043b\u0443/\u0441\u0435\u0433\u043c\u0435\u043d\u0442\u0443")
        elif rank >= 4:
            score -= 10
            reasons.append("\u0443\u043d\u0438\u0432\u0435\u0440\u0441\u0430\u043b\u044c\u043d\u044b\u0439 fallback")

        searchable_all = " ".join([
            str(review_group), str(review_company), str(review_networks), str(review_keywords), str(review_text)
        ])
        weak_overlap = category_words & words(searchable_all)
        if weak_overlap and not category_overlap:
            score += 5 * len(weak_overlap)
            reasons.append("\u0441\u043b\u0430\u0431\u043e\u0435 \u0441\u043e\u0432\u043f\u0430\u0434\u0435\u043d\u0438\u0435 \u043f\u043e \u0442\u0435\u043a\u0441\u0442\u0443")

        candidates.append({
            "score": score,
            "row": index,
            "company": review_company,
            "person": review_person,
            "text": review_text,
            "group": review_group,
            "company_type": review_company_type,
            "networks": review_networks,
            "network_tags": split_tags(review_networks),
            "price_segment": review_price_segment,
            "universality": review_universality,
            "universality_rank": rank,
            "category_matched": category_matched,
            "reason": "; ".join(reasons) or "\u043d\u0435\u0442 \u0441\u0438\u043b\u044c\u043d\u044b\u0445 \u0441\u043e\u0432\u043f\u0430\u0434\u0435\u043d\u0438\u0439, \u0432\u044b\u0431\u0440\u0430\u043d \u043a\u0430\u043a \u0434\u043e\u043f\u0443\u0441\u0442\u0438\u043c\u044b\u0439 \u043e\u0442\u0437\u044b\u0432",
        })

    candidates.sort(key=lambda item: (-item["score"], item["universality_rank"], item["row"]))
    return candidates


def review_result(candidate, top_candidates):
    source = "algorithm_exact_category" if candidate["category_matched"] else "algorithm_best_match"
    if candidate["universality_rank"] >= 4:
        source = "algorithm_universal_fallback"
    return {
        "company": candidate["company"],
        "person": candidate["person"],
        "text": candidate["text"],
        "source": source,
        "row": candidate["row"],
        "score": candidate["score"],
        "group": candidate["group"],
        "company_type": candidate["company_type"],
        "networks": candidate["networks"],
        "price_segment": candidate["price_segment"],
        "universality": candidate["universality"],
        "reason": candidate["reason"],
        "top_candidates": top_candidates,
    }


def choose_reviews(category_value, company_type_value="", price_category_value="", count=1):
    candidates = review_candidates(category_value, company_type_value, price_category_value)
    if not candidates:
        fallback = {
            "company": "\u0423\u0432\u0435\u043b\u043a\u0430", "person": "", "text": "", "source": "fallback_no_valid_reviews",
            "score": 0, "reason": "\u0412 \u0444\u0430\u0439\u043b\u0435 \u043e\u0442\u0437\u044b\u0432\u043e\u0432 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u043e \u0432\u0430\u043b\u0438\u0434\u043d\u044b\u0445 \u0441\u0442\u0440\u043e\u043a.", "top_candidates": []
        }
        return [fallback]

    selected = []
    used_companies = set()
    used_networks = set()
    phases = [
        lambda item: item["category_matched"] and item["company"].strip().lower() not in used_companies,
        lambda item: item["company"].strip().lower() not in used_companies and not (item["network_tags"] & used_networks),
        lambda item: item["company"].strip().lower() not in used_companies,
        lambda item: True,
    ]

    for phase in phases:
        for item in candidates:
            if len(selected) >= count:
                break
            if item in selected:
                continue
            if not phase(item):
                continue
            selected.append(item)
            if item["company"]:
                used_companies.add(item["company"].strip().lower())
            used_networks.update(item["network_tags"])
        if len(selected) >= count:
            break

    top = [
        {"row": item["row"], "company": item["company"], "group": item["group"], "score": item["score"], "universality": item["universality"], "reason": item["reason"]}
        for item in candidates[:max(3, count)]
    ]
    return [review_result(item, top) for item in selected[:count]]


def choose_review(category_value, company_type_value="", price_category_value=""):
    return choose_reviews(category_value, company_type_value, price_category_value, count=1)[0]


reviews = choose_reviews(category, company_type, price_category, count=3)
while len(reviews) < 3:
    reviews.append({"company": "", "person": "", "text": ""})
review = reviews[0]

payload = {
    "company": company,
    "contact": contact,
    "position": position,
    "category": category,
    "product_name": product_name,
    "product_photo": product_photo,
    "company_type": company_type,
    "price_category": price_category,
    "preferred_networks": client.get("Предпочтительные сети", "").strip(),
    "client_record": client,
    "review_source": review,
    "review_sources": reviews,
    "stats_source": stats_source,
    "replacements": {
        "{{block1}}": f"Как компании {company} получить 5-10 контрактов за 2 дня с розничными сетями, проведя 30-40 переговоров",
        "{{block2}}": category,
        "{{block3}}": str(stat.get("Переговоров", "")),
        "{{block4}}": str(stat.get("Интерес", "")),
        "{{block5}}": str(stat.get("Чемпионы", "")),
        "{{block6}}": review["company"],
        "{{block7}}": review["person"],
        "{{block8}}": review["text"],
        "{{block9}}": reviews[1]["company"],
        "{{block10}}": reviews[1]["person"],
        "{{block11}}": reviews[1]["text"],
        "{{block12}}": reviews[2]["company"],
        "{{block13}}": reviews[2]["person"],
        "{{block14}}": reviews[2]["text"],
    },
}

print(json.dumps(payload, ensure_ascii=False))
