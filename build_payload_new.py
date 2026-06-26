import json
import os
import re
import statistics
import sys
import zipfile
import xml.etree.ElementTree as ET
from argparse import ArgumentParser
from pathlib import Path

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

LOGO_DIR = Path("Логотипы поставщиков")
LOGO_FIELD_NAMES = (
    "Логотип",
    "Файл логотипа",
    "Логотип компании",
    "Название файла логотипа",
    "logo",
    "logo_file",
)


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

payload_override_raw = os.environ.get("PAYLOAD_OVERRIDE_JSON", "").strip()
if payload_override_raw:
    try:
        payload_override = json.loads(payload_override_raw)
    except json.JSONDecodeError:
        payload_override = {}
    if isinstance(payload_override, dict):
        client = {**client, **{str(key): str(value) for key, value in payload_override.items()}}

company = client.get("Название компании", "").strip()
contact = client.get("ФИО контакта", "").strip()
position = client.get("Должность контакта", "").strip()
category = client.get("Категория товара", "").strip() or "Молочная продукция Сыры"
product_name = client.get("Название товара", "").strip()
product_photo = client.get("Фото товара", "").strip()
company_type = client.get("Тип компании", "").strip()
price_category = client.get("Ценовая категория", "").strip()

stats_headers = [str(value).strip() for value in workbook["Статистика по группам"][0]]
stats_rows = workbook["Статистика по группам"][1:]
stats = []
for row in stats_rows:
    row = row + [""] * (len(stats_headers) - len(row))
    record = dict(zip(stats_headers, row))
    if record.get("Группа"):
        stats.append(record)

def category_parts(value):
    return [
        item.strip()
        for item in re.split(r"[;,\n\r]+", str(value or ""))
        if item.strip()
    ]


def stats_terms(value):
    text = str(value or "").lower()
    terms = {
        word
        for word in re.split(r"\W+", text)
        if len(word) >= 4
    }
    if "молоч" in text or "молок" in text or "сыр" in text or "творог" in text:
        terms.update({"молоко", "молочная", "молочной", "продукция", "сыры", "сыр", "творог"})
    if "мяс" in text or "колбас" in text or "деликатес" in text:
        terms.update({"мясные", "мясо", "мясопродукты", "деликатесы", "колбасные"})
    if "круп" in text or "бакале" in text or "зерно" in text:
        terms.update({"крупы", "бакалея", "зернопродукты"})
    if "космет" in text or "парфюм" in text or "уход" in text:
        terms.update({"косметика", "парфюмерия", "уходовая"})
    if "живот" in text or "зоо" in text or "корм" in text:
        terms.update({"товары", "животных", "зоотовары", "корма"})
    if "строй" in text or "строит" in text or "diy" in text:
        terms.update({"стройтовары", "строительные", "материалы", "diy"})
    return terms


def score_stats_record(category_value, record):
    category_lower = str(category_value or "").lower()
    group = str(record.get("Группа", ""))
    group_lower = group.lower()
    if not category_lower or not group_lower:
        return 0, ""
    if group_lower in category_lower or category_lower in group_lower:
        return 1000, "exact"
    overlap = stats_terms(category_lower) & stats_terms(group_lower)
    if overlap:
        return 100 * len(overlap), "word_match"
    return 0, ""


scored_stats = []
for part_index, part in enumerate(category_parts(category) or [category]):
    for record_index, record in enumerate(stats):
        score, source = score_stats_record(part, record)
        if score:
            scored_stats.append((score, part_index, record_index, source, record))

if scored_stats:
    scored_stats.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, _, _, stats_source, stat = scored_stats[0]
else:
    stat = None
    stats_source = "exact"

category_lower = category.lower()
if not stat and ("мясо" in category_lower or "мясопродукт" in category_lower):
    stat = next((record for record in stats if record.get("Группа") == "Мясные деликатесы"), None)
    if stat:
        stats_source = "mapped_from_data_group"

if not stat:
    stats_source = "average_food_benchmark"
    numeric_fields = [
        "Переговоров",
        "Интерес",
        "Чемпионы",
        "Контрактов",
        "Заключенных контрактов на одного поставщика",
        "Закупщики по вашей категории",
        "Федеральных сетей",
        "Локальный сетей",
        "региональных сетей",
        "Хорека",
        "Контрактов всего",
        "2025 переговоров",
        "2025 контрактов",
    ]
    stat = {
        field: round(statistics.mean([
            to_num(record.get(field)) for record in stats
            if to_num(record.get(field)) is not None
        ]), 2)
        for field in numeric_fields
        if any(to_num(record.get(field)) is not None for record in stats)
    }


def first_name(value):
    parts = [part for part in re.split(r"\s+", str(value or "").strip()) if part]
    patronymic_index = next(
        (
            index
            for index, part in enumerate(parts)
            if re.search(r"(ович|евич|ич|овна|евна|ична|инична)$", part.lower())
        ),
        None,
    )
    if patronymic_index is not None and patronymic_index > 0:
        return parts[patronymic_index - 1]
    return parts[0] if parts else ""


def first_name_patronymic(value):
    parts = [part for part in re.split(r"\s+", str(value or "").strip()) if part]
    patronymic_index = next(
        (
            index
            for index, part in enumerate(parts)
            if re.search(r"(ович|евич|ич|овна|евна|ична|инична)$", part.lower())
        ),
        None,
    )
    if patronymic_index is not None and patronymic_index > 0:
        return " ".join(parts[patronymic_index - 1:patronymic_index + 1])
    return " ".join(parts[:2]) if len(parts) >= 2 else (parts[0] if parts else "")


def quoted_category(value):
    text = str(value or "").strip()
    if not text:
        return ""
    if text.startswith("«") and text.endswith("»"):
        return text
    return f"«{text}»"


def recipient_block(contact_value, position_value):
    contact_value = str(contact_value or "").strip()
    position_value = str(position_value or "").strip()
    suffix = ", ".join(item for item in (contact_value, position_value) if item)
    return f"Получатель {suffix}".strip()


def first_slide_title(company_value):
    company_value = str(company_value or "").strip()
    return (
        f"Как компании «{company_value}» получить 5-10 контрактов "
        "с розничными сетями на Центре Закупок Сетей: WorldFood 2026"
    )


def stat_value(field, allow_decimal=False):
    value = stat.get(field, "")
    text = str(value).strip()
    if not text:
        return ""
    number_text = text.replace(" ", "").replace(",", ".")
    try:
        number = float(number_text)
    except ValueError:
        return text
    if allow_decimal:
        return f"{number:g}"
    return str(int(round(number)))


def average_check_value(field):
    value = stat.get(field, "")
    text = str(value).strip()
    if not text:
        return ""
    number_text = text.replace(" ", "").replace(",", ".")
    try:
        number = float(number_text)
    except ValueError:
        return text.replace(",", ".")
    return f"{number:.1f}"

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
        terms.update({"мясо", "мясные", "мясопродукты", "деликатесы", "food"})
    if "молоч" in text or "сыр" in text:
        terms.update({"молочная", "молоко", "сыр", "сыры", "food"})
    if "круп" in text:
        terms.update({"крупы", "бакалея", "food"})
    if "снек" in text or "батончик" in text or "зож" in text or "полезн" in text:
        terms.update({"снеки", "батончики", "перекус", "зож", "бакалея", "крупы", "food"})
    if "овощ" in text or "консерв" in text:
        terms.update({"овощная", "консервация", "консервы", "food"})
    if "космет" in text:
        terms.update({"косметика", "уходовая", "beauty"})
    if "живот" in text or "зоотовар" in text:
        terms.update({"зоотовары", "животных", "pet"})
    if "строй" in text or "diy" in text:
        terms.update({"стройтовары", "строительные", "diy"})
    return terms


def category_affinity(category_value, review_group):
    category_text = str(category_value or "").lower()
    group_text = str(review_group or "").lower()
    if not category_text or not group_text:
        return 0, ""
    if group_text in category_text or category_text in group_text:
        return 100, "точная/частичная товарная группа"

    snack_like = any(marker in category_text for marker in ["снек", "батончик", "зож", "полезн"])
    if snack_like:
        if "круп" in group_text:
            return 70, "близкая food-группа: снеки/батончики ближе к бакалее и крупам"
        if "овощ" in group_text or "консерв" in group_text:
            return 35, "смежная shelf-stable food-группа"
        if "мяс" in group_text:
            return 10, "дальняя food-группа"

    food_category = any(marker in category_text for marker in [
        "food", "питан", "продукт", "снек", "батончик", "круп", "мяс", "молоч", "сыр", "овощ", "консерв"
    ])
    food_group = any(marker in group_text for marker in ["круп", "мяс", "молоч", "сыр", "овощ", "консерв"])
    if food_category and food_group:
        return 20, "общая food-группа"
    return 0, ""


def split_tags(value):
    return {
        item.strip().lower()
        for item in re.split(r"[;,]", str(value or ""))
        if item.strip()
    }


def review_logo_file(record):
    for field_name in LOGO_FIELD_NAMES:
        value = str(record.get(field_name, "") or "").strip()
        if value:
            return value
    return ""


def review_logo_path(value):
    value = str(value or "").strip()
    if not value:
        return ""
    path = Path(value.replace("\\", "/"))
    if not path.is_absolute() and len(path.parts) == 1:
        path = LOGO_DIR / path
    search_name = path.name.lower()
    search_stem = path.stem.lower()
    search_root = path.parent if path.parent != Path(".") else LOGO_DIR
    if search_root.exists():
        for match in search_root.rglob("*"):
            if not match.is_file():
                continue
            if match.name.lower() == search_name or match.stem.lower() == search_stem:
                return str(match)
    return str(path)


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
        logo_file = review_logo_file(record)
        if not review_company and not review_text:
            continue

        score = 0
        reasons = []
        category_matched = False
        review_group_lower = str(review_group).strip().lower()
        category_lower_value = str(category_value).strip().lower()
        affinity_score, affinity_reason = category_affinity(category_value, review_group)
        if affinity_score:
            score += affinity_score
            category_matched = affinity_score >= 50
            reasons.append(affinity_reason)

        if review_group_lower and (
            review_group_lower == category_lower_value
            or review_group_lower in category_lower_value
            or category_lower_value in review_group_lower
        ):
            category_matched = True
            if affinity_score < 100:
                score += 100
                reasons.append("точная/частичная товарная группа")
        elif affinity_score and affinity_score < 50:
            score -= 8
            reasons.append("понижен приоритет дальней товарной группы")

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
            "logo_file": logo_file,
            "logo_path": review_logo_path(logo_file),
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
        "logo_file": candidate.get("logo_file", ""),
        "logo_path": candidate.get("logo_path", ""),
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


def review_option_payload(candidate):
    return {
        "id": f"review-row-{candidate['row']}",
        "row": candidate["row"],
        "company": candidate["company"],
        "person": candidate["person"],
        "text": candidate["text"],
        "logo_file": candidate.get("logo_file", ""),
        "logo_path": candidate.get("logo_path", ""),
        "group": candidate["group"],
        "company_type": candidate["company_type"],
        "networks": candidate["networks"],
        "price_segment": candidate["price_segment"],
        "universality": candidate["universality"],
        "score": candidate["score"],
        "reason": candidate["reason"],
        "category_matched": candidate["category_matched"],
    }


def select_review_candidates(candidates, count):
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

    return selected[:count]


def choose_review_option_groups(category_value, company_type_value="", price_category_value="", count=3):
    candidates = review_candidates(category_value, company_type_value, price_category_value)
    if not candidates:
        groups = []
        for slot in range(1, count + 1):
            fallback = {
                "id": f"review-fallback-{slot}",
                "row": 0,
                "company": "Увелка",
                "person": "",
                "text": "",
                "group": category_value,
                "company_type": company_type_value,
                "networks": "",
                "price_segment": price_category_value,
                "universality": "",
                "score": 0,
                "reason": "В файле отзывов не найдено валидных строк.",
                "category_matched": False,
            }
            groups.append(
                {
                    "slot": f"review{slot}",
                    "label": f"Отзыв {slot}",
                    "selectedOptionId": fallback["id"],
                    "options": [fallback],
                }
            )
        return groups

    primaries = select_review_candidates(candidates, count)
    used_rows = {item["row"] for item in primaries}
    used_alternates = set()
    groups = []

    for slot, primary in enumerate(primaries, start=1):
        options = [review_option_payload(primary)]
        primary_company = primary["company"].strip().lower()
        primary_networks = primary["network_tags"]
        alternate = None
        phases = [
            lambda item: item["row"] not in used_rows and item["row"] not in used_alternates and item["company"].strip().lower() != primary_company and not (item["network_tags"] & primary_networks),
            lambda item: item["row"] not in used_rows and item["row"] not in used_alternates and item["company"].strip().lower() != primary_company,
            lambda item: item["row"] not in used_rows and item["row"] not in used_alternates,
            lambda item: item["row"] != primary["row"],
        ]
        for phase in phases:
            alternate = next((item for item in candidates if phase(item)), None)
            if alternate:
                break
        if alternate:
            used_alternates.add(alternate["row"])
            options.append(review_option_payload(alternate))
        groups.append(
            {
                "slot": f"review{slot}",
                "label": f"Отзыв {slot}",
                "selectedOptionId": options[0]["id"],
                "options": options,
            }
        )

    return groups


def choose_reviews(category_value, company_type_value="", price_category_value="", count=1):
    candidates = review_candidates(category_value, company_type_value, price_category_value)
    if not candidates:
        fallback = {
            "company": "\u0423\u0432\u0435\u043b\u043a\u0430", "person": "", "text": "", "source": "fallback_no_valid_reviews",
            "score": 0, "reason": "\u0412 \u0444\u0430\u0439\u043b\u0435 \u043e\u0442\u0437\u044b\u0432\u043e\u0432 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u043e \u0432\u0430\u043b\u0438\u0434\u043d\u044b\u0445 \u0441\u0442\u0440\u043e\u043a.", "top_candidates": []
        }
        return [fallback]

    selected = select_review_candidates(candidates, count)

    top = [
        {"row": item["row"], "company": item["company"], "group": item["group"], "score": item["score"], "universality": item["universality"], "reason": item["reason"]}
        for item in candidates[:max(3, count)]
    ]
    return [review_result(item, top) for item in selected[:count]]


def choose_review(category_value, company_type_value="", price_category_value=""):
    return choose_reviews(category_value, company_type_value, price_category_value, count=1)[0]


all_review_candidates = review_candidates(category, company_type, price_category)
selected_review_candidates = select_review_candidates(all_review_candidates, 3)
review_candidate_pool = []
seen_review_rows = set()
for candidate in selected_review_candidates + all_review_candidates:
    row_number = candidate.get("row")
    if row_number in seen_review_rows:
        continue
    review_candidate_pool.append(review_option_payload(candidate))
    seen_review_rows.add(row_number)
    if len(review_candidate_pool) >= 6:
        break

reviews = choose_reviews(category, company_type, price_category, count=3)
while len(reviews) < 3:
    reviews.append({"company": "", "person": "", "text": ""})
review = reviews[0]
selected_review_ids = [f"review-row-{item['row']}" for item in selected_review_candidates[:3]]

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
    "review_candidates": review_candidate_pool,
    "selected_review_ids": selected_review_ids,
    "required_review_count": 3,
    "stats_source": stats_source,
    "stats_record": stat,
    "replacements": {
        "{{block1}}": recipient_block(contact, position),
        "{{block1 }}": recipient_block(contact, position),
        "{{block2}}": first_slide_title(company),
        "{{block3}}": quoted_category(category),
        "{{block4}}": f"переговоров с сетями по категории {quoted_category(category)}",
        "{{block5}}": f"{first_name_patronymic(contact)}, ознакомьтесь с итогами последнего ЦЗС™ по категории {quoted_category(category)}",
        "{{block6}}": stat_value("Контрактов"),
        "{{block8}}": stat_value("Заключенных контрактов на одного поставщика"),
        "{{block9}}": stat_value("Закупщики по вашей категории"),
        "{{block10}}": stat_value("Федеральных сетей"),
        "{{block11}}": stat_value("региональных сетей"),
        "{{block12}}": stat_value("Локальный сетей"),
        "{{block13}}": stat_value("Хорека"),
        "{{block14}}": stat_value("Контрактов всего"),
        "{{block15}}": reviews[0]["company"],
        "{{block16}}": reviews[0]["person"],
        "{{block17}}": reviews[0]["text"],
        "{{block18}}": reviews[1]["company"],
        "{{block19}}": reviews[1]["person"],
        "{{block20}}": reviews[1]["text"],
        "{{block21}}": reviews[2]["company"],
        "{{block22}}": reviews[2]["person"],
        "{{block23}}": reviews[2]["text"],
        "{{block24}}": first_name(contact),
        "{{block25}}": stat_value("2025 переговоров"),
        "{{block26}}": stat_value("2025 контрактов"),
        "{{block31}}": average_check_value("Средний чек"),
    },
}

print(json.dumps(payload, ensure_ascii=False))
