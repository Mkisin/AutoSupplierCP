import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from argparse import ArgumentParser
from pathlib import Path


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

NS = {"p": P, "a": A, "r": R, "rel": REL, "ct": CT}
IMAGE_TOKENS = [
    "{{pic2}}", "{{pic3}}", "{{pic4}}", "{{pic5}}", "{{pic6}}",
    "{{logo1}}", "{{logo2}}", "{{logo3}}", "{{logo4}}", "{{logo5}}", "{{logo6}}",
    "{{logo7}}", "{{logo8}}", "{{logo9}}", "{{logo10}}", "{{logo11}}",
]
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg"}

ET.register_namespace("p", P)
ET.register_namespace("a", A)
ET.register_namespace("r", R)
ET.register_namespace("", REL)

BASE_DIR = Path(__file__).resolve().parent


def qn(namespace, tag):
    return f"{{{namespace}}}{tag}"


def project_path(value):
    path = Path(str(value or "").strip().replace("\\", "/"))
    if path.is_absolute():
        return path
    return BASE_DIR / path


def read_payload(company):
    import subprocess

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    command = [sys.executable, str(BASE_DIR / "build_payload_new.py")]
    if company:
        command.append(company)
    result = subprocess.check_output(
        command,
        cwd=BASE_DIR,
        text=True,
        encoding="utf-8",
        env=env,
    )
    return json.loads(result)


def read_selection_override():
    raw = os.environ.get("PRESENTATION_SELECTION_JSON", "").strip()
    if not raw:
        return {}
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    return payload if isinstance(payload, dict) else {}


def image_ext(path):
    if not path.exists() or not path.is_file():
        return None
    header = path.read_bytes()[:16]
    if header.startswith(b"\xff\xd8\xff"):
        return ".jpg"
    if header.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if header.startswith(b"GIF8"):
        return ".gif"
    if header.startswith(b"RIFF") and b"WEBP" in header:
        return ".webp"
    if path.suffix.lower() == ".svg":
        return ".svg"
    return None


def is_image_filename(value):
    return Path(str(value or "")).suffix.lower() in IMAGE_SUFFIXES


def is_powerpoint_safe_image(path):
    return image_ext(path) in {".jpg", ".jpeg", ".png", ".gif", ".svg"}


def convert_to_png(source, destination):
    try:
        from PIL import Image, ImageOps

        with Image.open(source) as image:
            image = ImageOps.exif_transpose(image)
            if image.mode not in {"RGB", "RGBA"}:
                image = image.convert("RGBA" if "A" in image.getbands() else "RGB")
            image.save(destination, "PNG")
        return destination.exists() and destination.stat().st_size > 0 and image_ext(destination) == ".png"
    except ImportError:
        pass
    except Exception:
        if os.name != "nt":
            raise

    env = os.environ.copy()
    env["PPT_IMAGE_SOURCE"] = str(source.resolve())
    env["PPT_IMAGE_DESTINATION"] = str(destination.resolve())
    command = [
        "powershell",
        "-NoProfile",
        "-Command",
        (
            "$src=$env:PPT_IMAGE_SOURCE; "
            "$dst=$env:PPT_IMAGE_DESTINATION; "
            "Add-Type -AssemblyName PresentationCore; "
            "$uri=[System.Uri]::new($src); "
            "$decoder=[System.Windows.Media.Imaging.BitmapDecoder]::Create("
            "$uri, "
            "[System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, "
            "[System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad"
            "); "
            "$encoder=[System.Windows.Media.Imaging.PngBitmapEncoder]::new(); "
            "$encoder.Frames.Add($decoder.Frames[0]); "
            "$fs=[System.IO.File]::OpenWrite($dst); "
            "$encoder.Save($fs); "
            "$fs.Close()"
        ),
    ]
    subprocess.run(command, check=True, capture_output=True, text=True, env=env)
    return destination.exists() and destination.stat().st_size > 0 and image_ext(destination) == ".png"


def normalize_images_for_powerpoint(image_mapping):
    normalized = {}
    report = {}
    temp_files = []
    for index, (token, source) in enumerate(image_mapping.items(), start=1):
        fd, target_name = tempfile.mkstemp(prefix=f"ppt_image_{index}_", suffix=".png", dir=BASE_DIR)
        os.close(fd)
        target = Path(target_name)
        target.unlink(missing_ok=True)
        temp_files.append(target)
        original_ext = image_ext(source)
        try:
            if convert_to_png(source, target):
                normalized[token] = target
                report[token] = {
                    "source": str(source),
                    "used": str(target),
                    "source_format": original_ext,
                    "used_format": ".png",
                    "normalized": True,
                }
                continue
        except Exception as exc:
            if is_powerpoint_safe_image(source):
                normalized[token] = source
                report[token] = {
                    "source": str(source),
                    "used": str(source),
                    "source_format": original_ext,
                    "used_format": original_ext,
                    "normalized": False,
                    "warning": f"PNG normalization failed, original inserted: {exc}",
                }
            else:
                report[token] = {
                    "source": str(source),
                    "used": "",
                    "source_format": original_ext,
                    "used_format": "",
                    "normalized": False,
                    "error": f"Unsupported image format and conversion failed: {exc}",
                }
            continue

        if is_powerpoint_safe_image(source):
            normalized[token] = source
            report[token] = {
                "source": str(source),
                "used": str(source),
                "source_format": original_ext,
                "used_format": original_ext,
                "normalized": False,
            }
        else:
            report[token] = {
                "source": str(source),
                "used": "",
                "source_format": original_ext,
                "used_format": "",
                "normalized": False,
                "error": "Unsupported image format",
            }
    return normalized, report, temp_files


def first_file(pattern):
    files = sorted(BASE_DIR.glob(pattern))
    return next((path for path in files if image_ext(path)), None)


def data_path(value):
    return project_path(value)


def iter_images(folder):
    folder_path = project_path(folder)
    if not folder_path.exists():
        return []
    return sorted(
        [
            path
            for path in folder_path.rglob("*")
            if path.is_file() and image_ext(path)
        ],
        key=lambda path: str(path).lower(),
    )


def resolve_image_path(path):
    candidate = data_path(path)
    if image_ext(candidate):
        return candidate

    search_roots = []
    if len(candidate.parts) > 1:
        search_roots.append(candidate.parent)
    search_roots.append(BASE_DIR)

    seen = set()
    for root in search_roots:
        if not root.exists():
            continue
        root_key = str(root.resolve()).lower()
        if root_key in seen:
            continue
        seen.add(root_key)
        for match in root.rglob(candidate.name):
            if image_ext(match):
                return match
        if not candidate.suffix:
            for match in root.rglob("*"):
                if match.is_file() and match.stem.lower() == candidate.name.lower() and image_ext(match):
                    return match
    return candidate


def first_images(folder, count):
    return iter_images(folder)[:count]


def cell_col(ref):
    return re.sub(r"\d+", "", ref or "")


def col_to_index(col):
    result = 0
    for char in col:
        result = result * 26 + ord(char.upper()) - 64
    return result - 1


def read_xlsx(path):
    path = project_path(path)
    if not path.exists():
        return {}
    sheet_ns = {
        "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    with zipfile.ZipFile(path) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", sheet_ns):
                shared.append("".join(t.text or "" for t in item.findall(".//a:t", sheet_ns)))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall("rel:Relationship", sheet_ns)}

        result = {}
        for sheet in workbook.findall(".//a:sheet", sheet_ns):
            title = sheet.attrib["name"]
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            sheet_path = "xl/" + relmap[rid].lstrip("/")
            root = ET.fromstring(archive.read(sheet_path))
            rows = []
            for row in root.findall(".//a:row", sheet_ns):
                values = []
                for cell in row.findall("a:c", sheet_ns):
                    index = col_to_index(cell_col(cell.attrib.get("r", "")))
                    while len(values) <= index:
                        values.append("")
                    if cell.attrib.get("t") == "inlineStr":
                        value = "".join(t.text or "" for t in cell.findall(".//a:t", sheet_ns))
                    else:
                        value_node = cell.find("a:v", sheet_ns)
                        value = "" if value_node is None else value_node.text or ""
                        if cell.attrib.get("t") == "s" and value:
                            value = shared[int(value)]
                    values[index] = value.strip() if isinstance(value, str) else value
                if any(str(value).strip() for value in values):
                    rows.append(values)
            result[title] = rows
        return result


def normalize(value):
    value = str(value).lower()
    value = re.sub(r"\b(ооо|ао|пао|зао|оао|ип)\b", " ", value)
    value = re.sub(r"[^a-zа-я0-9]+", " ", value)
    return " ".join(value.split())


LOGO_ALIASES = {
    "«деликатесофф»": ["деликатессов", "деликатесофф"],
    "деликатесофф": ["деликатессов", "деликатесофф"],
    "natura siberica": ["сиберика", "natura siberica"],
    "кзвс краснодарзооветснаб": ["кзвс", "краснодарзооветснаб"],
    "кзвс": ["кзвс", "краснодарзооветснаб"],
    "инсепт": ["insept", "инсепт"],
    "td альянс": ["альянс"],
    "тд альянс": ["альянс"],
    "qharisma": ["qharisma", "harly"],
    "русское поле": ["русское поле", "уни пак", "уни-пак"],
    "райпищекомбинат мостовский": ["мостовский"],
    "экспресс кубань": ["экспресс кубань"],
    "мельник": ["melnik", "мельник"],
    "гельтек": ["geltek", "гельтек"],
    "titbit": ["titbit", "титбит"],
    "увелка": ["uvelka", "увелка"],
}


def logo_needles(company):
    cleaned = str(company or "").strip().strip("«»\"“”")
    normalized = normalize(cleaned)
    needles = [cleaned, normalized]
    needles.extend(LOGO_ALIASES.get(normalized, []))
    return [item for item in needles if item]


def score_file(path, needles):
    haystack = normalize(path.stem)
    score = 0
    for needle in needles:
        normalized = normalize(needle)
        if not normalized:
            continue
        if normalized == haystack:
            score += 100
        elif normalized in haystack or haystack in normalized:
            score += 50
        else:
            score += len(set(normalized.split()) & set(haystack.split())) * 10
    return score


def best_image(folder, needles):
    candidates = iter_images(folder)
    scored = [(score_file(path, needles), path) for path in candidates]
    scored = [item for item in scored if item[0] > 0]
    if not scored:
        return None
    scored.sort(key=lambda item: (-item[0], item[1].name))
    return scored[0][1]


def review_logo_image(review):
    for key in ("logo_path", "logo_file", "logo"):
        value = str((review or {}).get(key, "") or "").strip()
        if not value:
            continue
        path = resolve_image_path(value)
        if not image_ext(path) and len(data_path(value).parts) == 1:
            path = resolve_image_path(Path("Логотипы поставщиков") / value)
        if image_ext(path):
            return path

    company = str((review or {}).get("company", "") or "").strip()
    if not company:
        return None
    return best_image("Логотипы поставщиков", logo_needles(company))


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
    if "снек" in text or "батончик" in text or "зож" in text:
        terms.update({"снеки", "перекус", "зож", "food"})
    if "овощ" in text or "консерв" in text:
        terms.update({"овощная", "консервация", "консервы", "food"})
    if "космет" in text:
        terms.update({"косметика", "уходовая", "beauty"})
    if "живот" in text or "зоотовар" in text:
        terms.update({"зоотовары", "животных", "pet"})
    if "строй" in text or "diy" in text:
        terms.update({"стройтовары", "строительные", "diy", "non", "food"})
    return terms


def split_tags(value):
    return {
        item.strip().lower()
        for item in re.split(r"[;,\n\r]+", str(value or ""))
        if item.strip()
    }


def split_ordered(value):
    return [
        item.strip()
        for item in re.split(r"[;,\n\r]+", str(value or ""))
        if item.strip()
    ]


NETWORK_ALIASES = {
    "азбука вкуса": {"азбука", "azbuka", "azbuka vkusa"},
    "ашан": {"auchan"},
    "вкусвилл": {"вкус вилл", "vkusvill", "vkus vill"},
    "магнит": {"magnit"},
    "магнит сзфо": {"магнит спб", "магнит питер"},
    "метро": {"metro", "метро кэш энд керри", "metro cash carry"},
    "metro": {"метро", "метро кэш энд керри", "metro cash carry"},
    "окей": {"о кей", "o key", "okay", "okey", "оке"},
    "пятерочка": {"пятёрочка", "5ка", "5 ка", "пятерка", "пятёрка", "x5", "х5", "x5 group", "х5 group"},
    "перекресток": {"перекрёсток", "perekrestok", "x5", "х5", "x5 group", "х5 group"},
    "сбермаркет": {"сбер маркет", "sbermarket", "sber market", "купер", "kuper"},
}


def network_match_keys(value):
    normalized = normalize(value)
    keys = {normalized} if normalized else set()
    if not normalized:
        return keys

    for canonical, aliases in NETWORK_ALIASES.items():
        alias_values = {canonical, *aliases}
        alias_keys = {normalize(item) for item in alias_values if normalize(item)}
        if normalized in alias_keys or any(key and (key in normalized or normalized in key) for key in alias_keys):
            keys.update(alias_keys)
            keys.add(normalize(canonical))
    return {key for key in keys if key}


def preferred_network_matches(record, preferred_networks):
    network_name = str(record.get("Сеть", ""))
    if not network_name or not preferred_networks:
        return False
    normalized_name = normalize(network_name)
    network_keys = network_match_keys(network_name)
    preferred_keys = {
        key
        for item in preferred_networks
        for key in network_match_keys(item)
    }
    if network_keys & preferred_keys:
        return True
    network_words = words(network_name)
    return any(
        normalized_item in normalized_name
        or normalized_name in normalized_item
        or bool(network_words & words(item))
        for item in preferred_networks
        for normalized_item in [normalize(item)]
        if normalized_item
    )


def number_prefix(value, default=9):
    match = re.match(r"\s*(\d+)", str(value or ""))
    return int(match.group(1)) if match else default


def catalog_records(path):
    sheets = read_xlsx(path)
    rows = sheets.get("Сети") or next(iter(sheets.values()), [])
    if len(rows) <= 1:
        return []
    headers = [str(value).strip() for value in rows[0]]
    records = []
    for row in rows[1:]:
        row = row + [""] * (len(headers) - len(row))
        record = dict(zip(headers, row))
        file_value = record.get("Путь к файлу") or record.get("Файл") or record.get("Фото") or ""
        if not file_value:
            continue
        path_value = data_path(file_value)
        record["_path"] = path_value
        records.append(record)
    return records


def photo_score(record, payload):
    score = 0
    reasons = []
    category_words = category_terms(payload.get("category", ""))
    searchable = " ".join([
        str(record.get("Подходит для групп товаров", "")),
        str(record.get("Комментарий", "")),
        str(record.get("Тип сети", "")),
        str(record.get("Регион", "")),
        str(record.get("Сеть", "")),
    ])
    overlap = category_words & words(searchable)
    if overlap:
        score += 18 * len(overlap)
        reasons.append("категория")

    preferred = split_tags(payload.get("preferred_networks", ""))
    if preferred_network_matches(record, preferred):
        score += 70
        reasons.append("предпочтительная сеть")

    price_tags = split_tags(payload.get("price_category", ""))
    record_price_tags = split_tags(record.get("Ценовой сегмент", ""))
    if price_tags and record_price_tags and price_tags & record_price_tags:
        score += 25
        reasons.append("ценовой сегмент")

    company_type = str(payload.get("company_type", "")).lower()
    network_type = str(record.get("Тип сети", "")).lower()
    if "производитель" in company_type and ("федераль" in network_type or "регион" in network_type):
        score += 8
    if "премиум" in str(payload.get("price_category", "")).lower() and "преми" in network_type:
        score += 20

    priority = number_prefix(record.get("Приоритет"), default=9)
    score += max(0, 10 - priority) * 3

    universality = number_prefix(record.get("Уровень универсальности"), default=4)
    score += max(0, 5 - universality) * 5
    if universality >= 4:
        score -= 8
    return score, "; ".join(reasons) or "fallback"


def prioritized_photo_records(scored, preferred_networks):
    prioritized = []
    selected_paths = set()
    selected_networks = set()

    for score, priority, index, record, reason in scored:
        if not preferred_network_matches(record, preferred_networks):
            continue
        path_key = str(record["_path"].resolve()).lower()
        network_key = normalize(record.get("Сеть", ""))
        if path_key in selected_paths:
            continue
        if network_key and network_key in selected_networks:
            continue
        prioritized.append((score, priority, index, record, reason))
        selected_paths.add(path_key)
        if network_key:
            selected_networks.add(network_key)
    return prioritized


def choose_network_photos(payload, count):
    records = catalog_records("Каталог фотографий сетей.xlsx")
    if not records:
        photos = first_images("Фотки переговоров", count)
        return photos, {str(index): {"reason": "fallback_no_catalog"} for index, _ in enumerate(photos, start=1)}

    scored = []
    for index, record in enumerate(records, start=1):
        score, reason = photo_score(record, payload)
        priority = number_prefix(record.get("Приоритет"), default=9)
        scored.append((score, priority, index, record, reason))
    scored.sort(key=lambda item: (-item[0], item[1], item[2]))

    selected = []
    used_networks = set()
    used_paths = set()
    preferred = split_tags(payload.get("preferred_networks", ""))
    for score, priority, index, record, reason in prioritized_photo_records(scored, preferred):
        network_key = normalize(record.get("Сеть", ""))
        path_key = str(record["_path"].resolve()).lower()
        selected.append(record)
        if network_key:
            used_networks.add(network_key)
        used_paths.add(path_key)
        if len(selected) >= count:
            break

    for score, priority, index, record, reason in scored:
        network_key = normalize(record.get("Сеть", ""))
        path_key = str(record["_path"].resolve()).lower()
        if network_key and network_key in used_networks:
            continue
        if path_key in used_paths:
            continue
        selected.append(record)
        if network_key:
            used_networks.add(network_key)
        used_paths.add(path_key)
        if len(selected) >= count:
            break

    if len(selected) < count:
        for score, priority, index, record, reason in scored:
            if record in selected:
                continue
            network_key = normalize(record.get("Сеть", ""))
            path_key = str(record["_path"].resolve()).lower()
            if network_key and network_key in used_networks:
                continue
            if path_key in used_paths:
                continue
            selected.append(record)
            if network_key:
                used_networks.add(network_key)
            used_paths.add(path_key)
            if len(selected) >= count:
                break

    report = {}
    for offset, record in enumerate(selected[:count], start=1):
        match = next((item for item in scored if item[3] is record), None)
        report[str(offset)] = {
            "network": record.get("Сеть", ""),
            "score": match[0] if match else 0,
            "reason": match[4] if match else "",
        }
    return [resolve_image_path(record["_path"]) for record in selected[:count]], report


def choose_network_photo_candidates(payload, selected_count, candidate_count=12):
    records = catalog_records("Каталог фотографий сетей.xlsx")
    if not records:
        fallback_pool = first_images("Фотки переговоров", max(selected_count, candidate_count))
        candidates = []
        for index, photo in enumerate(fallback_pool[:candidate_count], start=1):
            candidates.append(
                {
                    "id": f"photo-{index}",
                    "path": str(photo),
                    "network": "",
                "network_type": "",
                "region": "",
                "goods_groups": "",
                "price_segment": "",
                "priority": "",
                    "universality": "",
                    "score": 0,
                    "reason": "fallback_no_catalog",
                }
            )
        selected_ids = [item["id"] for item in candidates[:selected_count]]
        selected_paths = [Path(item["path"]) for item in candidates[:selected_count]]
        report = {
            str(index): {"network": "", "score": 0, "reason": "fallback_no_catalog"}
            for index in range(1, len(selected_paths) + 1)
        }
        return selected_paths, report, candidates, selected_ids

    scored = []
    for index, record in enumerate(records, start=1):
        score, reason = photo_score(record, payload)
        priority = number_prefix(record.get("Приоритет"), default=9)
        scored.append((score, priority, index, record, reason))
    scored.sort(key=lambda item: (-item[0], item[1], item[2]))

    selected_records = []
    used_networks = set()
    used_paths = set()
    preferred = split_tags(payload.get("preferred_networks", ""))
    prioritized = prioritized_photo_records(scored, preferred)

    for score, priority, index, record, reason in prioritized:
        network_key = normalize(record.get("Сеть", ""))
        path_key = str(record["_path"].resolve()).lower()
        selected_records.append((score, record, reason))
        if network_key:
            used_networks.add(network_key)
        used_paths.add(path_key)
        if len(selected_records) >= selected_count:
            break

    for score, priority, index, record, reason in scored:
        network_key = normalize(record.get("Сеть", ""))
        path_key = str(record["_path"].resolve()).lower()
        if network_key and network_key in used_networks:
            continue
        if path_key in used_paths:
            continue
        selected_records.append((score, record, reason))
        if network_key:
            used_networks.add(network_key)
        used_paths.add(path_key)
        if len(selected_records) >= selected_count:
            break

    if len(selected_records) < selected_count:
        for score, priority, index, record, reason in scored:
            path_key = str(record["_path"].resolve()).lower()
            if path_key in used_paths:
                continue
            selected_records.append((score, record, reason))
            used_paths.add(path_key)
            if len(selected_records) >= selected_count:
                break

    candidate_records = []
    candidate_seen = set()
    for score, priority, index, record, reason in prioritized:
        path_key = str(record["_path"].resolve()).lower()
        if path_key in candidate_seen:
            continue
        candidate_records.append((score, record, reason))
        candidate_seen.add(path_key)
        if len(candidate_records) >= candidate_count:
            break

    for score, priority, index, record, reason in scored:
        path_key = str(record["_path"].resolve()).lower()
        if path_key in candidate_seen:
            continue
        candidate_records.append((score, record, reason))
        candidate_seen.add(path_key)
        if len(candidate_records) >= candidate_count:
            break

    for item in selected_records:
        path_key = str(item[1]["_path"].resolve()).lower()
        if path_key in candidate_seen:
            continue
        candidate_records.insert(len(selected_records), item)
        candidate_seen.add(path_key)

    candidates = []
    path_to_id = {}
    for index, (score, record, reason) in enumerate(candidate_records[:candidate_count], start=1):
        candidate_id = f"photo-{index}"
        path_to_id[str(record["_path"].resolve()).lower()] = candidate_id
        preferred_match = preferred_network_matches(record, preferred)
        candidates.append(
            {
                "id": candidate_id,
                "path": str(record["_path"]),
                "network": record.get("Сеть", ""),
                "network_type": record.get("Тип сети", ""),
                "region": record.get("Регион", ""),
                "goods_groups": record.get("Подходит для групп товаров", ""),
                "price_segment": record.get("Ценовой сегмент", ""),
                "priority": record.get("Приоритет", ""),
                "universality": record.get("Уровень универсальности", ""),
                "preferred_match": preferred_match,
                "score": score,
                "reason": (reason + "; обязательная предпочтительная сеть") if preferred_match else reason,
            }
        )

    selected_paths = [resolve_image_path(record["_path"]) for _, record, _ in selected_records[:selected_count]]
    selected_ids = []
    report = {}
    for offset, (score, record, reason) in enumerate(selected_records[:selected_count], start=1):
        path_key = str(record["_path"].resolve()).lower()
        selected_ids.append(path_to_id.get(path_key, ""))
        report[str(offset)] = {
            "network": record.get("Сеть", ""),
            "score": score,
            "reason": reason,
        }
    selected_ids = [item for item in selected_ids if item]
    return selected_paths, report, candidates, selected_ids


def image_map(payload):
    mapping = {}
    sources = {}
    photo_start = 2
    photo_limit = 5

    photo_tokens = [f"{{{{pic{index}}}}}" for index in range(photo_start, photo_start + photo_limit)]
    photos, photo_report, photo_candidates, selected_photo_ids = choose_network_photo_candidates(payload, photo_limit)
    for index, photo in enumerate(photos[:photo_limit], start=photo_start):
        token = f"{{{{pic{index}}}}}"
        mapping[token] = photo
        sources[token] = str(photo)
        report_key = str(index - photo_start + 1)
        if report_key in photo_report:
            sources[f"{token}_selection"] = photo_report[report_key]

    for index, review in enumerate(payload.get("review_sources", [])[:3], start=1):
        logo = review_logo_image(review)
        if logo:
            token = f"{{{{logo{index}}}}}"
            mapping[token] = logo
            sources[token] = str(logo)
    photo_networks = selected_photo_networks(photo_candidates, selected_photo_ids)
    network_logo_mapping, network_logo_sources = choose_network_logos(payload, count=8, selected_networks=photo_networks)
    mapping.update(network_logo_mapping)
    sources.update(network_logo_sources)
    return mapping, sources, photo_candidates, selected_photo_ids, photo_tokens


def selected_review_options(selection):
    if not isinstance(selection, dict):
        return []
    candidates = {
        str(item.get("id", "")): item
        for item in selection.get("review_candidates", [])
        if isinstance(item, dict)
    }
    selected = []
    for selected_id in selection.get("selected_review_ids", []):
        item = candidates.get(str(selected_id))
        if item:
            selected.append(item)
    return selected


def review_logo_mapping_from_replacements(replacements, selected_reviews=None):
    mapping = {}
    sources = {}
    review_company_tokens = {
        1: "{{block15}}",
        2: "{{block18}}",
        3: "{{block21}}",
    }
    selected_reviews = selected_reviews or []
    for index, company_token in review_company_tokens.items():
        logo_token = f"{{{{logo{index}}}}}"
        review = selected_reviews[index - 1] if index <= len(selected_reviews) else {}
        logo = review_logo_image(review) if review else None
        if logo:
            mapping[logo_token] = logo
            sources[logo_token] = str(logo)
            continue

        company = str(replacements.get(company_token, "")).strip()
        if not company:
            continue
        logo = review_logo_image({"company": company})
        if not logo:
            continue
        mapping[logo_token] = logo
        sources[logo_token] = str(logo)
    return mapping, sources


def apply_selection_override(replacements, images, image_sources, selection, payload=None):
    if not isinstance(selection, dict) or not selection:
        return replacements, images, image_sources

    review_replacement_tokens = {
        "{{block15}}",
        "{{block16}}",
        "{{block17}}",
        "{{block18}}",
        "{{block19}}",
        "{{block20}}",
        "{{block21}}",
        "{{block22}}",
        "{{block23}}",
    }
    override_replacements = selection.get("replacements")
    if isinstance(override_replacements, dict):
        for key, value in override_replacements.items():
            if (
                isinstance(key, str)
                and key in review_replacement_tokens
                and key.startswith("{{")
                and key.endswith("}}")
            ):
                replacements[key] = "" if value is None else str(value)

    override_images = selection.get("planned_images")
    network_logo_tokens = {f"{{{{logo{index}}}}}" for index in range(4, 12)}
    if isinstance(override_images, dict):
        for token, path_value in override_images.items():
            if not isinstance(token, str) or not token.startswith("{{"):
                continue
            if token in network_logo_tokens:
                continue
            path = resolve_image_path(path_value)
            if image_ext(path):
                images[token] = path
                image_sources[token] = str(path)

    # Review logos must always follow the actual companies selected for the review blocks.
    for token in ("{{logo1}}", "{{logo2}}", "{{logo3}}"):
        images.pop(token, None)
        image_sources.pop(token, None)
    logo_mapping, logo_sources = review_logo_mapping_from_replacements(
        replacements,
        selected_review_options(selection),
    )
    images.update(logo_mapping)
    image_sources.update(logo_sources)

    if isinstance(selection, dict):
        for token in network_logo_tokens:
            images.pop(token, None)
            image_sources.pop(token, None)
        photo_networks = selected_photo_networks(
            selection.get("photo_candidates") or [],
            selection.get("selected_photo_ids") or [],
        )
        network_logo_mapping, network_logo_sources = choose_network_logos(
            payload or {},
            count=8,
            selected_networks=photo_networks,
        )
        images.update(network_logo_mapping)
        image_sources.update(network_logo_sources)

    return replacements, images, image_sources


def network_logo_path(value):
    value = str(value or "").strip()
    if not value:
        return None
    raw_path = Path(value.replace("\\", "/"))
    if not raw_path.is_absolute() and len(raw_path.parts) == 1:
        path = project_path(Path("Логотипы сетей") / raw_path)
    else:
        path = project_path(raw_path)
    jpg_peer = path.with_suffix(".jpg")
    if image_ext(jpg_peer):
        return jpg_peer
    if image_ext(path):
        return path
    logo_dir = project_path("Логотипы сетей")
    if not logo_dir.exists():
        return None
    name = path.name.lower()
    stem = path.stem.lower()
    for match in logo_dir.rglob("*.jpg"):
        if match.stem.lower() == stem:
            return match
    for match in logo_dir.rglob("*"):
        if not match.is_file():
            continue
        if match.name.lower() == name or match.stem.lower() == stem:
            return match if image_ext(match) else None
    return None


def network_logo_score(record, preferred_networks):
    name = str(record.get("Компания", "") or record.get("Сеть", "")).strip()
    if not name:
        return 0
    if preferred_network_matches({"Сеть": name}, preferred_networks):
        return 1000
    name_keys = network_match_keys(name)
    preferred_keys = {
        key
        for item in preferred_networks
        for key in network_match_keys(item)
    }
    if name_keys & preferred_keys:
        return 900
    score = 0
    normalized_name = normalize(name)
    name_words = words(name)
    for item in preferred_networks:
        normalized_item = normalize(item)
        if not normalized_item:
            continue
        if normalized_item == normalized_name:
            score = max(score, 800)
        elif normalized_item in normalized_name or normalized_name in normalized_item:
            score = max(score, 600)
        else:
            score = max(score, len(name_words & words(item)) * 100)
    return score


def selected_photo_networks(photo_candidates, selected_photo_ids):
    candidates = {
        str(item.get("id", "")): item
        for item in photo_candidates or []
        if isinstance(item, dict)
    }
    networks = []
    seen = set()
    for selected_id in selected_photo_ids or []:
        item = candidates.get(str(selected_id))
        if not item:
            continue
        network = str(item.get("network", "")).strip()
        key = normalize(network)
        if network and key not in seen:
            networks.append(network)
            seen.add(key)
    return networks


def network_order_rank(name, ordered_networks):
    for index, network in enumerate(ordered_networks):
        if preferred_network_matches({"Сеть": name}, [network]):
            return index
    return len(ordered_networks) + 1000


def choose_network_logos(payload, count=8, selected_networks=None):
    sheets = read_xlsx("Список сетей.xlsx")
    rows = sheets.get("Shops") or next(iter(sheets.values()), [])
    if len(rows) <= 1:
        return {}, {}

    headers = [str(value).strip() for value in rows[0]]
    preferred = split_ordered(payload.get("preferred_networks", ""))
    selected_networks = [item for item in (selected_networks or []) if str(item).strip()]
    preferred_selected = [
        network
        for network in selected_networks
        if preferred_network_matches({"Сеть": network}, preferred)
    ]
    ordered_networks = []
    for network in [*preferred_selected, *selected_networks, *preferred]:
        key = normalize(network)
        if key and key not in {normalize(item) for item in ordered_networks}:
            ordered_networks.append(network)
    records = []
    for index, row in enumerate(rows[1:], start=2):
        row = row + [""] * (len(headers) - len(row))
        record = dict(zip(headers, row))
        logo = network_logo_path(record.get("Логотип", ""))
        if not logo:
            continue
        name = str(record.get("Компания", "") or record.get("Сеть", "")).strip()
        score = network_logo_score(record, preferred)
        if preferred_network_matches({"Сеть": name}, preferred_selected):
            score += 3000
        elif preferred_network_matches({"Сеть": name}, selected_networks):
            score += 2000
        order_rank = network_order_rank(name, ordered_networks)
        jpg_priority = 0 if logo.suffix.lower() in {".jpg", ".jpeg"} else 1
        records.append((score, order_rank, jpg_priority, index, record, logo))

    if ordered_networks:
        records.sort(key=lambda item: (-item[0], item[1], item[2], item[3]))
    else:
        records.sort(key=lambda item: (item[2], item[3]))

    selected = []
    used_paths = set()
    for score, order_rank, jpg_priority, index, record, logo in records:
        if preferred and score <= 0 and len(selected) < min(3, count):
            continue
        path_key = str(logo.resolve()).lower()
        if path_key in used_paths:
            continue
        selected.append((score, index, record, logo))
        used_paths.add(path_key)
        if len(selected) >= count:
            break

    if len(selected) < count:
        for score, order_rank, jpg_priority, index, record, logo in records:
            path_key = str(logo.resolve()).lower()
            if path_key in used_paths:
                continue
            selected.append((score, index, record, logo))
            used_paths.add(path_key)
            if len(selected) >= count:
                break

    mapping = {}
    sources = {}
    for offset, (score, index, record, logo) in enumerate(selected[:count], start=4):
        token = f"{{{{logo{offset}}}}}"
        mapping[token] = logo
        name = str(record.get("Компания", "") or record.get("Сеть", ""))
        sources[token] = {
            "path": str(logo),
            "network": name,
            "row": index,
            "score": score,
            "preferred_photo_match": preferred_network_matches({"Сеть": name}, preferred_selected),
            "selected_photo_match": preferred_network_matches({"Сеть": name}, selected_networks),
        }
    return mapping, sources


def next_media_name(existing, source):
    ext = image_ext(source) or source.suffix.lower()
    index = 1
    while True:
        candidate = f"ppt/media/generated_{index}{ext}"
        if candidate not in existing:
            existing.add(candidate)
            return candidate
        index += 1


def max_shape_id(root):
    values = []
    for node in root.findall(".//p:cNvPr", NS):
        try:
            values.append(int(node.attrib.get("id", "0")))
        except ValueError:
            pass
    return max(values, default=1000)


def next_rel_id(rels_root):
    values = []
    for rel in rels_root.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            values.append(int(rid[3:]))
    return f"rId{max(values, default=0) + 1}"


def ensure_content_type(content_types_root, ext):
    ext = ext.lower().lstrip(".")
    for node in content_types_root.findall("ct:Default", NS):
        if node.attrib.get("Extension", "").lower() == ext:
            return
    mime = mimetypes.types_map.get(f".{ext}", "image/jpeg")
    ET.SubElement(content_types_root, qn(CT, "Default"), {"Extension": ext, "ContentType": mime})


def make_picture(sp, rel_id, shape_id):
    c_nv_pr = sp.find(".//p:cNvPr", NS)
    name = c_nv_pr.attrib.get("name", "Generated image") if c_nv_pr is not None else "Generated image"
    xfrm = sp.find(".//a:xfrm", NS)

    pic = ET.Element(qn(P, "pic"))
    nv_pic_pr = ET.SubElement(pic, qn(P, "nvPicPr"))
    ET.SubElement(nv_pic_pr, qn(P, "cNvPr"), {"id": str(shape_id), "name": name})
    c_nv_pic_pr = ET.SubElement(nv_pic_pr, qn(P, "cNvPicPr"))
    ET.SubElement(c_nv_pic_pr, qn(A, "picLocks"), {"noChangeAspect": "1"})
    ET.SubElement(nv_pic_pr, qn(P, "nvPr"))

    blip_fill = ET.SubElement(pic, qn(P, "blipFill"))
    ET.SubElement(blip_fill, qn(A, "blip"), {qn(R, "embed"): rel_id})
    stretch = ET.SubElement(blip_fill, qn(A, "stretch"))
    ET.SubElement(stretch, qn(A, "fillRect"))

    sp_pr = ET.SubElement(pic, qn(P, "spPr"))
    if xfrm is not None:
        sp_pr.append(ET.fromstring(ET.tostring(xfrm, encoding="utf-8")))
    prst = ET.SubElement(sp_pr, qn(A, "prstGeom"), {"prst": "rect"})
    ET.SubElement(prst, qn(A, "avLst"))
    return pic


def shape_text(shape):
    return "".join(t.text or "" for t in shape.findall(".//a:t", NS)).strip()


def remove_empty_text_runs(shape):
    for paragraph in shape.findall(".//a:p", NS):
        for run in list(paragraph):
            if run.tag != qn(A, "r"):
                continue
            if "".join(node.text or "" for node in run.findall(".//a:t", NS)):
                continue
            paragraph.remove(run)


def replace_text(root, replacements):
    counts = {key: 0 for key in replacements}
    block14_value = replacements.get("{{block14}}")
    if block14_value is not None:
        for shape in root.findall(".//p:sp", NS):
            text_nodes = shape.findall(".//a:t", NS)
            if not text_nodes:
                continue
            combined = "".join(node.text or "" for node in text_nodes)
            if "{{block14" not in combined:
                continue
            counts["{{block14}}"] = counts.get("{{block14}}", 0) + combined.count("{{block14}}")
            text_nodes[0].text = str(block14_value)
            for node in text_nodes[1:]:
                node.text = ""

    for shape in root.findall(".//p:sp", NS):
        text_nodes = shape.findall(".//a:t", NS)
        if not text_nodes:
            continue
        combined = "".join(node.text or "" for node in text_nodes)
        stripped = combined.strip()
        for key, value in replacements.items():
            if stripped != key:
                continue
            counts[key] += combined.count(key)
            text_nodes[0].text = str(value)
            for node in text_nodes[1:]:
                node.text = ""
            remove_empty_text_runs(shape)
            break

    for text_node in root.findall(".//a:t", NS):
        if not text_node.text:
            continue
        new_text = text_node.text
        for key, value in replacements.items():
            if key in new_text:
                counts[key] += new_text.count(key)
                new_text = new_text.replace(key, value)
        if new_text != text_node.text:
            text_node.text = new_text

    # PowerPoint can split one visible placeholder across several text runs,
    # for example "{{" + "block9}}". Handle those at shape scope.
    for shape in root.findall(".//p:sp", NS):
        text_nodes = shape.findall(".//a:t", NS)
        if len(text_nodes) <= 1:
            continue
        combined = "".join(node.text or "" for node in text_nodes)
        if not combined:
            continue
        new_text = combined
        replaced_any = False
        for key, value in replacements.items():
            if key in new_text:
                counts[key] += new_text.count(key)
                new_text = new_text.replace(key, value)
                replaced_any = True
        if new_text != combined:
            if replaced_any and "{{" not in new_text:
                new_text = new_text.replace("\t", "").strip()
            text_nodes[0].text = new_text
            for node in text_nodes[1:]:
                node.text = ""
            remove_empty_text_runs(shape)
    return counts


def replace_image_placeholders(root, rels_root, image_mapping, media_entries):
    sp_tree = root.find(".//p:spTree", NS)
    if sp_tree is None:
        return [], 0

    additions = []
    replaced = {key: 0 for key in image_mapping}
    shape_id = max_shape_id(root) + 1
    children = list(sp_tree)

    for index, child in enumerate(children):
        if child.tag != qn(P, "sp"):
            continue
        token = shape_text(child)
        if token not in image_mapping:
            continue

        source = image_mapping[token]
        media_name = next_media_name(media_entries, source)
        rel_id = next_rel_id(rels_root)
        target = "../media/" + Path(media_name).name
        ET.SubElement(
            rels_root,
            qn(REL, "Relationship"),
            {
                "Id": rel_id,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "Target": target,
            },
        )
        pic = make_picture(child, rel_id, shape_id)
        shape_id += 1
        sp_tree.remove(child)
        sp_tree.insert(index, pic)
        additions.append((media_name, source))
        replaced[token] += 1

    return additions, replaced


def merge_counts(target, source):
    for key, value in source.items():
        target[key] = target.get(key, 0) + value


def missing_text_keys(counts):
    missing = []
    for key, value in counts.items():
        if value != 0:
            continue
        if key == "{{block1}}" and counts.get("{{block1 }}", 0) > 0:
            continue
        missing.append(key)
    return missing


def find_remaining_placeholders(root):
    found = set()
    for text_node in root.findall(".//a:t", NS):
        if text_node.text:
            found.update(re.findall(r"\{\{[^}]+\}\}", text_node.text))
    for shape in root.findall(".//p:sp", NS):
        combined = "".join(node.text or "" for node in shape.findall(".//a:t", NS))
        if combined:
            found.update(re.findall(r"\{\{[^}]+\}\}", combined))
    return found


def main():
    parser = ArgumentParser(description="Build PPTX presentation from template.")
    parser.add_argument("company", nargs="?", default="ГОСУДАРЕВ СТАНДАРТ")
    parser.add_argument("--dry-run", action="store_true", help="Only print payload and planned assets.")
    parser.add_argument("--template", default="Шаблон презентации новый.pptx", help="PPTX template path.")
    args = parser.parse_args()

    template = project_path(args.template)
    payload = read_payload(args.company)
    slug = re.sub(r"\W+", "_", payload["company"], flags=re.UNICODE).strip("_").lower()
    output = BASE_DIR / f"{slug}_ЦЗС_новый.pptx"
    if output.exists():
        index = 2
        while True:
            candidate = BASE_DIR / f"{slug}_ЦЗС_новый_v{index}.pptx"
            if not candidate.exists():
                output = candidate
                break
            index += 1
    replacements = payload["replacements"]
    images, image_sources, photo_candidates, selected_photo_ids, photo_tokens = image_map(payload)
    selection_override = read_selection_override()
    replacements, images, image_sources = apply_selection_override(
        replacements,
        images,
        image_sources,
        selection_override,
        payload,
    )

    if args.dry_run:
        print(json.dumps({
            "mode": "dry-run",
            "client": payload["company"],
            "contact": payload["contact"],
            "position": payload["position"],
            "category": payload["category"],
            "company_type": payload.get("company_type", ""),
            "price_category": payload.get("price_category", ""),
            "preferred_networks": payload.get("preferred_networks", ""),
            "client_record": payload.get("client_record", {}),
            "stats_source": payload["stats_source"],
            "review_source": payload.get("review_source", {}),
            "review_candidates": payload.get("review_candidates", []),
            "selected_review_ids": payload.get("selected_review_ids", []),
            "required_review_count": payload.get("required_review_count", 3),
            "photo_candidates": photo_candidates,
            "selected_photo_ids": selected_photo_ids,
            "required_photo_count": len(photo_tokens),
            "photo_tokens": photo_tokens,
            "replacements": replacements,
            "planned_images": image_sources,
            "missing_image_sources": [
                token for token in IMAGE_TOKENS
                if token not in images
            ],
            "output": str(output),
        }, ensure_ascii=False, indent=2))
        return

    temp = output.with_suffix(".tmp.pptx")
    if temp.exists():
        temp.unlink()

    text_replaced = {key: 0 for key in replacements}
    image_replaced = {key: 0 for key in images}
    added_media = []
    remaining_placeholders = set()

    normalized_images, normalization_report, temp_files = normalize_images_for_powerpoint(images)
    image_replaced = {key: 0 for key in normalized_images}

    try:
        with zipfile.ZipFile(template, "r") as zin:
            entries = {item.filename for item in zin.infolist()}
            media_entries = {name for name in entries if name.startswith("ppt/media/")}
            content_types_root = ET.fromstring(zin.read("[Content_Types].xml"))

            modified = {}
            for name in entries:
                if re.match(r"ppt/slides/slide\d+\.xml$", name):
                    root = ET.fromstring(zin.read(name))
                    merge_counts(text_replaced, replace_text(root, replacements))

                    rels_name = f"ppt/slides/_rels/{Path(name).name}.rels"
                    if rels_name in entries:
                        rels_root = ET.fromstring(zin.read(rels_name))
                    else:
                        rels_root = ET.Element(qn(REL, "Relationships"))

                    additions, count = replace_image_placeholders(root, rels_root, normalized_images, media_entries)
                    if any(count.values()):
                        merge_counts(image_replaced, count)
                        added_media.extend(additions)
                        modified[rels_name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

                    remaining_placeholders.update(find_remaining_placeholders(root))
                    modified[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

            for media_name, source in added_media:
                ensure_content_type(content_types_root, image_ext(source) or source.suffix)
            modified["[Content_Types].xml"] = ET.tostring(content_types_root, encoding="utf-8", xml_declaration=True)

            with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in modified:
                        zout.writestr(item, modified[item.filename])
                    else:
                        zout.writestr(item, zin.read(item.filename))
                for media_name, source in added_media:
                    zout.write(source, media_name)
    finally:
        for temp_file in temp_files:
            try:
                temp_file.unlink(missing_ok=True)
            except OSError:
                pass

    shutil.move(temp, output)
    print(json.dumps({
        "created": str(output),
        "text_replaced_total": sum(text_replaced.values()),
        "text_replaced": text_replaced,
        "image_replaced_total": sum(image_replaced.values()),
        "image_replaced": image_replaced,
        "planned_images": image_sources,
        "image_normalization": normalization_report,
        "missing_text_placeholders": missing_text_keys(text_replaced),
        "missing_image_placeholders": [key for key in images if image_replaced.get(key, 0) == 0],
        "missing_image_sources": [
            token for token in IMAGE_TOKENS
            if token not in images
        ],
        "remaining_placeholders": sorted(remaining_placeholders),
        "client": payload["company"],
        "contact": payload["contact"],
        "position": payload["position"],
        "category": payload["category"],
        "stats_source": payload["stats_source"],
        "review_source": payload.get("review_source", {}),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()


