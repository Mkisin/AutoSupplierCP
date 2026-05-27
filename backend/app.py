from __future__ import annotations

import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse
import urllib.request
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles


ROOT = Path(__file__).resolve().parents[1]
DATA_FILE = ROOT / "Данные для разработки.xlsx"
PHOTO_DIR = ROOT / "Фото товаров"
STATIC_DIR = ROOT / "frontend"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)

app = FastAPI(title="AutoSupplierCP")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

PRESENTATION_JOBS: dict[str, dict[str, Any]] = {}
PRESENTATION_JOBS_LOCK = threading.Lock()


def _safe_root_file(path_value: str) -> Path:
    path = Path(path_value)
    if not path.is_absolute():
        path = ROOT / path
    resolved = path.resolve()
    if ROOT.resolve() not in resolved.parents and resolved != ROOT.resolve():
        raise HTTPException(status_code=400, detail="Некорректный путь к файлу")
    return resolved


def _set_presentation_job(job_id: str, **values: Any) -> None:
    with PRESENTATION_JOBS_LOCK:
        job = PRESENTATION_JOBS.setdefault(job_id, {})
        job.update(values)
        job["updatedAt"] = datetime.now().isoformat(timespec="seconds")


def _get_presentation_job(job_id: str) -> dict[str, Any]:
    with PRESENTATION_JOBS_LOCK:
        job = PRESENTATION_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Задача сборки не найдена")
        return dict(job)


def _run_presentation_job(job_id: str, company: str) -> None:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    command = [sys.executable, "-X", "utf8", "build_presentation_direct.py", company]
    stdout_path: Path | None = None
    stderr_path: Path | None = None

    try:
        _set_presentation_job(job_id, status="running", progress=12, message="Готовлю данные компании")
        stdout_file = tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            delete=False,
            dir=ROOT,
            prefix="presentation_job_",
            suffix=".out",
        )
        stderr_file = tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            delete=False,
            dir=ROOT,
            prefix="presentation_job_",
            suffix=".err",
        )
        stdout_path = Path(stdout_file.name)
        stderr_path = Path(stderr_file.name)
        process = subprocess.Popen(
            command,
            cwd=ROOT,
            env=env,
            stdout=stdout_file,
            stderr=stderr_file,
            text=True,
            encoding="utf-8",
        )
        stdout_file.close()
        stderr_file.close()

        started = time.monotonic()
        while process.poll() is None:
            elapsed = time.monotonic() - started
            progress = min(86, 20 + int(elapsed * 8))
            _set_presentation_job(
                job_id,
                status="running",
                progress=progress,
                message="Собираю PPTX: тексты, изображения и связи слайдов",
            )
            time.sleep(0.75)
            if elapsed > 180:
                process.kill()
                raise TimeoutError("Сборка презентации заняла слишком много времени")

        process.wait()
        stdout = stdout_path.read_text(encoding="utf-8") if stdout_path else ""
        stderr = stderr_path.read_text(encoding="utf-8") if stderr_path else ""
        if process.returncode != 0:
            details = (stderr or stdout or "Не удалось сформировать презентацию").strip()
            raise RuntimeError(details)

        _set_presentation_job(job_id, status="running", progress=92, message="Проверяю созданный файл")
        report = json.loads(stdout)
        created = _safe_root_file(str(report.get("created", "")))
        if not created.exists():
            raise RuntimeError("Презентация не была создана")

        _set_presentation_job(
            job_id,
            status="done",
            progress=100,
            message="Презентация готова",
            fileName=created.name,
            downloadUrl=f"/api/presentations/{urllib.parse.quote(created.name)}",
            report=report,
        )
    except Exception as exc:
        _set_presentation_job(
            job_id,
            status="error",
            progress=100,
            message=str(exc) or "Не удалось сформировать презентацию",
        )
    finally:
        if stdout_path:
            stdout_path.unlink(missing_ok=True)
        if stderr_path:
            stderr_path.unlink(missing_ok=True)


HEADERS = [
    "N",
    "Название компании",
    "РРќРќ",
    "Тип компании",
    "Р¤РРћ контакта",
    "Должность контакта",
    "Категория товара",
    "Краткое описание продукции",
    "Отрасль",
    "Сфера деятельности",
    "Сайт ",
    "Страна ",
    "Город",
    "Название товара",
    "Категория по работе с сетями",
    "Предпочтительные сети",
    "Ценовая категория",
    "Фото товара",
    "Дата сохранения",
]

FIELD_TO_HEADER = {
    "companyName": "Название компании",
    "inn": "РРќРќ",
    "companyType": "Тип компании",
    "contactName": "Р¤РРћ контакта",
    "contactPosition": "Должность контакта",
    "productName": "Название товара",
    "productCategory": "Категория товара",
    "productDescription": "Краткое описание продукции",
    "industry": "Отрасль",
    "activity": "Сфера деятельности",
    "networkCategories": "Категория по работе с сетями",
    "website": "Сайт ",
    "country": "Страна ",
    "city": "Город",
    "preferredNetworks": "Предпочтительные сети",
    "priceCategory": "Ценовая категория",
}

DEFAULT_OPTIONS = {
    "companyTypes": [
        "Производитель конечного товара",
        "Дистрибьютор",
        "РРјРїРѕСЂС‚РµСЂ",
        "Торговая марка",
        "Фермерское хозяйство",
    ],
    "industries": ["Food", "Non-food", "DIY", "Beauty", "Pet", "Pharma"],
    "activities": [
        "Продукты питания\\напитки (фуд)",
        "Косметика и бытовая химия",
        "Товары для животных",
        "Строительные и хозяйственные товары",
        "Одежда и аксессуары",
    ],
    "networkWorkCategories": [
        "новичок",
        "растущий",
        "опытный",
    ],
    "preferredNetworkCategories": [
        "Федеральные сети",
        "Региональные сети",
        "HoReCa",
        "Маркетплейсы",
        "Специализированные сети",
        "Дискаунтеры",
    ],
    "productCategories": [
        "Мясо и мясопродукты",
        "Молочная продукция Сыры",
        "Крупы",
        "Косметика",
        "Товары для животных",
        "Стройтовары",
    ],
    "priceCategories": ["Эконом", "Средний", "Средний+", "Премиум", "Люкс"],
}


def _xlsx_parts() -> tuple[list[str], dict[str, bytes]]:
    if not DATA_FILE.exists():
        raise HTTPException(status_code=404, detail="Файл данных не найден")

    with zipfile.ZipFile(DATA_FILE, "r") as source:
        names = source.namelist()
        return names, {name: source.read(name) for name in names}


def _shared_strings(parts: dict[str, bytes]) -> list[str]:
    if "xl/sharedStrings.xml" not in parts:
        return []

    root = ET.fromstring(parts["xl/sharedStrings.xml"])
    ns = f"{{{MAIN_NS}}}"
    values = []
    for si in root.findall(f"{ns}si"):
        values.append("".join(t.text or "" for t in si.iter(f"{ns}t")))
    return values


def _sheet_target(parts: dict[str, bytes], sheet_name: str) -> str:
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

    ns = {"m": MAIN_NS, "r": REL_NS}
    for sheet in workbook.find("m:sheets", ns) or []:
        if sheet.attrib.get("name", "").lower() == sheet_name.lower():
            rid = sheet.attrib[f"{{{REL_NS}}}id"]
            target = rel_map[rid]
            return "xl/" + target.lstrip("/")

    raise HTTPException(status_code=404, detail=f"Лист {sheet_name} не найден")


def _cell_text(cell: ET.Element, shared: list[str]) -> str:
    ns = f"{{{MAIN_NS}}}"
    if cell.attrib.get("t") == "inlineStr":
        inline = cell.find(f"{ns}is")
        return "" if inline is None else "".join(t.text or "" for t in inline.iter(f"{ns}t"))

    value = cell.find(f"{ns}v")
    if value is None or value.text is None:
        return ""
    if cell.attrib.get("t") == "s":
        return shared[int(value.text)]
    return value.text


def _col_to_number(col: str) -> int:
    number = 0
    for char in col:
        number = number * 26 + ord(char.upper()) - 64
    return number


def _cell_ref(col: str, row: int) -> str:
    return f"{col}{row}"


def _inline_cell(col: str, row_number: int, value: Any) -> ET.Element:
    cell = ET.Element(f"{{{MAIN_NS}}}c", {"r": _cell_ref(col, row_number), "t": "inlineStr"})
    inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
    text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
    text.text = "" if value is None else str(value)
    return cell


def _row_values(row: ET.Element, shared: list[str]) -> dict[str, str]:
    ns = f"{{{MAIN_NS}}}"
    result = {}
    for cell in row.findall(f"{ns}c"):
        ref = cell.attrib.get("r", "")
        col = re.sub(r"\d+", "", ref)
        result[col] = _cell_text(cell, shared)
    return result


def _ensure_headers(root: ET.Element, sheet_data: ET.Element, shared: list[str]) -> None:
    ns = f"{{{MAIN_NS}}}"
    first_row = sheet_data.find(f"{ns}row[@r='1']")
    if first_row is None:
        first_row = ET.Element(f"{ns}row", {"r": "1"})
        sheet_data.insert(0, first_row)

    existing = _row_values(first_row, shared)
    for index, title in enumerate(HEADERS, start=1):
        col = _number_to_col(index)
        if existing.get(col) != title:
            for cell in list(first_row.findall(f"{ns}c")):
                if cell.attrib.get("r") == _cell_ref(col, 1):
                    first_row.remove(cell)
            first_row.append(_inline_cell(col, 1, title))

    _sort_cells(first_row)
    dimension = root.find(f"{ns}dimension")
    if dimension is not None:
        dimension.set("ref", f"A1:S{max(_max_used_row(sheet_data, shared), 1)}")


def _number_to_col(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _sort_cells(row: ET.Element) -> None:
    children = list(row)
    for child in children:
        row.remove(child)
    children.sort(key=lambda c: _col_to_number(re.sub(r"\d+", "", c.attrib.get("r", "A"))))
    for child in children:
        row.append(child)


def _sort_rows(sheet_data: ET.Element) -> None:
    rows = list(sheet_data)
    for row in rows:
        sheet_data.remove(row)
    rows.sort(key=lambda item: int(item.attrib.get("r", "0")))
    for row in rows:
        sheet_data.append(row)


def _max_used_row(sheet_data: ET.Element, shared: list[str]) -> int:
    ns = f"{{{MAIN_NS}}}"
    used = 1
    for row in sheet_data.findall(f"{ns}row"):
        row_number = int(row.attrib.get("r", "0"))
        values = _row_values(row, shared)
        if any(str(value).strip() for value in values.values()):
            used = max(used, row_number)
    return used


def _append_card(data: dict[str, str], photo_path: str) -> int:
    return _save_card_row(data, photo_path)


def _save_card_row(data: dict[str, str], photo_path: str, target_row_number: int | None = None) -> int:
    names, parts = _xlsx_parts()
    shared = _shared_strings(parts)
    sheet_path = _sheet_target(parts, "Карточка клиента")
    root = ET.fromstring(parts[sheet_path])
    ns = f"{{{MAIN_NS}}}"
    sheet_data = root.find(f"{ns}sheetData")
    if sheet_data is None:
        sheet_data = ET.SubElement(root, f"{ns}sheetData")

    _ensure_headers(root, sheet_data, shared)
    row_number = target_row_number or _max_used_row(sheet_data, shared) + 1
    existing_values: dict[str, str] = {}
    if target_row_number:
        existing_row = sheet_data.find(f"{ns}row[@r='{target_row_number}']")
        header_row = sheet_data.find(f"{ns}row[@r='1']")
        if existing_row is not None and header_row is not None:
            headers_by_col = _row_values(header_row, shared)
            values_by_col = _row_values(existing_row, shared)
            existing_values = {
                header.strip(): values_by_col.get(col, "")
                for col, header in headers_by_col.items()
                if header
            }

    row_values = [""] * len(HEADERS)
    row_values[0] = str(row_number - 1)
    header_indexes = {header: index for index, header in enumerate(HEADERS)}
    for field, header in FIELD_TO_HEADER.items():
        row_values[header_indexes[header]] = data.get(field, "").strip()
    row_values[header_indexes["Фото товара"]] = photo_path or existing_values.get("Фото товара", "")
    row_values[header_indexes["Дата сохранения"]] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row = ET.Element(f"{ns}row", {"r": str(row_number)})
    for index, value in enumerate(row_values, start=1):
        new_row.append(_inline_cell(_number_to_col(index), row_number, value))

    for existing in list(sheet_data.findall(f"{ns}row")):
        if existing.attrib.get("r") == str(row_number):
            sheet_data.remove(existing)
    sheet_data.append(new_row)
    _sort_rows(sheet_data)
    dimension = root.find(f"{ns}dimension")
    if dimension is not None:
        dimension.set("ref", f"A1:S{row_number}")

    parts[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    fd, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=ROOT)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
            for name in names:
                target.writestr(name, parts[name])
        temp_path.replace(DATA_FILE)
    except PermissionError as exc:
        temp_path.unlink(missing_ok=True)
        raise HTTPException(
            status_code=423,
            detail="Не удалось сохранить Excel. Закройте файл 'Данные для разработки.xlsx' и попробуйте снова.",
        ) from exc
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise

    return row_number


def _read_cards() -> list[dict[str, str]]:
    _, parts = _xlsx_parts()
    shared = _shared_strings(parts)
    sheet_path = _sheet_target(parts, "Карточка клиента")
    root = ET.fromstring(parts[sheet_path])
    ns = f"{{{MAIN_NS}}}"
    sheet_data = root.find(f"{ns}sheetData")
    if sheet_data is None:
        return []

    rows = sheet_data.findall(f"{ns}row")
    if not rows:
        return []

    headers = _row_values(rows[0], shared)
    cards = []
    for row in rows[1:]:
        raw = _row_values(row, shared)
        item = {}
        for col, header in headers.items():
            if header:
                item[header.strip()] = raw.get(col, "").strip()
        if any(item.values()):
            item["_rowNumber"] = row.attrib.get("r", "")
            cards.append(item)
    return cards


def _unique(values: list[str], fallback: list[str]) -> list[str]:
    result = []
    for value in fallback + values:
        normalized = value.strip()
        if normalized and normalized not in result:
            result.append(normalized)
    return result


def _save_photo(photo: UploadFile | None, company_name: str) -> str:
    if photo is None or not photo.filename:
        return ""

    PHOTO_DIR.mkdir(exist_ok=True)
    suffix = Path(photo.filename).suffix.lower()
    if suffix not in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
        raise HTTPException(status_code=400, detail="Фото должно быть в формате JPG, PNG, WEBP или GIF")

    safe_company = re.sub(r"[^A-Za-zА-Яа-я0-9_-]+", "_", company_name).strip("_") or "product"
    filename = f"{safe_company}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{suffix}"
    target = PHOTO_DIR / filename
    with target.open("wb") as file:
        shutil.copyfileobj(photo.file, file)
    return str(target.relative_to(ROOT))


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/options")
def options() -> dict[str, Any]:
    cards = _read_cards()
    return {
        "companyTypes": _unique([c.get("Тип компании", "") for c in cards], DEFAULT_OPTIONS["companyTypes"]),
        "industries": _unique([c.get("Отрасль", "") for c in cards], DEFAULT_OPTIONS["industries"]),
        "activities": _unique([c.get("Сфера деятельности", "") for c in cards], DEFAULT_OPTIONS["activities"]),
        "productCategories": _unique([c.get("Категория товара", "") for c in cards], DEFAULT_OPTIONS["productCategories"]),
        "networkWorkCategories": DEFAULT_OPTIONS["networkWorkCategories"],
        "preferredNetworkCategories": DEFAULT_OPTIONS["preferredNetworkCategories"],
        "priceCategories": DEFAULT_OPTIONS["priceCategories"],
        "cardCount": len(cards),
        "recentCards": cards[-6:],
    }


@app.get("/api/cards")
def cards() -> dict[str, Any]:
    return {"items": _read_cards()}


@app.get("/api/company-search")
def company_search(query: str) -> dict[str, Any]:
    query = query.strip()
    if len(query) < 3:
        return {"items": []}

    encoded = urllib.parse.urlencode({"query": query})
    request = urllib.request.Request(
        "https://egrul.nalog.ru/",
        data=encoded.encode("utf-8"),
        headers={
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "User-Agent": "Mozilla/5.0 AutoSupplierCP",
            "Accept": "application/json, text/javascript, */*; q=0.01",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=12) as response:
            token = json.loads(response.read().decode("utf-8")).get("t")
        if not token:
            return {"items": []}
        url = f"https://egrul.nalog.ru/search-result/{urllib.parse.quote(token)}"
        with urllib.request.urlopen(
            urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 AutoSupplierCP"}),
            timeout=12,
        ) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        raise HTTPException(
            status_code=502,
            detail="Не удалось получить данные из открытой базы ФНС. Можно ввести РРќРќ вручную.",
        ) from exc

    items = []
    for item in payload.get("rows", [])[:10]:
        items.append(
            {
                "name": item.get("n", ""),
                "inn": item.get("i", ""),
                "address": item.get("a", ""),
                "status": item.get("s", ""),
            }
        )
    return {"items": items}


@app.post("/api/cards")
async def save_card(
    payload: str = Form(...),
    photo: UploadFile | None = File(default=None),
) -> dict[str, Any]:
    try:
        data = json.loads(payload)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="Некорректные данные формы") from exc

    required = [
        "companyName",
        "companyType",
        "contactName",
        "contactPosition",
        "productName",
        "productCategory",
        "industry",
        "activity",
        "networkCategories",
        "website",
        "country",
        "city",
        "priceCategory",
    ]
    missing = [field for field in required if not str(data.get(field, "")).strip()]
    if missing:
        raise HTTPException(status_code=400, detail="Заполните обязательные поля")

    photo_path = _save_photo(photo, data.get("companyName", ""))
    row_number_raw = str(data.get("_rowNumber", "")).strip()
    target_row_number = int(row_number_raw) if row_number_raw.isdigit() and int(row_number_raw) > 1 else None
    row_number = _save_card_row(data, photo_path, target_row_number)
    return {"ok": True, "row": row_number, "mode": "updated" if target_row_number else "created", "photoPath": photo_path}


@app.post("/api/presentations")
def build_presentation(payload: dict[str, Any]) -> dict[str, Any]:
    company = str(payload.get("companyName", "")).strip()
    if not company:
        raise HTTPException(status_code=400, detail="Выберите или заполните компанию")

    job_id = uuid.uuid4().hex
    _set_presentation_job(
        job_id,
        id=job_id,
        companyName=company,
        status="queued",
        progress=5,
        message="Задача поставлена в очередь",
    )
    worker = threading.Thread(target=_run_presentation_job, args=(job_id, company), daemon=True)
    worker.start()
    return _get_presentation_job(job_id)


@app.get("/api/presentations/jobs/{job_id}")
def presentation_job(job_id: str) -> dict[str, Any]:
    return _get_presentation_job(job_id)


@app.get("/api/presentations/{filename}")
def download_presentation(filename: str) -> FileResponse:
    target = _safe_root_file(filename)
    if not target.exists() or target.suffix.lower() != ".pptx":
        raise HTTPException(status_code=404, detail="Презентация не найдена")
    return FileResponse(
        target,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=target.name,
    )


@app.get("/photo/{filename}")
def photo(filename: str) -> FileResponse:
    target = PHOTO_DIR / filename
    if not target.exists():
        raise HTTPException(status_code=404, detail="Фото не найдено")
    media_type = mimetypes.guess_type(target.name)[0]
    return FileResponse(target, media_type=media_type)
