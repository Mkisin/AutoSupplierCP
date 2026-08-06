from __future__ import annotations

import json
import mimetypes
import os
import re
import secrets
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles


ROOT = Path(__file__).resolve().parents[1]


def _load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


_load_env_file(ROOT / ".env")

DATA_FILE = ROOT / "Данные для разработки.xlsx"
PHOTO_CATALOG_FILE = ROOT / "Каталог фотографий сетей.xlsx"
PHOTO_DIR = ROOT / "Фото товаров"
STATIC_DIR = ROOT / "frontend"
AI_CONFIG_FILE = Path(os.environ.get("AI_CONFIG_FILE", ROOT / "ai_config.json"))
AI_PROMPT_FILE = Path(os.environ.get("AI_PROMPT_FILE", ROOT / "mainprompt.md"))
DADATA_API_KEY = os.environ.get("DADATA_API_KEY", "e827de6169ff2dd15ed29a07d4489511b80d1463")
DADATA_FIND_PARTY_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"
DADATA_FIND_OKVED_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/okved2"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)

app = FastAPI(title="AutoSupplierCP")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

PRESENTATION_JOBS: dict[str, dict[str, Any]] = {}
PRESENTATION_JOBS_LOCK = threading.Lock()
PRESENTATION_JOBS_DIR = ROOT / "data" / "presentation_jobs"
PRESENTATION_AUDIT_FILE = ROOT / "data" / "presentation_audit.jsonl"
PRESENTATION_AUDIT_LOCK = threading.Lock()
AI_SELECTIONS: dict[str, dict[str, Any]] = {}
AI_SELECTIONS_LOCK = threading.Lock()
DIRECT_HTTP_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

DATA_SOURCES_FILE = ROOT / "data_sources.json"
BITRIX_CONFIG_FILE = ROOT / "bitrix_config.json"
BITRIX_JOBS: dict[str, dict[str, Any]] = {}
BITRIX_JOBS_LOCK = threading.Lock()
APP_STORAGE_MODE = os.environ.get("APP_STORAGE_MODE", "").strip().lower()
BITRIX_DEAL_FIELD_MAP = {
    "companyName": "COMPANY_ID",
    "contactName": "CONTACT_ID",
    "contactPosition": "CONTACT_ID",
    "networkCategories": "UF_CRM_1662544467901",
    "website": "UF_CRM_1653999914176",
    "country": "UF_CRM_1660131191090",
    "city": "UF_CRM_1653999902075",
    "productName": "UF_CRM_1654516980198",
    "productCategory": "UF_CRM_1565770347",
    "productCompanyCategories": "UF_CRM_1556562617",
    "productNegotiationCategories": "UF_CRM_1654001230422",
    "priceCategory": "UF_CRM_1654001339275",
    "productDescription": "UF_CRM_5CBD83328D86E",
    "preferredNetworkCategories": "UF_CRM_1669628251",
    "preferredNetworks": "UF_CRM_1654093180996",
    "presentationFileField": "UF_CRM_1745830954",
}
BITRIX_LEAD_FIELD_MAP = {
    "companyName": "UF_CRM_1782290699215",
    "contactName": "NAME",
    "contactLastName": "LAST_NAME",
    "contactPosition": "POST",
    "website": "WEB",
    "country": "ADDRESS_COUNTRY",
    "city": "ADDRESS_CITY",
    "leadContactName": "UF_CRM_1443605291",
    "productCategory": "UF_CRM_1771319497733",
    "productNegotiationCategories": "UF_CRM_1782133009644",
    "priceCategory": "UF_CRM_1782134816704",
    "productDescription": "UF_CRM_1555689023348",
    "preferredNetworks": "UF_CRM_1782132125417",
    "conversationComment": "UF_CRM_1782135754570",
    "region": "UF_CRM_1782132366289",
    "presentationFileField": "UF_CRM_1782217139829",
}
BITRIX_FIELD_MAP = BITRIX_DEAL_FIELD_MAP
BITRIX_AFTER_PRESENTATION_STAGE_ID = "C148:FINAL_INVOICE"


def _media_folder_stats(path: Path) -> dict[str, int]:
    if not path.exists():
        return {"files": 0, "images": 0, "dirs": 0}
    image_suffixes = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg"}
    files = [item for item in path.rglob("*") if item.is_file()]
    return {
        "files": len(files),
        "images": sum(1 for item in files if item.suffix.lower() in image_suffixes),
        "dirs": sum(1 for item in path.rglob("*") if item.is_dir()),
    }


def _log_media_folder_stats(name: str, path: Path) -> None:
    stats = _media_folder_stats(path)
    print(
        f"[data] {name}: найдено файлов={stats['files']}, "
        f"изображений={stats['images']}, подпапок={stats['dirs']}",
        flush=True,
    )


def _download_data_sources() -> None:
    if not DATA_SOURCES_FILE.exists():
        return

    config = json.loads(DATA_SOURCES_FILE.read_text(encoding="utf-8"))

    for name, info in config.get("sheets", {}).items():
        local_path = ROOT / info["local_file"]
        if local_path.exists():
            continue
        try:
            print(f"[data] Скачиваю {name}...", flush=True)
            req = urllib.request.Request(
                info["url"], headers={"User-Agent": "Mozilla/5.0"}
            )
            with DIRECT_HTTP_OPENER.open(req, timeout=30) as resp:
                data = resp.read()
            tmp = local_path.with_name(local_path.name + ".tmp")
            tmp.write_bytes(data)
            tmp.replace(local_path)
            print(f"[data] {name}: готово ({len(data):,} байт)", flush=True)
        except Exception as exc:
            print(f"[data] {name}: ошибка — {exc}", flush=True)


@app.on_event("startup")
def _on_startup() -> None:
    _download_data_sources()


def _safe_root_file(path_value: str) -> Path:
    path = Path(path_value)
    if not path.is_absolute():
        path = ROOT / path
    resolved = path.resolve()
    if ROOT.resolve() not in resolved.parents and resolved != ROOT.resolve():
        raise HTTPException(status_code=400, detail="Некорректный путь к файлу")
    return resolved


def _resolve_preview_image(path_value: str) -> tuple[Path, str | None]:
    target = _safe_root_file(path_value)
    if target.exists() and target.is_file():
        try:
            from build_presentation_direct import image_ext

            return target, image_ext(target)
        except Exception:
            return target, None
    return target, None


def _presentation_job_path(job_id: str) -> Path | None:
    if not re.fullmatch(r"[0-9a-fA-F]{16,64}", str(job_id or "")):
        return None
    return PRESENTATION_JOBS_DIR / f"{job_id}.json"


def _save_presentation_job(job_id: str, job: dict[str, Any]) -> None:
    path = _presentation_job_path(job_id)
    if path is None:
        return
    try:
        PRESENTATION_JOBS_DIR.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(path)
    except Exception as exc:
        print(f"[jobs] presentation job save failed: {exc}", flush=True)


def _load_presentation_job(job_id: str) -> dict[str, Any] | None:
    path = _presentation_job_path(job_id)
    if path is None or not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"[jobs] presentation job load failed: {exc}", flush=True)
        return None
    return data if isinstance(data, dict) else None


def _set_presentation_job(job_id: str, **values: Any) -> None:
    with PRESENTATION_JOBS_LOCK:
        job = PRESENTATION_JOBS.setdefault(job_id, {})
        job.update(values)
        job["updatedAt"] = datetime.now().isoformat(timespec="seconds")
        _save_presentation_job(job_id, job)


def _get_presentation_job(job_id: str) -> dict[str, Any]:
    with PRESENTATION_JOBS_LOCK:
        job = PRESENTATION_JOBS.get(job_id)
        if not job:
            job = _load_presentation_job(job_id)
            if job:
                PRESENTATION_JOBS[job_id] = dict(job)
        if not job:
            raise HTTPException(status_code=404, detail="Задача сборки не найдена")
        return dict(job)


def _presentation_audit_item_from_job(
    job_id: str,
    job: dict[str, Any],
    event: str | None = None,
) -> dict[str, Any]:
    item = {
        "event": event or ("PRESENTATION_DONE" if job.get("status") == "done" else "PRESENTATION_ERROR"),
        "jobId": job_id,
        "createdAt": job.get("createdAt") or job.get("updatedAt") or datetime.now().isoformat(timespec="seconds"),
        "updatedAt": job.get("updatedAt"),
        "companyName": job.get("companyName", ""),
        "status": job.get("status", ""),
        "source": job.get("source", ""),
        "sourceEntityId": job.get("sourceEntityId", ""),
        "selectionProvider": job.get("selectionProvider", ""),
        "fileName": job.get("fileName"),
        "downloadUrl": job.get("downloadUrl"),
        "pdfFileName": job.get("pdfFileName"),
        "pdfDownloadUrl": job.get("pdfDownloadUrl"),
        "templatePdfFileName": job.get("templatePdfFileName"),
        "templatePdfDownloadUrl": job.get("templatePdfDownloadUrl"),
        "pdfError": job.get("pdfError"),
        "templatePdfError": job.get("templatePdfError"),
        "error": job.get("error"),
        "message": job.get("message"),
    }
    return {key: value for key, value in item.items() if value not in (None, "")}


def _write_presentation_audit(job_id: str, event: str) -> None:
    try:
        job = _get_presentation_job(job_id)
    except Exception:
        job = {"id": job_id, "status": "unknown"}
    item = _presentation_audit_item_from_job(job_id, job, event)
    line = json.dumps(item, ensure_ascii=False, separators=(",", ":"))
    log_parts = [
        event,
        str(item.get("createdAt") or ""),
        str(item.get("companyName") or ""),
        f"job={job_id}",
        f"source={item.get('source', '')}",
        f"pptx={item.get('fileName', '')}",
        f"pdf={item.get('pdfFileName', '')}",
        f"template_pdf={item.get('templatePdfFileName', '')}",
    ]
    if item.get("error"):
        log_parts.append(f"error={item.get('error')}")
    print(" | ".join(log_parts), flush=True)
    try:
        with PRESENTATION_AUDIT_LOCK:
            PRESENTATION_AUDIT_FILE.parent.mkdir(parents=True, exist_ok=True)
            with PRESENTATION_AUDIT_FILE.open("a", encoding="utf-8") as audit_file:
                audit_file.write(line + "\n")
    except Exception as exc:
        print(f"[audit] presentation audit write failed: {exc}", flush=True)


def _read_presentation_audit(limit: int = 100) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    if PRESENTATION_AUDIT_FILE.exists():
        try:
            lines = PRESENTATION_AUDIT_FILE.read_text(encoding="utf-8").splitlines()
        except Exception as exc:
            print(f"[audit] presentation audit read failed: {exc}", flush=True)
            lines = []
        for line in reversed(lines):
            if not line.strip():
                continue
            try:
                item = json.loads(line)
            except Exception:
                continue
            if isinstance(item, dict):
                items.append(item)
            if len(items) >= limit:
                break

    seen_job_ids = {str(item.get("jobId") or "") for item in items}
    if PRESENTATION_JOBS_DIR.exists() and len(items) < limit:
        job_files = sorted(
            PRESENTATION_JOBS_DIR.glob("*.json"),
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )
        for path in job_files:
            job_id = path.stem
            if job_id in seen_job_ids:
                continue
            job = _load_presentation_job(job_id)
            if not job or job.get("status") not in {"done", "error"}:
                continue
            items.append(_presentation_audit_item_from_job(job_id, job))
            seen_job_ids.add(job_id)
            if len(items) >= limit:
                break

    items.sort(key=lambda item: str(item.get("updatedAt") or item.get("createdAt") or ""), reverse=True)
    return items[:limit]


def _open_local_file(path: Path) -> str:
    try:
        os.startfile(str(path.resolve()))  # type: ignore[attr-defined]
        return "opened"
    except Exception as exc:
        return f"open_failed: {exc}"


def _find_office_converter() -> str | None:
    configured = os.environ.get("PPTX_TO_PDF_CONVERTER", "").strip()
    if configured:
        return configured
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    return None


def _http_json(
    url: str,
    payload: dict[str, Any] | None = None,
    headers: dict[str, str] | None = None,
    timeout: int = 120,
) -> Any:
    body = json.dumps(payload or {}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json", **(headers or {})},
        method="POST",
    )
    try:
        with DIRECT_HTTP_OPENER.open(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} from {url}: {details}") from exc
    return json.loads(raw) if raw else {}


def _http_get_json(
    url: str,
    headers: dict[str, str] | None = None,
    timeout: int = 120,
) -> Any:
    req = urllib.request.Request(
        url,
        headers=headers or {},
        method="GET",
    )
    try:
        with DIRECT_HTTP_OPENER.open(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} from {url}: {details}") from exc
    return json.loads(raw) if raw else {}


def _multipart_form_data(fields: dict[str, str], files: dict[str, tuple[str, bytes, str]]) -> tuple[bytes, str]:
    boundary = f"----AutoSupplierCP{uuid.uuid4().hex}"
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.extend(
            [
                f"--{boundary}\r\n".encode("ascii"),
                f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode("ascii"),
                str(value).encode("utf-8"),
                b"\r\n",
            ]
        )
    for name, (filename, content, content_type) in files.items():
        safe_filename = filename.encode("utf-8", errors="ignore").decode("utf-8")
        chunks.extend(
            [
                f"--{boundary}\r\n".encode("ascii"),
                (
                    f'Content-Disposition: form-data; name="{name}"; '
                    f'filename="{safe_filename}"\r\n'
                ).encode("utf-8"),
                f"Content-Type: {content_type}\r\n\r\n".encode("ascii"),
                content,
                b"\r\n",
            ]
        )
    chunks.append(f"--{boundary}--\r\n".encode("ascii"))
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def _convert_pptx_to_pdf_ilovepdf(pptx_path: Path) -> Path | None:
    public_key = os.environ.get("ILOVEPDF_PUBLIC_KEY", "").strip()
    if not public_key:
        return None

    auth = _http_json(
        "https://api.ilovepdf.com/v1/auth",
        {"public_key": public_key},
        timeout=30,
    )
    token = str(auth.get("token") or "").strip()
    if not token:
        raise RuntimeError("iLovePDF auth did not return a token")
    auth_headers = {"Authorization": f"Bearer {token}"}

    region = os.environ.get("ILOVEPDF_REGION", "eu").strip() or "eu"
    start = _http_get_json(
        f"https://api.ilovepdf.com/v1/start/officepdf/{region}",
        auth_headers,
        timeout=30,
    )
    server = str(start.get("server") or "").strip()
    task = str(start.get("task") or "").strip()
    if not server or not task:
        raise RuntimeError(f"iLovePDF start did not return server/task: {start}")
    server_url = server if server.startswith("http") else f"https://{server}"

    upload_body, upload_type = _multipart_form_data(
        {"task": task},
        {
            "file": (
                "source.pptx",
                pptx_path.read_bytes(),
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        },
    )
    upload_req = urllib.request.Request(
        f"{server_url}/v1/upload",
        data=upload_body,
        headers={**auth_headers, "Content-Type": upload_type},
        method="POST",
    )
    try:
        with DIRECT_HTTP_OPENER.open(upload_req, timeout=120) as resp:
            upload = json.loads(resp.read().decode("utf-8", errors="replace"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"iLovePDF upload failed HTTP {exc.code}: {details}") from exc

    server_filename = str(upload.get("server_filename") or "").strip()
    if not server_filename:
        raise RuntimeError(f"iLovePDF upload did not return server_filename: {upload}")

    process = _http_json(
        f"{server_url}/v1/process",
        {
            "task": task,
            "tool": "officepdf",
            "files": [{"server_filename": server_filename, "filename": pptx_path.name}],
        },
        auth_headers,
        timeout=180,
    )
    status = str(process.get("status") or "").strip()
    if status and status not in {"TaskSuccess", "TaskSuccessWithWarnings"}:
        raise RuntimeError(f"iLovePDF process failed: {process}")

    download_req = urllib.request.Request(
        f"{server_url}/v1/download/{task}",
        headers=auth_headers,
        method="GET",
    )
    try:
        with DIRECT_HTTP_OPENER.open(download_req, timeout=120) as resp:
            data = resp.read()
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"iLovePDF download failed HTTP {exc.code}: {details}") from exc
    if not data.startswith(b"%PDF"):
        raise RuntimeError(f"iLovePDF download did not return a PDF ({len(data)} bytes)")

    pdf_path = pptx_path.with_suffix(".pdf")
    pdf_path.write_bytes(data)
    return pdf_path


def _convert_pptx_to_pdf(pptx_path: Path) -> Path:
    if not pptx_path.exists():
        raise RuntimeError(f"PPTX file not found: {pptx_path}")
    if not zipfile.is_zipfile(pptx_path):
        raise RuntimeError(f"PPTX file is not a valid zip package: {pptx_path.name}")

    external_pdf = _convert_pptx_to_pdf_ilovepdf(pptx_path)
    if external_pdf is not None:
        return external_pdf

    pdf_path = pptx_path.with_suffix(".pdf")
    converter = _find_office_converter()
    if converter:
        started_at = time.time()
        with tempfile.TemporaryDirectory(prefix="pptx_to_pdf_") as temp_dir_raw:
            temp_dir = Path(temp_dir_raw)
            profile_dir = temp_dir / "lo-profile"
            profile_dir.mkdir(parents=True, exist_ok=True)
            temp_pptx = temp_dir / "source.pptx"
            temp_pdf = temp_dir / "source.pdf"
            shutil.copy2(pptx_path, temp_pptx)
            env = os.environ.copy()
            env["HOME"] = str(temp_dir)
            env["LANG"] = "C.UTF-8"
            env["LC_ALL"] = "C.UTF-8"
            base_command = [
                    converter,
                    "--headless",
                    "--nologo",
                    "--nodefault",
                    "--nofirststartwizard",
                    "--nolockcheck",
                    "--norestore",
                    f"-env:UserInstallation=file://{profile_dir.resolve().as_posix()}",
            ]
            process = subprocess.run(
                [
                    *base_command,
                    "--convert-to",
                    "pdf:impress_pdf_Export",
                    "--outdir",
                    str(temp_dir.resolve()),
                    str(temp_pptx.resolve()),
                ],
                cwd=ROOT,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                timeout=120,
            )
            if process.returncode == 0 and temp_pdf.exists():
                shutil.move(str(temp_pdf), str(pdf_path))
                return pdf_path
            first_stdout = process.stdout
            first_stderr = process.stderr
            process = subprocess.run(
                [
                    *base_command,
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(temp_dir.resolve()),
                    str(temp_pptx.resolve()),
                ],
                cwd=ROOT,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                timeout=120,
            )
            if process.returncode == 0 and temp_pdf.exists():
                shutil.move(str(temp_pdf), str(pdf_path))
                return pdf_path
            temp_files = ", ".join(item.name for item in temp_dir.iterdir())
        if process.returncode != 0:
            details = "\n".join(
                part.strip()
                for part in (first_stderr, first_stdout, process.stderr, process.stdout)
                if part and part.strip()
            )
            raise RuntimeError(
                f"PDF conversion failed for {pptx_path.name} "
                f"({pptx_path.stat().st_size} bytes): {details}"
            )
        if pdf_path.exists():
            return pdf_path
        candidates = [
            item
            for item in pptx_path.parent.glob("*.pdf")
            if item.is_file() and item.stat().st_mtime >= started_at - 2
        ]
        if candidates:
            candidates.sort(key=lambda item: (item.stem == pptx_path.stem, item.stat().st_mtime), reverse=True)
            created_pdf = candidates[0]
            if created_pdf != pdf_path:
                try:
                    created_pdf.replace(pdf_path)
                    return pdf_path
                except Exception:
                    return created_pdf
            return created_pdf
        details = "\n".join(
            part.strip()
            for part in (first_stderr, first_stdout, process.stderr, process.stdout)
            if part and part.strip()
        )
        raise RuntimeError(
            "PDF conversion finished, but PDF file was not created"
            + f" for {pptx_path.name} ({pptx_path.stat().st_size} bytes)"
            + (f"; temp files: {temp_files}" if temp_files else "")
            + (f": {details}" if details else "")
        )

    if os.name == "nt":
        try:
            import win32com.client  # type: ignore[import-not-found]
        except Exception:
            win32com = None

        if win32com is not None:
            powerpoint = None
            presentation = None
            try:
                powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
                powerpoint.Visible = True
                presentation = powerpoint.Presentations.Open(str(pptx_path.resolve()), True, False, True)
                presentation.SaveAs(str(pdf_path.resolve()), 32)
            finally:
                if presentation is not None:
                    presentation.Close()
                if powerpoint is not None:
                    powerpoint.Quit()
            if pdf_path.exists():
                return pdf_path

        script = """
param([string]$pptxArg, [string]$pdfArg)
$ErrorActionPreference = 'Stop'
$pptx = [System.IO.Path]::GetFullPath($pptxArg)
$pdf = [System.IO.Path]::GetFullPath($pdfArg)
$powerpoint = $null
$presentation = $null
try {
  $powerpoint = New-Object -ComObject PowerPoint.Application
  $powerpoint.Visible = -1
  $presentation = $powerpoint.Presentations.Open($pptx, $true, $false, $true)
  $presentation.SaveAs($pdf, 32)
} finally {
  if ($presentation -ne $null) { $presentation.Close() }
  if ($powerpoint -ne $null) { $powerpoint.Quit() }
}
"""
        script_path: Path | None = None
        try:
            script_file = tempfile.NamedTemporaryFile(
                "w",
                encoding="utf-8-sig",
                delete=False,
                dir=ROOT,
                prefix="pptx_to_pdf_",
                suffix=".ps1",
            )
            script_path = Path(script_file.name)
            script_file.write(script)
            script_file.close()
            process = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script_path),
                    str(pptx_path.resolve()),
                    str(pdf_path.resolve()),
                ],
                cwd=ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                timeout=120,
            )
        finally:
            if script_path:
                script_path.unlink(missing_ok=True)
        if process.returncode != 0:
            details = (process.stderr or process.stdout or "").strip()
            raise RuntimeError(f"PowerPoint PDF conversion failed: {details}")
        if pdf_path.exists():
            return pdf_path

    raise RuntimeError("PDF conversion requires LibreOffice/soffice or Microsoft PowerPoint")


def _apply_payload_override_env(env: dict[str, str], form_payload: dict[str, Any] | None = None) -> None:
    payload_override = _payload_override_from_form(form_payload)
    if payload_override:
        env["PAYLOAD_OVERRIDE_JSON"] = json.dumps(payload_override, ensure_ascii=False)


def _build_template_pdf(
    company: str,
    selection: dict[str, Any] | None = None,
    payload_override: dict[str, Any] | None = None,
) -> tuple[Path, dict[str, Any]]:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    if selection:
        env["PRESENTATION_SELECTION_JSON"] = json.dumps(selection, ensure_ascii=False)
    _apply_payload_override_env(env, payload_override)
    process = subprocess.run(
        [sys.executable, "-X", "utf8", "build_pdf_template_direct.py", company],
        cwd=ROOT,
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        timeout=180,
    )
    if process.returncode != 0:
        details = (process.stderr or process.stdout or "Не удалось сформировать PDF из шаблона").strip()
        raise RuntimeError(details)
    report = json.loads(process.stdout)
    created = _safe_root_file(str(report.get("created", "")))
    if not created.exists() or created.suffix.lower() != ".pdf":
        raise RuntimeError("PDF из шаблона не был создан")
    return created, report


def _run_presentation_job(
    job_id: str,
    company: str,
    selection: dict[str, Any] | None = None,
    make_pdf: bool = False,
    require_pdf: bool = False,
    payload_override: dict[str, Any] | None = None,
    source: str = "manual",
    source_entity_id: str = "",
) -> None:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    if selection:
        env["PRESENTATION_SELECTION_JSON"] = json.dumps(selection, ensure_ascii=False)
    _apply_payload_override_env(env, payload_override)
    command = [sys.executable, "-X", "utf8", "build_presentation_direct_new.py", company]
    stdout_path: Path | None = None
    stderr_path: Path | None = None
    _set_presentation_job(job_id, source=source, sourceEntityId=source_entity_id)

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

        open_status = _open_local_file(created)
        _set_presentation_job(
            job_id,
            status="running",
            progress=94 if make_pdf else 100,
            message="PPTX готов, начинаю конвертацию в PDF" if make_pdf else "Презентация готова",
            fileName=created.name,
            downloadUrl=f"/api/presentations/{urllib.parse.quote(created.name)}",
            openStatus=open_status,
            report=report,
        )

        pdf_path: Path | None = None
        pdf_error: str | None = None
        if make_pdf:
            _set_presentation_job(job_id, status="running", progress=96, message="Конвертирую презентацию в PDF")
            try:
                pdf_path = _convert_pptx_to_pdf(created)
            except Exception as exc:
                pdf_error = str(exc) or "PDF conversion failed"
                if require_pdf:
                    raise

        template_pdf_path: Path | None = None
        template_pdf_error: str | None = None
        template_pdf_report: dict[str, Any] | None = None
        _set_presentation_job(job_id, status="running", progress=98, message="Собираю PDF из PDF-шаблона")
        try:
            template_pdf_path, template_pdf_report = _build_template_pdf(company, selection, payload_override)
        except Exception as exc:
            template_pdf_error = str(exc) or "PDF template generation failed"

        _set_presentation_job(
            job_id,
            status="done",
            progress=100,
            message="Презентации готовы",
            fileName=created.name,
            downloadUrl=f"/api/presentations/{urllib.parse.quote(created.name)}",
            pdfFileName=pdf_path.name if pdf_path is not None else None,
            pdfDownloadUrl=(
                f"/api/presentations/{urllib.parse.quote(pdf_path.name)}"
                if pdf_path is not None
                else None
            ),
            pdfError=pdf_error,
            templatePdfFileName=template_pdf_path.name if template_pdf_path is not None else None,
            templatePdfDownloadUrl=(
                f"/api/presentations/{urllib.parse.quote(template_pdf_path.name)}"
                if template_pdf_path is not None
                else None
            ),
            templatePdfError=template_pdf_error,
            openStatus=open_status,
            source=source,
            sourceEntityId=source_entity_id,
            report={**report, "template_pdf": template_pdf_report},
        )
        _write_presentation_audit(job_id, "PRESENTATION_DONE")
    except Exception as exc:
        _set_presentation_job(
            job_id,
            status="error",
            progress=100,
            message=str(exc) or "Не удалось сформировать презентацию",
            error=str(exc) or "Не удалось сформировать презентацию",
        )
        _write_presentation_audit(job_id, "PRESENTATION_ERROR")
    finally:
        if stdout_path:
            stdout_path.unlink(missing_ok=True)
        if stderr_path:
            stderr_path.unlink(missing_ok=True)


HEADERS = [
    "N",
    "Название компании",
    "ИНН",
    "Тип компании",
    "ФИО контакта",
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
    "Предпочтительные категории сетей",
    "Предпочтительные сети",
    "Ценовая категория",
    "Фото товара",
    "Дата сохранения",
    "DaData: название",
    "DaData: статус",
    "DaData: юр адрес",
    "DaData: основной вид деятельности",
    "DaData: среднесписочная численность",
    "DaData: уставной капитал",
    "DaData: реестр МСП",
    "DaData: специальный налоговый режим",
    "DaData: доходы",
    "DaData: расходы",
    "DaData: недоимки",
    "DaData: штрафы",
    "DaData: ген. директор",
    "DaData: обновлено",
]

FIELD_TO_HEADER = {
    "companyName": "Название компании",
    "inn": "ИНН",
    "companyType": "Тип компании",
    "contactName": "ФИО контакта",
    "contactPosition": "Должность контакта",
    "productName": "Название товара",
    "productCategory": "Категория товара",
    "productDescription": "Краткое описание продукции",
    "industry": "Отрасль",
    "activity": "Сфера деятельности",
    "networkCategories": "Категория по работе с сетями",
    "preferredNetworkCategories": "Предпочтительные категории сетей",
    "website": "Сайт ",
    "country": "Страна ",
    "city": "Город",
    "preferredNetworks": "Предпочтительные сети",
    "priceCategory": "Ценовая категория",
    "dadataName": "DaData: название",
    "dadataStatus": "DaData: статус",
    "dadataLegalAddress": "DaData: юр адрес",
    "dadataMainActivity": "DaData: основной вид деятельности",
    "dadataEmployeeCount": "DaData: среднесписочная численность",
    "dadataCapital": "DaData: уставной капитал",
    "dadataSmb": "DaData: реестр МСП",
    "dadataTaxSystem": "DaData: специальный налоговый режим",
    "dadataIncome": "DaData: доходы",
    "dadataExpense": "DaData: расходы",
    "dadataDebt": "DaData: недоимки",
    "dadataPenalty": "DaData: штрафы",
    "dadataManager": "DaData: ген. директор",
    "dadataUpdatedAt": "DaData: обновлено",
}

DADATA_FIELDS = [
    "dadataName",
    "dadataStatus",
    "dadataLegalAddress",
    "dadataMainActivity",
    "dadataEmployeeCount",
    "dadataCapital",
    "dadataSmb",
    "dadataTaxSystem",
    "dadataIncome",
    "dadataExpense",
    "dadataDebt",
    "dadataPenalty",
    "dadataManager",
    "dadataUpdatedAt",
]

DEFAULT_OPTIONS = {
    "companyTypes": [
        "Производитель конечного товара",
        "Дистрибьютор",
        "Импортер",
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
        dimension.set("ref", f"A1:{_number_to_col(len(HEADERS))}{max(_max_used_row(sheet_data, shared), 1)}")


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


def _replace_row_values(row: ET.Element, row_number: int, values_by_col: dict[str, Any]) -> None:
    ns = f"{{{MAIN_NS}}}"
    for col, value in values_by_col.items():
        target_ref = _cell_ref(col, row_number)
        for cell in list(row.findall(f"{ns}c")):
            if cell.attrib.get("r") == target_ref:
                row.remove(cell)
        row.append(_inline_cell(col, row_number, value))
    _sort_cells(row)


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

    data = dict(data)
    if _digits(str(data.get("inn", ""))) and not any(str(data.get(field, "")).strip() for field in DADATA_FIELDS):
        try:
            data.update(_fetch_dadata_company(_digits(str(data.get("inn", "")))))
        except HTTPException:
            pass

    row_values = [""] * len(HEADERS)
    row_values[0] = str(row_number - 1)
    header_indexes = {header: index for index, header in enumerate(HEADERS)}
    for field, header in FIELD_TO_HEADER.items():
        if field in data:
            row_values[header_indexes[header]] = str(data.get(field, "")).strip()
        else:
            row_values[header_indexes[header]] = existing_values.get(header, "")
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
        dimension.set("ref", f"A1:{_number_to_col(len(HEADERS))}{row_number}")

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


def _update_card_inn(row_number: int, inn: str) -> None:
    inn = _digits(inn)
    if not _is_valid_inn(inn):
        raise HTTPException(status_code=400, detail="Некорректный ИНН")
    if row_number <= 1:
        raise HTTPException(status_code=400, detail="Некорректная строка карточки")

    names, parts = _xlsx_parts()
    shared = _shared_strings(parts)
    sheet_path = _sheet_target(parts, "Карточка клиента")
    root = ET.fromstring(parts[sheet_path])
    ns = f"{{{MAIN_NS}}}"
    sheet_data = root.find(f"{ns}sheetData")
    if sheet_data is None:
        raise HTTPException(status_code=404, detail="Лист карточек пуст")

    _ensure_headers(root, sheet_data, shared)
    header_row = sheet_data.find(f"{ns}row[@r='1']")
    target_row = sheet_data.find(f"{ns}row[@r='{row_number}']")
    if header_row is None or target_row is None:
        raise HTTPException(status_code=404, detail="Карточка не найдена")

    headers_by_col = _row_values(header_row, shared)
    inn_col = next((col for col, header in headers_by_col.items() if header.strip() == "ИНН"), None)
    if not inn_col:
        raise HTTPException(status_code=500, detail="Колонка ИНН не найдена")

    target_ref = _cell_ref(inn_col, row_number)
    for cell in list(target_row.findall(f"{ns}c")):
        if cell.attrib.get("r") == target_ref:
            target_row.remove(cell)
    target_row.append(_inline_cell(inn_col, row_number, inn))
    _sort_cells(target_row)

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
            detail="Не удалось сохранить ИНН. Закройте Excel-файл 'Данные для разработки.xlsx' и попробуйте снова.",
        ) from exc
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def _update_card_dadata(row_number: int, dadata: dict[str, str]) -> None:
    if row_number <= 1:
        raise HTTPException(status_code=400, detail="Некорректная строка карточки")

    names, parts = _xlsx_parts()
    shared = _shared_strings(parts)
    sheet_path = _sheet_target(parts, "Карточка клиента")
    root = ET.fromstring(parts[sheet_path])
    ns = f"{{{MAIN_NS}}}"
    sheet_data = root.find(f"{ns}sheetData")
    if sheet_data is None:
        raise HTTPException(status_code=404, detail="Лист карточек пуст")

    _ensure_headers(root, sheet_data, shared)
    header_row = sheet_data.find(f"{ns}row[@r='1']")
    target_row = sheet_data.find(f"{ns}row[@r='{row_number}']")
    if header_row is None or target_row is None:
        raise HTTPException(status_code=404, detail="Карточка не найдена")

    headers_by_col = _row_values(header_row, shared)
    values_by_col = {}
    for field in DADATA_FIELDS:
        header = FIELD_TO_HEADER[field]
        col = next((col for col, value in headers_by_col.items() if value.strip() == header), None)
        if col:
            values_by_col[col] = dadata.get(field, "")
    _replace_row_values(target_row, row_number, values_by_col)

    dimension = root.find(f"{ns}dimension")
    if dimension is not None:
        dimension.set("ref", f"A1:{_number_to_col(len(HEADERS))}{max(_max_used_row(sheet_data, shared), row_number)}")
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
            detail="Не удалось сохранить сведения DaData. Закройте Excel-файл и попробуйте снова.",
        ) from exc
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


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


def _read_sheet_rows(xlsx_path: Path, preferred_sheet_name: str | None = None) -> list[list[str]]:
    if not xlsx_path.exists():
        return []
    with zipfile.ZipFile(xlsx_path, "r") as source:
        parts = {name: source.read(name) for name in source.namelist()}
    shared = _shared_strings(parts)

    sheet_path = ""
    if preferred_sheet_name:
        try:
            sheet_path = _sheet_target(parts, preferred_sheet_name)
        except HTTPException:
            sheet_path = ""
    if not sheet_path:
        workbook = ET.fromstring(parts["xl/workbook.xml"])
        rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        ns = {"m": MAIN_NS, "r": REL_NS}
        first_sheet = next(iter(workbook.find("m:sheets", ns) or []), None)
        if first_sheet is None:
            return []
        rid = first_sheet.attrib[f"{{{REL_NS}}}id"]
        sheet_path = "xl/" + rel_map[rid].lstrip("/")

    root = ET.fromstring(parts[sheet_path])
    ns = f"{{{MAIN_NS}}}"
    rows: list[list[str]] = []
    for row in root.findall(f".//{ns}row"):
        values_by_col = _row_values(row, shared)
        if not values_by_col:
            continue
        max_index = max(_col_to_number(col) for col in values_by_col)
        ordered = [""] * (max_index + 1)
        for col, value in values_by_col.items():
            ordered[_col_to_number(col)] = value
        if any(str(value).strip() for value in ordered):
            rows.append(ordered)
    return rows


def _preferred_network_names() -> list[str]:
    rows = _read_sheet_rows(PHOTO_CATALOG_FILE, "Сети")
    if len(rows) <= 1:
        return []
    headers = [str(value).strip() for value in rows[0]]
    if "Сеть" not in headers:
        return []
    network_index = headers.index("Сеть")
    result: list[str] = []
    seen: set[str] = set()
    for row in rows[1:]:
        if network_index >= len(row):
            continue
        name = str(row[network_index]).strip()
        key = name.casefold()
        if name and key not in seen:
            result.append(name)
            seen.add(key)
    return result


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


def _digits(value: str) -> str:
    return re.sub(r"\D+", "", value or "")


def _is_valid_inn(inn: str) -> bool:
    inn = _digits(inn)
    if len(inn) == 10:
        weights = [2, 4, 10, 3, 5, 9, 4, 6, 8]
        return sum(int(inn[index]) * weights[index] for index in range(9)) % 11 % 10 == int(inn[9])
    if len(inn) == 12:
        weights_11 = [7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
        weights_12 = [3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
        check_11 = sum(int(inn[index]) * weights_11[index] for index in range(10)) % 11 % 10
        check_12 = sum(int(inn[index]) * weights_12[index] for index in range(11)) % 11 % 10
        return check_11 == int(inn[10]) and check_12 == int(inn[11])
    return False


def _normalize_website_url(website: str) -> str:
    website = website.strip()
    if not website:
        raise HTTPException(status_code=400, detail="Укажите сайт компании")
    if not re.match(r"^https?://", website, re.IGNORECASE):
        website = "https://" + website
    parsed = urllib.parse.urlparse(website)
    if not parsed.netloc:
        raise HTTPException(status_code=400, detail="Некорректный адрес сайта")
    return urllib.parse.urlunparse((parsed.scheme, parsed.netloc, parsed.path.rstrip("/"), "", "", ""))


def _site_probe_urls(website: str) -> list[str]:
    base = _normalize_website_url(website)
    parsed = urllib.parse.urlparse(base)
    root = f"{parsed.scheme}://{parsed.netloc}"
    paths = [
        parsed.path if parsed.path else "",
        "/contacts",
        "/contact",
        "/kontakty",
        "/rekvizity",
        "/requisites",
        "/about",
        "/privacy",
        "/privacy-policy",
        "/politika-konfidentsialnosti",
    ]
    urls: list[str] = []
    for path in paths:
        url = root + (path if path.startswith("/") else f"/{path}" if path else "")
        if url not in urls:
            urls.append(url)
    return urls[:10]


def _read_public_text(url: str) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 AutoSupplierCP",
            "Accept": "text/html,application/xhtml+xml,application/pdf,text/plain,*/*;q=0.8",
        },
    )
    with urllib.request.urlopen(request, timeout=7) as response:
        content_type = response.headers.get("Content-Type", "")
        if "pdf" in content_type.lower():
            return ""
        data = response.read(700_000)
    charset_match = re.search(r"charset=([\w\-]+)", content_type, re.IGNORECASE)
    encodings = [charset_match.group(1)] if charset_match else []
    encodings.extend(["utf-8", "windows-1251"])
    for encoding in encodings:
        try:
            return data.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", errors="ignore")


def _extract_inns(text: str) -> list[str]:
    candidates: list[str] = []
    normalized = re.sub(r"[\u00a0\s]+", " ", text)
    patterns = [
        r"(?i)(?:ИНН|INN)\D{0,40}((?:\d[\s\-]*){10,12})",
        r"\b((?:\d[\s\-]*){10}|(?:\d[\s\-]*){12})\b",
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, normalized):
            inn = _digits(match.group(1))
            if inn not in candidates and _is_valid_inn(inn):
                candidates.append(inn)
    return candidates


def _repair_mojibake(value: Any) -> str:
    text = str(value or "")
    if "Ð" not in text and "Ñ" not in text:
        return text
    try:
        return text.encode("latin1").decode("utf-8")
    except UnicodeError:
        return text


def _format_money(value: Any) -> str:
    if value in (None, ""):
        return ""
    try:
        amount = float(value)
    except (TypeError, ValueError):
        return str(value)
    if amount.is_integer():
        return f"{int(amount):,}".replace(",", " ")
    return f"{amount:,.2f}".replace(",", " ")


def _format_count(value: Any) -> str:
    if value in (None, ""):
        return ""
    try:
        return f"{int(value):,}".replace(",", " ")
    except (TypeError, ValueError):
        return str(value)


def _status_label(value: Any) -> str:
    labels = {
        "ACTIVE": "Действующая",
        "LIQUIDATING": "Ликвидируется",
        "LIQUIDATED": "Ликвидирована",
        "BANKRUPT": "Банкротство",
        "REORGANIZING": "Реорганизуется",
    }
    text = str(value or "")
    return labels.get(text, text)


def _tax_system_label(value: Any) -> str:
    labels = {
        "AUSN": "АУСН",
        "ESHN": "ЕСХН",
        "SRP": "СРП",
        "USN": "УСН",
    }
    text = str(value or "")
    return labels.get(text, text)


def _smb_label(smb: Any) -> str:
    if not isinstance(smb, dict):
        return ""
    labels = {"MICRO": "микро", "SMALL": "малое", "MEDIUM": "среднее"}
    category = labels.get(str(smb.get("category") or ""), str(smb.get("category") or ""))
    return category or "есть запись"


def _dadata_post(url: str, payload: dict[str, Any], timeout: int = 14) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Authorization": f"Token {DADATA_API_KEY}",
            "User-Agent": "AutoSupplierCP",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _okved_name(code: str) -> str:
    code = str(code or "").strip()
    if not code:
        return ""
    try:
        payload = _dadata_post(DADATA_FIND_OKVED_URL, {"query": code, "count": 1})
    except Exception:
        return ""
    suggestions = payload.get("suggestions") or []
    if not suggestions:
        return ""
    data = suggestions[0].get("data") or {}
    return str(data.get("name") or suggestions[0].get("value") or "")


def _main_okved(data: dict[str, Any]) -> str:
    for item in data.get("okveds") or []:
        if item.get("main"):
            code = str(item.get("code") or "")
            name = str(item.get("name") or "")
            return f"{code} {name}".strip()
    code = str(data.get("okved") or "")
    name = _okved_name(code)
    return f"{code} - {name}".strip(" -") if name else code


def _manager_name(data: dict[str, Any]) -> str:
    management = data.get("management") if isinstance(data.get("management"), dict) else {}
    name = str(management.get("name") or "")
    post = str(management.get("post") or "")
    if name:
        return f"{post}: {name}" if post else name

    managers = data.get("managers") or []
    if managers:
        manager = managers[0]
        fio = manager.get("fio") if isinstance(manager.get("fio"), dict) else {}
        source = fio.get("source") or manager.get("name") or ""
        post = manager.get("post") or ""
        return f"{post}: {source}" if post and source else str(source or "")
    return ""


def _dadata_company_from_suggestion(suggestion: dict[str, Any]) -> dict[str, str]:
    data = suggestion.get("data") or {}
    finance = data.get("finance") if isinstance(data.get("finance"), dict) else {}
    capital = data.get("capital") if isinstance(data.get("capital"), dict) else {}
    documents = data.get("documents") if isinstance(data.get("documents"), dict) else {}
    name = data.get("name") if isinstance(data.get("name"), dict) else {}
    state = data.get("state") if isinstance(data.get("state"), dict) else {}
    address = data.get("address") if isinstance(data.get("address"), dict) else {}
    capital_value = capital.get("value")
    capital_text = ""
    if capital_value not in (None, ""):
        capital_type = str(capital.get("type") or "Уставной капитал")
        capital_text = f"{capital_type}: {_format_money(capital_value)}"

    selection = {
        "dadataName": str(name.get("full_with_opf") or suggestion.get("value") or ""),
        "dadataStatus": _status_label(state.get("status")),
        "dadataLegalAddress": str(address.get("unrestricted_value") or address.get("value") or ""),
        "dadataMainActivity": _main_okved(data),
        "dadataEmployeeCount": _format_count(data.get("employee_count")),
        "dadataCapital": capital_text,
        "dadataSmb": _smb_label(documents.get("smb")),
        "dadataTaxSystem": _tax_system_label(finance.get("tax_system")),
        "dadataIncome": _format_money(finance.get("income")),
        "dadataExpense": _format_money(finance.get("expense")),
        "dadataDebt": _format_money(finance.get("debt")),
        "dadataPenalty": _format_money(finance.get("penalty")),
        "dadataManager": _manager_name(data),
        "dadataUpdatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _fetch_dadata_company(inn: str) -> dict[str, str]:
    inn = _digits(inn)
    if not _is_valid_inn(inn):
        raise HTTPException(status_code=400, detail="Некорректный ИНН для запроса DaData")
    if not DADATA_API_KEY:
        raise HTTPException(status_code=500, detail="Не задан API-ключ DaData")

    try:
        payload = _dadata_post(DADATA_FIND_PARTY_URL, {"query": inn, "count": 1, "branch_type": "MAIN"})
    except Exception as exc:
        raise HTTPException(status_code=502, detail="Не удалось получить сведения DaData") from exc

    suggestions = payload.get("suggestions") or []
    if not suggestions:
        raise HTTPException(status_code=404, detail="DaData не нашла компанию по этому ИНН")
    return _dadata_company_from_suggestion(suggestions[0])


def _find_site_inns(website: str) -> list[dict[str, str]]:
    found: dict[str, str] = {}
    for url in _site_probe_urls(website):
        try:
            text = _read_public_text(url)
        except Exception:
            continue
        for inn in _extract_inns(text):
            found.setdefault(inn, url)
        if found:
            break
    return [{"inn": inn, "sourceUrl": source_url} for inn, source_url in found.items()]


def _run_presentation_dry_run(company: str) -> dict[str, Any]:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    command = [sys.executable, "-X", "utf8", "build_presentation_direct_new.py", company, "--dry-run"]
    try:
        result = subprocess.run(
            command,
            cwd=ROOT,
            env=env,
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=90,
        )
    except subprocess.CalledProcessError as exc:
        details = (exc.stderr or exc.stdout or "Не удалось подготовить выбор материалов").strip()
        raise HTTPException(status_code=500, detail=details) from exc
    except subprocess.TimeoutExpired as exc:
        raise HTTPException(status_code=504, detail="Подготовка выбора материалов заняла слишком много времени") from exc

    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=500, detail="Сборщик вернул некорректный JSON dry-run") from exc


def _payload_override_from_form(form_payload: dict[str, Any] | None) -> dict[str, str]:
    source = form_payload or {}
    result: dict[str, str] = {}
    for key, value in source.items():
        if key not in FIELD_TO_HEADER:
            continue
        normalized = str(value or "").strip()
        if normalized:
            result[FIELD_TO_HEADER[key]] = normalized
    return result


def _run_presentation_dry_run_with_override(company: str, form_payload: dict[str, Any] | None = None) -> dict[str, Any]:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    payload_override = _payload_override_from_form(form_payload)
    if payload_override:
        env["PAYLOAD_OVERRIDE_JSON"] = json.dumps(payload_override, ensure_ascii=False)
    command = [sys.executable, "-X", "utf8", "build_presentation_direct_new.py", company, "--dry-run"]
    try:
        result = subprocess.run(
            command,
            cwd=ROOT,
            env=env,
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=90,
        )
    except subprocess.CalledProcessError as exc:
        details = (exc.stderr or exc.stdout or "Не удалось подготовить выбор материалов").strip()
        raise HTTPException(status_code=500, detail=details) from exc
    except subprocess.TimeoutExpired as exc:
        raise HTTPException(status_code=504, detail="Подготовка выбора материалов заняла слишком много времени") from exc

    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=500, detail="Сборщик вернул некорректный JSON dry-run") from exc


def _reviews_from_replacements(replacements: dict[str, Any]) -> list[dict[str, str]]:
    review_keys = [
        ("{{block15}}", "{{block16}}", "{{block17}}"),
        ("{{block18}}", "{{block19}}", "{{block20}}"),
        ("{{block21}}", "{{block22}}", "{{block23}}"),
    ]
    reviews = []
    for company_key, person_key, text_key in review_keys:
        if replacements.get(company_key) or replacements.get(text_key):
            reviews.append(
                {
                    "company": str(replacements.get(company_key, "")),
                    "person": str(replacements.get(person_key, "")),
                    "text": str(replacements.get(text_key, "")),
                }
            )
    return reviews


def _candidate_map(items: list[dict[str, Any]], key: str = "id") -> dict[str, dict[str, Any]]:
    mapped: dict[str, dict[str, Any]] = {}
    for item in items:
        item_key = str(item.get(key, "")).strip()
        if item_key:
            mapped[item_key] = item
    return mapped


def _normalize_selected_ids(
    requested_ids: list[str] | None,
    candidates: list[dict[str, Any]],
    default_ids: list[str],
    required_count: int,
) -> list[str]:
    valid_ids = {str(item.get("id", "")) for item in candidates if item.get("id")}
    selected: list[str] = []
    for item_id in requested_ids or []:
        normalized = str(item_id).strip()
        if normalized and normalized in valid_ids and normalized not in selected:
            selected.append(normalized)
        if len(selected) >= required_count:
            break
    for item_id in default_ids:
        normalized = str(item_id).strip()
        if normalized and normalized in valid_ids and normalized not in selected:
            selected.append(normalized)
        if len(selected) >= required_count:
            break
    return selected[:required_count]


def _normalize_photo_selected_ids(
    requested_ids: list[str] | None,
    candidates: list[dict[str, Any]],
    default_ids: list[str],
    required_count: int,
) -> list[str]:
    preferred_ids = [
        str(item.get("id", "")).strip()
        for item in candidates
        if item.get("preferred_match") and str(item.get("id", "")).strip()
    ]
    if not preferred_ids:
        return _normalize_selected_ids(requested_ids, candidates, default_ids, required_count)

    valid_ids = {str(item.get("id", "")) for item in candidates if item.get("id")}
    selected: list[str] = []
    for item_id in preferred_ids:
        normalized = str(item_id).strip()
        if normalized and normalized in valid_ids and normalized not in selected:
            selected.append(normalized)
        if len(selected) >= required_count:
            return selected[:required_count]
    for item_id in requested_ids or []:
        normalized = str(item_id).strip()
        if normalized and normalized in valid_ids and normalized not in selected:
            selected.append(normalized)
        if len(selected) >= required_count:
            break
    for item_id in default_ids:
        normalized = str(item_id).strip()
        if normalized and normalized in valid_ids and normalized not in selected:
            selected.append(normalized)
        if len(selected) >= required_count:
            break
    return selected[:required_count]


def _build_selection_from_choices(
    base_selection: dict[str, Any],
    photo_ids: list[str] | None = None,
    review_ids: list[str] | None = None,
) -> dict[str, Any]:
    selection = dict(base_selection)
    selection["planned_images"] = dict(base_selection.get("planned_images") or {})
    selection["replacements"] = dict(base_selection.get("replacements") or {})
    selection["photo_candidates"] = [dict(item) for item in (base_selection.get("photo_candidates") or [])]
    selection["review_candidates"] = [dict(item) for item in (base_selection.get("review_candidates") or [])]

    photo_candidates = _candidate_map(selection["photo_candidates"])
    review_candidates = _candidate_map(selection["review_candidates"])
    photo_tokens = [str(item) for item in (base_selection.get("photo_tokens") or []) if str(item).startswith("{{pic")]
    required_photo_count = int(base_selection.get("required_photo_count") or len(photo_tokens) or 0)
    required_review_count = int(base_selection.get("required_review_count") or 3)

    selection["selected_photo_ids"] = _normalize_photo_selected_ids(
        photo_ids,
        selection["photo_candidates"],
        [str(item) for item in (base_selection.get("selected_photo_ids") or [])],
        required_photo_count,
    )
    selection["selected_review_ids"] = _normalize_selected_ids(
        review_ids,
        selection["review_candidates"],
        [str(item) for item in (base_selection.get("selected_review_ids") or [])],
        required_review_count,
    )

    for token in photo_tokens:
        selection["planned_images"].pop(token, None)
    for token, selected_id in zip(photo_tokens, selection["selected_photo_ids"], strict=False):
        selected_option = photo_candidates.get(selected_id)
        if not selected_option:
            continue
        selection["planned_images"][token] = str(selected_option.get("path", ""))

    review_tokens = {
        0: ("{{block15}}", "{{block16}}", "{{block17}}"),
        1: ("{{block18}}", "{{block19}}", "{{block20}}"),
        2: ("{{block21}}", "{{block22}}", "{{block23}}"),
    }
    for company_key, person_key, text_key in review_tokens.values():
        selection["replacements"][company_key] = ""
        selection["replacements"][person_key] = ""
        selection["replacements"][text_key] = ""
    for index, selected_id in enumerate(selection["selected_review_ids"]):
        selected_option = review_candidates.get(selected_id)
        if not selected_option or index not in review_tokens:
            continue
        if not selected_option:
            continue
        company_key, person_key, text_key = review_tokens[index]
        selection["replacements"][company_key] = str(selected_option.get("company", ""))
        selection["replacements"][person_key] = str(selected_option.get("person", ""))
        selection["replacements"][text_key] = str(selected_option.get("text", ""))

    selection["reviews"] = _reviews_from_replacements(selection["replacements"])
    return selection


def _selected_photo_summary(selection: dict[str, Any]) -> list[dict[str, Any]]:
    candidates = _candidate_map([dict(item) for item in (selection.get("photo_candidates") or [])])
    summary = []
    for item_id in selection.get("selected_photo_ids") or []:
        candidate = candidates.get(str(item_id))
        if not candidate:
            continue
        summary.append(
            {
                "id": str(item_id),
                "network": str(candidate.get("network", "")),
                "path": str(candidate.get("path", "")),
                "preferredMatch": bool(candidate.get("preferred_match")),
            }
        )
    return summary


def _default_selection_from_report(report: dict[str, Any], provider: str) -> dict[str, Any]:
    planned_images = {
        key: value
        for key, value in (report.get("planned_images") or {}).items()
        if key.startswith("{{") and isinstance(value, str)
    }
    replacements = dict(report.get("replacements") or {})
    reviews = _reviews_from_replacements(replacements)

    selection = {
        "provider": provider,
        "source": "rules",
        "client": report.get("client", ""),
        "category": report.get("category", ""),
        "stats": {
            "source": report.get("stats_source", ""),
            "category": replacements.get("{{block3}}", report.get("category", "")),
            "contracts": replacements.get("{{block6}}", ""),
            "average_check": replacements.get("{{block31}}", ""),
            "contracts_per_supplier": replacements.get("{{block8}}", ""),
            "buyers": replacements.get("{{block9}}", ""),
        },
        "reviews": reviews,
        "planned_images": planned_images,
        "replacements": replacements,
        "photo_candidates": report.get("photo_candidates") or [],
        "review_candidates": report.get("review_candidates") or [],
        "selected_photo_ids": report.get("selected_photo_ids") or [],
        "selected_review_ids": report.get("selected_review_ids") or [],
        "required_photo_count": int(report.get("required_photo_count") or 0),
        "required_review_count": int(report.get("required_review_count") or 3),
        "photo_tokens": report.get("photo_tokens") or [],
        "rationale": {
            "summary": "Выбор подготовлен текущими правилами проекта.",
            "stats": "Статистика выбрана по товарной категории или ближайшему доступному бенчмарку.",
            "photos": "Фотографии выбраны по каталогу, приоритетам, предпочтительным сетям и универсальности.",
            "reviews": "Отзывы выбраны по близости категории, типу компании, ценовому сегменту и универсальности.",
        },
    }
    return _build_selection_from_choices(selection)


def _ai_provider_settings(provider: str) -> dict[str, str]:
    normalized = provider.strip().lower()
    if normalized not in {"openai", "deepseek"}:
        return {
            "provider": "rules",
            "api_key": "",
            "base_url": "",
            "model": "",
            "chat_completions_path": "/chat/completions",
            "timeout_seconds": "45",
            "temperature": "0.2",
            "max_tokens": "",
        }

    defaults = {
        "openai": {
            "api_key": os.environ.get("OPENAI_API_KEY", ""),
            "base_url": os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1"),
            "model": os.environ.get("OPENAI_MODEL", "gpt-4.1-mini"),
        },
        "deepseek": {
            "api_key": os.environ.get("DEEPSEEK_API_KEY", ""),
            "base_url": os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
            "model": os.environ.get("DEEPSEEK_MODEL", "deepseek-chat"),
        },
    }
    settings = {
        "provider": normalized,
        "chat_completions_path": "/chat/completions",
        "timeout_seconds": os.environ.get("AI_TIMEOUT_SECONDS", "120"),
        "temperature": os.environ.get("AI_TEMPERATURE", "0.2"),
        "max_tokens": os.environ.get("AI_MAX_TOKENS", "1200"),
        **defaults[normalized],
    }
    config = _load_ai_config().get("providers", {}).get(normalized, {})
    if isinstance(config, dict):
        for key in ("api_key", "base_url", "model", "chat_completions_path", "timeout_seconds", "temperature", "max_tokens"):
            value = config.get(key)
            if value is not None and str(value).strip():
                settings[key] = str(value).strip()
    return settings


def _load_ai_config() -> dict[str, Any]:
    if not AI_CONFIG_FILE.exists():
        return {}
    try:
        payload = json.loads(AI_CONFIG_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=500, detail=f"Некорректный JSON в {AI_CONFIG_FILE.name}") from exc
    return payload if isinstance(payload, dict) else {}


def _read_ai_prompt() -> str:
    if AI_PROMPT_FILE.exists():
        return AI_PROMPT_FILE.read_text(encoding="utf-8").strip()
    return (
        "Ты редактор B2B-презентаций. Выбери photo_ids и review_ids только из переданных кандидатов. "
        "Верни только валидный JSON."
    )


def _split_form_tags(value: Any) -> list[str]:
    result = []
    seen = set()
    for item in re.split(r"[;,\n\r]+", str(value or "")):
        normalized = item.strip()
        key = normalized.casefold()
        if normalized and key not in seen:
            result.append(normalized)
            seen.add(key)
    return result


def _client_prompt_payload(report: dict[str, Any]) -> dict[str, Any]:
    client_record = report.get("client_record") if isinstance(report.get("client_record"), dict) else {}
    return {
        "company": report.get("client", ""),
        "contact": report.get("contact", ""),
        "position": report.get("position", ""),
        "category": report.get("category", ""),
        "company_type": report.get("company_type") or client_record.get("Тип компании", ""),
        "region": client_record.get("Регион", "") or client_record.get("Город", ""),
        "city": client_record.get("Город", ""),
        "price_category": report.get("price_category") or client_record.get("Ценовая категория", ""),
        "preferred_networks": _split_form_tags(report.get("preferred_networks") or client_record.get("Предпочтительные сети", "")),
    }


def _ai_prompt_payload(report: dict[str, Any], selection: dict[str, Any]) -> dict[str, Any]:
    current_photo_ids, current_review_ids = _current_choice_ids(selection)
    return {
        "client": _client_prompt_payload(report),
        "stats": selection.get("stats"),
        "required_photo_count": selection.get("required_photo_count"),
        "required_review_count": selection.get("required_review_count"),
        "photo_candidates": selection.get("photo_candidates") or [],
        "review_candidates": selection.get("review_candidates") or [],
        "current_selection": {
            "photo_ids": current_photo_ids,
            "review_ids": current_review_ids,
        },
    }


def _request_ai_choice(
    provider: str,
    report: dict[str, Any],
    selection: dict[str, Any],
    api_key_override: str = "",
) -> tuple[dict[str, Any], dict[str, Any]]:
    settings = _ai_provider_settings(provider)
    if api_key_override.strip():
        settings["api_key"] = api_key_override.strip()
    if settings["provider"] == "rules":
        raise HTTPException(status_code=400, detail="Для ответа ИИ выберите OpenAI или DeepSeek")
    if not settings["api_key"]:
        raise HTTPException(status_code=400, detail=f"Не задан ключ для {settings['provider']}")
    body = _ai_request_body(settings, report, selection)

    try:
        timeout = int(float(settings.get("timeout_seconds", "45")))
    except ValueError:
        timeout = 120

    url = settings["base_url"].rstrip("/") + "/" + settings.get("chat_completions_path", "/chat/completions").strip("/")
    request = urllib.request.Request(
        url,
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Authorization": f"Bearer {settings['api_key']}",
            "User-Agent": "AutoSupplierCP",
        },
        method="POST",
    )
    try:
        with DIRECT_HTTP_OPENER.open(request, timeout=timeout) as response:
            raw_text = response.read().decode("utf-8", errors="replace")
            try:
                response_payload = json.loads(raw_text)
            except json.JSONDecodeError as exc:
                raise HTTPException(
                    status_code=502,
                    detail={
                        "message": f"ИИ-провайдер вернул не JSON: {exc}",
                        "rawResponse": {
                            "provider": settings["provider"],
                            "url": url,
                            "status": "invalid_json_response",
                            "rawText": raw_text,
                            "parse_error": str(exc),
                        },
                    },
                ) from exc
    except (TimeoutError, socket.timeout) as exc:
        raise HTTPException(
            status_code=504,
            detail=f"ИИ-провайдер не успел ответить за {timeout} секунд. Увеличьте timeout_seconds в ai_config.json или попробуйте модель быстрее.",
        ) from exc
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise HTTPException(
            status_code=502,
            detail={
                "message": f"ИИ-провайдер вернул ошибку HTTP {exc.code}: {details}",
                "rawResponse": {
                    "provider": settings["provider"],
                    "url": url,
                    "status": "http_error",
                    "http_status": exc.code,
                    "rawText": details,
                },
            },
        ) from exc
    except urllib.error.URLError as exc:
        raise HTTPException(
            status_code=502,
            detail={
                "message": f"ИИ-провайдер недоступен на сетевом уровне: {exc}",
                "rawResponse": {
                    "provider": settings["provider"],
                    "url": url,
                    "status": "network_error",
                    "error": str(exc),
                    "reason": str(getattr(exc, "reason", "")),
                },
            },
        ) from exc
    try:
        content = response_payload["choices"][0]["message"].get("content") or ""
    except Exception as exc:
        raise HTTPException(
            status_code=502,
            detail={
                "message": f"ИИ-провайдер вернул неожиданный формат ответа: {exc}",
                "rawResponse": {
                    "provider": settings["provider"],
                    "url": url,
                    "status": "unexpected_response_shape",
                    "response": response_payload,
                    "parse_error": str(exc),
                },
            },
        ) from exc
    raw_response = {
        "provider": settings["provider"],
        "url": url,
        "status": "ok",
        "content": content,
        "response": response_payload,
    }
    try:
        return _extract_json_object(content), raw_response
    except Exception as exc:
        raw_response["parse_error"] = str(exc)
        raise HTTPException(
            status_code=502,
            detail={
                "message": f"ИИ-провайдер ответил некорректно: {exc}",
                "rawResponse": raw_response,
            },
        ) from exc


def _ai_request_body(settings: dict[str, str], report: dict[str, Any], selection: dict[str, Any]) -> dict[str, Any]:
    try:
        temperature = float(settings.get("temperature", "0.2"))
    except ValueError:
        temperature = 0.2
    try:
        max_tokens = int(float(settings.get("max_tokens", "0")))
    except ValueError:
        max_tokens = 0

    body: dict[str, Any] = {
        "model": settings["model"],
        "messages": [
            {"role": "system", "content": _read_ai_prompt()},
            {"role": "user", "content": json.dumps(_ai_prompt_payload(report, selection), ensure_ascii=False)},
        ],
        "temperature": temperature,
    }
    if max_tokens > 0:
        body["max_tokens"] = max_tokens
    return body


def _ai_request_preview(provider: str, report: dict[str, Any], selection: dict[str, Any]) -> dict[str, Any]:
    settings = _ai_provider_settings(provider)
    if settings["provider"] == "rules":
        return {
            "provider": "rules",
            "message": "Для просмотра JSON запроса выберите OpenAI или DeepSeek.",
            "request": None,
        }
    return {
        "provider": settings["provider"],
        "url": settings["base_url"].rstrip("/") + "/" + settings.get("chat_completions_path", "/chat/completions").strip("/"),
        "timeout_seconds": settings.get("timeout_seconds"),
        "request": _ai_request_body(settings, report, selection),
    }


def _extract_json_object(text: str) -> dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, re.DOTALL)
        if not match:
            raise
        payload = json.loads(match.group(0))
    if not isinstance(payload, dict):
        raise ValueError("AI response must be a JSON object")
    return payload


def _merge_ai_selection(default_selection: dict[str, Any], ai_payload: dict[str, Any]) -> dict[str, Any]:
    photo_ids = ai_payload.get("photo_ids") if isinstance(ai_payload.get("photo_ids"), list) else []
    review_ids = ai_payload.get("review_ids") if isinstance(ai_payload.get("review_ids"), list) else []
    selection = _build_selection_from_choices(default_selection, photo_ids=photo_ids, review_ids=review_ids)
    selection["source"] = "ai"
    rationale = dict(default_selection.get("rationale") or {})
    if isinstance(ai_payload.get("rationale"), dict):
        rationale.update({key: str(value) for key, value in ai_payload["rationale"].items()})
    elif ai_payload.get("summary"):
        rationale["summary"] = str(ai_payload["summary"])
    selection["rationale"] = rationale
    return selection


def _call_ai_selector(provider: str, report: dict[str, Any], default_selection: dict[str, Any]) -> dict[str, Any]:
    settings = _ai_provider_settings(provider)
    if settings["provider"] == "rules":
        return default_selection
    if not settings["api_key"]:
        selection = dict(default_selection)
        selection["source"] = "rules_no_api_key"
        selection["rationale"] = {
            **selection.get("rationale", {}),
            "summary": f"Ключ для {settings['provider']} не задан, использован алгоритмический выбор.",
        }
        return selection
    try:
        ai_payload, raw_response = _request_ai_choice(provider, report, default_selection)
        default_selection["aiRawResponse"] = raw_response
        return _merge_ai_selection(default_selection, ai_payload)
    except Exception as exc:
        selection = dict(default_selection)
        selection["source"] = "rules_ai_error"
        selection["rationale"] = {
            **selection.get("rationale", {}),
            "summary": f"ИИ-провайдер ответил некорректно, использован алгоритмический выбор: {exc}",
        }
        return selection


def _current_choice_ids(selection: dict[str, Any]) -> tuple[list[str], list[str]]:
    return (
        [str(item) for item in (selection.get("selected_photo_ids") or []) if str(item).strip()],
        [str(item) for item in (selection.get("selected_review_ids") or []) if str(item).strip()],
    )


def _normalize_choice_payload(payload: dict[str, Any] | None) -> tuple[list[str], list[str]]:
    body = payload or {}
    photo_ids = [str(item) for item in (body.get("photoIds") or [])]
    review_ids = [str(item) for item in (body.get("reviewIds") or [])]
    return photo_ids, review_ids


def _selection_snapshot(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": item["id"],
        "approved": bool(item.get("approved")),
        "createdAt": item.get("createdAt"),
        "companyName": item.get("companyName", ""),
        "provider": item.get("provider", "rules"),
        "selection": item.get("selection"),
        "finalSelection": item.get("finalSelection"),
        "approvedSelection": item.get("approvedSelection"),
    }


def _call_ai_final_selector(
    provider: str,
    report: dict[str, Any],
    base_selection: dict[str, Any],
    api_key_override: str = "",
) -> dict[str, Any]:
    try:
        ai_payload, raw_response = _request_ai_choice(provider, report, base_selection, api_key_override)
        selection = _merge_ai_selection(base_selection, ai_payload)
        selection["aiRawResponse"] = raw_response
        return selection
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"ИИ-провайдер ответил некорректно: {exc}") from exc


def _store_ai_selection(company: str, provider: str, selection: dict[str, Any], report: dict[str, Any]) -> dict[str, Any]:
    selection_id = uuid.uuid4().hex
    item = {
        "id": selection_id,
        "companyName": company,
        "provider": provider,
        "approved": False,
        "createdAt": datetime.now().isoformat(timespec="seconds"),
        "selection": selection,
        "finalSelection": None,
        "approvedSelection": None,
        "report": report,
    }
    with AI_SELECTIONS_LOCK:
        AI_SELECTIONS[selection_id] = item
    return dict(item)


def _get_ai_selection(selection_id: str, require_approved: bool = False) -> dict[str, Any]:
    with AI_SELECTIONS_LOCK:
        item = AI_SELECTIONS.get(selection_id)
        if not item:
            raise HTTPException(status_code=404, detail="Выбор материалов не найден")
        if require_approved and not item.get("approved"):
            raise HTTPException(status_code=400, detail="Сначала одобрите выбор материалов")
        return dict(item)


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/styles.css")
def styles() -> FileResponse:
    return FileResponse(STATIC_DIR / "styles.css", media_type="text/css")


@app.get("/app.js")
def frontend_script() -> FileResponse:
    return FileResponse(STATIC_DIR / "app.js", media_type="application/javascript")


@app.get("/api/options")
def options() -> dict[str, Any]:
    cards = _read_cards()
    preferred_networks = _preferred_network_names()
    return {
        "companyTypes": _unique([c.get("Тип компании", "") for c in cards], DEFAULT_OPTIONS["companyTypes"]),
        "industries": _unique([c.get("Отрасль", "") for c in cards], DEFAULT_OPTIONS["industries"]),
        "activities": _unique([c.get("Сфера деятельности", "") for c in cards], DEFAULT_OPTIONS["activities"]),
        "productCategories": _unique([c.get("Категория товара", "") for c in cards], DEFAULT_OPTIONS["productCategories"]),
        "networkWorkCategories": DEFAULT_OPTIONS["networkWorkCategories"],
        "preferredNetworkCategories": DEFAULT_OPTIONS["preferredNetworkCategories"],
        "preferredNetworks": preferred_networks,
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
        payload = {"rows": []}
        for attempt in range(5):
            with urllib.request.urlopen(
                urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 AutoSupplierCP"}),
                timeout=12,
            ) as response:
                payload = json.loads(response.read().decode("utf-8"))
            if payload.get("rows") or attempt == 4:
                break
            time.sleep(0.7)
    except Exception as exc:
        raise HTTPException(
            status_code=502,
            detail="Не удалось получить данные из открытой базы ФНС. Можно ввести ИНН вручную.",
        ) from exc

    items = []
    for item in payload.get("rows", [])[:10]:
        items.append(
            {
                "name": _repair_mojibake(item.get("n", "")),
                "inn": item.get("i", ""),
                "address": _repair_mojibake(item.get("a", "")),
                "status": _repair_mojibake(item.get("s", "")),
            }
        )
    return {"items": items}


@app.get("/api/site-inn-search")
def site_inn_search(website: str, selectedInn: str = "") -> dict[str, Any]:
    items = _find_site_inns(website)
    selected = _digits(selectedInn)
    return {
        "items": items,
        "selectedInn": selected,
        "matchedSelectedInn": bool(selected and any(item["inn"] == selected for item in items)),
    }


@app.post("/api/dadata-company")
def dadata_company(payload: dict[str, Any]) -> dict[str, Any]:
    inn = _digits(str(payload.get("inn", "")))
    details = _fetch_dadata_company(inn)
    row_number_raw = str(payload.get("rowNumber", "")).strip()
    if row_number_raw.isdigit():
        _update_card_dadata(int(row_number_raw), details)
    return {"items": details}


@app.post("/api/ai-selection")
def prepare_ai_selection(payload: dict[str, Any]) -> dict[str, Any]:
    company = str(payload.get("companyName", "")).strip()
    provider = str(payload.get("provider", "rules")).strip().lower() or "rules"
    draft_payload = payload.get("draftPayload") if isinstance(payload.get("draftPayload"), dict) else {}
    if not company:
        raise HTTPException(status_code=400, detail="Выберите или заполните компанию")
    if provider not in {"rules", "openai", "deepseek"}:
        raise HTTPException(status_code=400, detail="Неизвестный ИИ-провайдер")

    report = _run_presentation_dry_run_with_override(company, draft_payload)
    default_selection = _default_selection_from_report(report, provider)
    stored = _store_ai_selection(company, provider, default_selection, report)
    return _selection_snapshot(stored)


@app.post("/api/ai-selection/{selection_id}/approve")
def approve_ai_selection(selection_id: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
    photo_ids, review_ids = _normalize_choice_payload(payload)
    mode = str((payload or {}).get("mode", "manual")).strip().lower()
    with AI_SELECTIONS_LOCK:
        item = AI_SELECTIONS.get(selection_id)
        if not item:
            raise HTTPException(status_code=404, detail="Выбор материалов не найден")
        if mode == "ai":
            if not item.get("finalSelection"):
                raise HTTPException(status_code=400, detail="Сначала получите окончательный выбор ИИ")
            approved_selection = dict(item["finalSelection"])
        else:
            approved_selection = _build_selection_from_choices(
                item.get("selection") or {},
                photo_ids=photo_ids,
                review_ids=review_ids,
            )
        item["selection"] = approved_selection
        item["approvedSelection"] = approved_selection
        item["approved"] = True
        item["approvedAt"] = datetime.now().isoformat(timespec="seconds")
        item["updatedAt"] = item["approvedAt"]
        return {
            "ok": True,
            "id": selection_id,
            "approved": True,
            "approvedAt": item["approvedAt"],
            "selection": item["selection"],
            "finalSelection": item.get("finalSelection"),
        }


@app.post("/api/ai-selection/{selection_id}/finalize")
def finalize_ai_selection(selection_id: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
    provider = str((payload or {}).get("provider", "")).strip().lower()
    api_key_override = str((payload or {}).get("apiKey", "")).strip()
    photo_ids, review_ids = _normalize_choice_payload(payload)
    with AI_SELECTIONS_LOCK:
        item = AI_SELECTIONS.get(selection_id)
        if not item:
            raise HTTPException(status_code=404, detail="Выбор материалов не найден")
        base_selection = _build_selection_from_choices(
            item.get("selection") or {},
            photo_ids=photo_ids,
            review_ids=review_ids,
        )
        chosen_provider = provider or str(item.get("provider", "rules")).strip().lower() or "rules"
        report = dict(item.get("report") or {})
    final_selection = _call_ai_final_selector(chosen_provider, report, base_selection, api_key_override)
    with AI_SELECTIONS_LOCK:
        item = AI_SELECTIONS.get(selection_id)
        if not item:
            raise HTTPException(status_code=404, detail="Выбор материалов не найден")
        item["provider"] = chosen_provider
        item["selection"] = base_selection
        item["finalSelection"] = final_selection
        item["updatedAt"] = datetime.now().isoformat(timespec="seconds")
        return _selection_snapshot(item)


@app.post("/api/ai-selection/{selection_id}/request-preview")
def ai_selection_request_preview(selection_id: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
    provider = str((payload or {}).get("provider", "")).strip().lower()
    photo_ids, review_ids = _normalize_choice_payload(payload)
    with AI_SELECTIONS_LOCK:
        item = AI_SELECTIONS.get(selection_id)
        if not item:
            raise HTTPException(status_code=404, detail="Выбор материалов не найден")
        base_selection = _build_selection_from_choices(
            item.get("selection") or {},
            photo_ids=photo_ids,
            review_ids=review_ids,
        )
        chosen_provider = provider or str(item.get("provider", "rules")).strip().lower() or "rules"
        report = dict(item.get("report") or {})
    return _ai_request_preview(chosen_provider, report, base_selection)


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


@app.patch("/api/cards/{row_number}/inn")
def update_card_inn(row_number: int, payload: dict[str, Any]) -> dict[str, Any]:
    inn = _digits(str(payload.get("inn", "")))
    _update_card_inn(row_number, inn)
    return {"ok": True, "row": row_number, "inn": inn}


@app.post("/api/presentations")
def build_presentation(payload: dict[str, Any]) -> dict[str, Any]:
    company = str(payload.get("companyName", "")).strip()
    selection_id = str(payload.get("approvedSelectionId", "")).strip()
    if not company:
        raise HTTPException(status_code=400, detail="Выберите или заполните компанию")
    if not selection_id:
        raise HTTPException(status_code=400, detail="Сначала подготовьте и одобрите выбор материалов ИИ")
    selection_item = _get_ai_selection(selection_id, require_approved=True)
    if str(selection_item.get("companyName", "")).strip() != company:
        raise HTTPException(status_code=400, detail="Одобренный выбор относится к другой компании")

    job_id = uuid.uuid4().hex
    _set_presentation_job(
        job_id,
        id=job_id,
        companyName=company,
        selectionId=selection_id,
        selectionProvider=selection_item.get("provider", ""),
        source="manual",
        createdAt=datetime.now().isoformat(timespec="seconds"),
        status="queued",
        progress=5,
        message="Задача поставлена в очередь",
    )
    worker = threading.Thread(
        target=_run_presentation_job,
        args=(job_id, company, selection_item.get("selection") or {}),
        kwargs={"make_pdf": True},
        daemon=True,
    )
    worker.start()
    return _get_presentation_job(job_id)


@app.get("/api/presentations/jobs/{job_id}")
def presentation_job(job_id: str) -> dict[str, Any]:
    return _get_presentation_job(job_id)


@app.get("/api/presentations/history")
def presentation_history(limit: int = 100) -> dict[str, Any]:
    limit = max(1, min(500, limit))
    return {"items": _read_presentation_audit(limit)}


@app.get("/api/presentations/{filename}")
def download_presentation(filename: str) -> FileResponse:
    target = _safe_root_file(filename)
    suffix = target.suffix.lower()
    if not target.exists() or suffix not in {".pptx", ".pdf"}:
        raise HTTPException(status_code=404, detail="Презентация не найдена")
    media_type = (
        "application/pdf"
        if suffix == ".pdf"
        else "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    return FileResponse(
        target,
        media_type=media_type,
        filename=target.name,
    )


@app.get("/photo/{filename}")
def photo(filename: str) -> FileResponse:
    target = PHOTO_DIR / filename
    if not target.exists():
        raise HTTPException(status_code=404, detail="Фото не найдено")
    media_type = mimetypes.guess_type(target.name)[0]
    return FileResponse(target, media_type=media_type)


@app.get("/api/image-preview")
def image_preview(path: str) -> FileResponse:
    target, actual_ext = _resolve_preview_image(path)
    if not target.exists() or not target.is_file():
        raise HTTPException(status_code=404, detail="Изображение не найдено")
    media_types_by_ext = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".gif": "image/gif",
        ".webp": "image/webp",
        ".svg": "image/svg+xml",
    }
    media_type = media_types_by_ext.get(actual_ext or "") or mimetypes.guess_type(target.name)[0]
    return FileResponse(target, media_type=media_type)


# ── Битрикс24 интеграция ────────────────────────────────────────────────

def _load_bitrix_config() -> dict[str, Any]:
    cfg: dict[str, Any] = {}
    if BITRIX_CONFIG_FILE.exists():
        cfg = json.loads(BITRIX_CONFIG_FILE.read_text(encoding="utf-8"))

    webhook_url = os.environ.get("BITRIX_WEBHOOK_URL", "").strip() or str(cfg.get("webhook_url", "")).strip()
    if not webhook_url or "ВАШ-ДОМЕН" in webhook_url or "WEBHOOK_TOKEN" in webhook_url or "ТОКЕН" in webhook_url:
        raise RuntimeError(
            "BITRIX_WEBHOOK_URL не настроен. "
            "На Render добавьте переменную окружения с URL входящего вебхука Битрикс24."
        )

    legacy_field_map = {k: v for k, v in (cfg.get("field_map") or {}).items() if v}
    deal_field_map = dict(BITRIX_DEAL_FIELD_MAP)
    deal_field_map.update(legacy_field_map)
    deal_field_map.update({k: v for k, v in (cfg.get("deal_field_map") or {}).items() if v})

    lead_field_map = dict(BITRIX_LEAD_FIELD_MAP)
    lead_field_map.update({k: v for k, v in (cfg.get("lead_field_map") or {}).items() if v})
    return {"webhook_url": webhook_url, "deal_field_map": deal_field_map, "lead_field_map": lead_field_map}


def _verify_bitrix_inbound_token(request: Request) -> None:
    expected = os.environ.get("BITRIX_INBOUND_TOKEN", "").strip()
    if not expected:
        return
    provided = (
        request.query_params.get("token")
        or request.headers.get("X-Bitrix-Token")
        or request.headers.get("X-AutoSupplierCP-Token")
        or ""
    ).strip()
    if not secrets.compare_digest(provided, expected):
        raise HTTPException(status_code=403, detail="Неверный токен входящего webhook")


def _bitrix_call(webhook_url: str, method: str, params: dict[str, Any] | None = None) -> Any:
    url = f"{webhook_url.rstrip('/')}/{method}.json"
    body = json.dumps(params or {}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with DIRECT_HTTP_OPENER.open(req, timeout=20) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    if "error" in data:
        raise RuntimeError(
            f"Битрикс24 API: {data['error']} — {data.get('error_description', '')}"
        )
    return data.get("result")


def _bitrix_field_value_to_text(fields_meta: dict[str, Any], field_id: str, raw: Any) -> str:
    if raw is None:
        return ""

    meta = fields_meta.get(field_id) or {}
    enum_values = {
        str(item.get("ID", "")): str(item.get("VALUE", ""))
        for item in meta.get("items", []) or []
    }

    def scalar_to_text(value: Any) -> str:
        if value is None:
            return ""
        if enum_values:
            mapped = enum_values.get(str(value))
            if mapped:
                return mapped
        if isinstance(value, dict):
            for key in ("VALUE", "value", "TEXT", "text", "ADDRESS", "address"):
                normalized = str(value.get(key) or "").strip()
                if normalized:
                    return normalized
            return "; ".join(str(v).strip() for v in value.values() if str(v or "").strip())
        return str(value).strip()

    if isinstance(raw, list):
        return "; ".join(text for value in raw if (text := scalar_to_text(value)))
    return scalar_to_text(raw)


def _join_unique_text_values(*values: Any) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        for item in re.split(r"[;,\n\r]+", str(value or "")):
            normalized = item.strip()
            if not normalized:
                continue
            key = normalized.lower()
            if key in seen:
                continue
            seen.add(key)
            result.append(normalized)
    return "; ".join(result)


def _bitrix_get_deal_data(deal_id: str, webhook_url: str, field_map: dict[str, str]) -> dict[str, Any]:
    deal = _bitrix_call(webhook_url, "crm.deal.get", {"id": deal_id})
    if not deal:
        raise RuntimeError(f"Сделка {deal_id} не найдена в Битрикс24")
    fields_meta = _bitrix_call(webhook_url, "crm.deal.fields", {}) or {}

    card: dict[str, Any] = {}

    company_id = str(deal.get("COMPANY_ID") or "").strip()
    if company_id and company_id != "0":
        company = _bitrix_call(webhook_url, "crm.company.get", {"id": company_id}) or {}
        card["companyName"] = str(company.get("TITLE") or "").strip()
        webs = company.get("WEB") or []
        if webs and not card.get("website"):
            card["website"] = str(webs[0].get("VALUE", "")).strip()

    company_field = field_map.get("companyName", "")
    if not card.get("companyName") and company_field:
        card["companyName"] = _bitrix_field_value_to_text(fields_meta, company_field, deal.get(company_field))
    if not card.get("companyName"):
        card["companyName"] = str(deal.get("TITLE") or "").strip()

    contact_id = str(deal.get("CONTACT_ID") or "").strip()
    if contact_id and contact_id != "0":
        contact = _bitrix_call(webhook_url, "crm.contact.get", {"id": contact_id}) or {}
        name_parts = [contact.get("NAME", ""), contact.get("LAST_NAME", "")]
        card["contactName"] = " ".join(p for p in name_parts if p).strip()
        card["contactPosition"] = str(contact.get("WORK_POSITION") or "").strip()

    for app_field, bitrix_field in field_map.items():
        if app_field in ("companyName", "contactName", "contactPosition", "presentationFileField"):
            continue
        if not bitrix_field or bitrix_field == "UF_CRM_XXXX":
            continue
        raw = deal.get(bitrix_field)
        if raw is None:
            continue
        value = _bitrix_field_value_to_text(fields_meta, bitrix_field, raw)
        if value:
            card[app_field] = value

    combined_product_category = _join_unique_text_values(
        card.get("productCompanyCategories", ""),
        card.get("productNegotiationCategories", ""),
        card.get("productCategory", ""),
    )
    if combined_product_category:
        card["productCategory"] = combined_product_category

    return card


def _bitrix_get_lead_data(lead_id: str, webhook_url: str, field_map: dict[str, str]) -> dict[str, Any]:
    lead = _bitrix_call(webhook_url, "crm.lead.get", {"id": lead_id})
    if not lead:
        raise RuntimeError(f"Lead {lead_id} not found in Bitrix24")
    fields_meta = _bitrix_call(webhook_url, "crm.lead.fields", {}) or {}

    card: dict[str, Any] = {}

    for app_field, bitrix_field in field_map.items():
        if app_field in ("presentationFileField", "contactLastName", "leadContactName"):
            continue
        if not bitrix_field or bitrix_field == "UF_CRM_XXXX":
            continue
        raw = lead.get(bitrix_field)
        if raw is None:
            continue
        value = _bitrix_field_value_to_text(fields_meta, bitrix_field, raw)
        if value:
            card[app_field] = value

    if not card.get("companyName"):
        card["companyName"] = str(
            lead.get("COMPANY_TITLE")
            or lead.get("TITLE")
            or lead.get("NAME")
            or ""
        ).strip()

    name_field = field_map.get("contactName", "NAME")
    last_name_field = field_map.get("contactLastName", "LAST_NAME")
    name = _bitrix_field_value_to_text(fields_meta, name_field, lead.get(name_field))
    last_name = _bitrix_field_value_to_text(fields_meta, last_name_field, lead.get(last_name_field))
    contact_name = " ".join(part for part in (name, last_name) if part).strip()
    lead_contact_field = field_map.get("leadContactName", "")
    if not contact_name and lead_contact_field:
        contact_name = _bitrix_field_value_to_text(fields_meta, lead_contact_field, lead.get(lead_contact_field))
    if contact_name:
        card["contactName"] = contact_name

    combined_product_category = _join_unique_text_values(
        card.get("productCompanyCategories", ""),
        card.get("productNegotiationCategories", ""),
        card.get("productCategory", ""),
    )
    if combined_product_category:
        card["productCategory"] = combined_product_category

    return card


def _bitrix_get_card_data(entity_type: str, entity_id: str, webhook_url: str, field_map: dict[str, str]) -> dict[str, Any]:
    if entity_type == "deal":
        return _bitrix_get_deal_data(entity_id, webhook_url, field_map)
    if entity_type == "lead":
        return _bitrix_get_lead_data(entity_id, webhook_url, field_map)
    raise RuntimeError(f"Unknown Bitrix24 entity type: {entity_type}")


def _bitrix_upload_file(webhook_url: str, entity_type: str, entity_id: str, file_path: Path, file_field: str) -> None:
    if not file_field or file_field == "UF_CRM_XXXX":
        raise RuntimeError("presentationFileField не настроен в bitrix_config.json")
    if entity_type not in {"deal", "lead"}:
        raise RuntimeError(f"Unknown Bitrix24 entity type: {entity_type}")
    import base64
    content_b64 = base64.b64encode(file_path.read_bytes()).decode("ascii")
    _bitrix_call(
        webhook_url,
        f"crm.{entity_type}.update",
        {
            "id": entity_id,
            "fields": {file_field: {"fileData": [file_path.name, content_b64]}},
        },
    )


def _bitrix_move_deal_stage(webhook_url: str, deal_id: str, stage_id: str) -> None:
    if not stage_id:
        return
    _bitrix_call(
        webhook_url,
        "crm.deal.update",
        {
            "id": deal_id,
            "fields": {"STAGE_ID": stage_id},
        },
    )


def _default_ai_provider() -> str:
    for provider in ("deepseek", "openai"):
        settings = _ai_provider_settings(provider)
        if settings.get("api_key"):
            return provider
    return "rules"


def _storage_mode() -> str:
    if APP_STORAGE_MODE in {"excel", "memory"}:
        return APP_STORAGE_MODE
    if os.environ.get("RENDER", "").strip().lower() in {"1", "true", "yes"}:
        return "memory"
    return "excel"


def _set_bitrix_job(job_id: str, **values: Any) -> None:
    with BITRIX_JOBS_LOCK:
        job = BITRIX_JOBS.setdefault(job_id, {})
        job.update(values)
        job["updatedAt"] = datetime.now().isoformat(timespec="seconds")


def _run_bitrix_pipeline(job_id: str, entity_type: str, entity_id: str) -> None:
    try:
        _set_bitrix_job(job_id, status="running", progress=5, message="Загружаю данные карточки из Битрикс24")

        cfg = _load_bitrix_config()
        webhook_url: str = cfg["webhook_url"]
        field_map: dict[str, str] = cfg.get(f"{entity_type}_field_map", {})
        file_field: str = field_map.get("presentationFileField", "")

        card_data = _bitrix_get_card_data(entity_type, entity_id, webhook_url, field_map)
        company = str(card_data.get("companyName") or "").strip()
        if not company:
            raise RuntimeError("Не удалось получить название компании из карточки Битрикс24")

        storage_mode = _storage_mode()
        _set_bitrix_job(
            job_id,
            progress=15,
            message=f"Сохраняю карточку: {company}" if storage_mode == "excel" else f"Данные карточки получены: {company}",
            storageMode=storage_mode,
            cardData=card_data,
        )

        for fallback_field, fallback_value in (("companyType", "Производитель"),):
            if not card_data.get(fallback_field):
                card_data[fallback_field] = fallback_value

        if storage_mode == "excel":
            try:
                _save_card_row(card_data, None, None)
            except Exception as exc:
                print(f"[bitrix] предупреждение: не удалось сохранить карточку: {exc}", flush=True)

        _set_bitrix_job(job_id, progress=25, message="Подготавливаю материалы для ИИ")

        try:
            report = _run_presentation_dry_run_with_override(company, card_data)
        except HTTPException as exc:
            raise RuntimeError(str(exc.detail)) from exc

        ai_provider = _default_ai_provider()
        default_selection = _default_selection_from_report(report, ai_provider)

        _set_bitrix_job(job_id, progress=45, message=f"Запрашиваю ИИ ({ai_provider})")

        try:
            final_selection = _call_ai_final_selector(ai_provider, report, default_selection)
        except HTTPException as exc:
            print(f"[bitrix] ИИ недоступен, использую правила: {exc.detail}", flush=True)
            final_selection = default_selection

        _set_bitrix_job(job_id, progress=60, message="Собираю презентацию")

        pptx_job_id = uuid.uuid4().hex
        _set_presentation_job(
            pptx_job_id,
            id=pptx_job_id,
            companyName=company,
            source=f"bitrix_{entity_type}",
            sourceEntityId=entity_id,
            createdAt=datetime.now().isoformat(timespec="seconds"),
            status="queued",
            progress=5,
            message="Задача от Битрикс24",
        )
        _run_presentation_job(
            pptx_job_id,
            company,
            final_selection,
            make_pdf=True,
            require_pdf=True,
            payload_override=card_data,
            source=f"bitrix_{entity_type}",
            source_entity_id=entity_id,
        )
        pptx_job = _get_presentation_job(pptx_job_id)

        if pptx_job.get("status") != "done":
            raise RuntimeError(
                pptx_job.get("error")
                or pptx_job.get("message")
                or "Презентация не была собрана"
            )

        pdf_file_name = str(pptx_job.get("pdfFileName") or "").strip()
        if not pdf_file_name:
            raise RuntimeError("PDF presentation was not created")
        pdf_path = ROOT / pdf_file_name

        _set_bitrix_job(job_id, progress=90, message="Загружаю презентацию в Битрикс24")

        _bitrix_upload_file(webhook_url, entity_type, entity_id, pdf_path, file_field)
        _set_bitrix_job(job_id, progress=95, message="Finishing Bitrix24 update")
        if entity_type == "deal":
            _bitrix_move_deal_stage(webhook_url, entity_id, BITRIX_AFTER_PRESENTATION_STAGE_ID)

        _set_bitrix_job(
            job_id,
            status="done",
            progress=100,
            message="Презентация готова и загружена в Битрикс24",
            fileName=pptx_job["fileName"],
            downloadUrl=pptx_job.get("downloadUrl"),
            pdfFileName=pdf_file_name,
            pdfDownloadUrl=pptx_job.get("pdfDownloadUrl"),
            selectedPhotos=_selected_photo_summary(final_selection),
        )

    except Exception as exc:
        _set_bitrix_job(job_id, status="error", progress=0, message=str(exc), error=str(exc))
        print(f"[bitrix] ошибка pipeline: {exc}", flush=True)


def _extract_bitrix_deal_id(payload: Any) -> str:
    if not isinstance(payload, dict):
        return ""

    def normalize(value: Any) -> str:
        if isinstance(value, list):
            for item in value:
                normalized = normalize(item)
                if normalized:
                    return normalized
            return ""
        if isinstance(value, dict):
            return _extract_bitrix_deal_id(value)
        raw = str(value or "").strip()
        if not raw:
            return ""
        match = re.search(r"(?:DEAL_|D_)?(\d+)$", raw, re.IGNORECASE)
        return match.group(1) if match else raw

    for key in ("deal_id", "DEAL_ID", "id", "ID", "document_id", "DOCUMENT_ID"):
        normalized = normalize(payload.get(key))
        if normalized:
            return normalized

    for key in ("data", "DATA", "FIELDS", "fields"):
        normalized = normalize(payload.get(key))
        if normalized:
            return normalized

    for key, value in payload.items():
        if str(key).lower() in {
            "data[fields][id]",
            "data[fields][deal_id]",
            "document_id[2]",
            "document_id",
        }:
            normalized = normalize(value)
            if normalized:
                return normalized
    return ""


def _extract_bitrix_lead_id(payload: Any) -> str:
    if not isinstance(payload, dict):
        return ""

    def normalize(value: Any) -> str:
        if isinstance(value, list):
            for item in value:
                normalized = normalize(item)
                if normalized:
                    return normalized
            return ""
        if isinstance(value, dict):
            return _extract_bitrix_lead_id(value)
        raw = str(value or "").strip()
        if not raw:
            return ""
        match = re.search(r"(?:LEAD_|L_)?(\d+)$", raw, re.IGNORECASE)
        return match.group(1) if match else raw

    for key in ("lead_id", "LEAD_ID", "id", "ID", "document_id", "DOCUMENT_ID"):
        normalized = normalize(payload.get(key))
        if normalized:
            return normalized

    for key in ("data", "DATA", "FIELDS", "fields"):
        normalized = normalize(payload.get(key))
        if normalized:
            return normalized

    for key, value in payload.items():
        if str(key).lower() in {
            "data[fields][id]",
            "data[fields][lead_id]",
            "document_id[2]",
            "document_id",
        }:
            normalized = normalize(value)
            if normalized:
                return normalized
    return ""


async def _extract_bitrix_request_entity_id(request: Request, entity_type: str) -> str:
    extractor = _extract_bitrix_lead_id if entity_type == "lead" else _extract_bitrix_deal_id
    entity_id = extractor(dict(request.query_params))
    content_type = request.headers.get("content-type", "")
    if entity_id:
        return entity_id
    if "application/json" in content_type:
        try:
            body = await request.json()
        except Exception:
            body = {}
        return extractor(body)
    form = await request.form()
    return extractor(dict(form))


async def _bitrix_webhook_for_entity(request: Request, entity_type: str) -> dict[str, Any]:
    _verify_bitrix_inbound_token(request)

    entity_id = await _extract_bitrix_request_entity_id(request, entity_type)
    if not entity_id:
        raise HTTPException(status_code=400, detail=f"{entity_type}_id not provided")

    job_id = uuid.uuid4().hex
    _set_bitrix_job(
        job_id,
        id=job_id,
        entity_type=entity_type,
        entity_id=entity_id,
        status="queued",
        progress=0,
        message="Р—Р°РґР°С‡Р° РїРѕСЃС‚Р°РІР»РµРЅР° РІ РѕС‡РµСЂРµРґСЊ",
    )
    threading.Thread(target=_run_bitrix_pipeline, args=(job_id, entity_type, entity_id), daemon=True).start()
    return {"ok": True, "job_id": job_id, "entity_type": entity_type, "entity_id": entity_id}


@app.post("/api/bitrix/deal/webhook")
async def bitrix_deal_webhook(request: Request) -> dict[str, Any]:
    return await _bitrix_webhook_for_entity(request, "deal")


@app.post("/api/bitrix/lead/webhook")
async def bitrix_lead_webhook(request: Request) -> dict[str, Any]:
    return await _bitrix_webhook_for_entity(request, "lead")


@app.post("/api/bitrix/webhook")
async def bitrix_webhook(request: Request) -> dict[str, Any]:
    _verify_bitrix_inbound_token(request)

    deal_id = _extract_bitrix_deal_id(dict(request.query_params))
    content_type = request.headers.get("content-type", "")
    if deal_id:
        pass
    elif "application/json" in content_type:
        try:
            body = await request.json()
        except Exception:
            body = {}
        deal_id = _extract_bitrix_deal_id(body)
    else:
        form = await request.form()
        deal_id = _extract_bitrix_deal_id(dict(form))

    if not deal_id:
        raise HTTPException(status_code=400, detail="deal_id не передан")

    job_id = uuid.uuid4().hex
    _set_bitrix_job(
        job_id,
        id=job_id,
        deal_id=deal_id,
        status="queued",
        progress=0,
        message="Задача поставлена в очередь",
    )
    threading.Thread(target=_run_bitrix_pipeline, args=(job_id, "deal", deal_id), daemon=True).start()
    return {"ok": True, "job_id": job_id}


@app.get("/api/bitrix/jobs/{job_id}")
def bitrix_job_status(job_id: str, request: Request) -> dict[str, Any]:
    _verify_bitrix_inbound_token(request)

    with BITRIX_JOBS_LOCK:
        job = BITRIX_JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Задача не найдена")
    return dict(job)
