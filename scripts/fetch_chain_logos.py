from __future__ import annotations

import contextlib
import html
import mimetypes
import re
import shutil
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Список сетей.xlsx"
LOGO_DIR = ROOT / "Логотипы сетей"
BACKUP = ROOT / "Список сетей.before_logos.xlsx"
SHEET_NAME = "Shops"
LOGO_HEADER = "Логотип"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36"
)
SSL_CONTEXT = ssl._create_unverified_context()


RU = {
    "а": "a",
    "б": "b",
    "в": "v",
    "г": "g",
    "д": "d",
    "е": "e",
    "ё": "e",
    "ж": "zh",
    "з": "z",
    "и": "i",
    "й": "y",
    "к": "k",
    "л": "l",
    "м": "m",
    "н": "n",
    "о": "o",
    "п": "p",
    "р": "r",
    "с": "s",
    "т": "t",
    "у": "u",
    "ф": "f",
    "х": "h",
    "ц": "c",
    "ч": "ch",
    "ш": "sh",
    "щ": "sch",
    "ъ": "",
    "ы": "y",
    "ь": "",
    "э": "e",
    "ю": "yu",
    "я": "ya",
}


@dataclass
class Shop:
    row: int
    company: str
    site: str


@dataclass
class Candidate:
    url: str
    score: int
    source: str


def q(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch.upper()) - 64
    return n


def cell_col(ref: str) -> str:
    return re.match(r"[A-Z]+", ref).group(0)


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    kind = cell.attrib.get("t")
    value = cell.find(q("v"))
    if kind == "s" and value is not None and value.text is not None:
        return shared_strings[int(value.text)]
    if kind == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(f".//{q('t')}"))
    return "" if value is None or value.text is None else value.text


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings = []
    for si in root.findall(q("si")):
        strings.append("".join(t.text or "" for t in si.iter(q("t"))))
    return strings


def sheet_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    for sheet in wb.find(q("sheets")):
        if sheet.attrib["name"] == sheet_name:
            rid = sheet.attrib[f"{{{NS_REL}}}id"]
            target = targets[rid].lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    raise RuntimeError(f"Sheet not found: {sheet_name}")


def load_shops() -> tuple[list[Shop], dict[int, str]]:
    with zipfile.ZipFile(WORKBOOK) as zf:
        shared_strings = read_shared_strings(zf)
        path = sheet_path(zf, SHEET_NAME)
        root = ET.fromstring(zf.read(path))
    rows = root.find(q("sheetData")).findall(q("row"))
    header = {}
    for cell in rows[0].findall(q("c")):
        header[cell_text(cell, shared_strings).strip()] = cell_col(cell.attrib["r"])
    company_col = header["Компания"]
    site_col = header["Сайт"]
    logo_by_row = {}
    if LOGO_HEADER in header:
        logo_col = header[LOGO_HEADER]
        for row in rows[1:]:
            for cell in row.findall(q("c")):
                if cell_col(cell.attrib["r"]) == logo_col:
                    logo_by_row[int(row.attrib["r"])] = cell_text(cell, shared_strings).strip()
    shops = []
    for row in rows[1:]:
        row_num = int(row.attrib["r"])
        values = {cell_col(c.attrib["r"]): cell_text(c, shared_strings).strip() for c in row.findall(q("c"))}
        company = values.get(company_col, "").strip()
        site = values.get(site_col, "").strip()
        if company and site:
            shops.append(Shop(row_num, company, site))
    return shops, logo_by_row


def slugify(text: str) -> str:
    text = text.split(",")[0].strip().lower()
    text = "".join(RU.get(ch, ch) for ch in text)
    text = re.sub(r"[^a-z0-9]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-")
    return text[:60] or "logo"


def request(url: str, timeout: int = 12) -> urllib.response.addinfourl:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    return urllib.request.urlopen(req, timeout=timeout, context=SSL_CONTEXT)


def fetch_html(site: str) -> tuple[str, str] | None:
    site = site.strip()
    if not site or site == "-":
        return None
    if not re.match(r"https?://", site, re.I):
        urls = [f"https://{site}", f"http://{site}"]
    else:
        urls = [site]
    for url in urls:
        try:
            with request(url) as resp:
                raw = resp.read(1_500_000)
                final_url = resp.geturl()
                charset = resp.headers.get_content_charset() or "utf-8"
                return raw.decode(charset, errors="replace"), final_url
        except Exception:
            continue
    return None


def attr(tag: str, name: str) -> str:
    m = re.search(rf"""\b{name}\s*=\s*(['"])(.*?)\1""", tag, re.I | re.S)
    return html.unescape(m.group(2).strip()) if m else ""


def candidates_from_html(page: str, base_url: str) -> list[Candidate]:
    out: list[Candidate] = []
    seen = set()

    def add(raw: str, score: int, source: str) -> None:
        if not raw or raw.startswith(("data:", "javascript:", "mailto:", "#")):
            return
        url = urllib.parse.urljoin(base_url, raw)
        key = url.split("#", 1)[0]
        if key in seen:
            return
        seen.add(key)
        lower = key.lower()
        if "logo" in lower:
            score += 50
        if lower.endswith(".svg"):
            score += 12
        if any(lower.endswith(ext) for ext in (".png", ".webp", ".jpg", ".jpeg")):
            score += 8
        out.append(Candidate(key, score, source))

    for tag in re.findall(r"<link\b[^>]*>", page, re.I | re.S):
        rel = attr(tag, "rel").lower()
        if "icon" in rel or "apple-touch" in rel:
            add(attr(tag, "href"), 20, "icon")

    for tag in re.findall(r"<meta\b[^>]*>", page, re.I | re.S):
        prop = (attr(tag, "property") or attr(tag, "name")).lower()
        if prop in {"og:image", "twitter:image", "twitter:image:src"}:
            add(attr(tag, "content"), 30, prop)

    for tag in re.findall(r"<img\b[^>]*>", page, re.I | re.S):
        marker = " ".join([attr(tag, "src"), attr(tag, "alt"), attr(tag, "class"), attr(tag, "id")]).lower()
        src = attr(tag, "src") or attr(tag, "data-src") or attr(tag, "data-lazy-src")
        if "logo" in marker or "brand" in marker:
            add(src, 80, "img-logo")
        elif src and any(x in src.lower() for x in ("logo", "logotype")):
            add(src, 70, "img-src")

    return sorted(out, key=lambda c: c.score, reverse=True)


def extension(url: str, content_type: str) -> str:
    suffix = Path(urllib.parse.urlparse(url).path).suffix.lower()
    if suffix in {".svg", ".png", ".jpg", ".jpeg", ".webp", ".gif", ".ico"}:
        return ".jpg" if suffix == ".jpeg" else suffix
    if "svg" in content_type:
        return ".svg"
    guessed = mimetypes.guess_extension(content_type.split(";", 1)[0].strip())
    return ".jpg" if guessed == ".jpe" else (guessed or ".png")


def download(candidate: Candidate, company: str) -> str | None:
    try:
        with request(candidate.url) as resp:
            data = resp.read(4_000_000)
            ctype = resp.headers.get("Content-Type", "").lower()
    except Exception:
        return None
    if len(data) < 200:
        return None
    path_ext = Path(urllib.parse.urlparse(candidate.url).path).suffix.lower()
    if not (ctype.startswith("image/") or path_ext == ".svg" or data[:5].lower().startswith(b"<svg")):
        return None
    ext = extension(candidate.url, ctype)
    name = f"{slugify(company)}{ext}"
    path = LOGO_DIR / name
    idx = 2
    while path.exists() and path.read_bytes() != data:
        name = f"{slugify(company)}-{idx}{ext}"
        path = LOGO_DIR / name
        idx += 1
    path.write_bytes(data)
    return name


def find_logo(shop: Shop) -> str | None:
    sites = [part.strip() for part in re.split(r"[;\n]+", shop.site) if part.strip()]
    for site in sites:
        page = fetch_html(site)
        if not page:
            continue
        html_text, final_url = page
        candidates = candidates_from_html(html_text, final_url)
        candidates.append(Candidate(urllib.parse.urljoin(final_url, "/favicon.ico"), 5, "favicon-fallback"))
        for candidate in candidates[:14]:
            logo = download(candidate, shop.company)
            if logo:
                return logo
    return None


def set_inline_text(row: ET.Element, ref: str, text: str) -> None:
    for cell in row.findall(q("c")):
        if cell.attrib.get("r") == ref:
            row.remove(cell)
            break
    cell = ET.Element(q("c"), {"r": ref, "t": "inlineStr"})
    inline = ET.SubElement(cell, q("is"))
    t = ET.SubElement(inline, q("t"))
    t.text = text
    inserted = False
    for idx, existing in enumerate(row.findall(q("c"))):
        if col_to_num(cell_col(existing.attrib["r"])) > col_to_num(cell_col(ref)):
            row.insert(idx, cell)
            inserted = True
            break
    if not inserted:
        row.append(cell)


def update_workbook(logos: dict[int, str]) -> None:
    if not BACKUP.exists():
        shutil.copy2(WORKBOOK, BACKUP)
    temp = WORKBOOK.with_suffix(".logos.tmp.xlsx")
    with zipfile.ZipFile(WORKBOOK, "r") as zin, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
        sheet = sheet_path(zin, SHEET_NAME)
        root = ET.fromstring(zin.read(sheet))
        sheet_data = root.find(q("sheetData"))
        rows = sheet_data.findall(q("row"))
        header_cells = rows[0].findall(q("c"))
        logo_col = None
        with zipfile.ZipFile(WORKBOOK) as zread:
            shared_strings = read_shared_strings(zread)
        for cell in header_cells:
            if cell_text(cell, shared_strings).strip() == LOGO_HEADER:
                logo_col = cell_col(cell.attrib["r"])
                break
        if logo_col is None:
            max_col_num = max(col_to_num(cell_col(cell.attrib["r"])) for cell in header_cells)
            logo_col = chr(ord("A") + max_col_num)
            set_inline_text(rows[0], f"{logo_col}1", LOGO_HEADER)
        for row in rows[1:]:
            row_num = int(row.attrib["r"])
            if row_num in logos:
                set_inline_text(row, f"{logo_col}{row_num}", logos[row_num])
        dim = root.find(q("dimension"))
        if dim is not None:
            max_row = max(int(row.attrib["r"]) for row in rows)
            dim.attrib["ref"] = f"A1:{logo_col}{max_row}"
        xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        for item in zin.infolist():
            if item.filename == sheet:
                zout.writestr(item, xml)
            else:
                zout.writestr(item, zin.read(item.filename))
    temp.replace(WORKBOOK)


def main() -> None:
    LOGO_DIR.mkdir(exist_ok=True)
    shops, existing = load_shops()
    logos: dict[int, str] = {row: value for row, value in existing.items() if value}
    todo = [shop for shop in shops if not logos.get(shop.row)]
    print(f"shops={len(shops)} todo={len(todo)} existing={len(logos)}")
    for idx, shop in enumerate(todo, 1):
        logo = find_logo(shop)
        if logo:
            logos[shop.row] = logo
            print(f"{idx}/{len(todo)} OK row={shop.row} {shop.company!r} -> {logo}")
        else:
            print(f"{idx}/{len(todo)} MISS row={shop.row} {shop.company!r} site={shop.site}")
        time.sleep(0.2)
    update_workbook(logos)
    print(f"updated={len(logos)} backup={BACKUP.name}")


if __name__ == "__main__":
    main()
