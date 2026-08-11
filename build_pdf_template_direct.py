from __future__ import annotations

import json
import math
import os
import re
import sys
import textwrap
from argparse import ArgumentParser
from pathlib import Path
from typing import Any

import fitz

from build_presentation_direct_new import (
    apply_selection_override,
    image_map,
    normalize_images_for_powerpoint,
    read_payload,
    read_selection_override,
)


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = BASE_DIR / "Шаблон презентации новый.pdf"
PAGE_WIDTH = 842.25
PAGE_HEIGHT = 595.5
REGULAR_FONT_NAMES = (
    "DejaVuSans.ttf",
    "LiberationSans-Regular.ttf",
    "Arial.ttf",
    "arial.ttf",
    "calibri.ttf",
)
BOLD_FONT_NAMES = (
    "DejaVuSans-Bold.ttf",
    "LiberationSans-Bold.ttf",
    "Arial Bold.ttf",
    "arialbd.ttf",
    "calibrib.ttf",
)


def slugify(value: str) -> str:
    slug = re.sub(r"\W+", "_", value, flags=re.UNICODE).strip("_").lower()
    return slug or "presentation"


def configured_font_path(env_name: str) -> Path | None:
    value = os.environ.get(env_name, "").strip()
    if not value:
        return None
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = BASE_DIR / path
    return path


def first_existing_path(paths: list[Path | None]) -> Path | None:
    return next((path for path in paths if path is not None and path.is_file()), None)


def find_installed_font(names: tuple[str, ...]) -> Path | None:
    wanted = {name.lower() for name in names}
    roots = [
        Path("/usr/share/fonts"),
        Path("/usr/local/share/fonts"),
        Path.home() / ".fonts",
    ]
    for root in roots:
        if not root.exists():
            continue
        try:
            for path in root.rglob("*"):
                if path.is_file() and path.name.lower() in wanted:
                    return path
        except OSError:
            continue
    return None


def resolve_template_fonts() -> tuple[Path, Path]:
    regular_candidates = [
        configured_font_path("PDF_TEMPLATE_FONT"),
        BASE_DIR / "fonts" / "DejaVuSans.ttf",
        BASE_DIR / "fonts" / "LiberationSans-Regular.ttf",
        Path(r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\calibri.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
        Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"),
    ]
    bold_candidates = [
        configured_font_path("PDF_TEMPLATE_BOLD_FONT"),
        BASE_DIR / "fonts" / "DejaVuSans-Bold.ttf",
        BASE_DIR / "fonts" / "LiberationSans-Bold.ttf",
        Path(r"C:\Windows\Fonts\arialbd.ttf"),
        Path(r"C:\Windows\Fonts\calibrib.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"),
    ]
    regular_font = first_existing_path(regular_candidates) or find_installed_font(REGULAR_FONT_NAMES)
    if regular_font is None:
        raise RuntimeError(
            "PDF template font not found. Set PDF_TEMPLATE_FONT to a readable .ttf/.otf file "
            "or install fonts-dejavu/fonts-liberation."
        )
    bold_font = first_existing_path(bold_candidates) or find_installed_font(BOLD_FONT_NAMES) or regular_font
    return regular_font, bold_font


def next_available_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(2, 1000):
        candidate = path.with_name(f"{path.stem}_v{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Не удалось подобрать свободное имя файла для {path.name}")


def project_path(value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return BASE_DIR / path


def rgb(value: str) -> tuple[float, float, float]:
    value = value.strip().lstrip("#")
    if len(value) != 6:
        return (0, 0, 0)
    return tuple(int(value[index : index + 2], 16) / 255 for index in (0, 2, 4))  # type: ignore[return-value]


def rect(values: list[float]) -> fitz.Rect:
    x0, y0, x1, y1 = values
    return fitz.Rect(
        max(0, min(PAGE_WIDTH, x0)),
        max(0, min(PAGE_HEIGHT, y0)),
        max(0, min(PAGE_WIDTH, x1)),
        max(0, min(PAGE_HEIGHT, y1)),
    )


def expand(values: list[float], pad_x: float = 4, pad_y: float = 3) -> list[float]:
    return [values[0] - pad_x, values[1] - pad_y, values[2] + pad_x, values[3] + pad_y]


FIELDS: dict[str, dict[str, Any]] = {
    "{{block1}}": {
        "page": 1,
        "type": "text",
        "rect": [43.0, 135.0, 799.0, 152.0],
        "font_size": 9.5,
        "min_font_size": 7,
        "color": "#ffffff",
        "fill": None,
        "align": "left",
        "valign": "top",
        "max_lines": 1,
    },
    "{{block2}}": {
        "page": 1,
        "type": "text",
        "rect": [43.0, 153.0, 798.0, 282.0],
        "font_size": 30,
        "min_font_size": 18,
        "color": "#ffffff",
        "fill": None,
        "align": "left",
        "valign": "middle",
        "bold": True,
        "max_lines": 3,
    },
    "{{block3}}": {
        "page": 1,
        "type": "text",
        "rect": [277.0, 303.0, 781.0, 321.0],
        "font_size": 11,
        "min_font_size": 8,
        "color": "#191919",
        "fill": "#ffffff",
        "align": "left",
        "valign": "middle",
        "prefix": "Категория ",
        "max_lines": 1,
    },
    "{{block4}}": {
        "page": 2,
        "type": "text",
        "rect": [438.0, 29.0, 772.0, 59.0],
        "font_size": 17,
        "min_font_size": 12,
        "color": "#ffffff",
        "fill": "#0a5f81",
        "align": "left",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block5}}": {
        "page": 3,
        "type": "text",
        "rect": [282.0, 78.0, 728.0, 106.0],
        "font_size": 16,
        "min_font_size": 10,
        "color": "#262626",
        "fill": "#ffffff",
        "align": "left",
        "valign": "middle",
        "max_lines": 2,
    },
    "{{block6}}": {
        "page": 3,
        "type": "text",
        "rect": [265.0, 139.0, 480.0, 196.0],
        "font_size": 38,
        "min_font_size": 24,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block8}}": {
        "page": 3,
        "type": "text",
        "rect": [451.0, 139.0, 685.0, 212.0],
        "font_size": 38,
        "min_font_size": 24,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block31}}": {
        "page": 3,
        "type": "text",
        "rect": [608.0, 139.0, 842.0, 212.0],
        "font_size": 24,
        "min_font_size": 16,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block9}}": {
        "page": 3,
        "type": "text",
        "rect": [252.0, 296.0, 411.0, 337.0],
        "font_size": 28,
        "min_font_size": 18,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block10}}": {
        "page": 3,
        "type": "text",
        "rect": [421.0, 315.0, 584.0, 358.0],
        "font_size": 20,
        "min_font_size": 14,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block11}}": {
        "page": 3,
        "type": "text",
        "rect": [421.0, 371.0, 584.0, 419.0],
        "font_size": 20,
        "min_font_size": 14,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block12}}": {
        "page": 3,
        "type": "text",
        "rect": [595.0, 315.0, 802.0, 358.0],
        "font_size": 20,
        "min_font_size": 14,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block13}}": {
        "page": 3,
        "type": "text",
        "rect": [595.0, 371.0, 802.0, 420.0],
        "font_size": 20,
        "min_font_size": 14,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block14}}": {
        "page": 5,
        "type": "text",
        "rect": [566.0, 62.0, 831.0, 188.0],
        "font_size": 30,
        "min_font_size": 18,
        "color": "#184c6b",
        "fill": "#ffffff",
        "align": "center",
        "valign": "middle",
        "max_lines": 1,
    },
    "{{block15}}": {"page": 7, "type": "text", "rect": [289.0, 121.0, 505.0, 151.0], "font_size": 14, "min_font_size": 9, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "middle", "max_lines": 1},
    "{{block16}}": {"page": 7, "type": "text", "rect": [277.0, 134.0, 469.0, 158.0], "font_size": 9, "min_font_size": 7, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "middle", "max_lines": 1},
    "{{block17}}": {"page": 7, "type": "text", "rect": [277.0, 171.0, 536.0, 281.0], "font_size": 8.7, "min_font_size": 6, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "top", "max_lines": 8},
    "{{block18}}": {"page": 7, "type": "text", "rect": [559.0, 121.0, 752.0, 151.0], "font_size": 14, "min_font_size": 9, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "middle", "max_lines": 1},
    "{{block19}}": {"page": 7, "type": "text", "rect": [559.0, 140.0, 753.0, 164.0], "font_size": 9, "min_font_size": 7, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "middle", "max_lines": 1},
    "{{block20}}": {"page": 7, "type": "text", "rect": [559.0, 165.0, 812.0, 269.0], "font_size": 8.7, "min_font_size": 6, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "top", "max_lines": 8},
    "{{block21}}": {"page": 7, "type": "text", "rect": [277.0, 345.0, 506.0, 374.0], "font_size": 14, "min_font_size": 9, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "middle", "max_lines": 1},
    "{{block22}}": {"page": 7, "type": "text", "rect": [277.0, 368.0, 469.0, 392.0], "font_size": 9, "min_font_size": 7, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "middle", "max_lines": 1},
    "{{block23}}": {"page": 7, "type": "text", "rect": [277.0, 393.0, 542.0, 495.0], "font_size": 8.7, "min_font_size": 6, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "top", "max_lines": 8},
    "{{block24}}": {"page": 8, "type": "text", "rect": [282.0, 22.0, 701.0, 98.0], "font_size": 23, "min_font_size": 14, "color": "#1d1d1d", "fill": "#ffffff", "align": "left", "valign": "middle", "max_lines": 2, "suffix": ", Вы в одном клике\nот входа в сети!"},
    "{{block25}}": {"page": 8, "type": "skip"},
    "{{block26}}": {"page": 8, "type": "skip"},
    "{{block3_quote}}": {"page": 8, "type": "text", "rect": [438.0, 188.0, 758.0, 276.0], "font_size": 10.5, "min_font_size": 7.5, "color": "#222222", "fill": "#ffffff", "align": "left", "valign": "middle", "max_lines": 5},
    "{{event_date}}": {"page": 1, "type": "text", "rect": [284.0, 528.0, 672.0, 553.0], "font_size": 14, "min_font_size": 10, "color": "#ffffff", "fill": None, "align": "left", "valign": "middle", "max_lines": 1},
}

for index, x0 in enumerate([265.0, 331.0, 397.0, 463.0, 529.0, 595.0, 661.0, 727.0], start=4):
    FIELDS[f"{{{{logo{index}}}}}"] = {
        "page": 1,
        "type": "image",
        "rect": [x0, 405.0, x0 + 60.0, 466.0],
        "fit": "contain",
        "fill": "#f5f8fa",
    }

for token, values in {
    "{{pic2}}": [271.0, 141.0, 620.0, 397.0],
    "{{pic3}}": [637.0, 141.0, 831.0, 263.0],
    "{{pic4}}": [270.0, 417.0, 458.0, 551.0],
    "{{pic5}}": [469.0, 418.0, 653.0, 552.0],
    "{{pic6}}": [660.0, 415.0, 837.0, 550.0],
}.items():
    FIELDS[token] = {"page": 6, "type": "image", "rect": values, "fit": "cover", "fill": "#ffffff"}

for token, values in {
    "{{logo1}}": [469.0, 111.0, 542.0, 158.0],
    "{{logo2}}": [751.0, 111.0, 823.0, 158.0],
    "{{logo3}}": [463.0, 333.0, 535.0, 380.0],
}.items():
    FIELDS[token] = {"page": 7, "type": "image", "rect": values, "fit": "contain", "fill": "#ffffff"}


def display_value(token: str, replacements: dict[str, object]) -> str:
    if token == "{{block3_quote}}":
        block25 = str(replacements.get("{{block25}}") or "").strip()
        block26 = str(replacements.get("{{block26}}") or "").strip()
        block3 = str(replacements.get("{{block3}}") or "").strip()
        return (
            f"«В 2025 году наши клиенты провели {block25} переговоров "
            f"и заключили {block26} предварительных контрактов по категории "
            f"{block3}. Уверены, вместе мы займем полки в нужных вам сетях»."
        )
    if token == "{{event_date}}":
        return "15 и 16 сентября 2026, Москва, Крокус Экспо"
    value = str(replacements.get(token) or "").strip()
    field = FIELDS.get(token, {})
    return f"{field.get('prefix', '')}{value}{field.get('suffix', '')}"


def align_flag(value: str) -> int:
    return {
        "left": fitz.TEXT_ALIGN_LEFT,
        "center": fitz.TEXT_ALIGN_CENTER,
        "right": fitz.TEXT_ALIGN_RIGHT,
        "justify": fitz.TEXT_ALIGN_JUSTIFY,
    }.get(value, fitz.TEXT_ALIGN_LEFT)


def draw_cover(page: fitz.Page, image_path: Path, target: fitz.Rect, fit: str) -> None:
    source = fitz.Pixmap(str(image_path))
    try:
        iw, ih = source.width, source.height
    finally:
        source = None
    if iw <= 0 or ih <= 0 or target.width <= 0 or target.height <= 0:
        return
    if fit == "cover":
        scale = max(target.width / iw, target.height / ih)
    else:
        scale = min(target.width / iw, target.height / ih)
    width = iw * scale
    height = ih * scale
    placed = fitz.Rect(
        target.x0 + (target.width - width) / 2,
        target.y0 + (target.height - height) / 2,
        target.x0 + (target.width + width) / 2,
        target.y0 + (target.height + height) / 2,
    )
    page.insert_image(placed, filename=str(image_path), keep_proportion=True, overlay=True)


def write_text(
    page: fitz.Page,
    target: fitz.Rect,
    text: str,
    field: dict[str, Any],
    font_file: Path | None,
    font_name: str = "pdfgenfont",
) -> dict[str, Any]:
    warnings: list[str] = []
    fontsize = float(field.get("font_size", 12))
    min_font_size = float(field.get("min_font_size", max(6, fontsize - 4)))
    max_lines = int(field.get("max_lines", 99))
    color = rgb(str(field.get("color", "#000000")))
    align = align_flag(str(field.get("align", "left")))
    lineheight = 1.12

    def estimated_lines(size: float) -> list[str]:
        if not text:
            return [""]
        source_lines = text.splitlines() or [text]
        chars_per_line = max(1, int(target.width / max(1, size * 0.47)))
        lines: list[str] = []
        for source_line in source_lines:
            wrapped = textwrap.wrap(
                source_line,
                width=chars_per_line,
                break_long_words=False,
                replace_whitespace=False,
            )
            lines.extend(wrapped or [""])
        return lines

    def fits(size: float) -> bool:
        lines = estimated_lines(size)
        if len(lines) > max_lines:
            return False
        return len(lines) * size * lineheight <= target.height + 0.1

    def draw(size: float) -> float:
        lines = estimated_lines(size)
        required = len(lines) * size * lineheight
        y0 = target.y0
        if str(field.get("valign", "top")) == "middle" and target.height > required:
            y0 += (target.height - required) / 2
        box = fitz.Rect(target.x0, y0, target.x1, target.y1)
        return page.insert_textbox(
            box,
            text,
            fontsize=size,
            fontname=font_name,
            fontfile=str(font_file) if font_file else None,
            color=color,
            align=align,
            overlay=True,
        )

    selected = min_font_size
    size = fontsize
    while size >= min_font_size:
        if fits(size):
            selected = size
            break
        size -= 0.5
    if selected < fontsize:
        warnings.append(f"Font size reduced from {fontsize:g} to {selected:g}")
    spare = draw(selected)
    if spare < -0.1:
        warnings.append("Text may overflow")
    return {"font_size": selected, "warnings": warnings}


def cover_field(page: fitz.Page, field: dict[str, Any], target: fitz.Rect) -> None:
    fill = field.get("fill")
    if not fill:
        return
    page.draw_rect(target, color=None, fill=rgb(str(fill)), overlay=True)


def build_pdf(
    company: str,
    output_path: Path | None = None,
    template_path: Path = DEFAULT_TEMPLATE,
    debug_layout: bool = False,
) -> dict[str, Any]:
    if not template_path.exists():
        raise RuntimeError(f"PDF template not found: {template_path}")
    payload = read_payload(company)
    replacements = dict(payload["replacements"])
    images, image_sources, photo_candidates, selected_photo_ids, photo_tokens = image_map(payload)
    replacements, images, image_sources = apply_selection_override(
        replacements,
        images,
        image_sources,
        read_selection_override(),
        payload,
    )
    if output_path is None:
        output_path = next_available_path(BASE_DIR / f"{slugify(str(payload['company']))}_ЦЗС_pdf_template.pdf")

    font_file, bold_font_file = resolve_template_fonts()
    normalized_images, normalization_report, temp_files = normalize_images_for_powerpoint(images)
    warnings: list[dict[str, Any]] = []
    text_written: dict[str, int] = {}
    image_written: dict[str, int] = {}

    doc = fitz.open(template_path)
    try:
        for field in FIELDS.values():
            if field.get("type") == "skip":
                continue
            page = doc[int(field["page"]) - 1]
            target = rect(expand(list(field["rect"]), 3, 2))
            fill = field.get("fill", "#ffffff")
            page.add_redact_annot(target, fill=rgb(str(fill)) if fill else None)
        for page in doc:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        for token, field in FIELDS.items():
            if field.get("type") == "skip":
                continue
            page = doc[int(field["page"]) - 1]
            target = rect(list(field["rect"]))
            if field["type"] == "text":
                text = display_value(token, replacements)
                result = write_text(
                    page,
                    target,
                    text,
                    field,
                    bold_font_file if field.get("bold") else font_file,
                    "pdfgenbold" if field.get("bold") else "pdfgenfont",
                )
                text_written[token] = 1 if text else 0
                for message in result["warnings"]:
                    warnings.append({"field": token, "message": message})
            elif field["type"] == "image":
                image_path = normalized_images.get(token)
                if image_path and Path(image_path).exists():
                    draw_cover(page, Path(image_path), target, str(field.get("fit", "contain")))
                    image_written[token] = 1
                else:
                    image_written[token] = 0
                    warnings.append({"field": token, "message": "Image source is missing"})
            if debug_layout:
                page.draw_rect(target, color=(1, 0, 0), width=0.7, overlay=True)
                page.insert_text(
                    fitz.Point(target.x0, max(10, target.y0 - 2)),
                    token,
                    fontsize=6,
                    color=(1, 0, 0),
                    overlay=True,
                )
        doc.save(output_path, garbage=4, deflate=True)
    finally:
        doc.close()
        for temp_file in temp_files:
            try:
                temp_file.unlink(missing_ok=True)
            except OSError:
                pass

    return {
        "created": str(output_path),
        "template": str(template_path),
        "client": payload["company"],
        "contact": payload["contact"],
        "position": payload["position"],
        "category": payload["category"],
        "text_replaced_total": sum(text_written.values()),
        "text_replaced": text_written,
        "image_replaced_total": sum(image_written.values()),
        "image_replaced": image_written,
        "planned_images": image_sources,
        "image_normalization": normalization_report,
        "photo_candidates": photo_candidates,
        "selected_photo_ids": selected_photo_ids,
        "photo_tokens": photo_tokens,
        "warnings": warnings,
    }


def main() -> None:
    parser = ArgumentParser(description="Build PDF presentation directly from PDF template.")
    parser.add_argument("company", nargs="?", default="ГОСУДАРЕВ СТАНДАРТ")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="PDF template path.")
    parser.add_argument("--output", default="", help="Output PDF path.")
    parser.add_argument("--debug-layout", action="store_true", help="Draw field rectangles and names.")
    args = parser.parse_args()

    output = project_path(args.output) if args.output else None
    report = build_pdf(
        args.company,
        output_path=output,
        template_path=project_path(args.template),
        debug_layout=args.debug_layout,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
