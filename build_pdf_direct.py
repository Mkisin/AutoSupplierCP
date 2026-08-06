from __future__ import annotations

import base64
import html
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
from argparse import ArgumentParser
from pathlib import Path

from build_presentation_direct_new import (
    apply_selection_override,
    image_map,
    read_payload,
    read_selection_override,
)


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
HTML_TEMPLATE = TEMPLATE_DIR / "pdf_presentation.html"
CSS_TEMPLATE = TEMPLATE_DIR / "pdf_presentation.css"


def e(value: object) -> str:
    return html.escape(str(value or ""), quote=True)


def slugify(value: str) -> str:
    slug = re.sub(r"\W+", "_", value, flags=re.UNICODE).strip("_").lower()
    return slug or "presentation"


def next_available_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(2, 1000):
        candidate = path.with_name(f"{path.stem}_v{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Не удалось подобрать свободное имя файла для {path.name}")


def data_uri(path: Path | None) -> str:
    if path is None or not path.exists() or not path.is_file():
        return ""
    content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{content_type};base64,{encoded}"


def first_nonempty(*values: object) -> str:
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def replacement(replacements: dict[str, object], key: str) -> str:
    return str(replacements.get(key) or "").strip()


def image_tag(images: dict[str, Path], token: str, class_name: str = "") -> str:
    uri = data_uri(images.get(token))
    if not uri:
        return '<div class="empty">Изображение не найдено</div>'
    class_attr = f' class="{e(class_name)}"' if class_name else ""
    return f'<img{class_attr} src="{uri}" alt="">'


def metric(value: object, label: str) -> str:
    text = first_nonempty(value, "—")
    return (
        '<div class="panel metric">'
        f'<div class="metric-value">{e(text)}</div>'
        f'<div class="metric-label">{e(label)}</div>'
        "</div>"
    )


def logo_tile(images: dict[str, Path], token: str) -> str:
    uri = data_uri(images.get(token))
    if not uri:
        return '<div class="logo-tile"><span class="empty">Нет логотипа</span></div>'
    return f'<div class="logo-tile"><img src="{uri}" alt=""></div>'


def slide(inner: str) -> str:
    return (
        '<section class="slide">'
        '<div class="band-top"></div>'
        '<div class="band-side"></div>'
        f'<div class="content">{inner}</div>'
        "</section>"
    )


def build_slides(payload: dict[str, object], replacements: dict[str, object], images: dict[str, Path]) -> str:
    company = str(payload.get("company") or "").strip()
    category = str(payload.get("category") or "").strip()
    contact_line = replacement(replacements, "{{block1}}")
    title = first_nonempty(replacement(replacements, "{{block2}}"), company)
    intro = first_nonempty(replacement(replacements, "{{block5}}"), category)

    slides: list[str] = []
    slides.append(
        slide(
            '<div class="hero-layout">'
            '<div class="hero-text">'
            '<div class="kicker">Центр Закупок Сетей</div>'
            f'<h1 class="title">{e(title)}</h1>'
            f'<p class="subtitle">{e(intro)}</p>'
            f'<p class="meta">{e(contact_line)}</p>'
            "</div>"
            '<div class="hero-photo">'
            f'<div class="image-tile primary">{image_tag(images, "{{pic2}}")}</div>'
            f'<div class="image-tile">{image_tag(images, "{{pic3}}")}</div>'
            f'<div class="image-tile">{image_tag(images, "{{pic4}}")}</div>'
            "</div>"
            "</div>"
        )
    )

    slides.append(
        slide(
            f'<h2 class="section-title">Потенциал переговоров по категории {e(category)}</h2>'
            f'<p class="section-note">{e(replacement(replacements, "{{block4}}"))}</p>'
            '<div class="stats-grid">'
            + metric(replacement(replacements, "{{block6}}"), "контрактов")
            + metric(replacement(replacements, "{{block9}}"), "закупщики в категории")
            + metric(replacement(replacements, "{{block10}}"), "федеральные сети")
            + metric(replacement(replacements, "{{block14}}"), "контрактов всего")
            + metric(replacement(replacements, "{{block11}}"), "региональные сети")
            + metric(replacement(replacements, "{{block12}}"), "локальные сети")
            + metric(replacement(replacements, "{{block13}}"), "HoReCa")
            + metric(replacement(replacements, "{{block31}}"), "средний чек")
            + "</div>"
        )
    )

    slides.append(
        slide(
            '<h2 class="section-title">Фото сетей и переговорной среды</h2>'
            '<p class="section-note">Подборка используется как визуальное подтверждение релевантных каналов продаж.</p>'
            '<div class="photo-wall">'
            f'<div class="image-tile wide">{image_tag(images, "{{pic2}}")}</div>'
            f'<div class="image-tile">{image_tag(images, "{{pic3}}")}</div>'
            f'<div class="image-tile">{image_tag(images, "{{pic4}}")}</div>'
            f'<div class="image-tile">{image_tag(images, "{{pic5}}")}</div>'
            f'<div class="image-tile">{image_tag(images, "{{pic6}}")}</div>'
            "</div>"
        )
    )

    slides.append(
        slide(
            '<h2 class="section-title">Сети, с которыми может быть релевантен диалог</h2>'
            '<p class="section-note">Логотипы подбираются из текущего каталога и предпочтений компании.</p>'
            '<div class="logos-grid">'
            + "".join(logo_tile(images, f"{{{{logo{index}}}}}") for index in range(4, 12))
            + "</div>"
        )
    )

    review_cards = []
    review_tokens = [
        ("{{logo1}}", "{{block15}}", "{{block16}}", "{{block17}}"),
        ("{{logo2}}", "{{block18}}", "{{block19}}", "{{block20}}"),
        ("{{logo3}}", "{{block21}}", "{{block22}}", "{{block23}}"),
    ]
    for logo_token, company_token, person_token, text_token in review_tokens:
        review_cards.append(
            '<div class="panel review">'
            f'<div class="review-logo">{image_tag(images, logo_token)}</div>'
            f'<div class="review-company">{e(replacement(replacements, company_token))}</div>'
            f'<div class="review-person">{e(replacement(replacements, person_token))}</div>'
            f'<div class="review-text">{e(replacement(replacements, text_token))}</div>'
            "</div>"
        )
    slides.append(
        slide(
            '<h2 class="section-title">Отзывы поставщиков</h2>'
            '<p class="section-note">Три релевантных кейса из базы отзывов.</p>'
            '<div class="review-grid">'
            + "".join(review_cards)
            + "</div>"
        )
    )

    slides.append(
        slide(
            '<div class="closing">'
            '<div>'
            '<div class="kicker">Следующий шаг</div>'
            f'<h2 class="title">Готовы обсудить вход в сети, {e(replacement(replacements, "{{block24}}"))}?</h2>'
            f'<p class="subtitle">{e(replacement(replacements, "{{block5}}"))}</p>'
            "</div>"
            '<div class="callout">'
            f'{e(replacement(replacements, "{{block25}}"))} переговоров и '
            f'{e(replacement(replacements, "{{block26}}"))} контрактов в 2025 году помогают оценить практический потенциал категории.'
            "</div>"
            "</div>"
        )
    )
    return "\n".join(slides)


def render_html(company: str, output_html: Path) -> dict[str, object]:
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

    css = CSS_TEMPLATE.read_text(encoding="utf-8")
    template = HTML_TEMPLATE.read_text(encoding="utf-8")
    document = (
        template.replace("{{title}}", e(f"PDF презентация: {payload['company']}"))
        .replace("{{css}}", css)
        .replace("{{slides}}", build_slides(payload, replacements, images))
    )
    output_html.write_text(document, encoding="utf-8")
    return {
        "html": str(output_html),
        "client": payload["company"],
        "contact": payload["contact"],
        "position": payload["position"],
        "category": payload["category"],
        "planned_images": image_sources,
        "photo_candidates": photo_candidates,
        "selected_photo_ids": selected_photo_ids,
        "photo_tokens": photo_tokens,
    }


def browser_candidates() -> list[str]:
    candidates = [
        os.environ.get("PDF_BROWSER", "").strip(),
        shutil.which("chromium") or "",
        shutil.which("chromium-browser") or "",
        shutil.which("google-chrome") or "",
        shutil.which("chrome") or "",
        shutil.which("msedge") or "",
        shutil.which("MicrosoftEdge") or "",
    ]
    if os.name == "nt":
        candidates.extend(
            [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            ]
        )
    return [item for item in candidates if item and Path(item).exists()]


def export_pdf_with_browser(input_html: Path, output_pdf: Path) -> str | None:
    browsers = browser_candidates()
    if not browsers:
        return "Chrome/Edge/Chromium не найден; HTML создан, PDF можно напечатать из браузера вручную."
    browser = browsers[0]
    url = input_html.resolve().as_uri()
    profile_dir = Path(tempfile.mkdtemp(prefix="pdf_browser_profile_"))
    try:
        command = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-crash-reporter",
            "--disable-breakpad",
            "--disable-dev-shm-usage",
            f"--user-data-dir={profile_dir.resolve()}",
            f"--print-to-pdf={output_pdf.resolve()}",
            "--print-to-pdf-no-header",
            url,
        ]
        process = subprocess.run(
            command,
            cwd=BASE_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            timeout=90,
        )
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)
    if process.returncode != 0 or not output_pdf.exists():
        details = (process.stderr or process.stdout or "").strip()
        return f"Не удалось экспортировать PDF через {browser}: {details}"
    return None


def main() -> None:
    parser = ArgumentParser(description="Build direct HTML/PDF presentation.")
    parser.add_argument("company", nargs="?", default="ГОСУДАРЕВ СТАНДАРТ")
    parser.add_argument("--html-only", action="store_true", help="Only create HTML, skip PDF export.")
    args = parser.parse_args()

    payload = read_payload(args.company)
    slug = slugify(str(payload["company"]))
    output_pdf = next_available_path(BASE_DIR / f"{slug}_ЦЗС_pdf_direct.pdf")
    output_html = output_pdf.with_suffix(".html")

    report = render_html(args.company, output_html)
    report["pdf"] = None
    report["pdf_error"] = None
    if not args.html_only:
        error = export_pdf_with_browser(output_html, output_pdf)
        if error:
            report["pdf_error"] = error
        else:
            report["pdf"] = str(output_pdf)

    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
