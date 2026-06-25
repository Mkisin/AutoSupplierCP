from __future__ import annotations

import subprocess
import urllib.parse
from pathlib import Path

from fetch_chain_logos import LOGO_DIR, update_workbook, load_shops


ROOT = Path(__file__).resolve().parents[1]
SCRIPT_DIR = Path(__file__).resolve().parent
CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
SIZE = 640


def file_uri(path: Path) -> str:
    return path.resolve().as_uri()


def target_for(source: Path) -> Path:
    target = source.with_suffix(".jpg")
    if source.suffix.lower() in {".jpg", ".jpeg"}:
        return target
    if target.exists():
        return source.with_name(f"{source.stem}-converted.jpg")
    return target


def render_to_png(source: Path, png_path: Path, html_path: Path) -> None:
    html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
html, body {{
  margin: 0;
  width: {SIZE}px;
  height: {SIZE}px;
  overflow: hidden;
  background: #fff;
}}
body {{
  display: flex;
  align-items: center;
  justify-content: center;
}}
img {{
  max-width: {int(SIZE * 0.82)}px;
  max-height: {int(SIZE * 0.82)}px;
  object-fit: contain;
}}
</style>
</head>
<body><img src="{file_uri(source)}"></body>
</html>
"""
    html_path.write_text(html, encoding="utf-8")
    subprocess.run(
        [
            str(CHROME),
            "--headless=new",
            "--disable-gpu",
            "--hide-scrollbars",
            f"--window-size={SIZE},{SIZE}",
            "--force-device-scale-factor=1",
            "--allow-file-access-from-files",
            f"--screenshot={png_path}",
            file_uri(html_path),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def png_to_jpg(png_path: Path, jpg_path: Path) -> None:
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(SCRIPT_DIR / "png_to_jpg.ps1"),
            str(png_path.resolve()),
            str(jpg_path.resolve()),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def convert(source: Path) -> Path:
    target = target_for(source)
    if source.suffix.lower() in {".jpg", ".jpeg"}:
        if source.suffix.lower() == ".jpeg":
            target.write_bytes(source.read_bytes())
        return target
    safe = source.stem.replace(" ", "_")
    png = LOGO_DIR / f"__render_{safe}.png"
    html = LOGO_DIR / f"__render_{safe}.html"
    try:
        render_to_png(source, png, html)
        png_to_jpg(png, target)
    finally:
        for temp_file in (png, html):
            if temp_file.exists():
                temp_file.unlink()
    return target


def main() -> None:
    if not CHROME.exists():
        raise RuntimeError(f"Chrome not found: {CHROME}")
    shops, logos = load_shops()
    converted_by_name: dict[str, str] = {}
    updated = dict(logos)
    for shop in shops:
        name = logos.get(shop.row)
        if not name:
            continue
        source = LOGO_DIR / name
        if not source.exists():
            print(f"MISS file row={shop.row} {name}")
            continue
        if name not in converted_by_name:
            target = convert(source)
            converted_by_name[name] = target.name
            print(f"{name} -> {target.name}")
        updated[shop.row] = converted_by_name[name]
    update_workbook(updated)
    jpg_count = sum(1 for shop in shops if updated.get(shop.row, "").lower().endswith(".jpg"))
    print(f"updated_rows={jpg_count} unique_converted={len(converted_by_name)}")


if __name__ == "__main__":
    main()
