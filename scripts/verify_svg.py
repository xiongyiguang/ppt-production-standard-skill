#!/usr/bin/env python3
"""Check editable slide SVG sources for common delivery failures."""

from __future__ import annotations

import argparse
import re
from collections import Counter
from pathlib import Path
import xml.etree.ElementTree as ET

XLINK = "{http://www.w3.org/1999/xlink}href"
MOJIBAKE = re.compile(r"\uFFFD|Ã.|Â.|锟斤拷|浣犲ソ|鎴戜滑|鍙戝竷")
NUMBER = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)")


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def numeric(value: str | None) -> float | None:
    if not value:
        return None
    match = NUMBER.match(value.strip())
    return float(match.group()) if match else None


def canvas(root: ET.Element) -> tuple[float, float] | None:
    view_box = root.attrib.get("viewBox", "").replace(",", " ").split()
    if len(view_box) == 4:
        try:
            width, height = float(view_box[2]), float(view_box[3])
            if width > 0 and height > 0:
                return width, height
        except ValueError:
            pass
    width, height = numeric(root.attrib.get("width")), numeric(root.attrib.get("height"))
    return (width, height) if width and height and width > 0 and height > 0 else None


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("svg", nargs="+", type=Path)
    parser.add_argument("--require-embedded-images", action="store_true")
    parser.add_argument("--forbid-full-page-raster", action="store_true")
    parser.add_argument("--allow-no-text", action="store_true")
    parser.add_argument("--strict-export-safe", action="store_true")
    args = parser.parse_args()

    errors: list[str] = []
    warnings: list[str] = []
    for path in args.svg:
        try:
            raw = path.read_text(encoding="utf-8")
        except (OSError, UnicodeError) as exc:
            errors.append(f"{path}: cannot read UTF-8 SVG: {exc}")
            continue
        if MOJIBAKE.search(raw):
            errors.append(f"{path}: suspected mojibake")
        try:
            root = ET.fromstring(raw)
        except ET.ParseError as exc:
            errors.append(f"{path}: XML parse error: {exc}")
            continue
        page_size = canvas(root)
        if "viewBox" not in root.attrib:
            errors.append(f"{path}: missing viewBox")

        icon_keys: list[str] = []
        text_count = 0
        for element in root.iter():
            name = local_name(element.tag)
            if name == "text":
                text_count += 1
            key = element.attrib.get("data-icon-key")
            if key:
                icon_keys.append(key)
            if name == "image":
                href = element.attrib.get("href") or element.attrib.get(XLINK)
                if not href:
                    errors.append(f"{path}: image element has no href")
                elif args.require_embedded_images and not href.startswith("data:image/"):
                    errors.append(f"{path}: external image is not embedded: {href}")
                if args.forbid_full_page_raster and page_size:
                    width, height = numeric(element.attrib.get("width")), numeric(element.attrib.get("height"))
                    if width and height and width >= page_size[0] * 0.9 and height >= page_size[1] * 0.9:
                        errors.append(f"{path}: image covers at least 90% of the canvas in both dimensions")
            if name in {"script", "foreignObject", "filter", "mask"}:
                message = f"{path}: export-risk element <{name}>"
                (errors if args.strict_export_safe else warnings).append(message)

        duplicates = [key for key, count in Counter(icon_keys).items() if count > 1]
        if duplicates:
            errors.append(f"{path}: duplicate data-icon-key values: {', '.join(sorted(duplicates))}")
        if text_count == 0 and not args.allow_no_text:
            errors.append(f"{path}: no editable SVG text elements")

    for warning in warnings:
        print("WARNING: " + warning)
    if errors:
        print("\n".join(errors))
        raise SystemExit(1)
    print(f"Validated {len(args.svg)} SVG file(s); warnings: {len(warnings)}")


if __name__ == "__main__":
    main()
