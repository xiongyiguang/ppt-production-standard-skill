#!/usr/bin/env python3
"""Embed icon files into SVG <image data-icon-key="..."> elements."""

from __future__ import annotations

import argparse
import base64
import json
import mimetypes
from pathlib import Path
import xml.etree.ElementTree as ET

XLINK = "http://www.w3.org/1999/xlink"
ET.register_namespace("xlink", XLINK)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("svg", type=Path)
    parser.add_argument("mapping", type=Path, help="JSON object mapping data-icon-key to an image path")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--allow-unused", action="store_true", help="allow mapping keys that are not present in the SVG")
    args = parser.parse_args()

    mapping = json.loads(args.mapping.read_text(encoding="utf-8"))
    if not isinstance(mapping, dict) or not all(isinstance(key, str) and isinstance(value, str) for key, value in mapping.items()):
        raise SystemExit("Mapping must be a JSON object of string icon keys to string file paths")
    root = ET.parse(args.svg)
    unresolved = []
    used = set()
    embedded = 0
    for element in root.iter():
        key = element.attrib.get("data-icon-key")
        if not key:
            continue
        if key not in mapping:
            unresolved.append(key)
            continue
        used.add(key)
        icon_path = Path(mapping[key])
        if not icon_path.is_absolute():
            icon_path = (args.mapping.parent / icon_path).resolve()
        if not icon_path.is_file():
            raise SystemExit(f"Mapped icon does not exist or is not a file: {icon_path}")
        mime = mimetypes.guess_type(icon_path.name)[0] or "image/png"
        if not mime.startswith("image/"):
            raise SystemExit(f"Mapped asset is not an image: {icon_path}")
        data = base64.b64encode(icon_path.read_bytes()).decode("ascii")
        uri = f"data:{mime};base64,{data}"
        element.set("href", uri)
        element.set(f"{{{XLINK}}}href", uri)
        embedded += 1
    if unresolved:
        raise SystemExit("Unresolved icon keys: " + ", ".join(sorted(set(unresolved))))
    unused = sorted(set(mapping) - used)
    if unused and not args.allow_unused:
        raise SystemExit("Unused mapping keys: " + ", ".join(unused))
    if embedded == 0:
        raise SystemExit("No data-icon-key elements were embedded")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    root.write(args.output, encoding="utf-8", xml_declaration=True)
    print(f"Embedded {embedded} icon images into {args.output}")


if __name__ == "__main__":
    main()
