#!/usr/bin/env python3
"""Scan text-bearing PPTX parts for encoding damage and placeholder residue."""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

MOJIBAKE = re.compile(r"\uFFFD|Ã.|Â.|锟斤拷|浣犲ソ|鎴戜滑|鍙戝竷")
PLACEHOLDER = re.compile(
    r"\b(?:lorem|ipsum|todo)\b|\[\s*insert\b|click\s+to\s+add|单击此处(?:添加|输入)",
    re.IGNORECASE,
)
TEXT_PART = re.compile(
    r"ppt/(?:slides|notesSlides|charts|diagrams)/[^/]+\.xml$"
)
DRAWING_TEXT = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--check-placeholders", action="store_true")
    parser.add_argument("--forbid-text", action="append", default=[], help="exact text fragment that must not remain; repeatable")
    args = parser.parse_args()

    problems: list[str] = []
    scanned = 0
    slide_count = 0
    with zipfile.ZipFile(args.pptx) as archive:
        corrupt = archive.testzip()
        if corrupt:
            raise SystemExit(f"Corrupt ZIP member: {corrupt}")
        for name in sorted(archive.namelist()):
            if not TEXT_PART.fullmatch(name):
                continue
            scanned += 1
            if name.startswith("ppt/slides/slide"):
                slide_count += 1
            data = archive.read(name)
            try:
                raw = data.decode("utf-8")
            except UnicodeDecodeError as exc:
                problems.append(f"{name}: invalid UTF-8: {exc}")
                continue
            if MOJIBAKE.search(raw):
                problems.append(f"{name}: suspected mojibake")
            try:
                root = ET.fromstring(data)
                text = "\n".join((node.text or "") for node in root.iter(DRAWING_TEXT))
            except ET.ParseError as exc:
                problems.append(f"{name}: XML parse error: {exc}")
                continue
            if args.check_placeholders and PLACEHOLDER.search(text):
                problems.append(f"{name}: placeholder-like text remains")
            for fragment in args.forbid_text:
                if fragment in text:
                    problems.append(f"{name}: forbidden text remains: {fragment!r}")

    if problems:
        print("\n".join(problems))
        raise SystemExit(1)
    print(f"Validated {slide_count} slides across {scanned} text-bearing XML parts")


if __name__ == "__main__":
    main()
