#!/usr/bin/env python3
"""Extract visually separated icons from a white-background generated icon sheet."""

from __future__ import annotations

import argparse
import json
from collections import deque
from pathlib import Path

import numpy as np
from PIL import Image


def dilate(mask: np.ndarray, rounds: int) -> np.ndarray:
    result = mask.copy()
    for _ in range(max(0, rounds)):
        padded = np.pad(result, 1, constant_values=False)
        result = np.logical_or.reduce(
            [padded[y:y + result.shape[0], x:x + result.shape[1]] for y in range(3) for x in range(3)]
        )
    return result


def components(mask: np.ndarray, min_area: int) -> list[tuple[int, int, int, int, int]]:
    h, w = mask.shape
    seen = np.zeros_like(mask, dtype=bool)
    found: list[tuple[int, int, int, int, int]] = []
    for y in range(h):
        for x in range(w):
            if not mask[y, x] or seen[y, x]:
                continue
            q = deque([(x, y)])
            seen[y, x] = True
            min_x = max_x = x
            min_y = max_y = y
            area = 0
            while q:
                cx, cy = q.popleft()
                area += 1
                min_x, max_x = min(min_x, cx), max(max_x, cx)
                min_y, max_y = min(min_y, cy), max(max_y, cy)
                for nx, ny in ((cx - 1, cy), (cx + 1, cy), (cx, cy - 1), (cx, cy + 1)):
                    if 0 <= nx < w and 0 <= ny < h and mask[ny, nx] and not seen[ny, nx]:
                        seen[ny, nx] = True
                        q.append((nx, ny))
            if area >= min_area:
                found.append((min_x, min_y, max_x + 1, max_y + 1, area))
    return found


def whiten_to_alpha(crop: Image.Image, threshold: int) -> Image.Image:
    rgba = np.asarray(crop.convert("RGBA")).copy()
    rgb = rgba[:, :, :3].astype(np.int16)
    distance = 255 - rgb.min(axis=2)
    alpha = np.clip(distance * (255 / max(1, 255 - threshold)), 0, 255).astype(np.uint8)
    rgba[:, :, 3] = alpha
    return Image.fromarray(rgba, "RGBA")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("sheet", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--white-threshold", type=int, default=245)
    parser.add_argument("--min-area", type=int, default=400)
    parser.add_argument("--dilate", type=int, default=2)
    parser.add_argument("--padding", type=int, default=16)
    parser.add_argument(
        "--expected-count",
        type=int,
        help="fail before writing files when the detected component count differs from the semantic inventory",
    )
    args = parser.parse_args()

    image = Image.open(args.sheet).convert("RGB")
    array = np.asarray(image)
    foreground = np.any(array < args.white_threshold, axis=2)
    boxes = components(dilate(foreground, args.dilate), args.min_area)
    boxes.sort(key=lambda b: (b[1], b[0]))
    if not boxes:
        raise SystemExit("No icon components detected")
    if args.expected_count is not None and len(boxes) != args.expected_count:
        raise SystemExit(
            f"Detected {len(boxes)} components; expected {args.expected_count}. "
            "Regenerate with wider spacing or adjust extraction parameters."
        )
    args.out_dir.mkdir(parents=True, exist_ok=True)
    manifest = []
    for index, (x1, y1, x2, y2, area) in enumerate(boxes, start=1):
        x1 = max(0, x1 - args.padding)
        y1 = max(0, y1 - args.padding)
        x2 = min(image.width, x2 + args.padding)
        y2 = min(image.height, y2 + args.padding)
        icon = whiten_to_alpha(image.crop((x1, y1, x2, y2)), args.white_threshold)
        name = f"{args.sheet.stem}_c{index:02d}.png"
        icon.save(args.out_dir / name)
        manifest.append({"key": f"{args.sheet.stem}.c{index:02d}", "file": name, "box": [x1, y1, x2, y2], "area": area})
    (args.out_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Extracted {len(manifest)} components to {args.out_dir}")


if __name__ == "__main__":
    main()
