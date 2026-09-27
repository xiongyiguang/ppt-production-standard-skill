#!/usr/bin/env python3
"""Read-only structural audit for the user's PowerPoint production standard."""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET


A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/package/2006/relationships}"

ALLOWED_EXPLICIT_FONTS = {
    "Arial",
    "Microsoft YaHei",
    "Microsoft YaHei UI",
    "微软雅黑",
}
THEME_FONT_PREFIXES = ("+mj", "+mn")
PROTECTED_TEMPLATE_PREFIXES = (
    "ppt/slideMasters/",
    "ppt/slideLayouts/",
    "ppt/theme/",
    "ppt/handoutMasters/",
    "ppt/notesMasters/",
)


def local_name(tag: str) -> str:
    return tag.split("}")[-1]


def relationship_part(part: str) -> str:
    return posixpath.join(
        posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
    )


def related_part(archive: zipfile.ZipFile, part: str, suffix: str) -> str | None:
    rels = relationship_part(part)
    if rels not in archive.namelist():
        return None
    root = ET.fromstring(archive.read(rels))
    for rel in root.findall(R + "Relationship"):
        if rel.attrib.get("Type", "").endswith("/" + suffix):
            return posixpath.normpath(
                posixpath.join(posixpath.dirname(part), rel.attrib["Target"])
            )
    return None


def slide_number(name: str) -> int:
    stem = posixpath.basename(name).removeprefix("slide").removesuffix(".xml")
    return int(stem)


def color_token(node: ET.Element) -> str:
    value = node.attrib.get("val", "")
    transforms = [
        f"{local_name(child.tag)}={child.attrib.get('val', '')}" for child in node
    ]
    return value + ("[" + ",".join(transforms) + "]" if transforms else "")


def paragraph_spacing_nonzero(node: ET.Element) -> bool:
    for tag in ("spcBef", "spcAft"):
        spacing = node.find(A + tag)
        if spacing is None:
            continue
        for child in spacing:
            try:
                if int(child.attrib.get("val", "0")) != 0:
                    return True
            except ValueError:
                return True
    return False


def inspect_slide(data: bytes, slide: str) -> dict:
    root = ET.fromstring(data)
    direct_rgb = Counter()
    scheme_colors = Counter()
    sizes_below_12: list[dict] = []
    noninteger_sizes: list[dict] = []
    disallowed_fonts = Counter()
    nonmiddle_text_shapes: list[dict] = []
    nonzero_paragraph_spacing = 0
    unusual_explicit_line_spacing = Counter()
    text_boxes = 0
    text_autoshapes = 0
    shape_geometries = Counter()

    common_slide = root.find(P + "cSld")
    visible_roots = []
    if common_slide is not None:
        for child_name in ("bg", "spTree"):
            child = common_slide.find(P + child_name)
            if child is not None:
                visible_roots.append(child)
    for visible_root in visible_roots:
        for node in visible_root.iter(A + "srgbClr"):
            direct_rgb[node.attrib.get("val", "").upper()] += 1
        for node in visible_root.iter(A + "schemeClr"):
            scheme_colors[color_token(node)] += 1

    for properties_name in ("rPr", "defRPr", "endParaRPr"):
        for properties in root.iter(A + properties_name):
            size_value = properties.attrib.get("sz")
            if size_value:
                try:
                    points = int(size_value) / 100
                except ValueError:
                    points = -1
                record = {"slide": slide, "size_pt": points}
                if 0 <= points < 12:
                    sizes_below_12.append(record)
                if points >= 0 and not points.is_integer():
                    noninteger_sizes.append(record)

            for font_tag in ("latin", "ea", "cs"):
                font_node = properties.find(A + font_tag)
                if font_node is None:
                    continue
                typeface = font_node.attrib.get("typeface", "").strip()
                if not typeface or typeface.startswith(THEME_FONT_PREFIXES):
                    continue
                if typeface not in ALLOWED_EXPLICIT_FONTS:
                    disallowed_fonts[typeface] += 1

    for paragraph_properties in root.iter(A + "pPr"):
        if paragraph_spacing_nonzero(paragraph_properties):
            nonzero_paragraph_spacing += 1
        line_spacing = paragraph_properties.find(A + "lnSpc")
        if line_spacing is not None and len(line_spacing):
            child = line_spacing[0]
            token = f"{local_name(child.tag)}={child.attrib.get('val', '')}"
            if token not in {"spcPct=100000", "spcPct=120000"}:
                unusual_explicit_line_spacing[token] += 1

    for shape in root.iter(P + "sp"):
        geometry = shape.find(P + "spPr/" + A + "prstGeom")
        if geometry is not None:
            shape_geometries[geometry.attrib.get("prst", "unknown")] += 1

        text = "".join((node.text or "") for node in shape.iter(A + "t")).strip()
        if not text:
            continue
        shape_props = shape.find(P + "nvSpPr")
        common_props = shape_props.find(P + "cNvSpPr") if shape_props is not None else None
        is_text_box = common_props is not None and common_props.attrib.get("txBox") == "1"
        if is_text_box:
            text_boxes += 1
        else:
            text_autoshapes += 1
            body = shape.find(".//" + A + "bodyPr")
            anchor = body.attrib.get("anchor") if body is not None else None
            if anchor != "ctr":
                name_node = shape.find(".//" + P + "cNvPr")
                nonmiddle_text_shapes.append(
                    {
                        "slide": slide,
                        "shape": name_node.attrib.get("name", "") if name_node is not None else "",
                        "anchor": anchor or "default",
                        "text": text[:80],
                    }
                )

    return {
        "direct_rgb": dict(direct_rgb),
        "scheme_colors": dict(scheme_colors),
        "sizes_below_12": sizes_below_12,
        "noninteger_sizes": noninteger_sizes,
        "disallowed_fonts": dict(disallowed_fonts),
        "nonmiddle_text_shapes": nonmiddle_text_shapes,
        "nonzero_paragraph_spacing": nonzero_paragraph_spacing,
        "unusual_explicit_line_spacing": dict(unusual_explicit_line_spacing),
        "text_boxes": text_boxes,
        "text_autoshapes": text_autoshapes,
        "shape_geometries": dict(shape_geometries),
        "straight_rectangles": shape_geometries["rect"],
        "rounded_rectangles": sum(
            count
            for geometry, count in shape_geometries.items()
            if "round" in geometry.lower()
        ),
    }


def merge_slide_results(results: list[dict]) -> dict:
    direct_rgb = Counter()
    scheme_colors = Counter()
    disallowed_fonts = Counter()
    unusual_spacing = Counter()
    shape_geometries = Counter()
    merged = {
        "direct_rgb": {},
        "scheme_colors": {},
        "sizes_below_12": [],
        "noninteger_sizes": [],
        "disallowed_fonts": {},
        "nonmiddle_text_shapes": [],
        "nonzero_paragraph_spacing": 0,
        "unusual_explicit_line_spacing": {},
        "text_boxes": 0,
        "text_autoshapes": 0,
        "shape_geometries": {},
        "straight_rectangles": 0,
        "rounded_rectangles": 0,
    }
    for result in results:
        direct_rgb.update(result["direct_rgb"])
        scheme_colors.update(result["scheme_colors"])
        disallowed_fonts.update(result["disallowed_fonts"])
        unusual_spacing.update(result["unusual_explicit_line_spacing"])
        shape_geometries.update(result["shape_geometries"])
        for key in ("sizes_below_12", "noninteger_sizes", "nonmiddle_text_shapes"):
            merged[key].extend(result[key])
        for key in (
            "nonzero_paragraph_spacing",
            "text_boxes",
            "text_autoshapes",
            "straight_rectangles",
            "rounded_rectangles",
        ):
            merged[key] += result[key]
    merged["direct_rgb"] = dict(direct_rgb)
    merged["scheme_colors"] = dict(scheme_colors)
    merged["disallowed_fonts"] = dict(disallowed_fonts)
    merged["unusual_explicit_line_spacing"] = dict(unusual_spacing)
    merged["shape_geometries"] = dict(shape_geometries)
    return merged


def layout_name(archive: zipfile.ZipFile, layout: str | None) -> str:
    if not layout:
        return ""
    root = ET.fromstring(archive.read(layout))
    common_slide = root.find(P + "cSld")
    return common_slide.attrib.get("name", "") if common_slide is not None else ""


def parse_slide_spec(specification: str) -> set[int]:
    slides: set[int] = set()
    for token in specification.split(","):
        token = token.strip()
        if not token:
            continue
        if "-" in token:
            start_text, end_text = token.split("-", 1)
            start = int(start_text)
            end = int(end_text)
            if start <= 0 or end < start:
                raise ValueError(f"invalid slide range: {token}")
            slides.update(range(start, end + 1))
        else:
            slide = int(token)
            if slide <= 0:
                raise ValueError(f"invalid slide number: {token}")
            slides.add(slide)
    return slides


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def protected_template_parts(archive: zipfile.ZipFile) -> set[str]:
    return {
        name
        for name in archive.namelist()
        if name.startswith(PROTECTED_TEMPLATE_PREFIXES)
    }


def compare_template_integrity(
    archive: zipfile.ZipFile, baseline: zipfile.ZipFile
) -> dict:
    current_parts = protected_template_parts(archive)
    baseline_parts = protected_template_parts(baseline)
    missing = sorted(baseline_parts - current_parts)
    extra = sorted(current_parts - baseline_parts)
    changed = []
    for part in sorted(current_parts & baseline_parts):
        current_hash = sha256_bytes(archive.read(part))
        baseline_hash = sha256_bytes(baseline.read(part))
        if current_hash != baseline_hash:
            changed.append(
                {
                    "part": part,
                    "current_sha256": current_hash,
                    "baseline_sha256": baseline_hash,
                }
            )
    return {
        "ok": not missing and not extra and not changed,
        "missing_parts": missing,
        "extra_parts": extra,
        "changed_parts": changed,
    }


def audit(
    path: Path,
    require_richinfo_theme: bool,
    required_layout: str | None,
    required_color_scheme: str | None,
    required_theme_name: str | None,
    content_slides: set[int],
    required_template: Path | None,
) -> dict:
    with zipfile.ZipFile(path, "r") as archive:
        corrupt_member = archive.testzip()
        slides = sorted(
            (
                name
                for name in archive.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ),
            key=slide_number,
        )
        slide_results = [
            inspect_slide(archive.read(slide), f"slide{slide_number(slide)}")
            for slide in slides
        ]
        slide_summary = merge_slide_results(slide_results)
        inherited_slide_parts: set[str] = set()
        template_integrity = None
        if required_template is not None:
            with zipfile.ZipFile(required_template, "r") as baseline:
                template_integrity = compare_template_integrity(archive, baseline)
                baseline_members = set(baseline.namelist())
                inherited_slide_parts = {
                    slide
                    for slide in slides
                    if slide in baseline_members
                    and archive.read(slide) == baseline.read(slide)
                }
        compliance_results = [
            result
            for slide, result in zip(slides, slide_results)
            if slide not in inherited_slide_parts
        ]
        compliance_summary = merge_slide_results(compliance_results)

        supporting_parts = sorted(
            name
            for name in archive.namelist()
            if name.endswith(".xml")
            and (
                name.startswith("ppt/slideLayouts/slideLayout")
                or name.startswith("ppt/slideMasters/slideMaster")
                or name.startswith("ppt/handoutMasters/handoutMaster")
                or name.startswith("ppt/notesMasters/notesMaster")
            )
        )
        supporting_results = [
            inspect_slide(archive.read(part), part) for part in supporting_parts
        ]
        supporting_parts_summary = merge_slide_results(supporting_results)
        merged = merge_slide_results(slide_results + supporting_results)
        slide_shape_summary = {}
        rounded_rectangle_review_notes = []
        for slide, result in zip(slides, slide_results):
            slide_label = f"slide{slide_number(slide)}"
            straight = result["straight_rectangles"]
            rounded = result["rounded_rectangles"]
            total_rectangles = straight + rounded
            ratio = rounded / total_rectangles if total_rectangles else 0
            slide_shape_summary[slide_label] = {
                "straight_rectangles": straight,
                "rounded_rectangles": rounded,
                "rounded_ratio": round(ratio, 4),
            }
            if rounded:
                rounded_rectangle_review_notes.append(
                    {
                        "slide": slide_label,
                        "straight_rectangles": straight,
                        "rounded_rectangles": rounded,
                        "rounded_ratio": round(ratio, 4),
                        "message": (
                            "counts are descriptive only; verify outer section and "
                            "module frames are straight rectangles, while rounded "
                            "rectangles are used as inner elements"
                        ),
                    }
                )

        layout_parts = sorted(
            (
                name
                for name in archive.namelist()
                if name.startswith("ppt/slideLayouts/slideLayout")
                and name.endswith(".xml")
            )
        )
        available_layouts = {
            layout: layout_name(archive, layout) for layout in layout_parts
        }

        theme_names = Counter()
        color_scheme_names = Counter()
        theme_parts = Counter()
        slide_theme_map = {}
        slide_theme_name_map = {}
        slide_layout_map = {}
        slide_color_scheme_map = {}
        for slide in slides:
            layout = related_part(archive, slide, "slideLayout")
            layout_label = layout_name(archive, layout)
            master = related_part(archive, layout, "slideMaster") if layout else None
            theme = related_part(archive, master, "theme") if master else None
            slide_theme_map[f"slide{slide_number(slide)}"] = theme
            slide_layout_map[f"slide{slide_number(slide)}"] = {
                "part": layout,
                "name": layout_label,
            }
            if theme:
                theme_parts[theme] += 1
                theme_root = ET.fromstring(archive.read(theme))
                scheme = theme_root.find(".//" + A + "clrScheme")
                theme_name = theme_root.attrib.get("name", "")
                scheme_name = scheme.attrib.get("name", "") if scheme is not None else ""
                theme_names[f"{theme_name} / {scheme_name}"] += 1
                color_scheme_names[scheme_name] += 1
                slide_theme_name_map[f"slide{slide_number(slide)}"] = theme_name
                slide_color_scheme_map[f"slide{slide_number(slide)}"] = scheme_name
            else:
                slide_theme_name_map[f"slide{slide_number(slide)}"] = ""
                slide_color_scheme_map[f"slide{slide_number(slide)}"] = ""

        richinfo_theme_ok = bool(color_scheme_names) and all(
            "彩讯科技" in name for name in color_scheme_names
        )
        required_layout_ok = (
            required_layout is None
            or required_layout in set(available_layouts.values())
        )
        required_color_scheme_ok = (
            required_color_scheme is None
            or (
                bool(slide_color_scheme_map)
                and all(
                    scheme == required_color_scheme
                    for scheme in slide_color_scheme_map.values()
                )
            )
        )
        required_theme_name_ok = (
            required_theme_name is None
            or (
                bool(slide_theme_name_map)
                and all(
                    name == required_theme_name
                    for name in slide_theme_name_map.values()
                )
            )
        )
        content_layout_failures = []
        for content_slide in sorted(content_slides):
            slide_label = f"slide{content_slide}"
            actual = slide_layout_map.get(slide_label)
            if actual is None:
                content_layout_failures.append(
                    {
                        "slide": slide_label,
                        "expected": required_layout,
                        "actual": "missing slide",
                    }
                )
            elif actual["name"] != required_layout:
                content_layout_failures.append(
                    {
                        "slide": slide_label,
                        "expected": required_layout,
                        "actual": actual["name"],
                    }
                )
        strict_failures = []
        if corrupt_member:
            strict_failures.append(f"corrupt ZIP member: {corrupt_member}")
        if compliance_summary["direct_rgb"]:
            strict_failures.append(
                "new or changed slide objects contain fixed RGB colors"
            )
        if compliance_summary["sizes_below_12"]:
            strict_failures.append("explicit font sizes below 12 pt")
        if compliance_summary["noninteger_sizes"]:
            strict_failures.append("explicit non-integer font sizes")
        if compliance_summary["disallowed_fonts"]:
            strict_failures.append("explicit fonts outside the standard set")
        if template_integrity is not None and not template_integrity["ok"]:
            strict_failures.append(
                "template master, layout, theme, or related protected parts changed"
            )
        if require_richinfo_theme and not richinfo_theme_ok:
            strict_failures.append("slides are not all bound to the 彩讯科技 theme")
        if not required_layout_ok:
            strict_failures.append(
                f"required layout is missing: {required_layout}"
            )
        if not required_color_scheme_ok:
            strict_failures.append(
                f"slides are not all bound to color scheme: {required_color_scheme}"
            )
        if not required_theme_name_ok:
            strict_failures.append(
                f"slides are not all bound to theme name: {required_theme_name}"
            )
        if content_layout_failures:
            strict_failures.append(
                f"content slides are not all bound to layout: {required_layout}"
            )

        return {
            "file": str(path.resolve()),
            "valid_zip": corrupt_member is None,
            "slide_count": len(slides),
            "available_layouts": available_layouts,
            "slide_layout_map": slide_layout_map,
            "slide_theme_map": slide_theme_map,
            "slide_theme_name_map": slide_theme_name_map,
            "slide_color_scheme_map": slide_color_scheme_map,
            "theme_names": dict(theme_names),
            "color_scheme_names": dict(color_scheme_names),
            "richinfo_theme_ok": richinfo_theme_ok,
            "require_richinfo_theme": require_richinfo_theme,
            "required_layout": required_layout,
            "required_layout_ok": required_layout_ok,
            "required_color_scheme": required_color_scheme,
            "required_color_scheme_ok": required_color_scheme_ok,
            "required_theme_name": required_theme_name,
            "required_theme_name_ok": required_theme_name_ok,
            "content_slides": sorted(content_slides),
            "content_layout_failures": content_layout_failures,
            "required_template": (
                str(required_template.resolve())
                if required_template is not None
                else None
            ),
            "template_integrity": template_integrity,
            "inherited_slide_parts": sorted(inherited_slide_parts),
            "slide_shape_summary": slide_shape_summary,
            "rounded_rectangle_review_notes": rounded_rectangle_review_notes,
            "slide_summary": slide_summary,
            "compliance_summary": compliance_summary,
            "supporting_parts_summary": supporting_parts_summary,
            "summary": merged,
            "strict_failures": strict_failures,
        }


def print_human(report: dict) -> None:
    summary = report["compliance_summary"]
    slide_summary = report["slide_summary"]
    print(f"File: {report['file']}")
    print(f"Slides: {report['slide_count']} | ZIP valid: {report['valid_zip']}")
    print(
        "Available layouts: "
        + json.dumps(report["available_layouts"], ensure_ascii=False)
    )
    print(
        "Slide layout map: "
        + json.dumps(report["slide_layout_map"], ensure_ascii=False)
    )
    print(f"Themes: {json.dumps(report['theme_names'], ensure_ascii=False)}")
    if report["template_integrity"] is not None:
        print(
            "Template protected parts unchanged: "
            f"{report['template_integrity']['ok']}"
        )
        print(
            "Inherited unchanged slides: "
            + json.dumps(report["inherited_slide_parts"], ensure_ascii=False)
        )
    print(
        "Color schemes: "
        + json.dumps(report["color_scheme_names"], ensure_ascii=False)
    )
    print(f"彩讯科技 theme: {report['richinfo_theme_ok']}")
    print(f"Fixed RGB colors: {json.dumps(summary['direct_rgb'], ensure_ascii=False)}")
    print(f"Theme color references: {sum(summary['scheme_colors'].values())}")
    print(f"Font sizes below 12 pt: {len(summary['sizes_below_12'])}")
    print(f"Non-integer font sizes: {len(summary['noninteger_sizes'])}")
    print(f"Non-standard explicit fonts: {json.dumps(summary['disallowed_fonts'], ensure_ascii=False)}")
    print(f"Non-middle text autoshapes: {len(summary['nonmiddle_text_shapes'])}")
    print(f"Paragraphs with nonzero before/after spacing: {summary['nonzero_paragraph_spacing']}")
    print(f"Text boxes / text autoshapes: {summary['text_boxes']} / {summary['text_autoshapes']}")
    print(
        "Straight / rounded rectangles: "
        f"{slide_summary['straight_rectangles']} / {slide_summary['rounded_rectangles']}"
    )
    if report["rounded_rectangle_review_notes"]:
        print(
            "Rounded rectangle hierarchy review: "
            + json.dumps(
                report["rounded_rectangle_review_notes"], ensure_ascii=False
            )
        )
    else:
        print("Rounded rectangle hierarchy review: no rounded rectangles")
    if report["content_layout_failures"]:
        print(
            "Content layout failures: "
            + json.dumps(report["content_layout_failures"], ensure_ascii=False)
        )
    elif report["content_slides"]:
        print("Content layout failures: none")
    if report["strict_failures"]:
        print("Strict failures:")
        for failure in report["strict_failures"]:
            print(f"- {failure}")
    else:
        print("Strict failures: none")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("presentation", type=Path)
    parser.add_argument("--require-richinfo-theme", action="store_true")
    parser.add_argument("--require-layout")
    parser.add_argument("--require-color-scheme")
    parser.add_argument("--require-theme-name")
    parser.add_argument(
        "--require-template",
        type=Path,
        help="baseline template whose master, layout, theme, and related protected parts must remain byte-identical",
    )
    parser.add_argument(
        "--content-slides",
        help='comma-separated slide numbers or ranges, for example "2-5,8"',
    )
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args()

    try:
        content_slides = (
            parse_slide_spec(args.content_slides) if args.content_slides else set()
        )
        if content_slides and not args.require_layout:
            parser.error("--content-slides requires --require-layout")
        report = audit(
            args.presentation,
            args.require_richinfo_theme,
            args.require_layout,
            args.require_color_scheme,
            args.require_theme_name,
            content_slides,
            args.require_template,
        )
    except (
        FileNotFoundError,
        ValueError,
        zipfile.BadZipFile,
        ET.ParseError,
    ) as error:
        print(f"Audit failed: {error}", file=sys.stderr)
        return 3

    if args.as_json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_human(report)
    return 2 if args.strict and report["strict_failures"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
