#!/usr/bin/env python3
"""Read-only grouping/line-spacing audit; rendering evidence is required for soft wraps."""
import argparse
import hashlib
import json
import posixpath
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def paragraphs(body):
    return [p for p in body.findall('a:p', NS)
            if ''.join(p.itertext()).strip() or p.find('a:br', NS) is not None]


def spacing(p):
    node = p.find('a:pPr/a:lnSpc/a:spcPct', NS)
    return int(node.get('val')) / 100000 if node is not None else None


def inspect(root, page, measurements=None):
    issues, pending, limitations = [], [], []
    text_count = 0
    for obj in root.iter():
        kind = obj.tag.split('}')[-1]
        if kind not in {'sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'}:
            continue
        nv = next((c for c in obj if c.tag.split('}')[-1].startswith('nv')), None)
        if nv is None:
            continue
        ident = nv.find('p:cNvPr', NS)
        if ident is None:
            continue
        sid, name = ident.get('id'), ident.get('name', '')
        base = {'slide': page, 'shape_id': sid, 'name': name}
        if any(n.get('noGrp') in {'1', 'true'} for n in nv.iter()):
            issues.append(dict(base, code='grouping_locked'))
        if nv.find('.//p:ph', NS) is not None:
            limitations.append(dict(base, code='placeholder_grouping_limit'))
        if obj.find('.//a:tbl', NS) is not None and kind == 'graphicFrame':
            limitations.append(dict(base, code='table_grouping_limit'))
        bodies = []
        body = obj.find('p:txBody', NS)
        if body is not None:
            bodies.append((sid, body))
        if kind == 'graphicFrame':
            for row, tr in enumerate(obj.findall('.//a:tbl/a:tr', NS), 1):
                for col, tc in enumerate(tr.findall('a:tc', NS), 1):
                    b = tc.find('a:txBody', NS)
                    if b is not None:
                        bodies.append((f'{sid}/r{row}c{col}', b))
        for key, body in bodies:
            ps = paragraphs(body)
            if not ps:
                continue
            text_count += 1
            detail = dict(base, text_key=key)
            measured = (measurements or {}).get(f'{page}:{key}')
            if measured is not None:
                lines = measured.get('lines')
                effective = measured.get('paragraphs')
                if not isinstance(lines, int) or isinstance(lines, bool) or lines < 1 or not isinstance(effective, list) or len(effective) != len(ps):
                    pending.append(dict(detail, code='invalid_layout_measurement'))
                    continue
                expected = 1.2 if lines > 1 else 1.0
                for index, fmt in enumerate(effective, 1):
                    if not isinstance(fmt, dict) or any(not isinstance(fmt.get(k), (int, float)) for k in ('within', 'before', 'after')):
                        pending.append(dict(detail, paragraph=index, code='invalid_paragraph_measurement'))
                        continue
                    if fmt.get('multiple') is not True or abs(fmt['within'] - expected) > .005:
                        issues.append(dict(detail, paragraph=index, code='wrong_rendered_line_spacing', expected=expected, actual=fmt))
                    if abs(fmt['before']) > .005 or abs(fmt['after']) > .005:
                        issues.append(dict(detail, paragraph=index, code='nonzero_paragraph_spacing'))
            else:
                # This lower bound proves multi-line, but cannot prove single-line.
                minimum = len(ps) + sum(len(p.findall('a:br', NS)) for p in ps)
                for index, p in enumerate(ps, 1):
                    value = spacing(p)
                    if minimum > 1 and value is not None and abs(value - 1.2) > .005:
                        issues.append(dict(detail, paragraph=index, code='wrong_explicit_multiline_spacing', expected=1.2, actual=value))
                    for tag in ('spcBef', 'spcAft'):
                        node = p.find(f'a:pPr/a:{tag}', NS)
                        if node is not None and any(float(x.get('val', '0')) != 0 for x in node):
                            issues.append(dict(detail, paragraph=index, code='nonzero_paragraph_spacing'))
                pending.append(dict(detail, code='needs_rendered_lines_and_effective_spacing'))
    return {'issues': issues, 'pending': pending, 'limitations': limitations, 'text_count': text_count}


def slide_parts(z):
    root = ET.fromstring(z.read('ppt/presentation.xml'))
    rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    targets = {r.get('Id'): r.get('Target') for r in rels if r.get('TargetMode') != 'External'}
    result = []
    for node in root.findall('p:sldIdLst/p:sldId', NS):
        target = targets[node.get('{' + NS['r'] + '}id')]
        result.append(posixpath.normpath(target.lstrip('/') if target.startswith('/') else 'ppt/' + target))
    return result


def audit(path, evidence=None, selected=None):
    digest = hashlib.sha256(Path(path).read_bytes()).hexdigest()
    measurements = None
    if evidence is not None:
        if evidence.get('sha256', '').lower() != digest:
            raise ValueError('layout evidence SHA256 does not match presentation')
        if evidence.get('source') != 'PowerPoint.TextRange' or not isinstance(evidence.get('measurements'), dict):
            raise ValueError('invalid PowerPoint layout evidence')
        measurements = evidence['measurements']
    report = {'sha256': digest, 'issues': [], 'pending': [], 'limitations': [], 'text_count': 0}
    with zipfile.ZipFile(path) as z:
        parts = slide_parts(z)
        if selected and not selected.issubset(set(range(1, len(parts) + 1))):
            raise ValueError('selected slide does not exist')
        for index, part in enumerate(parts, 1):
            if selected and index not in selected:
                continue
            found = inspect(ET.fromstring(z.read(part)), index, measurements)
            for key in ('issues', 'pending', 'limitations'):
                report[key].extend(found[key])
            report['text_count'] += found['text_count']
    report['pass'] = not report['issues'] and not report['pending']
    report['scope'] = sorted(selected) if selected else list(range(1, len(parts) + 1))
    report['note'] = 'Does not test interactive grouping, overflow, visual quality, charts or SmartArt internals.'
    return report


def parse_slides(value):
    result = set()
    for item in value.split(','):
        edges = [int(n) for n in item.split('-')]
        if len(edges) > 2 or min(edges) < 1 or edges[-1] < edges[0]:
            raise ValueError('invalid slide range')
        result.update(range(edges[0], edges[-1] + 1))
    return result


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('presentation', type=Path)
    p.add_argument('--layout-evidence', type=Path)
    p.add_argument('--slides')
    p.add_argument('--strict', action='store_true')
    p.add_argument('--json', action='store_true')
    args = p.parse_args()
    try:
        evidence = json.loads(args.layout_evidence.read_text(encoding='utf-8-sig')) if args.layout_evidence else None
        report = audit(args.presentation, evidence, parse_slides(args.slides) if args.slides else None)
    except (OSError, ValueError, KeyError, TypeError, AttributeError, zipfile.BadZipFile, ET.ParseError) as error:
        print(f'Audit failed: {error}', file=sys.stderr)
        return 3
    print(json.dumps(report, ensure_ascii=False, indent=2) if args.json else
          f"Pass: {report['pass']} | issues: {len(report['issues'])} | pending: {len(report['pending'])} | limitations: {len(report['limitations'])}")
    return 2 if args.strict and not report['pass'] else 0


if __name__ == '__main__':
    sys.exit(main())
