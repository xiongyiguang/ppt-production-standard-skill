"""In-memory regression fixtures; no files are deleted or presentations modified."""
import hashlib
import io
import unittest
import zipfile
from xml.etree import ElementTree as ET
from audit_editing_quality import NS, inspect, slide_parts, audit, parse_slides
from unittest.mock import patch


def slide(body, lock='', placeholder='', wrapper=False):
    shape = f'<p:sp><p:nvSpPr><p:cNvPr id="7" name="body"/><p:cNvSpPr>{lock}</p:cNvSpPr><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:txBody>{body}</p:txBody></p:sp>'
    if wrapper:
        shape = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="group"/></p:nvGrpSpPr>' + shape + '</p:grpSp>'
    return ET.fromstring(f'<p:sld xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:cSld><p:spTree>{shape}</p:spTree></p:cSld></p:sld>')


def para(value=100000, br=False):
    pr = f'<a:pPr><a:lnSpc><a:spcPct val="{value}"/></a:lnSpc></a:pPr>' if value else ''
    return '<a:p>' + pr + '<a:r><a:t>Example</a:t></a:r>' + ('<a:br/><a:r><a:t>next</a:t></a:r>' if br else '') + '</a:p>'


def evidence(lines=1, value=1.0, count=1, before=0):
    return {'1:7': {'lines': lines, 'paragraphs': [dict(multiple=True, within=value, before=before, after=0) for _ in range(count)]}}


class EditingQualityTests(unittest.TestCase):
    def test_explicit_break(self):
        r = inspect(slide(para(br=True)), 1)
        self.assertEqual(r['issues'][0]['code'], 'wrong_explicit_multiline_spacing')

    def test_paragraph_multiline(self):
        self.assertEqual(len(inspect(slide(para() * 2), 1)['issues']), 2)

    def test_soft_wrap(self):
        r = inspect(slide(para()), 1, evidence(2))
        self.assertEqual(r['issues'][0]['code'], 'wrong_rendered_line_spacing')

    def test_correct_multiline(self):
        r = inspect(slide(para(120000)), 1, evidence(2, 1.2))
        self.assertFalse(r['issues'] or r['pending'])

    def test_singleline(self):
        self.assertFalse(inspect(slide(para()), 1, evidence())['issues'])

    def test_singleline_wrong_12(self):
        self.assertTrue(inspect(slide(para(120000)), 1, evidence(1, 1.2))['issues'])

    def test_inheritance_requires_measurement(self):
        self.assertTrue(inspect(slide(para(None)), 1)['pending'])

    def test_effective_inheritance(self):
        r = inspect(slide(para(None)), 1, evidence(2, 1.2))
        self.assertFalse(r['pending'] or r['issues'])

    def test_nested_lock(self):
        r = inspect(slide(para(), '<a:spLocks noGrp="1"/>', wrapper=True), 1)
        self.assertEqual([i['shape_id'] for i in r['issues']], ['7'])

    def test_placeholder_limitation(self):
        self.assertEqual(inspect(slide(para(), placeholder='<p:ph/>'), 1)['limitations'][0]['code'], 'placeholder_grouping_limit')

    def test_nonzero_paragraph_spacing(self):
        self.assertEqual(inspect(slide(para()), 1, evidence(before=5))['issues'][0]['code'], 'nonzero_paragraph_spacing')

    def test_incomplete_measurements(self):
        self.assertTrue(inspect(slide(para() * 2), 1, evidence())['pending'])

    def test_table_cells(self):
        root = ET.fromstring(f'<p:sld xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody>{para()}</a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:sld>')
        r = inspect(root, 1)
        self.assertEqual(r['limitations'][0]['code'], 'table_grouping_limit')
        self.assertEqual(r['pending'][0]['text_key'], '9/r1c1')

    def test_order_not_filename(self):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w') as z:
            z.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{NS["p"]}" xmlns:r="{NS["r"]}"><p:sldIdLst><p:sldId r:id="r2"/><p:sldId r:id="r1"/></p:sldIdLst></p:presentation>')
            z.writestr('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="r1" Target="slides/slide1.xml"/><Relationship Id="r2" Target="/ppt/slides/slide4.xml"/></Relationships>')
        with zipfile.ZipFile(buf) as z:
            self.assertEqual(slide_parts(z), ['ppt/slides/slide4.xml', 'ppt/slides/slide1.xml'])

    def test_stale_evidence_rejected(self):
        with patch('pathlib.Path.read_bytes', return_value=b'current'):
            with self.assertRaisesRegex(ValueError, 'SHA256'):
                audit('unused.pptx', {'sha256': hashlib.sha256(b'old').hexdigest()})

    def test_slide_scope(self):
        self.assertEqual(parse_slides('1,3-5'), {1, 3, 4, 5})
        with self.assertRaises(ValueError):
            parse_slides('5-3')


if __name__ == '__main__':
    unittest.main()
