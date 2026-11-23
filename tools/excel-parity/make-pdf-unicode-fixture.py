#!/usr/bin/env python3
"""Writes the deterministic worksheet control for PDF text extraction of the three
hyphen-looking characters Excel/Word can print: ASCII hyphen-minus U+002D,
Unicode hyphen U+2010, and minus sign U+2212.  Calibri draws at least the first
two with one glyph, so a PDF whose ToUnicode map has one entry per glyph cannot
say which character a run held.  The fixture also carries a formatted negative
number and a date whose format code supplies ASCII hyphens, and runs of one /
two / three characters in one cell and in adjacent rich-text runs.

This is public generated data (no Excel-saved file); running it twice writes
the same bytes.  It is read by tests/native-pdf-unicode.fsx.

    python3 tools/excel-parity/make-pdf-unicode-fixture.py OUT.xlsx
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

ASCII_HYPHEN = '\u002d'
UNICODE_HYPHEN = '\u2010'
MINUS_SIGN = '\u2212'
MIXED = ASCII_HYPHEN + UNICODE_HYPHEN + MINUS_SIGN + ASCII_HYPHEN + UNICODE_HYPHEN + MINUS_SIGN

# (label, cell XML for column B)
TEXT_CASES = [
    ('ascii', f'<c r="B2" t="inlineStr"><is><t>{ASCII_HYPHEN}</t></is></c>'),
    ('u2010', f'<c r="B3" t="inlineStr"><is><t>{UNICODE_HYPHEN}</t></is></c>'),
    ('u2212', f'<c r="B4" t="inlineStr"><is><t>{MINUS_SIGN}</t></is></c>'),
    ('repeat-ascii', f'<c r="B5" t="inlineStr"><is><t>{ASCII_HYPHEN * 3}</t></is></c>'),
    ('repeat-u2010', f'<c r="B6" t="inlineStr"><is><t>{UNICODE_HYPHEN * 3}</t></is></c>'),
    ('repeat-u2212', f'<c r="B7" t="inlineStr"><is><t>{MINUS_SIGN * 3}</t></is></c>'),
    ('mixed', f'<c r="B8" t="inlineStr"><is><t>{MIXED}</t></is></c>'),
    ('negative-number', '<c r="B9" s="1"><v>-1234.5</v></c>'),
    ('date', '<c r="B10" s="2"><v>45322</v></c>'),
    # Six rich-text runs, so the exporter cannot treat the cell as one text run.
    ('mixed-runs', '<c r="B11" t="inlineStr"><is>'
     + ''.join(f'<r><t>{c}</t></r>' for c in MIXED)
     + '</is></c>'),
    # The same character repeated, each occurrence its own run.
    ('repeat-runs', '<c r="B12" t="inlineStr"><is>'
     + ''.join(f'<r><t>{c}</t></r>' for c in ASCII_HYPHEN * 3)
     + '</is></c>'),
    ('repeat-runs-u2010', '<c r="B13" t="inlineStr"><is>'
     + ''.join(f'<r><t>{c}</t></r>' for c in UNICODE_HYPHEN * 3)
     + '</is></c>'),
]


def build(out):
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    xfs = [
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
        '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    ]
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    # One row per case: the row numbers inside TEXT_CASES are fixed.
    rows = ('<row r="1"><c r="A1" t="inlineStr"><is><t>case</t></is></c>'
            '<c r="B1" t="inlineStr"><is><t>value</t></is></c></row>'
            + ''.join(
                f'<row r="{i + 2}">'
                f'<c r="A{i + 2}" t="inlineStr"><is><t>{escape(label)}</t></is></c>'
                f'{cell}</row>'
                for i, (label, cell) in enumerate(TEXT_CASES)))
    sheet = DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols><col min="1" max="1" width="20" customWidth="1"/>'
        '<col min="2" max="2" width="22" customWidth="1"/></cols>'
        f'<sheetData>{rows}</sheetData>'
        '<printOptions gridLines="0"/>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '<pageSetup paperSize="1" orientation="portrait" scale="100"/></worksheet>')
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{CT}.worksheet+xml"/></Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews>'
            '<sheets><sheet name="Hyphens" sheetId="1" r:id="rId1"/></sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            f'<Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
        'xl/worksheets/sheet1.xml': sheet,
    }
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(out)


if __name__ == '__main__':
    build(sys.argv[1])
