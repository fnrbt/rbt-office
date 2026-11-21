#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-text-zoom.xlsx, a synthetic workbook (not saved by Excel) with the
cases of how Excel 2019 sizes and places a cell's text, at 100 % and at zooms (checked in Excel; xlsx-browser check
`grid-text-excel-parity`; docs/excel-parity-tracker.md, "Cell text at a zoom"):

- "Z100", "Z80", "Z85", "Z67" (Calibri 11 at 100 %, 80 %, 85 %, 67 %): B1 a pangram, left; B2 the same, right;
  B3 twenty zeros; B4 the pangram centred; B5 the pangram bold. Column B is 425 px. Excel draws the text at
  15 px (100 %), 12 px (80 % and 85 % alike), in Small Fonts' 10-px face at 67 %; twenty zeros are 140, 120, 120
  and 120 px; a left-aligned line starts 2 px into the cell, a right-aligned one ends 3 px before its end.
- "Arial" (100 %): B1 Arial 10 (13 px, not 13.33: twenty zeros are 140 px), B2 Arial 7 (9 px: Small Fonts).
- "Numbers": Calibri 11 at 75 % in columns of 40, 54 and 68 px (30, 41 and 51 on screen): 1234567 (General),
  12345678 (format "0"). Excel shows `#####`, `1E+06`, `1234567` and `#####`, `#######`, `#########`; at
  100 % `1E+06`, `1234567`, `1234567` and `#####`, `#######`, `12345678`.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-text-zoom-fixture.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-text-zoom.xlsx')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
PANGRAM = 'The quick brown fox jumps over the lazy dog'

STYLES = DECL + (
    f'<styleSheet xmlns="{MAIN}">'
    '<fonts count="4"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><sz val="10"/><name val="Arial"/><family val="2"/></font>'
    '<font><sz val="7"/><name val="Arial"/><family val="2"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="8">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'                                                     # 0
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>'   # 1
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>'  # 2
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>' # 3
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left"/></xf>'  # 4
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left"/></xf>'  # 5
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left"/></xf>'  # 6
    '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'                                 # 7 "0"
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def inline(ref, style, text):
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def sheet(zoom, rows, cols, selected=False):
    return DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if selected else ""}'
        f' zoomScale="{zoom}" zoomScaleNormal="{zoom}" workbookViewId="0"/></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols><sheetData>{"".join(rows)}</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


WIDE = '<col min="2" max="2" width="60.7109375" customWidth="1"/>'
TEXT_ROWS = [f'<row r="1">{inline("B1", 1, PANGRAM)}</row>', f'<row r="2">{inline("B2", 2, PANGRAM)}</row>',
             f'<row r="3">{inline("B3", 1, "0" * 20)}</row>', f'<row r="4">{inline("B4", 3, PANGRAM)}</row>',
             f'<row r="5">{inline("B5", 4, PANGRAM)}</row>']
ARIAL_ROWS = [f'<row r="1">{inline("B1", 5, "0" * 20)}</row>', f'<row r="2">{inline("B2", 6, "0" * 20)}</row>']
NUMBER_COLS = ''.join(f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>' for c, w in ((2, 5.7109375), (3, 7.7109375), (4, 9.7109375)))
NUMBER_ROWS = ['<row r="1">' + ''.join(f'<c r="{c}1"><v>1234567</v></c>' for c in 'BCD') + '</row>',
               '<row r="2">' + ''.join(f'<c r="{c}2" s="7"><v>12345678</v></c>' for c in 'BCD') + '</row>']
SHEETS = [('Z100', sheet(100, TEXT_ROWS, WIDE, True)), ('Z80', sheet(80, TEXT_ROWS, WIDE)), ('Z85', sheet(85, TEXT_ROWS, WIDE)),
          ('Z67', sheet(67, TEXT_ROWS, WIDE)), ('Arial', sheet(100, ARIAL_ROWS, WIDE)), ('Numbers', sheet(75, NUMBER_ROWS, NUMBER_COLS))]

FILES = {
    '[Content_Types].xml': DECL + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                  for i in range(len(SHEETS)))
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'),
    '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
                               + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (name, _) in enumerate(SHEETS))
                               + '</sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': DECL + (
        f'<Relationships xmlns="{PKG}">'
        + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(SHEETS)))
        + f'<Relationship Id="rId{len(SHEETS) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
    'xl/styles.xml': STYLES,
}
for i, (_, xml) in enumerate(SHEETS):
    FILES[f'xl/worksheets/sheet{i + 1}.xml'] = xml

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in FILES.items():
        info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, text)
print(OUT)
