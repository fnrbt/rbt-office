#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-small-fonts.xlsx, a synthetic workbook (not saved by Excel): from what size on
screen Excel 2019 draws a cell's text in its own TrueType font rather than in Windows' raster Small Fonts (checked in
Excel on the VM: pictures with `shoot-sheets.ps1 -KeepZoom`, a raster font's text has no grey, anti-aliased, pixels;
docs/excel-parity-tracker.md, "Cell text at a zoom").

One sheet a zoom ("Z50" .. "Z100"); row r a font and size (FACES x SIZES), column B a text, column C a thin bottom
border on every row (to find the rows in a picture).

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-small-fonts-fixture.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-small-fonts.xlsx')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

FACES = ['Calibri', 'Arial']
SIZES = [7, 7.5, 8, 8.5, 9, 10, 10.5, 11, 12, 13, 14, 15, 16]
ZOOMS = [50, 53, 55, 56, 57, 58, 60, 62, 64, 65, 66, 67, 68, 70, 72, 73, 75, 76, 78, 79, 80, 85, 100]
TEXT = 'Levered Cash Flow 0123'
CASES = [(face, size) for face in FACES for size in SIZES]


def build():
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>'] + [f'<sz val="{s}"/><name val="{f}"/><family val="2"/>' for f, s in CASES]
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>']
    xfs += [f'<xf numFmtId="0" fontId="{i + 1}" fillId="0" borderId="0" xfId="0" applyFont="1"/>' for i in range(len(CASES))]
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left/><right/><top/><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    rows = ''.join(f'<row r="{i + 1}" ht="24" customHeight="1"><c r="B{i + 1}" s="{i + 2}" t="inlineStr"><is><t>{TEXT}</t></is></c>'
                   f'<c r="C{i + 1}" s="1"/></row>' for i in range(len(CASES)))
    cols = '<cols><col min="2" max="2" width="40" customWidth="1"/></cols>'
    sheets = [(f'Z{z}', DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if i == 0 else ""}'
        f' zoomScale="{z}" zoomScaleNormal="{z}" workbookViewId="0"/></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="15"/>{cols}<sheetData>{rows}</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'))
              for i, z in enumerate(ZOOMS)]
    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>' for i in range(len(sheets)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
                                   + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (name, _) in enumerate(sheets))
                                   + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
            + f'<Relationship Id="rId{len(sheets) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, (_, xml) in enumerate(sheets):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(OUT)


if __name__ == '__main__':
    build()
