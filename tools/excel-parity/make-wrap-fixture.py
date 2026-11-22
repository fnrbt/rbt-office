#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-wrap-budget.xlsx, a synthetic workbook (not saved by Excel): where Excel
2019 breaks a wrapped cell's text on screen, and how many lines it fits a row to when it opens the file (checked in
Excel on the VM: screenshots, `shoot-sheets.ps1 -KeepZoom`, and `Rows(r).Height`; docs/excel-parity-tracker.md,
"Wrapped text").

Each sheet has one text, in columns B.. of widths a pixel apart (at 100 %), around the width of its first line:
row 1 tall (custom height, top aligned) to see the lines drawn, row 2 of no height for Excel to fit on opening;
from row 6 the text in one column a row (B6, C7 ...), each row fitted on opening to its one cell.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-wrap-fixture.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-wrap-budget.xlsx')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

FONTS = ['<sz val="11"/><name val="Calibri"/><family val="2"/>',
         '<b/><sz val="11"/><name val="Calibri"/><family val="2"/>',
         '<sz val="10"/><name val="Arial"/><family val="2"/>',
         '<sz val="11"/><name val="Arial"/><family val="2"/>']
# (sheet, zoom, font index, text, first column's width in pixels at 100 %, columns)
SHEETS = [
    ('B100', 100, 1, 'Releasing Cost per Unit', 84, 18),
    ('R100', 100, 0, 'The quick brown fox jumps', 96, 18),
    ('A100', 100, 2, 'The quick brown fox jumps', 88, 18),
    ('L100', 100, 3, 'The quick brown fox jumps', 96, 18),
    ('B85', 85, 1, 'Releasing Cost per Unit', 78, 22),
    ('R85', 85, 0, 'The quick brown fox jumps', 98, 22),
    ('R70', 70, 0, 'The quick brown fox jumps', 120, 22),
]


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def width_of(px):
    """A column width Excel shows as `px` pixels at a maximum digit width of 7 (Calibri 11)."""
    return round(px / 7, 6)


def sheet_xml(zoom, font, text, first, count, selected):
    cols = ''.join(f'<col min="{i + 2}" max="{i + 2}" width="{width_of(first + i)}" customWidth="1"/>' for i in range(count))
    style = font + 1
    cells1 = ''.join(f'<c r="{col_letter(i + 2)}1" s="{style}" t="inlineStr"><is><t>{text}</t></is></c>' for i in range(count))
    cells2 = ''.join(f'<c r="{col_letter(i + 2)}2" s="{style}" t="inlineStr"><is><t>{text}</t></is></c>' for i in range(count))
    labels = ''.join(f'<c r="{col_letter(i + 2)}4"><v>{first + i}</v></c>' for i in range(count))
    diagonal = ''.join(f'<row r="{6 + i}"><c r="{col_letter(i + 2)}{6 + i}" s="{style}" t="inlineStr"><is><t>{text}</t></is></c></row>'
                       for i in range(count))
    return DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if selected else ""}'
        f' zoomScale="{zoom}" zoomScaleNormal="{zoom}" workbookViewId="0"/></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols><sheetData>'
        f'<row r="1" ht="60" customHeight="1">{cells1}</row><row r="2">{cells2}</row><row r="4">{labels}</row>{diagonal}'
        '</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


def build():
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'] + [
        f'<xf numFmtId="0" fontId="{i}" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">'
        '<alignment vertical="top" wrapText="1"/></xf>' for i in range(len(FONTS))]
    thin = '<left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top/><bottom/><diagonal/>'
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<fonts count="{len(FONTS)}">' + ''.join(f'<font>{f}</font>' for f in FONTS) + '</fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        f'<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>{thin}</border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    sheets = [(name, sheet_xml(zoom, font, text, first, count, i == 0)) for i, (name, zoom, font, text, first, count) in enumerate(SHEETS)]
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
