#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-font-faces.xlsx, a synthetic workbook (not saved by Excel): how Excel 2019
draws fonts the grid's tables did not have (checked in Excel on the VM: `Range.Text` with probe-display.vbs and pictures
with `shoot-sheets.ps1 -KeepZoom`; docs/excel-parity-tracker.md, "Cell text at a zoom"):

- "Black" (100 %): Segoe UI Black 9, 10, 11 and Segoe UI 10 (rows 1-4), $1,592,811 (format "$"#,##0) in columns B..K of
  40 .. 85 px - which show it and which `#`s - and in M (400 px) twenty zeros, to measure.
- "Z57" (57 %): Aptos Narrow 14 and Calibri 14 (7.98 points on screen: Small Fonts for Calibri), a text and twenty
  zeros in a 400-px column; Aptos Narrow 11 (6.27 points).
- "Room" (100 %): a number too wide for any of columns B.. (40 .. 85 px, a pixel apart) in Segoe UI Black 9 - 14, Arial
  Black 11, Segoe UI 11 bold and regular, Calibri 11 bold and regular: how many #s Excel fills each with; rows 12-15
  numbers that fit from some column on; "Room50" ..
  "Room200" the same at 50, 60, 67, 75, 85, 125, 150 and 200 %.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-font-faces-fixture.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-font-faces.xlsx')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

FONTS = ['<sz val="11"/><name val="Calibri"/><family val="2"/>',
         '<sz val="9"/><name val="Segoe UI Black"/><family val="2"/>',
         '<sz val="10"/><name val="Segoe UI Black"/><family val="2"/>',
         '<sz val="11"/><name val="Segoe UI Black"/><family val="2"/>',
         '<sz val="10"/><name val="Segoe UI"/><family val="2"/>',
         '<sz val="14"/><name val="Aptos Narrow"/><family val="2"/>',
         '<sz val="14"/><name val="Calibri"/><family val="2"/>',
         '<sz val="11"/><name val="Aptos Narrow"/><family val="2"/>',
         '<sz val="12"/><name val="Segoe UI Black"/><family val="2"/>',
         '<sz val="14"/><name val="Segoe UI Black"/><family val="2"/>',
         '<sz val="11"/><name val="Arial Black"/><family val="2"/>',
         '<b/><sz val="11"/><name val="Segoe UI"/><family val="2"/>',
         '<b/><sz val="11"/><name val="Calibri"/><family val="2"/>',
         '<sz val="11"/><name val="Segoe UI"/><family val="2"/>']
ROOM_FONTS = [1, 2, 3, 8, 9, 10, 11, 12, 13, 0]
TEXT = 'Levered Cash Flow 0123456789'


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def inline(ref, style, text):
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{text}</t></is></c>'


def sheet(zoom, cols, rows, selected=False):
    return DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if selected else ""}'
        f' zoomScale="{zoom}" zoomScaleNormal="{zoom}" workbookViewId="0"/></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols><sheetData>{"".join(rows)}</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


def build():
    # cellXfs: per font i, 2i+0 general, 2i+1 "$"#,##0 (numFmt 164)
    xfs = []
    for i in range(len(FONTS)):
        xfs.append(f'<xf numFmtId="0" fontId="{i}" fillId="0" borderId="0" xfId="0" applyFont="1"/>')
        xfs.append(f'<xf numFmtId="164" fontId="{i}" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>')
    # Then #,##0_) and #,##0_W (a width placeholder) in Calibri 11 and Segoe UI Black 11.
    placeholder = len(xfs)
    for fmt in (165, 166):
        for font in (0, 3):
            xfs.append(f'<xf numFmtId="{fmt}" fontId="{font}" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>')
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}"><numFmts count="3"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0"/>'
        '<numFmt numFmtId="165" formatCode="#,##0_)"/><numFmt numFmtId="166" formatCode="#,##0_W"/></numFmts>'
        f'<fonts count="{len(FONTS)}">' + ''.join(f'<font>{f}</font>' for f in FONTS) + '</fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    black_cols = ''.join(f'<col min="{i + 2}" max="{i + 2}" width="{round((40 + 5 * i) / 7, 6)}" customWidth="1"/>' for i in range(10))
    black_cols += '<col min="13" max="13" width="57.142857" customWidth="1"/>'
    black_rows = []
    for r, font in enumerate([1, 2, 3, 4]):
        cells = ''.join(f'<c r="{col_letter(i + 2)}{r + 1}" s="{2 * font + 1}"><v>1592811</v></c>' for i in range(10))
        black_rows.append(f'<row r="{r + 1}" ht="21" customHeight="1">{cells}{inline(f"M{r + 1}", 2 * font, "0" * 20)}</row>')
    z_cols = '<col min="2" max="2" width="57.142857" customWidth="1"/><col min="3" max="3" width="57.142857" customWidth="1"/>'
    z_rows = []
    for r, font in enumerate([5, 6, 7]):
        z_rows.append(f'<row r="{r + 1}" ht="21" customHeight="1">{inline(f"B{r + 1}", 2 * font, TEXT)}{inline(f"C{r + 1}", 2 * font, "0" * 20)}</row>')
    room_cols = ''.join(f'<col min="{i + 2}" max="{i + 2}" width="{round((40 + i) / 7, 6)}" customWidth="1"/>' for i in range(46))
    room_rows = []
    for r, font in enumerate(ROOM_FONTS):
        cells = ''.join(f'<c r="{col_letter(i + 2)}{r + 1}" s="{2 * font + 1}"><v>1E+20</v></c>' for i in range(46))
        room_rows.append(f'<row r="{r + 1}" ht="21" customHeight="1">{cells}</row>')
    # Rows 12-15: numbers that fit from some column on - $1,234,567 in Calibri 11, $12,345 and $123,456 in Segoe UI
    # Black 11 ("$"#,##0), 1234567 in Calibri 11 General.
    # Rows 16-19: 1,234,567 in Calibri 11 and Segoe UI Black 11 as #,##0_) and #,##0_W (the placeholder as wide as its
    # character, not a space).
    for r, (style, value) in enumerate([(1, 1234567), (7, 12345), (7, 123456), (0, 1234567),
                                        (placeholder, 1234567), (placeholder + 1, 1234567), (placeholder + 2, 1234567), (placeholder + 3, 1234567)]):
        cells = ''.join(f'<c r="{col_letter(i + 2)}{r + 12}" s="{style}"><v>{value}</v></c>' for i in range(46))
        room_rows.append(f'<row r="{r + 12}" ht="21" customHeight="1">{cells}</row>')
    sheets = [('Black', sheet(100, black_cols, black_rows, True)), ('Z57', sheet(57, z_cols, z_rows)), ('Room', sheet(100, room_cols, room_rows))] + [
        (f'Room{z}', sheet(z, room_cols, room_rows)) for z in (50, 60, 67, 75, 85, 125, 150, 200)]
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
