#!/usr/bin/env python3
"""Writes a synthetic workbook that holds out a fresh face/size/scale/format for the print italic reserve
(`FontMetrics.printItalicReservePt`). Blind: this generator and `print-italic-holdout-predictions.tsv` are written
before any Excel export of the workbook.

Each pair is two sheets that differ only in italic (same face, size, print scale, column width, number format and
values), with the value column filled #FFF2CC and a thin right border. The metric is Poppler's `pdftotext -bbox`
word-box xMax to the printed fill's right edge, as in `measure-print-right-edge.py`.

    python3 tools/excel-parity/make-print-italic-holdout.py OUT.xlsx
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
ACCOUNTING = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'

# name, face, points, italic, numFmtId, alignment, column width, values, print scale
SHEETS = [
    ('H01_verdana13_85_upright', 'Verdana', 13, False, 10, 'right', 18, [-0.9841, -1, 0.5, 1234567], 85),
    ('H02_verdana13_85_italic',  'Verdana', 13, True,  10, 'right', 18, [-0.9841, -1, 0.5, 1234567], 85),
    ('H03_georgia15_upright',    'Georgia', 15, False, 204, 'right', 18, [-0.9841, -1, 0.5, 1234567], 100),
    ('H04_georgia15_italic',     'Georgia', 15, True,  204, 'right', 18, [-0.9841, -1, 0.5, 1234567], 100),
    ('H05_segoe10p5_90_upright', 'Segoe UI', 10.5, False, 200, 'right', 18, [-0.9841, -1, 0.5, 1234567], 90),
    ('H06_segoe10p5_90_italic',  'Segoe UI', 10.5, True,  200, 'right', 18, [-0.9841, -1, 0.5, 1234567], 90),
    ('H07_verdana13_85_hash_upright', 'Verdana', 13, False, 205, 'right', 18, [-0.9841, -1, 0.5, 1234567], 85),
    ('H08_verdana13_85_hash_italic',  'Verdana', 13, True,  205, 'right', 18, [-0.9841, -1, 0.5, 1234567], 85),
    ('H09_aptos14_acct_upright', 'Aptos Narrow', 14, False, 201, 'right', 18, [-0.9841, -1, 0.5, 1234567], 100),
    ('H10_aptos14_acct_italic',  'Aptos Narrow', 14, True,  201, 'right', 18, [-0.9841, -1, 0.5, 1234567], 100),
    ('H11_aptos14_acctlit_upright', 'Aptos Narrow', 14, False, 201, 'right', 18, ['(0.98)', '(1.00)', '0.50', '1,234,567.00'], 100),
    ('H12_aptos14_acctlit_italic',  'Aptos Narrow', 14, True,  201, 'right', 18, ['(0.98)', '(1.00)', '0.50', '1,234,567.00'], 100),
    ('H13_aptos14_overflow_upright', 'Aptos Narrow', 14, False, 206, 'right', 10, [-0.9841, -1, 0.5, 1234567], 100),
    ('H14_aptos14_overflow_italic',  'Aptos Narrow', 14, True,  206, 'right', 10, [-0.9841, -1, 0.5, 1234567], 100),
]
NUMFMTS = [(200, '0.00'), (201, ACCOUNTING), (204, '#,##0.00'), (205, '"#######"'), (206, '0.000%')]


def attr(text):
    return escape(text, {'"': '&quot;'})


def cell_xml(ref, style, value):
    if isinstance(value, str):
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
    return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'


def build(out):
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    for _, face, size, italic, _, _, _, _, _ in SHEETS:
        fonts.append(f'{"<i/>" if italic else ""}<sz val="{size}"/><name val="{face}"/><family val="2"/>')
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>')
    borders = ('<border><left/><right/><top/><bottom/><diagonal/></border>'
               '<border><left/><right style="thin"><color rgb="FF808080"/></right><top/><bottom/><diagonal/></border>')
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    for i, (_, _, _, _, fmt, align, _, _, _) in enumerate(SHEETS):
        xfs.append(f'<xf numFmtId="{fmt}" fontId="{i + 1}" fillId="2" borderId="1" xfId="0" '
                   f'applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1">'
                   f'<alignment horizontal="{align}"/></xf>')
    numfmts = ''.join(f'<numFmt numFmtId="{i}" formatCode="{attr(code)}"/>' for i, code in NUMFMTS)
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<numFmts count="{len(NUMFMTS)}">{numfmts}</numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="3">{fills}</fills>'
        f'<borders count="2">{borders}</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    sheet_files = []
    for i, (name, face, size, italic, fmt, align, width, values, scale) in enumerate(SHEETS):
        style = 1 + i
        rows = ''.join(f'<row r="{r}">{cell_xml(f"B{r}", style, v)}</row>' for r, v in enumerate(values, 1))
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if i == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="{width}" customWidth="1"/></cols>'
            f'<sheetData>{rows}</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            f'<pageSetup paperSize="1" orientation="portrait" scale="{scale}"/></worksheet>')
        sheet_files.append((name, xml))

    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>'
                   for i in range(len(sheet_files)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
            + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (name, _) in enumerate(sheet_files))
            + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
                      for i in range(len(sheet_files)))
            + f'<Relationship Id="rId{len(sheet_files) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, (_, xml) in enumerate(sheet_files):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)


if __name__ == '__main__':
    build(sys.argv[1])
