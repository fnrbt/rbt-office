#!/usr/bin/env python3
"""Writes a synthetic workbook for the right edge of printed, right-aligned numbers: how far the last character's
Poppler word box (pdftotext -bbox xMax) stands from the cell's right border, for an upright and an italic number
(docs/excel-parity-tracker.md, "Found, not matched": oyabayo's right-aligned numbers). Checked against a fresh Excel
2019 direct PDF export (`excel-print-pages.sh`, before any preview or sheet activation) and
tools/spreadsheet-pdf-runner.

One sheet per property, each changing ONE thing from S00: font face and size, italic, number format, column width,
alignment, a size sweep, and held-out format/width pairs. Column B holds the printed values (4 rows); it is filled
#FFF2CC so the printed fill's right edge is the cell's right border, which the oracle measures directly. The metric
is a word box, so the gap includes the font's right side bearing and any trailing format pad; differences between
word-box gaps are observed placement differences, not measured glyph outlines.

    python3 tools/excel-parity/make-print-right-edge-fixture.py OUT.xlsx
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'


def attr(text):
    """Escape an XML attribute value (quotes included)."""
    return escape(text, {'"': '&quot;'})
# name, face, size, italic, format id, align, column width in characters, values, pageSetup scale
SHEETS = [
    ('S00_base',              'Aptos Narrow', 14, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S01_italic',            'Aptos Narrow', 14, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S02_plainformat',       'Aptos Narrow', 14, False, 200, 'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S03_italic_plainformat','Aptos Narrow', 14, True,  200, 'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S04_calibri11',         'Calibri',      11, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S05_calibri11_italic',  'Calibri',      11, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S06_italic_wide',       'Aptos Narrow', 14, True,  10,  'right',  30, [-0.9841, -1, 0.5, 1234567], 100),
    ('S07_italic_accounting', 'Aptos Narrow', 14, True,  201, 'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S08_italic_center',     'Aptos Narrow', 14, True,  10,  'center', 18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S09_italic_left',       'Aptos Narrow', 14, True,  10,  'left',   18, [-0.9841, -1, 0.5, 1234567], 100),
    # a size sweep of the reserve, upright and italic pairs
    ('S10_aptos9',            'Aptos Narrow',  9, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S11_aptos9_italic',     'Aptos Narrow',  9, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S12_aptos11',           'Aptos Narrow', 11, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S13_aptos11_italic',    'Aptos Narrow', 11, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S14_aptos20',           'Aptos Narrow', 20, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S15_aptos20_italic',    'Aptos Narrow', 20, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S16_calibri8',          'Calibri',       8, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S17_calibri8_italic',   'Calibri',       8, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S18_calibri14',         'Calibri',      14, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S19_calibri14_italic',  'Calibri',      14, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S20_calibri20',         'Calibri',      20, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S21_calibri20_italic',  'Calibri',      20, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    # held out after the rule: not used to derive anything; upright twins give same-string reserves
    ('S22_holdout',           'Aptos Narrow', 14, True,  202, 'right',  10, [-0.9841, -1, 0.5, 1234567], 100),
    ('S23_holdout2',          'Calibri',      12, True,  203, 'right',  24, [-0.9841, -1, 0.5, 1234567], 100),
    ('S24_holdout_upright',   'Aptos Narrow', 14, False, 202, 'right',  10, [-0.9841, -1, 0.5, 1234567], 100),
    ('S25_holdout2_upright',  'Calibri',      12, False, 203, 'right',  24, [-0.9841, -1, 0.5, 1234567], 100),
    # out of sample for the tan(angle) x selected-ascent candidate: other installed families
    ('S26_arial8',            'Arial',         8, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S27_arial8_italic',     'Arial',         8, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S28_arial14',           'Arial',        14, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S29_arial14_italic',    'Arial',        14, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S30_arial20',           'Arial',        20, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S31_arial20_italic',    'Arial',        20, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S32_arial14_digit',     'Arial',        14, False, 200, 'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S33_arial14_digit_italic','Arial',      14, True,  200, 'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S34_times14',           'Times New Roman', 14, False, 10, 'right', 18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S35_times14_italic',    'Times New Roman', 14, True, 10,  'right', 18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S36_cambria14',         'Cambria',      14, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    ('S37_cambria14_italic',  'Cambria',      14, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
    # print scale 75 %: the candidate predicts 0.75 x the 100 % reserve
    ('S38_aptos14_75',        'Aptos Narrow', 14, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 75),
    ('S39_aptos14_italic_75', 'Aptos Narrow', 14, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 75),
    ('S40_arial14_75',        'Arial',        14, False, 10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 75),
    ('S41_arial14_italic_75', 'Arial',        14, True,  10,  'right',  18, [-0.9841, -1, 0.5, 1234567], 75),
    ('S42_accounting_upright','Aptos Narrow', 14, False, 201, 'right',  18, [-0.9841, -1, 0.5, 1234567], 100),
]
NUMFMTS = [(200, '0.00'), (201, ACCOUNTING), (202, '0.000%'), (203, '#,##0.0')]


def build(out):
    # fonts: 0 Normal, then one per sheet
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    for _, face, size, italic, _, _, _, _, _ in SHEETS:
        fonts.append(f'{"<i/>" if italic else ""}<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>')
    # fills: 0 none, 1 gray125, 2 the measuring fill #FFF2CC
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>')
    # borders: 0 none, 1 a thin right line at the cell's right edge
    borders = ('<border><left/><right/><top/><bottom/><diagonal/></border>'
               '<border><left/><right style="thin"><color rgb="FF808080"/></right><top/><bottom/><diagonal/></border>')
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>']
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
        style = 2 + i
        rows = ''.join(f'<row r="{r}"><c r="B{r}" s="{style}"><v>{v}</v></c></row>' for r, v in enumerate(values, 1))
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
