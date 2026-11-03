#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-table-styles.xlsx: one sheet of small tables, each in a
built-in table style of every family (Light, Medium, Dark; the first of each seven and some with an
accent) with header rows, total rows, banded rows and columns and first / last columns turned on and off,
and one table in a custom style (`<tableStyles>` of styles.xml) that uses every element a table style has:
the whole table's lines, stripes two rows wide, a column stripe, the header and total rows and their first
and last cells. The theme is the Office theme of pandas-test_decimal.xlsx. Deterministic: run it again and
the bytes are the same.

    python3 tools/excel-parity/make-table-styles-fixture.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-table-styles.xlsx')
THEME_SOURCE = os.path.join(REPO, 'tests', 'test-documents-excel', 'pandas-test_decimal.xlsx')

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# (style name, show first column, show last column, row stripes, column stripes, total row)
TABLES = [
    ('TableStyleLight1', False, False, True, False, False),
    ('TableStyleLight9', True, False, True, False, True),
    ('TableStyleLight15', False, True, True, True, False),
    ('TableStyleMedium2', False, False, True, False, True),
    ('TableStyleMedium7', False, False, False, True, False),
    ('TableStyleMedium9', True, True, True, False, True),
    ('TableStyleMedium16', True, False, True, False, False),
    ('TableStyleMedium28', False, False, True, False, True),
    ('TableStyleDark1', True, False, True, False, True),
    ('TableStyleDark9', False, False, True, False, True),
    ('Boxed', True, True, True, True, True),
]
COLUMNS = ['Region', 'Q1', 'Q2', 'Q3']
ROWS = [('North', 12, 15, 9), ('South', 7, 11, 14), ('East', 20, 4, 8), ('West', 5, 9, 13), ('Centre', 16, 6, 10)]
WIDTH = len(COLUMNS)


def col(index):
    name = ''
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        name = chr(65 + rest) + name
    return name


def esc(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


strings = []


def sst(text):
    if text not in strings:
        strings.append(text)
    return strings.index(text)


# Tables two to a band of rows: A..D and F..I, each band header + 5 rows + total + 2 blank rows.
cells = {}  # (row, col) -> xml of the <c> without the r attribute
tables = []
for number, (style, first, last, rows, cols, total) in enumerate(TABLES):
    band, side = divmod(number, 2)
    top = band * 10 + 3           # 1-based row of the header (row 1 holds the title, the row above the name)
    left = side * (WIDTH + 1)     # 0-based column
    name = 'T%d_%s' % (number + 1, style)
    cells[(top - 1, left)] = '<c r="%s%d" t="s"><v>%d</v></c>' % (col(left), top - 1, sst(style))
    for k, title in enumerate(COLUMNS):
        cells[(top, left + k)] = '<c r="%s%d" t="s"><v>%d</v></c>' % (col(left + k), top, sst(title))
    for i, row in enumerate(ROWS):
        r = top + 1 + i
        cells[(r, left)] = '<c r="%s%d" t="s"><v>%d</v></c>' % (col(left), r, sst(row[0]))
        for k in range(1, WIDTH):
            cells[(r, left + k)] = '<c r="%s%d"><v>%d</v></c>' % (col(left + k), r, row[k])
    bottom = top + len(ROWS)
    if total:
        bottom += 1
        cells[(bottom, left)] = '<c r="%s%d" t="s"><v>%d</v></c>' % (col(left), bottom, sst('Total'))
        for k in range(1, WIDTH):
            value = sum(row[k] for row in ROWS)
            cells[(bottom, left + k)] = '<c r="%s%d"><f>SUBTOTAL(109,%s[%s])</f><v>%d</v></c>' % (col(left + k), bottom, name, COLUMNS[k], value)
    ref = '%s%d:%s%d' % (col(left), top, col(left + WIDTH - 1), bottom)
    filter_ref = '%s%d:%s%d' % (col(left), top, col(left + WIDTH - 1), top + len(ROWS))
    columns = ''.join(
        '<tableColumn id="%d" name="%s"%s/>' % (
            k + 1, COLUMNS[k],
            ((' totalsRowLabel="Total"' if k == 0 else ' totalsRowFunction="sum"') if total else ''))
        for k in range(WIDTH))
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<table xmlns="%s" id="%d" name="%s" displayName="%s" ref="%s"%s>'
           '<autoFilter ref="%s"/><tableColumns count="%d">%s</tableColumns>'
           '<tableStyleInfo name="%s" showFirstColumn="%d" showLastColumn="%d" showRowStripes="%d" showColumnStripes="%d"/></table>') % (
        MAIN, number + 1, name, name, ref, (' totalsRowCount="1"' if total else ' totalsRowShown="0"'),
        filter_ref, WIDTH, columns, style, first, last, rows, cols)
    tables.append(xml)

cells[(1, 0)] = '<c r="A1" t="s"><v>%d</v></c>' % sst('Table styles: built-in palettes and a custom style')

rows_xml = []
for r in sorted({r for r, _ in cells}):
    row_cells = ''.join(cells[(r, c)] for c in sorted(c for rr, c in cells if rr == r))
    rows_xml.append('<row r="%d">%s</row>' % (r, row_cells))
last_row = max(r for r, _ in cells)
sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<worksheet xmlns="%s" xmlns:r="%s"><dimension ref="A1:%s%d"/>'
         '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
         '<cols><col min="1" max="%d" width="11" customWidth="1"/></cols>'
         '<sheetData>%s</sheetData>'
         '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
         '<tableParts count="%d">%s</tableParts></worksheet>') % (
    MAIN, REL, col(2 * WIDTH), last_row, 2 * WIDTH + 1, ''.join(rows_xml), len(tables),
    ''.join('<tablePart r:id="rId%d"/>' % (i + 1) for i in range(len(tables))))

# The custom style's formats (dxfs), by element.
dxfs = [
    # 0 wholeTable: a medium outline in accent 5, thin inner lines in its light tint
    '<dxf><font><color theme="1"/></font><border><left style="medium"><color theme="8"/></left><right style="medium"><color theme="8"/></right>'
    '<top style="medium"><color theme="8"/></top><bottom style="medium"><color theme="8"/></bottom>'
    '<vertical style="thin"><color theme="8" tint="0.39997558519241921"/></vertical><horizontal style="thin"><color theme="8" tint="0.39997558519241921"/></horizontal></border></dxf>',
    # 1 headerRow: dark fill, white bold text
    '<dxf><font><b/><color theme="0"/></font><fill><patternFill patternType="solid"><bgColor theme="8" tint="-0.249977111117893"/></patternFill></fill></dxf>',
    # 2 totalRow: italic, a double line above
    '<dxf><font><i/></font><border><top style="double"><color theme="8"/></top></border></dxf>',
    # 3 firstColumn: bold
    '<dxf><font><b/></font></dxf>',
    # 4 lastColumn: red text
    '<dxf><font><color rgb="FFC00000"/></font></dxf>',
    # 5 firstRowStripe (two rows): light fill
    '<dxf><fill><patternFill patternType="solid"><bgColor theme="8" tint="0.79998168889431442"/></patternFill></fill></dxf>',
    # 6 firstColumnStripe: a light grey fill
    '<dxf><fill><patternFill patternType="solid"><bgColor theme="0" tint="-4.9989318521683403E-2"/></patternFill></fill></dxf>',
    # 7 firstHeaderCell: an orange fill
    '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFC000"/></patternFill></fill></dxf>',
    # 8 lastTotalCell: a green fill, bold
    '<dxf><font><b/></font><fill><patternFill patternType="solid"><bgColor theme="9" tint="0.59999389629810485"/></patternFill></fill></dxf>',
]
custom = ('<tableStyle name="Boxed" pivot="0" count="9">'
          '<tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="headerRow" dxfId="1"/>'
          '<tableStyleElement type="totalRow" dxfId="2"/><tableStyleElement type="firstColumn" dxfId="3"/>'
          '<tableStyleElement type="lastColumn" dxfId="4"/><tableStyleElement type="firstRowStripe" size="2" dxfId="5"/>'
          '<tableStyleElement type="firstColumnStripe" dxfId="6"/><tableStyleElement type="firstHeaderCell" dxfId="7"/>'
          '<tableStyleElement type="lastTotalCell" dxfId="8"/></tableStyle>')
styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<styleSheet xmlns="%s"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
          '<dxfs count="%d">%s</dxfs>'
          '<tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">%s</tableStyles>'
          '</styleSheet>') % (MAIN, len(dxfs), ''.join(dxfs), custom)

shared = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>') % (
    MAIN, len(strings), len(strings), ''.join('<si><t>%s</t></si>' % esc(s) for s in strings))

workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="%s" xmlns:r="%s"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews>'
            '<sheets><sheet name="Styles" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>') % (MAIN, REL)

with zipfile.ZipFile(THEME_SOURCE) as source:
    theme = source.read('xl/theme/theme1.xml')

content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                 '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                 '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                 '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                 '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                 '%s</Types>') % ''.join(
    '<Override PartName="/xl/tables/table%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' % (i + 1)
    for i in range(len(tables)))
root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/></Relationships>') % REL
workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>'
                 '<Relationship Id="rId2" Type="%s/theme" Target="theme/theme1.xml"/>'
                 '<Relationship Id="rId3" Type="%s/styles" Target="styles.xml"/>'
                 '<Relationship Id="rId4" Type="%s/sharedStrings" Target="sharedStrings.xml"/></Relationships>') % (REL, REL, REL, REL)
sheet_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>') % ''.join(
    '<Relationship Id="rId%d" Type="%s/table" Target="../tables/table%d.xml"/>' % (i + 1, REL, i + 1) for i in range(len(tables)))

parts = [
    ('[Content_Types].xml', content_types),
    ('_rels/.rels', root_rels),
    ('xl/workbook.xml', workbook),
    ('xl/_rels/workbook.xml.rels', workbook_rels),
    ('xl/worksheets/sheet1.xml', sheet),
    ('xl/worksheets/_rels/sheet1.xml.rels', sheet_rels),
    ('xl/theme/theme1.xml', theme),
    ('xl/styles.xml', styles),
    ('xl/sharedStrings.xml', shared),
] + [('xl/tables/table%d.xml' % (i + 1), xml) for i, xml in enumerate(tables)]

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as package:
    for name, data in parts:
        info = zipfile.ZipInfo(name, date_time=(2026, 11, 3, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        package.writestr(info, data if isinstance(data, bytes) else data.encode('utf-8'))
print(OUT)
