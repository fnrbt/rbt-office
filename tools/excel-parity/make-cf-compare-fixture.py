#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-cf-compare.xlsx, a synthetic workbook (not saved by Excel): how a
conditional format's "cell value" rule (`cellIs`) compares numbers in Excel 2019 (checked in Excel on the VM,
`Range.DisplayFormat.Interior.Color` of each cell; docs/excel-parity-tracker.md, "Conditional formats"):

Each row r of "Cmp" has a value (or formula) in A, one `cellIs` rule on A r that fills the cell red when it holds,
and in B what Excel shows for it (1: filled, 0: not). Excel compares as its `=` operator does: two numbers rounded
to 15 significant digits are equal, otherwise the binary order, no absolute tolerance (1E-10 > 0, 1E-10 <> 0).

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-cf-compare-fixture.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-cf-compare.xlsx')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

# (A: a number "v:..." or a formula "f:...", operator, operand formulas, Excel 2019's result: 1 filled)
CASES = [
    ('f:0.1+0.2', 'equal', ['0.3'], 1),
    ('f:0.1+0.2', 'greaterThan', ['0.3'], 0),
    ('f:0.1+0.2', 'lessThanOrEqual', ['0.3'], 1),
    ('f:0.1+0.2', 'notEqual', ['0.3'], 0),
    ('v:1E-10', 'greaterThan', ['0'], 1),
    ('v:1E-16', 'greaterThan', ['0'], 1),
    ('v:1E-10', 'equal', ['0'], 0),
    ('v:-1E-10', 'lessThan', ['0'], 1),
    ('f:1*(0.1+0.2-0.3)', 'equal', ['0'], 0),
    ('f:1*(0.1+0.2-0.3)', 'notEqual', ['0'], 1),
    ('f:0.1+0.2', 'between', ['0.3', '0.5'], 1),
    ('f:0.1+0.2', 'notBetween', ['0', '0.3'], 0),
    ('f:1+22*2^-52', 'equal', ['1'], 1),
    ('f:1+23*2^-52', 'equal', ['1'], 0),
    ('f:1E15+5', 'equal', ['1E15+10'], 1),
    ('v:0.5', 'lessThan', ['0.5+1E-10'], 1),
    ('v:0.5', 'greaterThanOrEqual', ['0.5+1E-10'], 0),
    ('v:1E-300', 'greaterThan', ['0'], 1),
    ('v:100', 'between', ['100.0000000001', '200'], 0),
    ('v:100', 'between', ['100.00000000000001', '200'], 1),
]


def cell_a(r, spec):
    kind, text = spec.split(':', 1)
    return f'<c r="A{r}"><f>{text}</f></c>' if kind == 'f' else f'<c r="A{r}"><v>{text}</v></c>'


def build():
    rows, cfs = [], []
    for i, (a, op, operands, excel) in enumerate(CASES):
        r = i + 1
        rows.append(f'<row r="{r}">{cell_a(r, a)}<c r="B{r}"><v>{excel}</v></c></row>')
        formulas = ''.join(f'<formula>{o}</formula>' for o in operands)
        cfs.append(f'<conditionalFormatting sqref="A{r}"><cfRule type="cellIs" dxfId="0" priority="{r}" operator="{op}">{formulas}</cfRule></conditionalFormatting>')
    sheet = DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="24" customWidth="1"/></cols>'
        f'<sheetData>{"".join(rows)}</sheetData>{"".join(cfs)}'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>'
        '</styleSheet>')
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{CT}.worksheet+xml"/></Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews>'
                                   '<sheets><sheet name="Cmp" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
        'xl/worksheets/sheet1.xml': sheet,
    }
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(OUT)


if __name__ == '__main__':
    build()
