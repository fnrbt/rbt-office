#!/usr/bin/env python3
"""Generate synthetic cache-open workbooks: python3 make-cache-open-probe.py DIR."""
import sys
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
package = 'http://schemas.openxmlformats.org/package/2006/relationships'
office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
content = 'http://schemas.openxmlformats.org/package/2006/content-types'
for iteration in (False, True):
    rows = []
    for r, label in enumerate(('absent', 'empty', 'zero', 'str-empty', 'nonzero', 'str-absent'), 1):
        attr = ' t="str"' if r in (4, 6) else ''
        cache = ('', '<v/>', '<v>0</v>', '<v/>', '<v>7</v>', '')[r - 1]
        rows.append(f'<row r="{r}"><c r="A{r}"{attr}><f>B{r}+1</f>{cache}</c>'
                    f'<c r="B{r}"><f>A{r}/2</f><v>2</v></c>'
                    f'<c r="C{r}"><f>A{r}+10</f>{"<v/>" if len(sys.argv) > 2 and sys.argv[2] == "uncached-dependent" else "<v>99</v>"}</c>'
                    f'<c r="D{r}"><f>1+2</f><v>77</v></c>'
                    f'<c r="E{r}" t="inlineStr"><is><t>{label}</t></is></c></row>')
    parts = {
        '[Content_Types].xml': f'<Types xmlns="{content}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{package}"><Relationship Id="rId1" Type="{office}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{package}"><Relationship Id="rId1" Type="{office}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{main}" xmlns:r="{office}"><sheets><sheet name="Caches" sheetId="1" r:id="rId1"/></sheets><calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1" iterate="{int(iteration)}" iterateCount="2" iterateDelta="0.001"/></workbook>',
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{main}"><sheetData>{"".join(rows)}</sheetData></worksheet>',
    }
    path = out / ('cache-open-iter.xlsx' if iteration else 'cache-open-noiter.xlsx')
    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
