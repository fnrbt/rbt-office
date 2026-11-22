#!/usr/bin/env python3
"""Create the public input for column-edit.vbs.

    python3 tools/word-parity/fixtures/column-edit.py /tmp/column-edit.docx
    tools/word-parity/run-word-vbs-on-vm.sh tools/word-parity/column-edit.vbs /tmp/column-edit.docx /tmp/column-edit.tsv
    # Repeat with each --case; Word references are in
    # tests/word/fixtures/column-edit/word/<case>.tsv.

Each case is a fixed 2x2 table, with zero cell margins and distinct cell fills.
Columns are 75/150 pt (base), 150/75 pt (reversed), or 200/225 pt (near-page).
The wide-table and narrow-table cases deliberately mismatch tblW and grid:
4500 vs 1500+1500 twips, and 3000 vs 1500+3000 twips, respectively.
The text area is 468 pt wide. The Word probe inserts beside each column in
either direction, then deletes the inserted column. Positions and widths are points.
Fill colors in the TSV are Word's BGR decimal color values: red 255, green
65280, blue 16711680, yellow 65535. The XML rows give twips.
"""
import argparse
import zipfile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
parser = argparse.ArgumentParser()
parser.add_argument('output')
parser.add_argument('--case', choices=('base', 'reversed', 'near-page', 'wide-table', 'narrow-table'), default='base')
args = parser.parse_args()
widths = {'base': (1500, 3000), 'reversed': (3000, 1500), 'near-page': (4000, 4500), 'wide-table': (1500, 1500), 'narrow-table': (1500, 3000)}[args.case]
table_width = {'wide-table': 4500, 'narrow-table': 3000}.get(args.case, sum(widths))
fills = (('FF0000', '00FF00'), ('0000FF', 'FFFF00'))
grid = ''.join(f'<w:gridCol w:w="{width}"/>' for width in widths)
rows = []
for row in range(1, 3):
    cells = []
    for col, width in enumerate(widths, 1):
        fill = fills[row - 1][col - 1]
        cells.append(f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:shd w:val="clear" w:fill="{fill}"/></w:tcPr>'
                     f'<w:p><w:r><w:t>r{row}c{col}</w:t></w:r></w:p></w:tc>')
    rows.append('<w:tr>' + ''.join(cells) + '</w:tr>')
doc = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body><w:tbl><w:tblPr><w:tblW w:w="{table_width}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{''.join(rows)}</w:tbl><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'''
ct = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'''
rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc)
