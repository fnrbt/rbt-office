#!/usr/bin/env python3
"""Fixtures for how table borders and cell margins take vertical space in Word (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/table-border.py [out dir]   (default tests/word/fixtures/table-border)

then, to ask Word (winvm) and read where it put every line:

    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/table-border --refresh-word
    python3 tools/word-parity/fixtures/para-border.py --positions tests/word/fixtures/table-border

Every line is one label in Arial 10pt at exactly 20pt a line, no space before or
after, on a Letter page with 1in margins. A fixture is a plain paragraph a01, a
table of one column (two where named) whose rows hold one line each (r01, r02 ...;
the second column's s01 ...), and a plain paragraph c01: the gap above r01, between
the rows and below the last, less the 20pt line, is what the borders and the cell
margins took. A fixture name ending in -c14 is the same document in Word 2010's
compatibility mode (w:compatSetting compatibilityMode 14) instead of 15.
--positions is para-border.py's: <dir>/word/positions.tsv, every label's page and
the top of its glyphs (pt). tests/word/table-border.fsx checks ours against it.
"""
import os
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def para(label, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{label}</w:t></w:r></w:p>'


def bdr(side, sz, val='single'):
    return f'<w:{side} w:val="{val}" w:sz="{sz}" w:space="0" w:color="000000"/>'


def borders(tag, sides, sz, val='single'):
    if not sides:
        return ''
    return f'<w:{tag}>' + ''.join(bdr(s, sz, val) for s in sides) + f'</w:{tag}>'


ALL = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']


def table(rows=3, cols=1, sides=ALL, sz=4, val='single', mar=(0, 0), style=None, tc=None, tr=None, ppr=''):
    """tc: {(row, col): tcBorders xml}; tr: trPr xml for every row."""
    width = 4000
    mar_xml = (f'<w:tblCellMar><w:top w:w="{mar[0]}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
               f'<w:bottom w:w="{mar[1]}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>')
    tbl_pr = ((f'<w:tblStyle w:val="{style}"/>' if style else '') + f'<w:tblW w:w="{width * cols}" w:type="dxa"/>'
              + ('' if style else borders('tblBorders', [s for s in sides if cols > 1 or s != 'insideV'], sz, val) + mar_xml))
    out = []
    for r in range(1, rows + 1):
        cells = []
        for c in range(1, cols + 1):
            tcb = (tc or {}).get((r, c), '')
            label = f'{"rs"[c - 1]}{r:02d}'
            cells.append(f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{tcb}</w:tcPr>{para(label, ppr)}</w:tc>')
        out.append(f'<w:tr>{"<w:trPr>" + tr + "</w:trPr>" if tr else ""}{"".join(cells)}</w:tr>')
    grid = ''.join(f'<w:gridCol w:w="{width}"/>' for _ in range(cols))
    return f'<w:tbl><w:tblPr>{tbl_pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{"".join(out)}</w:tbl>'


def probe(tbl):
    return para('a01') + tbl + para('c01')


def document(body):
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
            'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:customStyle="1" w:styleId="Grid43"><w:name w:val="Grid 43"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="400" w:lineRule="exact"/></w:pPr><w:tblPr><w:tblBorders>{"".join(bdr(s, 4) for s in ALL)}</w:tblBorders><w:tblCellMar><w:top w:w="43" w:type="dxa"/><w:left w:w="115" w:type="dxa"/><w:bottom w:w="43" w:type="dxa"/><w:right w:w="115" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>'''


def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting '
            f'w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')


TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''


def write(path, body, mode=15):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document(body))
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', settings(mode))


def families():
    """name -> body, one entry per fixture (each also as -c14)."""
    out = {'none': probe(table(sides=[]))}
    # The width of every border: what each takes above, between and below the rows.
    for sz in (4, 12, 24, 48):
        out[f'all-sz{sz}'] = probe(table(sz=sz))
    out['insideH-sz24'] = probe(table(sides=['insideH'], sz=24))
    out['topbottom-sz24'] = probe(table(sides=['top', 'bottom'], sz=24))
    out['sides-sz24'] = probe(table(sides=['left', 'right'], sz=24))
    out['double-sz6'] = probe(table(sz=6, val='double'))
    # Cell margins with and without borders.
    out['mar43-none'] = probe(table(sides=[], mar=(43, 43)))
    out['mar43-sz4'] = probe(table(mar=(43, 43)))
    out['mar43-sz24'] = probe(table(sz=24, mar=(43, 43)))
    out['mar100-sz24'] = probe(table(sz=24, mar=(100, 0)))
    # A table style's borders and margins (docx4j's ElementTable on Table Grid).
    out['style-grid43'] = probe(table(style='Grid43'))
    out['style-grid43-2col'] = probe(table(style='Grid43', cols=2))
    # One cell's own border.
    out['cell-top24'] = probe(table(tc={(2, 1): borders('tcBorders', ['top'], 24)}))
    out['cell-bottom24'] = probe(table(tc={(2, 1): borders('tcBorders', ['bottom'], 24)}))
    out['cell2col-bottom24'] = probe(table(cols=2, tc={(2, 1): borders('tcBorders', ['bottom'], 24)}))
    out['cell2col-nob-bottom24'] = probe(table(cols=2, sides=[], tc={(2, 1): borders('tcBorders', ['bottom'], 24)}))
    # Row heights with borders.
    out['exact-sz24'] = probe(table(sz=24, tr='<w:trHeight w:val="600" w:hRule="exact"/>'))
    out['atleast-sz24'] = probe(table(sz=24, tr='<w:trHeight w:val="600" w:hRule="atLeast"/>'))
    out['atleast-low-sz24'] = probe(table(sz=24, tr='<w:trHeight w:val="300" w:hRule="atLeast"/>'))
    return out


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/table-border')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    all_ = families()
    for name, body in all_.items():
        write(os.path.join(out_dir, name + '.docx'), body)
        write(os.path.join(out_dir, name + '-c14.docx'), body, mode=14)
    print(f'{2 * len(all_)} fixtures in {out_dir}')
