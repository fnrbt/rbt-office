#!/usr/bin/env python3
"""Fixture family: how Word decides a table's column widths (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/table-width.py [out dir]   (default tests/word/fixtures/table-width)

then, to ask Word (winvm) and read where it put every cell's edges:

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/<own dir> \
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/table-width --refresh-word
    python3 tools/word-parity/fixtures/table-width.py --edges tests/word/fixtures/table-width/word/*.pdf \
      > tests/word/fixtures/table-width/word/edges.tsv

Every document: Letter, 1in margins (text 468pt = 9360 twips wide), Arial 10pt at exactly 20pt a line,
nothing before or after, and one table (the default cell margins 108 twips, no borders) between the
paragraphs a01 and c01. Cell (r, c) holds a left-aligned paragraph "Lrc" and a right-aligned one "Rrc"
(a content paragraph before them where the name says): "Lrc"'s left is the cell's left plus its left
margin, "Rrc"'s right its right less its right margin. --edges writes, per PDF, for every cell:
<doc> TAB <row> TAB <col> TAB <page> TAB <left of Lrc> TAB <right of Rrc>   (pt from the page's left)

The names (each grid / tcW / tblW in twips; "pct" 5000 = 100%; -fixed: tblLayout fixed, -auto: none):
  agree-*            grid 3000/6360, tcW the same, tblW auto
  grideq-tcw-*       grid 4680/4680, tcW 3000/6360 (the showcase's annual-results: python-docx's grid)
  tcweq-grid-*       grid 3000/6360, tcW 4680/4680
  grid-notcw-*       grid 3000/6360, no tcW
  tcw-nogrid-*       tcW 3000/6360, no tblGrid
  narrow-grid-*      grid 2000/2000, no tcW (the grid narrower than the text)
  narrow-tcw-*       grid 2000/2000, tcW 2000/2000
  narrow-tblw-*      that, tblW 9360 dxa;  narrow-pct-*: tblW 5000 pct;  narrow-pct-notcw-*: no tcW
  narrow-tcw4-*      grid 2000/2000, tcW 4000/4000 (the cells wider than the grid)
  small-tblw-*       grid 3000/6360, tcW the same, tblW 6000 dxa (the table narrower than its cells)
  pct-grid-*         grid 1155/8101, tcW 1000/4000 pct, tblW 5000 pct (docx4j-headers-footers)
  tcwpct-*           grid 4680/4680, tcW 1000/4000 pct, tblW auto
  tcwauto-*          grid 3000/6360, tcW auto
  wide-*             grid 6000/6000, tcW the same (wider than the text);  wide-tblw-*: tblW 12000 dxa;
                     wide-grid-*: no tcW;  wide-pct-*: tblW 5000 pct
  ind720-*           grid 4680/4680, tcW the same, tblInd 720
  mar0-*, mar288-*   grid 4680/4680, tcW the same, cell margins left/right 0 / 288
  span-*             grid 2000/2000/5360; row 1 a gridSpan 2 cell (tcW 4000) and one of 5360, row 2 three
                     cells 2000/2000/5360;  span-conflict-*: row 1's spanning cell tcW 6000, its other 3360
  cols2-*            two columns 720 apart (each 4320 wide):  -full: grid 9360, tcW 9360 (field-guide's
                     cards); -small: grid 2000/2000 tcW the same; -pct: tblW 5000 pct, grid 4680/4680
  nest-*             a one-cell table (grid 9360, tcW 9360) holding a table: -full: grid 9360 tcW 9360;
                     -small: grid 3000/3000 tcW the same; -pct: tblW 5000 pct, grid 4680/4680, tcW 2340
  content-*          grid 1000/1000, tcW 1000/1000, cell 1 holds "Northgate Offshore" (two words wider
                     than 50pt), cell 2 "x";  content-long-*: cell 1 a 60-word sentence;
                     content-tcwauto-*: tcW auto, the long sentence; content-nogrid-*: no tcW, no grid
  content-short-*    no grid, no tcW, cell 1 "Northgate Offshore" (short of the line)
  rows-swap-*        row 1 tcW 3000/6360, row 2 6360/3000;  rows-mixed-*: row 2 4000/5360;
                     rows-second-*: row 1 no tcW / 6360, row 2 4000/5360
  vmerge-*           rows-swap with row 1's first cell a vMerge restart, row 2's its continuation
  span-small-*       span's row 1 with its spanning cell's tcW 1000 (less than the columns it spans)
  widen-*            grid and tcW 1000/3000, tblW 9360 dxa (widen-pct: 5000 pct)
  shrink3-*          tcW 6000/3000/3000, cell 2 a 30-letter word;  shrink3-tblw-*: tblW 7000 dxa
  tcwpct-tblw-*      tcW 1000/4000 pct, tblW 6000 dxa
  <name>-c14         the same in Word 2010's layout (compatibilityMode 14)
"""
import os
import re
import subprocess
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def para(text, jc=None):
    ppr = f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ''
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def width(tag, w):
    if w is None:
        return ''
    kind, value = w if isinstance(w, tuple) else ('dxa', w)
    return f'<w:{tag} w:w="{value}" w:type="{kind}"/>'


def table(grid, tcw=None, tblw=('auto', 0), layout=None, ind=None, mar=None, rows=None, content=None, tag='', inner=None):
    """grid: gridCol widths (None: no tblGrid); tcw: per cell widths (None: no tcW) or a list per row;
    rows: list of rows, each a list of (span, tcW) - overrides grid/tcw cells; content: {(r, c): xml}."""
    pr = width('tblW', tblw)
    if ind is not None:
        pr += f'<w:tblInd w:w="{ind}" w:type="dxa"/>'
    if layout:
        pr += f'<w:tblLayout w:type="{layout}"/>'
    if mar is not None:
        pr += f'<w:tblCellMar><w:left w:w="{mar}" w:type="dxa"/><w:right w:w="{mar}" w:type="dxa"/></w:tblCellMar>'
    if rows is None:
        n = len(grid) if grid else len(tcw)
        rows = [[(1, tcw[c] if tcw else None) for c in range(n)]]
    out = []
    for r, row in enumerate(rows, 1):
        cells = []
        for c, (span, w, *extra) in enumerate(row, 1):
            tcpr = width('tcW', w) + (f'<w:gridSpan w:val="{span}"/>' if span > 1 else '') + ''.join(extra)
            body = (content or {}).get((r, c), '')
            if inner and (r, c) == (1, 1):
                body += inner + para('')
            cells.append(f'<w:tc><w:tcPr>{tcpr}</w:tcPr>{body}{para(f"L{tag}{r}{c}")}{para(f"R{tag}{r}{c}", "right")}</w:tc>')
        out.append(f'<w:tr>{"".join(cells)}</w:tr>')
    grid_xml = '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{g}"/>' for g in grid) + '</w:tblGrid>' if grid else ''
    return f'<w:tbl><w:tblPr>{pr}</w:tblPr>{grid_xml}{"".join(out)}</w:tbl>'


def document(body, columns=None):
    cols = f'<w:cols w:num="{columns[0]}" w:space="{columns[1]}"/>' if columns else ''
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
            f'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>{cols}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>'''


def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting '
            f'w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')


TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''


def write(path, body, mode=15, columns=None):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document(body, columns))
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', settings(mode))


def probe(tbl):
    return para('a01') + tbl + para('c01')


LONG = ' '.join(f'w{i:02d}' for i in range(1, 61))


def families():
    """name -> (body, columns), one per fixture and layout (each also as -c14)."""
    out = {}
    for lay, layout in (('auto', None), ('fixed', 'fixed')):
        def add(name, columns=None, **kw):
            out[f'{name}-{lay}'] = (probe(table(layout=layout, **kw)), columns)
        add('agree', grid=[3000, 6360], tcw=[3000, 6360])
        add('grideq-tcw', grid=[4680, 4680], tcw=[3000, 6360])
        add('tcweq-grid', grid=[3000, 6360], tcw=[4680, 4680])
        add('grid-notcw', grid=[3000, 6360])
        add('tcw-nogrid', grid=None, tcw=[3000, 6360])
        add('narrow-grid', grid=[2000, 2000])
        add('narrow-tcw', grid=[2000, 2000], tcw=[2000, 2000])
        add('narrow-tblw', grid=[2000, 2000], tcw=[2000, 2000], tblw=9360)
        add('narrow-pct', grid=[2000, 2000], tcw=[2000, 2000], tblw=('pct', 5000))
        add('narrow-pct-notcw', grid=[2000, 2000], tblw=('pct', 5000))
        add('narrow-tcw4', grid=[2000, 2000], tcw=[4000, 4000])
        add('small-tblw', grid=[3000, 6360], tcw=[3000, 6360], tblw=6000)
        add('pct-grid', grid=[1155, 8101], tcw=[('pct', 1000), ('pct', 4000)], tblw=('pct', 5000))
        add('tcwpct', grid=[4680, 4680], tcw=[('pct', 1000), ('pct', 4000)])
        add('tcwauto', grid=[3000, 6360], tcw=[('auto', 0), ('auto', 0)])
        add('wide', grid=[6000, 6000], tcw=[6000, 6000])
        add('wide-tblw', grid=[6000, 6000], tcw=[6000, 6000], tblw=12000)
        add('wide-grid', grid=[6000, 6000])
        add('wide-pct', grid=[6000, 6000], tcw=[6000, 6000], tblw=('pct', 5000))
        add('ind720', grid=[4680, 4680], tcw=[4680, 4680], ind=720)
        add('mar0', grid=[4680, 4680], tcw=[4680, 4680], mar=0)
        add('mar288', grid=[4680, 4680], tcw=[4680, 4680], mar=288)
        add('span', grid=[2000, 2000, 5360], rows=[[(2, 4000), (1, 5360)], [(1, 2000), (1, 2000), (1, 5360)]])
        add('span-conflict', grid=[2000, 2000, 5360], rows=[[(2, 6000), (1, 3360)], [(1, 2000), (1, 2000), (1, 5360)]])
        add('cols2-full', columns=(2, 720), grid=[9360], tcw=[9360])
        add('cols2-small', columns=(2, 720), grid=[2000, 2000], tcw=[2000, 2000])
        add('cols2-pct', columns=(2, 720), grid=[4680, 4680], tblw=('pct', 5000))
        for nest, inner in (('full', table([9360], [9360], layout=layout, tag='i')),
                            ('small', table([3000, 3000], [3000, 3000], layout=layout, tag='i')),
                            ('pct', table([4680, 4680], [2340, 2340], tblw=('pct', 5000), layout=layout, tag='i'))):
            add(f'nest-{nest}', grid=[9360], tcw=[9360], inner=inner)
        add('content', grid=[1000, 1000], tcw=[1000, 1000], content={(1, 1): para('Northgate Offshore'), (1, 2): para('x')})
        add('content-long', grid=[1000, 1000], tcw=[1000, 1000], content={(1, 1): para(LONG), (1, 2): para('x')})
        add('content-tcwauto', grid=[1000, 1000], tcw=[('auto', 0), ('auto', 0)], content={(1, 1): para(LONG), (1, 2): para('x')})
        add('content-nogrid', grid=None, tcw=None, rows=[[(1, None), (1, None)]], content={(1, 1): para(LONG), (1, 2): para('x')})
        add('content-short', grid=None, rows=[[(1, None), (1, None)]], content={(1, 1): para('Northgate Offshore'), (1, 2): para('x')})
        add('rows-swap', grid=[4680, 4680], rows=[[(1, 3000), (1, 6360)], [(1, 6360), (1, 3000)]])
        add('rows-mixed', grid=[4680, 4680], rows=[[(1, 3000), (1, 6360)], [(1, 4000), (1, 5360)]])
        add('rows-second', grid=[4680, 4680], rows=[[(1, None), (1, 6360)], [(1, 4000), (1, 5360)]])
        add('vmerge', grid=[4680, 4680], rows=[[(1, 3000, '<w:vMerge w:val="restart"/>'), (1, 6360)], [(1, 6360, '<w:vMerge/>'), (1, 3000)]])
        add('span-small', grid=[2000, 2000, 5360], rows=[[(2, 1000), (1, 5360)], [(1, 2000), (1, 2000), (1, 5360)]])
        add('widen', grid=[1000, 3000], tcw=[1000, 3000], tblw=9360)
        add('widen-pct', grid=[1000, 3000], tcw=[1000, 3000], tblw=('pct', 5000))
        add('shrink3', grid=[6000, 3000, 3000], tcw=[6000, 3000, 3000], content={(1, 2): para('abcdefghijklmnopqrstuvwxyzabcd')})
        add('shrink3-tblw', grid=[6000, 3000, 3000], tcw=[6000, 3000, 3000], tblw=7000, content={(1, 2): para('abcdefghijklmnopqrstuvwxyzabcd')})
        add('tcwpct-tblw', grid=[4680, 4680], tcw=[('pct', 1000), ('pct', 4000)], tblw=6000)
    return out


def edges(pdfs):
    for pdf in pdfs:
        doc = os.path.splitext(os.path.basename(pdf))[0]
        xml = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        page = 0
        cells = {}
        for line in xml.splitlines():
            if '<page ' in line:
                page += 1
            m = re.search(r'xMin="([\d.]+)" yMin="[\d.]+" xMax="([\d.]+)" yMax="[\d.]+">([LR])(i?)(\d)(\d)</word>', line)
            if m:
                key = (m.group(4), int(m.group(5)), int(m.group(6)))
                entry = cells.setdefault(key, [page, None, None])
                if m.group(3) == 'L':
                    entry[1] = float(m.group(1))
                else:
                    entry[2] = float(m.group(2))
        for (tag, r, c), (pg, left, right) in sorted(cells.items()):
            print(f'{doc}\t{tag}{r}\t{c}\t{pg}\t{left:.2f}\t{right:.2f}' if left is not None and right is not None
                  else f'{doc}\t{tag}{r}\t{c}\t{pg}\t{left}\t{right}')


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--edges':
        edges(sys.argv[2:])
        sys.exit(0)
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/table-width')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    all_ = families()
    for name, (body, columns) in all_.items():
        write(os.path.join(out_dir, name + '.docx'), body, columns=columns)
        write(os.path.join(out_dir, name + '-c14.docx'), body, mode=14, columns=columns)
    print(f'{2 * len(all_)} fixtures in {out_dir}')
