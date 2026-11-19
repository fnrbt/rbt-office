#!/usr/bin/env python3
"""Fixtures for where Word breaks a table row across pages (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/row-break.py [out dir]   (default tests/word/fixtures/row-break)

then, to ask Word (winvm) and compare with ours:

    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/row-break --refresh-word
    python3 tools/word-parity/fixtures/row-break.py --labels tests/word/fixtures/row-break

Every fixture is measured in whole lines: Arial 10pt at exactly 20pt a line, no
space before or after, table cells without top or bottom margins or borders, a
Letter page with 1in margins (648pt of body: 32 lines, the 33rd does not fit).
Every line is one label (f01 for the fillers before the table, then the
family's own), each paragraph's lines joined with w:br, so where Word puts each
line does not depend on how text is measured. A family varies one thing - the
number of filler lines (the `_fNN` suffix) moves the table's top down a line at
a time, so the page bottom crosses the row at every line.

--labels reads Word's PDFs (<dir>/word/<name>.pdf, written by --refresh-word)
and writes <dir>/word/labels.tsv: for every fixture the page each label is on
(a fixture without a PDF keeps its rows; the PDFs are not committed). Name
filters after --refresh-word ask Word about some fixtures only.
tests/word/row-break.fsx lays the fixtures out and checks them against it; the
families, what each asks and what Word answered are in its header comment.
"""
import os
import re
import subprocess
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def run(label_lines, widow=False, extra_ppr=''):
    """One paragraph, a line per label (w:br between them)."""
    runs = '<w:r><w:br/></w:r>'.join(f'<w:r><w:t>{t}</w:t></w:r>' for t in label_lines)
    wc = '' if widow is None else '<w:widowControl/>' if widow else '<w:widowControl w:val="0"/>'
    return f'<w:p><w:pPr>{wc}{extra_ppr}</w:pPr>{runs}</w:p>'


def paras(labels, **kw):
    return ''.join(run([t], **kw) for t in labels)


def cell(width, content, vmerge=None):
    vm = '' if vmerge is None else ('<w:vMerge w:val="restart"/>' if vmerge == 'restart' else '<w:vMerge/>')
    if not content:
        content = '<w:p><w:pPr><w:widowControl w:val="0"/></w:pPr></w:p>'
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{vm}</w:tcPr>{content}</w:tc>'


def row(cells, cant_split=False):
    trpr = '<w:trPr><w:cantSplit/></w:trPr>' if cant_split else ''
    return f'<w:tr>{trpr}{"".join(cells)}</w:tr>'


def table(widths, rows):
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            f'<w:tblGrid>{grid}</w:tblGrid>{"".join(rows)}</w:tbl>')


def nested(rows, merge=None, prefix='n', lines=1, cant_split_row=None, width=(1500, 1500)):
    """A two-column nested table of `rows` rows: column 1 merged over rows merge=(first, last)
    (1-based) holding one label m01, else n01.. ; column 2 o01.. (each `lines` lines)."""
    out = []
    for i in range(1, rows + 1):
        lab = lambda p: [f'{p}{i:02d}' + (chr(96 + k) if lines > 1 else '') for k in range(1, lines + 1)]
        if merge and merge[0] <= i <= merge[1]:
            c1 = cell(width[0], run(['m01']) if i == merge[0] else '', 'restart' if i == merge[0] else 'cont')
        else:
            c1 = cell(width[0], run(lab(prefix)))
        c2 = cell(width[1], run(lab('o')))
        out.append(row([c1, c2], cant_split=(cant_split_row == i)))
    return table(list(width), out) + run(['x01'])  # (a cell ends with a paragraph after its table)


def outer(cells):
    """One outer row: cells is [(width, content)]."""
    return table([w for w, _ in cells], [row([cell(w, c) for w, c in cells])])


def filler(n):
    return paras([f'f{i:02d}' for i in range(1, n + 1)])


def document(body, page_h=15840):
    sect = (f'<w:sectPr><w:pgSz w:w="12240" w:h="{page_h}"/><w:pgMar w:top="1440" w:right="1440" '
            'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{run(["end"])}{sect}</w:body></w:document>'


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>'''
SETTINGS = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''


def write(path, body, **kw):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document(body, **kw))
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)


def lines(prefix, n):
    return [f'{prefix}{i:02d}' for i in range(1, n + 1)]


def families():
    """name -> body, one entry per fixture."""
    out = {}
    # Calibration: 32 body lines a page; 32 table rows of a line a page.
    out['calib-body'] = filler(40)
    out['calib-rows'] = filler(10) + table([3000], [row([cell(3000, run([t]))]) for t in lines('r', 40)])
    # Plain cells split at the page bottom (control): one row, one cell of 20 lines.
    for f in range(20, 31, 2):
        out[f'cell-text_f{f}'] = filler(f) + outer([(3000, run(lines('b', 20)))])
    # Two text cells of different line pitch: does each cell break at its own last line that fits?
    for f in range(24, 31):
        out[f'cell-two_f{f}'] = filler(f) + outer([(3000, run(lines('a', 12))),
                                                    (3000, run(lines('b', 6), extra_ppr='<w:spacing w:line="800" w:lineRule="exact"/>'))])
    # A top-level vertical merge over rows 3-6 of 8 (one line each; the merged cell holds m01, or 4 lines).
    for tall in (False, True):
        for f in range(25, 31):
            rows = []
            for i in range(1, 9):
                if 3 <= i <= 6:
                    content = (run(lines('m', 4)) if tall else run(['m01'])) if i == 3 else ''
                    c1 = cell(1500, content, 'restart' if i == 3 else 'cont')
                else:
                    c1 = cell(1500, run([f'n{i:02d}']))
                rows.append(row([c1, cell(1500, run([f'r{i:02d}']))]))
            out[f'top-merge{"-tall" if tall else ""}_f{f}'] = filler(f) + table([1500, 1500], rows)
    # An outer row holding a nested table of 10 one-line rows, with no merge or a merge over rows a-b.
    for merge in (None, (3, 6), (1, 2), (7, 9), (1, 10)):
        tag = 'nest-plain' if merge is None else f'nest-merge{merge[0]}-{merge[1]}'
        for f in range(22, 31):
            out[f'{tag}_f{f}'] = filler(f) + outer([(3000, nested(10, merge)), (1500, run(['b01']))])
    # The same with a text cell of 20 lines beside the nested table: does the text cell break where the
    # nested table does? Nested rows of two lines (break points every other line), and nested with a merge.
    for f in range(22, 31):
        out[f'nest2-text_f{f}'] = filler(f) + outer([(3000, nested(10, lines=2)), (1500, run(lines('b', 20)))])
        out[f'nestmerge-text_f{f}'] = filler(f) + outer([(3000, nested(10, (3, 6))), (1500, run(lines('b', 20)))])
    # A cantSplit row of a nested table over the break (rows 3-6 are one row of 4 lines).
    for f in range(22, 31):
        rows = []
        for i in range(1, 8):
            n = 4 if i == 3 else 1
            rows.append(row([cell(1500, run([f'n{i:02d}' + (chr(96 + k) if n > 1 else '') for k in range(1, n + 1)])),
                             cell(1500, run([f'o{i:02d}']))], cant_split=(i == 3)))
        out[f'nest-cantsplit_f{f}'] = filler(f) + outer([(3000, table([1500, 1500], rows) + run(['x01'])), (1500, run(['b01']))])
    # Two levels: an outer row, a nested table whose second row holds a nested table with a merge.
    for f in range(22, 31):
        inner = nested(6, (2, 5), prefix='k')
        mid = table([3200, 800], [row([cell(3200, run(['p01'])), cell(800, run(['q01']))]),
                                  row([cell(3200, inner), cell(800, run(['q02']))]),
                                  row([cell(3200, run(['p03'])), cell(800, run(['q03']))])]) + run(['y01'])
        out[f'depth2-merge_f{f}'] = filler(f) + outer([(4200, mid), (1500, run(['b01']))])
    # Top-level cantSplit row of 5 lines between rows of one (control).
    for f in range(24, 31):
        out[f'top-cantsplit_f{f}'] = filler(f) + table([3000], [row([cell(3000, run(['r01']))]),
                                                                 row([cell(3000, run(lines('c', 5)))], cant_split=True),
                                                                 row([cell(3000, run(['r03']))])])
    # Taller than a page: a cantSplit row, a nested merge group, a top-level merge group, a plain cell.
    out['tall-cantsplit_f10'] = filler(10) + table([3000], [row([cell(3000, run(lines('c', 40)))], cant_split=True)])
    out['tall-nestmerge_f10'] = filler(10) + outer([(3000, nested(40, (1, 40))), (1500, run(['b01']))])
    out['tall-nestmerge-text_f10'] = filler(10) + outer([(3000, nested(40, (1, 40))), (1500, run(lines('b', 40)))])
    out['tall-text_f10'] = filler(10) + outer([(3000, run(lines('b', 40)))])
    # Widow/orphan control in a cell: a 10-line paragraph with it on, off, and beside a cell without it.
    for f in range(21, 32):
        out[f'widow-on_f{f}'] = filler(f) + outer([(3000, run(lines('w', 10), widow=True))])
        out[f'widow-body_f{f}'] = filler(f) + run(lines('w', 10), widow=True)
        out[f'widow-two_f{f}'] = filler(f) + outer([(3000, run(lines('w', 10), widow=True)), (3000, run(lines('b', 10)))])
    for f in (22, 30, 31):
        out[f'widow-off_f{f}'] = filler(f) + outer([(3000, run(lines('w', 10)))])
    # A cell that would keep nothing on the page: a first line taller than the room (60pt lines),
    # spacing before its first line (40pt), a nested table whose first row is cantSplit, or whose
    # first row has a cell orphan control empties - beside a cell of 20 plain lines.
    tall_line = '<w:spacing w:line="1200" w:lineRule="exact"/>'
    space_before = '<w:spacing w:before="800"/>'
    for f in range(26, 32):
        out[f'first-tall_f{f}'] = filler(f) + outer([(3000, run(lines('a', 20))), (3000, run(lines('b', 3), extra_ppr=tall_line))])
        out[f'first-space_f{f}'] = filler(f) + outer([(3000, run(lines('a', 20))), (3000, run(lines('b', 5), extra_ppr=space_before))])
        first_cant = table([1500, 1500], [row([cell(1500, run(lines('n', 4))), cell(1500, run(['o01']))], cant_split=True),
                                          row([cell(1500, run(['n05'])), cell(1500, run(['o05']))])]) + run(['x01'])
        out[f'nest-first-cantsplit_f{f}'] = filler(f) + outer([(3000, first_cant), (3000, run(lines('b', 10)))])
        orphan_row = lambda: row([cell(1500, run(lines('w', 4), widow=True)), cell(1500, run(lines('v', 4)))])
        first_orphan = table([1500, 1500], [orphan_row(), row([cell(1500, run(['n05'])), cell(1500, run(['o05']))])]) + run(['x01'])
        out[f'nest-first-orphan_f{f}'] = filler(f) + outer([(3000, first_orphan), (3000, run(lines('b', 10)))])
        later_orphan = table([1500, 1500], [row([cell(1500, run(['n01'])), cell(1500, run(['o01']))]), orphan_row(),
                                            row([cell(1500, run(['n05'])), cell(1500, run(['o05']))])]) + run(['x01'])
        out[f'nest-orphan_f{f}'] = filler(f) + outer([(3000, later_orphan), (3000, run(lines('b', 10)))])
        # Lines of exactly 40pt in the body: does the whole line have to fit, or its text?
        out[f'exact-body_f{f}'] = filler(f) + run(lines('e', 4), extra_ppr='<w:spacing w:line="800" w:lineRule="exact"/>')
    # No w:widowControl anywhere (styles and docDefaults have none either): on or off?
    for f in (23, 31):
        out[f'widow-unset_f{f}'] = filler(f) + outer([(3000, run(lines('w', 10), widow=None)), (3000, run(lines('b', 10)))])
        out[f'widow-unset-body_f{f}'] = filler(f) + run(lines('w', 10), widow=None)
    # Keep lines together (w:keepLines) in a cell beside a plain one: does it make a cell keep nothing?
    for f in range(26, 31):
        out[f'keeplines-two_f{f}'] = filler(f) + outer([(3000, run(lines('k', 5), extra_ppr='<w:keepLines/>')), (3000, run(lines('b', 10)))])
    # Behind a repeated header row: the row a cell of which keeps nothing goes to the next page,
    # under the header again.
    for f in range(27, 32):
        header = '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + cell(3000, run(['h01'])) + cell(3000, run(['h02'])) + '</w:tr>'
        rows = [header, row([cell(3000, run(['r01'])), cell(3000, run(['s01']))]),
                row([cell(3000, run(lines('a', 10))), cell(3000, run(lines('b', 3), extra_ppr=tall_line))])]
        out[f'header-first-tall_f{f}'] = filler(f) + table([3000, 3000], rows)
    return out


SMALL = {  # A half-height page (7in: body 360pt, 18 lines): the merge rule does not depend on the page.
    f'small-nest-merge3-6_f{f}': (lambda f: filler(f) + outer([(3000, nested(10, (3, 6))), (1500, run(['b01']))]))(f)
    for f in range(5, 13)
} | {
    f'small-nest-plain_f{f}': (lambda f: filler(f) + outer([(3000, nested(10)), (1500, run(['b01']))]))(f)
    for f in range(5, 13)
}


def labels(folder):
    """<folder>/word/labels.tsv from Word's PDFs: name, label, page."""
    word = os.path.join(folder, 'word')
    rows, seen = [], set()
    for name in sorted(f[:-5] for f in os.listdir(folder) if f.endswith('.docx')):
        pdf = os.path.join(word, name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        seen.add(name)
        text = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True, check=True).stdout
        for page, chunk in enumerate(text.split('\f'), 1):
            for lab in re.findall(r'\b([a-z]\d\d[a-z]?|end)\b', chunk):
                rows.append(f'{name}\t{lab}\t{page}')
    # (Fixtures whose PDF is not here keep what labels.tsv had for them.)
    tsv = os.path.join(word, 'labels.tsv')
    if os.path.exists(tsv):
        rows += [l for l in open(tsv).read().splitlines() if l and l.split('\t')[0] not in seen]
    rows.sort(key=lambda l: l.split('\t')[0])
    with open(tsv, 'w') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'{len(rows)} labels')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--labels':
        labels(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/row-break')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    all_ = families()
    for name, body in all_.items():
        write(os.path.join(out_dir, name + '.docx'), body)
    for name, body in SMALL.items():
        write(os.path.join(out_dir, name + '.docx'), body, page_h=10080)
    print(f'{len(all_) + len(SMALL)} fixtures in {out_dir}')
