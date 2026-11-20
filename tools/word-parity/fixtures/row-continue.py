#!/usr/bin/env python3
"""Fixtures for where the part of a table row broken across pages starts on the next page
(docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/row-continue.py [out dir]   (default tests/word/fixtures/row-continue)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/row-continue \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/row-continue --refresh-word
    python3 tools/word-parity/fixtures/row-continue.py --tops tests/word/fixtures/row-continue

--tops reads Word's PDFs (<dir>/word/<name>.pdf, not committed) and writes <dir>/word/tops.tsv: every
label's fixture, page and top (pdftotext's box, points). tests/word/row-continue.fsx checks ours.

The showcase handbook's fault-code table (Word 2010's layout): the part of a row that goes on on
the next page starts as far below the repeated header row as a row of its own does there - its
cell's top margin and the row's top border above its text again. As the row-break fixtures (whose
helpers this uses): Arial 10pt at exactly 20pt a line, Letter, 1in margins, 28 filler lines (f01 ..),
then a table of two fixed columns: a header row (h01, w:tblHeader) or none, then a row of a cell of 6
lines (r01 - r06, one paragraph) beside one of one (s01), then a row of one line (t01 / u01). It varies:
  h0 / h1       no header row / a repeated header row
  m0 / m100     the cells' top and bottom margins (tblCellMar), twips
  b0 / b4 / b24 no borders / single borders of 4 or 24 eighths (all six of tblBorders)
  c14 / c15     compatibility mode
"""
import importlib.util
import os
import re
import subprocess
import sys
import zipfile

here = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('row_break', os.path.join(here, 'row-break.py'))
rb = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rb)


def table(rows, mar, border):
    borders = ''
    if border:
        borders = '<w:tblBorders>' + ''.join(
            f'<w:{s} w:val="single" w:sz="{border}" w:space="0" w:color="000000"/>'
            for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) + '</w:tblBorders>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>{borders}'
            f'<w:tblCellMar><w:top w:w="{mar}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
            f'<w:bottom w:w="{mar}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' + ''.join(rows) + '</w:tbl>')


def body(header, mar, border):
    rows = []
    if header:
        rows.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>' + rb.cell(3000, rb.run(['h01'])) + rb.cell(3000, rb.run(['h02'])) + '</w:tr>')
    rows.append(rb.row([rb.cell(3000, rb.run(rb.lines('r', 6))), rb.cell(3000, rb.run(['s01']))]))
    rows.append(rb.row([rb.cell(3000, rb.run(['t01'])), rb.cell(3000, rb.run(['u01']))]))
    return rb.filler(28) + table(rows, mar, border)


def settings(mode):
    return rb.SETTINGS.replace('w:val="15"', f'w:val="{mode}"')


def write(path, xml, mode):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in (('[Content_Types].xml', rb.TYPES), ('_rels/.rels', rb.RELS), ('word/_rels/document.xml.rels', rb.DOC_RELS),
                           ('word/document.xml', rb.document(xml)), ('word/styles.xml', rb.STYLES), ('word/settings.xml', settings(mode))):
            z.writestr(zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0)), data, zipfile.ZIP_DEFLATED)


def docs():
    out = {}
    for mode in (14, 15):
        for header in (0, 1):
            for mar in (0, 100):
                for border in (0, 4, 24):
                    out[f'h{header}-m{mar}-b{border}-c{mode}'] = (body(header, mar, border), mode)
    return out


def tops(folder):
    rows = []
    for name in sorted(f[:-5] for f in os.listdir(folder) if f.endswith('.docx')):
        pdf = os.path.join(folder, 'word', name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        for page, chunk in enumerate(out.split('<page ')[1:], 1):
            for m in re.finditer(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">(.*?)</word>', chunk):
                rows.append(f'{name}\t{page}\t{float(m.group(1)):.2f}\t{float(m.group(2)):.2f}\t{m.group(3)}')
    with open(os.path.join(folder, 'word', 'tops.tsv'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'{len(rows)} words')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(here, '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--tops':
        tops(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/row-continue')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    ds = docs()
    for name, (xml, mode) in ds.items():
        write(os.path.join(out_dir, name + '.docx'), xml, mode)
    print(f'{len(ds)} fixtures in {out_dir}')
