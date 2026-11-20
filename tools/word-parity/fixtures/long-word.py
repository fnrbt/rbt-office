#!/usr/bin/env python3
"""Fixtures for how Word breaks a word longer than its line (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/long-word.py [out dir]   (default tests/word/fixtures/long-word)

then, to ask Word (winvm) and compare with ours:

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/long-word \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/long-word --refresh-word
    python3 tools/word-parity/fixtures/long-word.py --lines tests/word/fixtures/long-word

--lines reads Word's PDFs (<dir>/word/<name>.pdf, not committed) and writes <dir>/word/lines.tsv:
for every fixture, every line of its body in order - page, the line's top in points, its text
(pdftotext's words joined by a space). tests/word/long-word.fsx lays the fixtures out and checks
our lines against it.

Every fixture but the hf-* ones is in Courier New 10pt - every character 6.0pt (1229/2048 em), no
kerning pairs - at exactly 20pt a line, nothing before or after, on a Letter page with 1in margins.
A line holds 30 characters and not 31: the body paragraphs have a right indent leaving 183pt; a
table cell is 183pt wide inside its margins (fixed layout). The text: a prefix word of p 'a's, a
space, a word of L characters (the digits 0-9 over and over), a space and "end":

  body-p<p>-L<L>, cell-p<p>-L<L>   p = 0, 1, 10, 20, 28 (one character's room left), 29 (none);
                                   L = 29, 30 (fits a line), 31, 45, 70 (longer than a line)
  body-p10-L45-c14, cell-...-c14   the same in compatibility mode 14
  runs-p10-L45                     the long word in a run of its own, the prefix's run ending in its space
  glued-p10-L45                    the long word's first character in the prefix's run (a word over two runs)
  sect-L45                         "aaaa " then "§" in one run and the long word in the next (no space between)
  hyph-p5                          "aaaaa " then 35 digits, a hyphen and 10 more (the part before the hyphen is too long)
  two-long                         "aa ", 40 digits, a space, 40 digits
  foot_f<f>                        f filler lines (f27 ... f32: room for 5 ... 0 lines), then a one-cell row
                                   holding "aaaa " + 130 digits + " end" (6 lines) with widow control
  fld-<kind>-<pre>, fldcell-...    five 10-letter words as a QUOTE field's result (kind quote), a link's
                                   text (link) or plain text (plain), after nothing, "aaaa " or a glued "\u00a7"
  fld-quote-glued-tworuns          the same, the result in two runs
  fldq-p<p>-L<L>, fldqcell-...     body-p<p>-L<L>'s text as a QUOTE field's result (p 0, 10, 28; L 29 - 45)
  fldq-glued-L45                   "aaaa \u00a7" then the 45 digits as a field's result
  hf-*                             docx4j-headers-footers' narrow cell: Calibri 11pt, a table of grid 8101 /
                                   1155 twips at 5000 pct, cell margins 115, "\u00a7" then "Error! Reference
                                   source not found." in bold - hf-quote as a QUOTE field's result (regular in
                                   hf-quote-regular), hf-plain as plain text, hf-kern2 with w:kern 2 on it
                                   (kerning at 1pt and up), hf-foot_f<f> the hf-quote row under f lines of
                                   13.43pt (the page bottom crossing the cell)

A REF field is not in the family: Word's PDF export updates it (a REF to a missing bookmark becomes
"Error! Reference source not found." again) and lays the new result out its own way - in
docx4j-headers-footers' cell "\u00a7Error! R" / "eference", where Word's layout (Information(5/6) of
each character, and its pages) has "\u00a7Error! " / "Reference" as for plain text. The PDF of a
REF field's line is not Word's layout of it.
No dependencies: the package is written as XML.
"""
import os
import re
import subprocess
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
DIGITS = '0123456789'


def digits(n):
    return (DIGITS * (n // 10 + 1))[:n]


def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def r(text, rpr=''):
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


BODY_IND = '<w:ind w:right="5700"/>'  # 468 - 183 = 285pt


def p(runs, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>'


def cell_table(content, width=3660, margin=108, extra_tblpr=''):
    total = width + 2 * margin
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:tblLayout w:type="fixed"/>{extra_tblpr}'
            f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="{margin}" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="{margin}" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{total}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="{total}" w:type="dxa"/></w:tcPr>'
            f'{content}</w:tc></w:tr></w:tbl>')


def filler(n, ppr=''):
    return ''.join(p(r(f'f{i:02d}'), ppr) for i in range(1, n + 1))


def sect(page_h=15840):
    return (f'<w:sectPr><w:pgSz w:w="12240" w:h="{page_h}"/><w:pgMar w:top="1440" w:right="1440" '
            'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


COURIER_STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:eastAsia="Courier New"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:widowControl/><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>'''

# docx4j-headers-footers' styles, cut down to what its element tables use.
HF_STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-AU"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders><w:tblCellMar><w:top w:w="43" w:type="dxa"/><w:left w:w="115" w:type="dxa"/><w:bottom w:w="43" w:type="dxa"/><w:right w:w="115" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>'''


def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>'
            '</w:compat></w:settings>')


TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''


def write(path, body, styles=COURIER_STYLES, mode=15):
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect()}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/settings.xml', settings(mode))


def text(pre, length):
    return ('a' * pre + ' ' if pre else '') + digits(length) + ' end'


def hf_cell(result_runs):
    """docx4j-headers-footers' element table row: a wide cell and the narrow one holding `result_runs`."""
    cell1 = '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="pct"/></w:tcPr>' + p(r('altChunk (Anchor for Imported External Content)')) + '</w:tc>'
    cell2 = '<w:tc><w:tcPr><w:tcW w:w="500" w:type="pct"/></w:tcPr>' + p(r('\u00a7') + result_runs) + '</w:tc>'
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="8101"/><w:gridCol w:w="1155"/></w:tblGrid>'
            f'<w:tr>{cell1}{cell2}</w:tr></w:tbl>')


RESULT = 'Error! Reference source not found.'
BOLD = '<w:b/><w:bCs/><w:lang w:val="en-US"/>'


def complex_field(code, result_runs):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve">{esc(code)}</w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>{result_runs}<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def families():
    """name -> (body, styles, mode)"""
    out = {}
    for pre in (0, 1, 10, 20, 28, 29):
        for length in (29, 30, 31, 45, 70):
            t = text(pre, length)
            out[f'body-p{pre}-L{length}'] = (p(r(t), BODY_IND), COURIER_STYLES, 15)
            out[f'cell-p{pre}-L{length}'] = (cell_table(p(r(t))) + p(r('x')), COURIER_STYLES, 15)
    t = text(10, 45)
    out['body-p10-L45-c14'] = (p(r(t), BODY_IND), COURIER_STYLES, 14)
    out['cell-p10-L45-c14'] = (cell_table(p(r(t))) + p(r('x')), COURIER_STYLES, 14)
    out['runs-p10-L45'] = (p(r('a' * 10 + ' ') + r(digits(45)) + r(' end'), BODY_IND), COURIER_STYLES, 15)
    out['glued-p10-L45'] = (p(r('a' * 10 + ' 0') + r(digits(45)[1:] + ' end'), BODY_IND), COURIER_STYLES, 15)
    out['sect-L45'] = (p(r('aaaa \u00a7') + r(digits(45) + ' end', '<w:b/>'), BODY_IND), COURIER_STYLES, 15)
    out['hyph-p5'] = (p(r('aaaaa ' + digits(35) + '-' + digits(10) + ' end'), BODY_IND), COURIER_STYLES, 15)
    out['two-long'] = (p(r('aa ' + digits(40) + ' ' + digits(40) + ' end'), BODY_IND), COURIER_STYLES, 15)
    for f in range(27, 33):
        out[f'foot_f{f:02d}'] = (filler(f) + cell_table(p(r(text(4, 130)), '<w:widowControl/>')) + p(r('x')), COURIER_STYLES, 15)
    # Field results (Courier): five 10-letter words, with nothing, "aaaa " or a glued "\u00a7" before the field.
    words = ' '.join(c * 10 for c in 'bcdef')
    kinds = {
        'quote': lambda t: complex_field(' QUOTE "x" ', r(t)),
        'link': lambda t: f'<w:hyperlink w:anchor="missing">{r(t)}</w:hyperlink>',
        'plain': lambda t: r(t),
    }
    for kind, make in kinds.items():
        for pre, before in (('none', ''), ('space', 'aaaa '), ('glued', '\u00a7')):
            out[f'fld-{kind}-{pre}'] = (p((r(before) if before else '') + make(words) + r(' end'), BODY_IND), COURIER_STYLES, 15)
            out[f'fldcell-{kind}-{pre}'] = (cell_table(p((r(before) if before else '') + make(words) + r(' end'))) + p(r('x')), COURIER_STYLES, 15)
    # The field result's first word does not fit after the text before it, in a separate run of the result.
    out['fld-quote-glued-tworuns'] = (p(r('\u00a7') + complex_field(' QUOTE "x" ', r('bbbbbbbbbb ') + r(' '.join(c * 10 for c in 'cdef'))) + r(' end'), BODY_IND), COURIER_STYLES, 15)
    # A QUOTE field's result (Word keeps it as cached; a REF to a missing bookmark it makes "Error! Reference
    # source not found." again): the text of body-p<p>-L<L> as the result, " end" after the field.
    quote = lambda t: complex_field(' QUOTE "x" ', r(t))
    for pre in (0, 10, 28):
        for length in (29, 30, 31, 45):
            t = ('a' * pre + ' ' if pre else '') + digits(length)
            out[f'fldq-p{pre}-L{length}'] = (p(quote(t) + r(' end'), BODY_IND), COURIER_STYLES, 15)
    for length in (30, 31):
        out[f'fldqcell-p10-L{length}'] = (cell_table(p(quote('a' * 10 + ' ' + digits(length)) + r(' end'))) + p(r('x')), COURIER_STYLES, 15)
    out['fldq-glued-L45'] = (p(r('aaaa \u00a7') + quote(digits(45)) + r(' end'), BODY_IND), COURIER_STYLES, 15)
    out['hf-quote'] = (hf_cell(complex_field(' QUOTE "x" ', r(RESULT, BOLD))) + p(r('x')), HF_STYLES, 14)
    out['hf-quote-regular'] = (hf_cell(complex_field(' QUOTE "x" ', r(RESULT))) + p(r('x')), HF_STYLES, 14)
    out['hf-plain'] = (hf_cell(r(RESULT, BOLD)) + p(r('x')), HF_STYLES, 14)
    out['hf-kern2'] = (hf_cell(r(RESULT, BOLD + '<w:kern w:val="2"/>')) + p(r('x')), HF_STYLES, 14)
    # What of a cell fits at the foot of the page: the cell's bottom margin (top 0), widow control unset in
    # compatibility mode 14 and 15. Lines w01 ... of a paragraph joined with w:br.
    def brs(labels, ppr=''):
        return p('<w:r><w:br/></w:r>'.join(r(t) for t in labels), ppr)
    eight = [f'w{i:02d}' for i in range(1, 9)]
    for m in (0, 43, 200, 400):
        for f in (27, 28):
            out[f'mar-b{m}_f{f}'] = (filler(f) + cell_table(brs(eight, '<w:widowControl w:val="0"/>'), extra_tblpr='')
                                     .replace('<w:bottom w:w="0" w:type="dxa"/><w:right', f'<w:bottom w:w="{m}" w:type="dxa"/><w:right')
                                     + p(r('x')), COURIER_STYLES, 15)
    ten = [f'w{i:02d}' for i in range(1, 11)]
    nowc = COURIER_STYLES.replace('<w:widowControl/>', '')
    for mode in (14, 15):
        for f in (23, 31):
            out[f'wc-unset-c{mode}-cell_f{f}'] = (filler(f) + cell_table(brs(ten)) + p(r('x')), nowc, mode)
            out[f'wc-unset-c{mode}-body_f{f}'] = (filler(f) + brs(ten) + p(r('x')), nowc, mode)
    for f in (23, 31):
        out[f'wc-on-c14-cell_f{f}'] = (filler(f) + cell_table(brs(ten, '<w:widowControl/>')) + p(r('x')), nowc, 14)
        out[f'wc-on-c14-body_f{f}'] = (filler(f) + brs(ten, '<w:widowControl/>') + p(r('x')), nowc, 14)
        out[f'wc-on-c12-cell_f{f}'] = (filler(f) + cell_table(brs(ten, '<w:widowControl/>')) + p(r('x')), nowc, 12)
    # The cell's top below f single Calibri lines of 13.43pt in a 1in-margin Letter page (720pt down).
    single = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
    for f in (44, 45, 46, 47):
        out[f'hf-foot_f{f}'] = (filler(f, single) + hf_cell(complex_field(' QUOTE "x" ', r(RESULT, BOLD))) + p(r('x')), HF_STYLES, 14)
    for f in (45, 46):
        cell = hf_cell(complex_field(' QUOTE "x" ', r(RESULT, BOLD)))
        out[f'hf-foot-c15_f{f}'] = (filler(f, single) + cell + p(r('x')), HF_STYLES, 15)
        out[f'hf-foot-wc_f{f}'] = (filler(f, single) + cell.replace('<w:p><w:pPr></w:pPr><w:r><w:t xml:space="preserve">\u00a7', '<w:p><w:pPr><w:widowControl/></w:pPr><w:r><w:t xml:space="preserve">\u00a7') + p(r('x')), HF_STYLES, 14)
        out[f'hf-foot-nomar_f{f}'] = (filler(f, single) + cell.replace('<w:tblW w:w="5000" w:type="pct"/>', '<w:tblW w:w="5000" w:type="pct"/><w:tblCellMar><w:top w:w="43" w:type="dxa"/><w:left w:w="115" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="115" w:type="dxa"/></w:tblCellMar>') + p(r('x')), HF_STYLES, 14)
    return out


def lines(folder):
    """<folder>/word/lines.tsv from Word's PDFs: name, page, top, text."""
    word = os.path.join(folder, 'word')
    rows, seen = [], set()
    for name in sorted(f[:-5] for f in os.listdir(folder) if f.endswith('.docx')):
        pdf = os.path.join(word, name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        seen.add(name)
        out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        page = 0
        for chunk in out.split('<page ')[1:]:
            page += 1
            words = [(float(m.group(2)), float(m.group(1)), m.group(3)) for m in
                     re.finditer(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">(.*?)</word>', chunk)]
            # A line: the words whose tops are within 2pt, left to right.
            groups = []
            for top, x, t in sorted(words):
                if groups and abs(groups[-1][0] - top) < 2.0:
                    groups[-1][1].append((x, t))
                else:
                    groups.append((top, [(x, t)]))
            for top, ws in groups:
                rows.append(f'{name}\t{page}\t{top:.2f}\t{" ".join(t for _, t in sorted(ws))}')
    tsv = os.path.join(word, 'lines.tsv')
    if os.path.exists(tsv):
        rows += [l for l in open(tsv, encoding='utf-8').read().splitlines() if l and l.split('\t')[0] not in seen]
    rows.sort(key=lambda l: (l.split('\t')[0], int(l.split('\t')[1]), float(l.split('\t')[2])))
    with open(tsv, 'w', encoding='utf-8') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'{len(rows)} lines')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--lines':
        lines(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/long-word')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    all_ = families()
    for name, (body, styles, mode) in all_.items():
        write(os.path.join(out_dir, name + '.docx'), body, styles, mode)
    print(f'{len(all_)} fixtures in {out_dir}')
