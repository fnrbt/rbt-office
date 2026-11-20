#!/usr/bin/env python3
"""Fixtures for how far Word narrows a justified line's spaces to fit more text on it (Word 2013's
layout; docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/justify-squeeze.py [out dir]   (default tests/word/fixtures/justify-squeeze)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/justify-squeeze \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/justify-squeeze --refresh-word
    python3 tools/word-parity/fixtures/justify-squeeze.py --table tests/word/fixtures/justify-squeeze

--table reads Word's PDFs (<dir>/word/<name>.pdf) - or with --pdfs <dir> another set of PDFs of the same
documents (tools/pdf-runner's) - and writes <dir>/word/squeeze.tsv: per probe whether its first line held
all its tokens (1) or gave the last one to the next line (0), and where the first line's last word ends.

A probe is a paragraph: n spaces between n + 1 tokens of digits (equal widths in every font here, never
hyphenated) as wide as whole digits take them to 360pt at most, then " end". A right indent (twips) makes
the line narrower than that text by n x o, o the probe's overflow per space: the first line holds every
token only when Word lets each space give up o. (Not character spacing on a digit: ours does not count
it after a run's last character - TextBuffer.characterSpacingSlots -, Word does.)
A paragraph's spacing after (one line) parts the probes. Probes o = s x 0.00, 0.02 ... 0.40 (s: the
space's width), or for `fine-*` 0.240 ... 0.262 of it in steps of 0.002, round the boundary found.

  <font><size>-n<n>[-m<mode>]   fonts cour (Courier New: a 0.6em space), cal (Calibri 0.226em), times
                                (0.25em), arial (0.278em); sizes 10 and 20pt; n 1, 2, 4, 9 spaces; mode
                                15 (default), 14, 12, none (no compatibilityMode)
  jc-<jc>[-m14]                 cour10-n4 under left / center / right / distribute / thaiDistribute
  last-<font><size>-n<n>        no " end": the probe's line is its paragraph's last
  hyph-<jc>                     cour10-n4 with w:autoHyphenation, the last token accessibility (broken
                                ac-ces-si-bil-ity); does Word narrow spaces or hyphenate?
  hyph-part                     the same, the line "accessibil-" over by n x o (the whole word never fits)
  hyph-part<n>-<jc>             "accessibil-" over by 0 ... 9pt in all (0.5pt steps), n 1, 2, 4, 9, both / left
  cell-cour10-n4 / cols-cour10-n4   the same in a fixed 360pt table cell / in the first of two 360pt columns
No dependencies: the package is written as XML.
"""
import os
import re
import subprocess
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
LINE = 360.0
# Advances (units per 2048 em) of the space and of the digits (all one width) - the fonts' hmtx.
FONTS = {'cour': ('Courier New', 1229, 1229), 'cal': ('Calibri', 463, 1038),
         'times': ('Times New Roman', 512, 1024), 'arial': ('Arial', 569, 1139)}
STEPS = [i * 0.02 for i in range(21)]


def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def r(text, rpr=''):
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def probe_runs(font, size, n, over, last=False, word=None, part=None):
    """The probe's runs, its right indent (twips) and its overflow: n + 1 tokens as wide as whole digits
    take them to LINE at most, the line narrowed by a right indent to the text's width less n x over.
    word: the last token (its width in digits' advances: Courier only); part: the characters of it (its
    hyphen included) that are that much over the line instead."""
    _, sp, dg = FONTS[font]
    s, d = sp * size / 2048, dg * size / 2048
    tail = part if part else len(word) if word else 0
    digits = int((LINE - n * s) // d) - tail
    tokens = n + 1 - (1 if word else 0)
    sizes = [digits // tokens + (1 if i < digits % tokens else 0) for i in range(tokens)]
    natural = (digits + tail) * d + n * s
    indent = round((LINE - natural + n * over) * 20)
    text = ' '.join(str(i % 10) * k for i, k in enumerate(sizes)) + (' ' + word if word else '')
    return r(text + ('' if last else ' end')), indent, natural - (LINE - indent / 20), s


def ppr(jc='both', after=True, size=10, indent=0):
    a = f' w:after="{size * 40}"' if after else ''
    return f'<w:pPr><w:spacing w:before="0"{a}/><w:ind w:right="{indent}"/><w:jc w:val="{jc}"/></w:pPr>'


def styles(font, size):
    face = FONTS[font][0]
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
            f'<w:rPrDefault><w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:cs="{face}" w:eastAsia="{face}"/>'
            f'<w:sz w:val="{size * 2}"/><w:szCs w:val="{size * 2}"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
            f'<w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{size * 40}" '
            'w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>'
            '<w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
            '</w:styles>')


def settings(mode=15, hyph=False):
    h = '<w:autoHyphenation/>' if hyph else ''
    c = (f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" '
         f'w:val="{mode}"/></w:compat>') if mode else ''
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>{h}{c}</w:settings>'


def sect(cols=False):
    if cols:
        # Two 360pt columns 18pt apart: a page 882pt wide with 1in margins.
        return ('<w:sectPr><w:pgSz w:w="17640" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
                'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="360"/></w:sectPr>')
    # A 360pt line: Letter, 1in margins, 108pt off the right.
    return ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="3600" w:bottom="1440" '
            'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''


def write(path, body, st, se, cols=False):
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect(cols)}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', st)
        z.writestr('word/settings.xml', se)


def cell(content):
    total = int(LINE * 20) + 216
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{total}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="{total}" w:type="dxa"/></w:tcPr>'
            f'{content}</w:tc></w:tr></w:tbl><w:p/>')


def family(out, name, font, size, n, mode=15, jc='both', last=False, word=None, hyph=False, where='body', steps=None, part=None):
    """One document of probes; returns its manifest rows."""
    rows, paras = [], []
    _, sp, _ = FONTS[font]
    s = sp * size / 2048
    for i, f in enumerate(steps if steps is not None else STEPS):
        over = f * s if steps is None else f
        runs, indent, actual, s = probe_runs(font, size, n, over, last, word, part)
        paras.append(f'<w:p>{ppr(jc, size=size, indent=indent)}{runs}</w:p>')
        rows.append(f'{name}\t{i}\t{font}\t{size}\t{n}\t{s:.4f}\t{actual / n:.4f}\t{jc}\t{mode}\t{int(last)}\t{where}')
    body = ''.join(paras)
    if where == 'cell':
        body = cell(body)
    write(os.path.join(out, name + '.docx'), body, styles(font, size), settings(mode, hyph), cols=(where == 'cols'))
    return rows


def generate(out):
    os.makedirs(out, exist_ok=True)
    for f in os.listdir(out):
        if f.endswith('.docx'):
            os.remove(os.path.join(out, f))
    rows = []
    for font, size, ns in [('cour', 10, (1, 2, 4, 9)), ('cour', 20, (4,)), ('cal', 10, (4, 9)), ('cal', 20, (4,)),
                           ('times', 10, (4, 9)), ('arial', 10, (4,))]:
        for n in ns:
            rows += family(out, f'{font}{size}-n{n}', font, size, n)
    for mode in (14, 12, None):
        rows += family(out, f'cour10-n4-m{mode or "none"}', 'cour', 10, 4, mode=mode)
        rows += family(out, f'cal10-n4-m{mode or "none"}', 'cal', 10, 4, mode=mode)
    for jc in ('left', 'center', 'right', 'distribute', 'thaiDistribute'):
        rows += family(out, f'jc-{jc}', 'cour', 10, 4, jc=jc)
    rows += family(out, 'jc-distribute-m14', 'cour', 10, 4, jc='distribute', mode=14)
    rows += family(out, 'last-cour10-n4', 'cour', 10, 4, last=True)
    rows += family(out, 'last-cal10-n4', 'cal', 10, 4, last=True)
    for jc in ('both', 'left'):
        rows += family(out, f'hyph-{jc}', 'cour', 10, 4, jc=jc, word='accessibility', hyph=True)
    rows += family(out, 'hyph-both-m14', 'cour', 10, 4, word='accessibility', hyph=True, mode=14)
    # The whole word never fits; "accessibil-" (11 characters) is over the line by n x the probe's overflow.
    rows += family(out, 'hyph-part', 'cour', 10, 4, word='accessibility', part=11, hyph=True)
    # The same over 0 ... 9pt in all (0.5pt steps), 1, 2, 4 and 9 spaces, justified and left: what lets
    # "accessibil-" stay - its spaces narrowed (a quarter of 6pt each), or its hyphen not counted (6pt)?
    for n in (1, 2, 4, 9):
        for jc in ('both', 'left'):
            rows += family(out, f'hyph-part{n}-{jc}', 'cour', 10, n, jc=jc, word='accessibility', part=11, hyph=True,
                           steps=[0.5 * i / n for i in range(19)])
    for font, size in [('cour', 10), ('cour', 20), ('cal', 10), ('times', 10), ('arial', 10)]:
        s = FONTS[font][1] * size / 2048
        rows += family(out, f'fine-{font}{size}-n4', font, size, 4, steps=[s * (0.240 + 0.002 * i) for i in range(12)])
    rows += family(out, 'fine-cal10-n9', 'cal', 10, 9, steps=[FONTS['cal'][1] * 10 / 2048 * (0.240 + 0.002 * i) for i in range(12)])
    rows += family(out, 'cell-cour10-n4', 'cour', 10, 4, where='cell')
    rows += family(out, 'cols-cour10-n4', 'cour', 10, 4, where='cols')
    with open(os.path.join(out, 'probes.tsv'), 'w') as f:
        f.write('doc\tprobe\tfont\tsize\tspaces\tspace\toverflow_per_space\tjc\tmode\tlast\twhere\n' + '\n'.join(rows) + '\n')
    print(len(set(r.split('\t')[0] for r in rows)), 'documents,', len(rows), 'probes')


def table(folder, pdfs=None):
    """<folder>/word/squeeze.tsv: per probe its first line's word count against its tokens, and its right end."""
    probes = [l.split('\t') for l in open(os.path.join(folder, 'probes.tsv')).read().splitlines()[1:]]
    by = {}
    for p in probes:
        by.setdefault(p[0], []).append(p)
    out = []
    src = pdfs or os.path.join(folder, 'word')
    for name, ps in by.items():
        pdf = os.path.join(src, name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        text = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        # Lines in reading order (page, column - the second starts right of the first's 360pt -, top); a
        # probe's paragraph ends with its line ending in "end" - a last-line probe (no " end") at a gap of
        # more than 1.5 lines.
        ordered, page = [], 0
        last_doc = ps[0][9] == '1'
        for chunk in text.split('<page ')[1:]:
            page += 1
            words = [(float(m.group(2)), float(m.group(1)), float(m.group(3)), m.group(4)) for m in
                     re.finditer(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="[\d.]+">(.*?)</word>', chunk)]
            lines = {}
            for top, x0, x1, t in sorted(words):
                col = 1 if x0 > 72 + LINE + 9 else 0
                key = next((k for k in lines if k[0] == col and abs(k[1] - top) < 2.0), (col, top))
                lines.setdefault(key, []).append((x0, x1, t))
            for (col, top) in sorted(lines):
                ordered.append((page, col, top, sorted(lines[(col, top)])))
        step = int(ps[0][3]) * 2
        paras = []
        for page, col, top, ws in ordered:
            if last_doc:
                if paras and paras[-1][0] == (page, col) and top - paras[-1][1] < step * 1.5:
                    paras[-1][2].append(ws)
                    paras[-1][1] = top
                else:
                    paras.append([(page, col), top, [ws]])
            else:
                if not paras or paras[-1][3]:
                    paras.append([(page, col), top, [], False])
                paras[-1][2].append(ws)
                paras[-1][3] = ''.join(w[2] for w in ws).endswith('end')  # (distribute spreads e n d)
        for p, para in zip(ps, paras):
            first = para[2][0]
            # (The probe fits when the line after it is " end" alone - or, for a last line, when there is none.)
            last = p[9] == '1'
            fit = (1 if len(para[2]) == 1 else 0) if last else (1 if ''.join(w[2] for w in para[2][1]) == 'end' else 0) if len(para[2]) > 1 else 0
            if name.startswith('hyph-part') and len(para[2]) > 1:
                # (Its first line kept "accessibil-": the next starts "ity".)
                fit = 1 if para[2][1][0][2].startswith('ity') else 0
            out.append('\t'.join(p[:1] + [p[1], p[6], str(fit), f'{first[0][0]:.2f}', f'{first[-1][1]:.2f}',
                                          str(len(para[2])), ' / '.join(' '.join(w[2] for w in l) for l in para[2])]))
        if len(paras) != len(ps):
            print(f'{name}: {len(paras)} paragraphs for {len(ps)} probes', file=sys.stderr)
    dest = os.path.join(folder, 'word', 'squeeze.tsv') if not pdfs else os.path.join(pdfs, 'squeeze.tsv')
    with open(dest, 'w') as f:
        f.write('doc\tprobe\toverflow_per_space\tfits\tleft\tright\tlines\ttext\n' + '\n'.join(out) + '\n')
    # The boundary per document: the largest overflow per space that fits, the smallest that does not.
    for name, ps in by.items():
        rs = [o for o in out if o.split('\t')[0] == name]
        fits = [float(o.split('\t')[2]) for o in rs if o.split('\t')[3] == '1']
        nots = [float(o.split('\t')[2]) for o in rs if o.split('\t')[3] == '0']
        s = float(ps[0][5])
        mf, mn = max(fits, default=None), min(nots, default=None)
        print(f'{name:22s} space {s:6.3f}  fits up to {mf if mf is None else round(mf, 3)} '
              f'({"-" if mf is None else f"{mf / s:.3f}"} of the space), not from {mn if mn is None else round(mn, 3)}'
              + ('' if mf is None or mn is None or mf < mn else '  (not monotonic)'))


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--table':
        table(sys.argv[2], sys.argv[4] if len(sys.argv) > 4 and sys.argv[3] == '--pdfs' else None)
        sys.exit(0)
    generate(sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/justify-squeeze'))
