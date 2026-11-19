#!/usr/bin/env python3
"""Fixtures for how paragraph borders take vertical space in Word (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/para-border.py [out dir]   (default tests/word/fixtures/para-border)

then, to ask Word (winvm) and read where it put every line:

    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/para-border --refresh-word
    python3 tools/word-parity/fixtures/para-border.py --positions tests/word/fixtures/para-border

Every line is one label in Arial 10pt at exactly 20pt a line, no space before or
after unless the fixture is about it, on a Letter page with 1in margins; so where
Word puts a line depends only on the borders (and the spacing a fixture names),
not on how text is measured. A fixture is a plain paragraph a01, the paragraphs
under test b01, b02 ... and a plain paragraph c01: the gap above b01, between the
b's and below the last one, less the 20pt line, is what the borders took.

--positions reads Word's PDFs (<dir>/word/<name>.pdf, written by --refresh-word)
and writes <dir>/word/positions.tsv: for every fixture, every label's page and the
top of its glyphs in points (pdftotext); the PDFs are not committed.
tests/word/para-border.fsx lays the fixtures out and checks the gaps against it;
the families, what each asks and what Word answered are in its header comment.
"""
import os
import re
import subprocess
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def bdr(side, sz=4, space=1, val='single', color='auto'):
    return f'<w:{side} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'


def box(sz=4, space=1, sides='top left bottom right', **kw):
    """A pBdr of the named sides, all alike (left/right with space 4, as Word writes a box)."""
    order = ['top', 'left', 'bottom', 'right', 'between']
    out = ''.join(bdr(s, sz, space if s in ('top', 'bottom', 'between') else 4, **kw) for s in order if s in sides.split())
    return f'<w:pBdr>{out}</w:pBdr>'


def para(label, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{label}</w:t></w:r></w:p>'


def probe(*bs):
    """a01, then one paragraph per pPr in bs (b01, b02 ...), then c01."""
    return para('a01') + ''.join(para(f'b{i:02d}', ppr) for i, ppr in enumerate(bs, 1)) + para('c01')


def document(body):
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
            'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Boxed"><w:name w:val="Boxed"/><w:basedOn w:val="Normal"/><w:pPr>{box()}</w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="BoxedLast"><w:name w:val="Boxed Last"/><w:basedOn w:val="Boxed"/><w:pPr><w:spacing w:after="200"/></w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="BoxedOther"><w:name w:val="Boxed Other"/><w:basedOn w:val="Normal"/><w:pPr>{box()}</w:pPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="BoxedCtx"><w:name w:val="Boxed Ctx"/><w:basedOn w:val="Normal"/><w:pPr><w:contextualSpacing/><w:spacing w:before="200" w:after="200"/>{box()}</w:pPr></w:style>
</w:styles>'''
SETTINGS = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''


def write(path, body):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', document(body))
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)


def style(sid):
    return f'<w:pStyle w:val="{sid}"/>'


def ind(left=0, right=0, hanging=0):
    return f'<w:ind w:left="{left}" w:right="{right}"' + (f' w:hanging="{hanging}"' if hanging else '') + '/>'


def spacing(before=0, after=0):
    return f'<w:spacing w:before="{before}" w:after="{after}"/>'


def families():
    """name -> body, one entry per fixture."""
    out = {'none': probe('', '', '')}
    # What one border takes: its width (sz, eighths of a point) and its space (points), top and bottom alone.
    for side in ('top', 'bottom'):
        for sz in (4, 12, 24, 48):
            for space in (0, 1, 4, 10):
                out[f'{side}-sz{sz}-sp{space}'] = probe(f'<w:pBdr>{bdr(side, sz, space)}</w:pBdr>')
    for val in ('double', 'thickThinSmallGap'):
        for side in ('top', 'bottom'):
            out[f'{side}-{val}'] = probe(f'<w:pBdr>{bdr(side, 6, 1, val)}</w:pBdr>')
    # Line styles: what a compound or patterned line takes, by its width (space 0).
    for val in ('double', 'triple', 'thickThinSmallGap', 'thinThickSmallGap', 'thickThinMediumGap', 'thickThinLargeGap',
                'thinThickThinSmallGap', 'dotted', 'dashed', 'wave', 'doubleWave', 'threeDEmboss', 'outset'):
        for sz in (6, 12, 24):
            out[f'line-{val}-sz{sz}'] = probe(f'<w:pBdr>{bdr("top", sz, 0, val)}</w:pBdr>')
    # One box alone and three alike: is a run of paragraphs with the same borders one box?
    out['box1'] = probe(box())
    out['box3'] = probe(box(), box(), box())
    out['box3-sz24-sp4'] = probe(box(24, 4), box(24, 4), box(24, 4))
    out['topbottom3'] = probe(*[box(sides='top bottom')] * 3)
    out['between3'] = probe(*[box(sides='top left bottom right between')] * 3)
    out['between3-sz24-sp4'] = probe(*[box(24, 4, sides='top left bottom right between')] * 3)
    # What ends a group: one thing of b02 differs from b01 and b03.
    out['grp-style-two'] = probe(style('Boxed'), style('BoxedOther'), style('Boxed'))
    out['grp-style-based'] = probe(style('Boxed'), style('Boxed'), style('BoxedLast'))
    out['grp-style-direct'] = probe(style('Boxed'), box(), style('Boxed'))
    out['grp-indent-hang'] = probe(ind(180, hanging=180) + box(), ind(540, hanging=540) + box(), ind(180, hanging=180) + box())
    out['grp-indent-left'] = probe(box(), ind(720) + box(), box())
    out['grp-indent-first'] = probe(box(), '<w:ind w:left="0" w:firstLine="720"/>' + box(), box())
    out['grp-indent-right'] = probe(box(), ind(0, 720) + box(), box())
    out['grp-color'] = probe(box(), box(color='FF0000'), box())
    out['grp-width'] = probe(box(), box(12), box())
    out['grp-space'] = probe(box(), box(4, 4), box())
    out['grp-val'] = probe(box(), box(val='double'), box())
    out['grp-sides'] = probe(box(), box(sides='top bottom'), box())
    out['grp-left-only'] = probe(box(), box(sides='top left bottom'), box())
    out['grp-after'] = probe(*[spacing(0, 200) + box()] * 3)
    out['grp-before'] = probe(*[spacing(200, 0) + box()] * 3)
    out['grp-ctx'] = probe(style('BoxedCtx'), style('BoxedCtx'), style('BoxedCtx'))
    # The text's left edge, not the left indent: b02's first line or its other lines start where b01's do, or not.
    out['grp-edge-hang'] = probe(box(), ind(720, hanging=720) + box(), box())
    out['grp-edge-hang-off'] = probe(box(), ind(360, hanging=180) + box(), box())
    out['grp-edge-first-off'] = probe(box(), '<w:ind w:left="360" w:firstLine="360"/>' + box(), box())
    out['grp-edge-both'] = probe(ind(360) + box(), ind(720, hanging=360) + box(), '<w:ind w:left="360" w:firstLine="720"/>' + box())
    out['grp-between-off'] = probe(box(sides='top left bottom right between'), box(), box(sides='top left bottom right between'))
    for sz, sp in ((4, 1), (24, 4), (4, 10)):
        bx, bt = box(sz, sp), box(sz, sp, sides='top left bottom right between')
        out[f'between-first-sz{sz}-sp{sp}'] = probe(bt, bx, bx)
        out[f'between-second-sz{sz}-sp{sp}'] = probe(bx, bt, bx)
    # (Not borders: space after and the next one's space before - added or the larger?)
    out['space-sum'] = para('a01', spacing(0, 200)) + para('b01', spacing(100, 0)) + para('c01', spacing(300, 0)) + para('d01')
    out['grp-shade'] =probe(box(), box() + '<w:shd w:val="clear" w:color="auto" w:fill="E0E0E0"/>', box())
    # Spacing outside a single box: does space before/after go outside the border?
    out['box1-before'] = probe(spacing(200, 0) + box())
    out['box1-after'] = probe(spacing(0, 200) + box())
    # The document's case: SchemaFragment (box, keepNext, after 0) then SchemaFragmentLast (based on it, after 200), hanging indents.
    out['schema'] = probe(style('Boxed') + ind(180, hanging=180), style('Boxed') + ind(540, hanging=540),
                          style('Boxed') + ind(900, hanging=900), style('BoxedLast') + ind(180, hanging=180))
    return out


def positions(folder):
    """<folder>/word/positions.tsv from Word's PDFs: name, label, page, top of the label's glyphs (pt)."""
    word = os.path.join(folder, 'word')
    rows, seen = [], set()
    for name in sorted(f[:-5] for f in os.listdir(folder) if f.endswith('.docx')):
        pdf = os.path.join(word, name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        seen.add(name)
        xml = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        page = 0
        for m in re.finditer(r'<page |<word xMin="[\d.]+" yMin="([\d.]+)"[^>]*>([^<]*)</word>', xml):
            if m.group(0).startswith('<page'):
                page += 1
            elif re.fullmatch(r'[a-z]\d\d', m.group(2)):
                rows.append(f'{name}\t{m.group(2)}\t{page}\t{float(m.group(1)):.3f}')
    tsv = os.path.join(word, 'positions.tsv')
    if os.path.exists(tsv):
        rows += [l for l in open(tsv).read().splitlines() if l and l.split('\t')[0] not in seen]
    rows.sort(key=lambda l: (l.split('\t')[0], l.split('\t')[1]))
    with open(tsv, 'w') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'{len(rows)} positions')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--positions':
        positions(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/para-border')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    all_ = families()
    for name, body in all_.items():
        write(os.path.join(out_dir, name + '.docx'), body)
    print(f'{len(all_)} fixtures in {out_dir}')
