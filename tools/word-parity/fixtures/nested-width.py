#!/usr/bin/env python3
"""Fixture family: how wide Word makes a column whose cell holds a nested table, when the table is
narrower than its cells' widths and its columns give way (docs/layout-parity-flow.md; the calibre
family's centred 70% table, a nested 80% table in its left cell).

    python3 tools/word-parity/fixtures/nested-width.py [out dir]   (default tests/word/fixtures/nested-width)

then, to ask Word (winvm) and read where it put every cell's edges:

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/<own dir> \
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/nested-width --refresh-word
    python3 tools/word-parity/fixtures/table-width.py --edges tests/word/fixtures/nested-width/word/*.pdf \
      > tests/word/fixtures/nested-width/word/edges.tsv

The page, font and labels are table-width.py's (Arial 10pt, cells' "Lrc" left-aligned and "Rrc"
right-aligned; a nested table's "Lirc" / "Rirc"). The outer table: two columns, grid and tcW 4680/4680,
tblW 6000 dxa (so its columns give way by 3360 twips), the nested table in cell (1, 1) unless a name says:
  text            no nested table; cell (1, 1) holds the word WORD (a 26-letter word)
  nest            a nested table, grid and tcW 1000/1000 (each cell its labels only)
  nest-word       that, its cell (1, 1) holding WORD as well
  nest-wide       a nested table of grid and tcW 2600/2600 (its widths over its least)
  nest-pct        nest with tblW 4000 pct (80%)
  nest-pct-word   nest-word with tblW 4000 pct
  nest-fixed      nest-word laid out fixed (tblLayout fixed)
  nest-tblw       nest-word with tblW 4000 dxa
  nest-deep       a nested table (grid and tcW 2000) holding nest-word in its cell
  nest-right      nest-word in cell (1, 2) instead
  nest-both       nest-word in both cells
  nest-vmerge     a nested table of two rows (grid and tcW 1000/1000), its first column merged, the merge's
                  cell holding "x" and WORD (the import hands WORD to the continuation)
  vmerge-word     no nested table; cell (1, 1) holds "x" and WORD (text's control)
  calibre         the calibre table: tblW 3500 pct, jc center, grid 3488/3215, tcW 4788/4788, borders
                  single 4; its cell (1, 1) a 4000 pct centred table, grid 1343/1267, tcW 1678/1629 with
                  "Three" in its first cell
  center-pct100   a 100% table (grid and tcW 4680/4680), centred;  right-pct50: a 50% one, right-aligned;
                  center-long: an autofit centred table, tcW auto, its cell (1, 1) a 60-word sentence
  <name>-outfixed the outer table laid out fixed (not the three above)
  <name>-c14      the same in Word 2010's layout (compatibilityMode 14)
"""
import importlib.util
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('table_width', os.path.join(HERE, 'table-width.py'))
tw = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tw)

WORD = 'abcdefghijklmnopqrstuvwxyz'
BORDERS = ('<w:tblBorders>' + ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
           for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) + '</w:tblBorders>')


def with_pr(tbl, extra):
    """Adds properties at the end of a table's w:tblPr."""
    return tbl.replace('</w:tblPr>', extra + '</w:tblPr>', 1)


def nested(word=False, grid=(1000, 1000), tblw=('auto', 0), layout=None):
    content = {(1, 1): tw.para(WORD)} if word else None
    return tw.table(list(grid), list(grid), tblw=tblw, layout=layout, tag='i', content=content)


def families():
    out = {}
    for lay, layout in (('', None), ('-outfixed', 'fixed')):
        def add(name, inner=None, content=None, inner2=None):
            tbl = tw.table([4680, 4680], [4680, 4680], tblw=6000, layout=layout, inner=inner, content=content)
            if inner2 is not None:
                # the nested table in cell (1, 2) too: before its labels
                marker = '<w:p><w:r><w:t xml:space="preserve">L12</w:t>'
                tbl = tbl.replace(marker, inner2 + tw.para('') + marker, 1)
            out[name + lay] = tw.probe(tbl)
        add('text', content={(1, 1): tw.para(WORD)})
        add('nest', inner=nested())
        add('nest-word', inner=nested(word=True))
        add('nest-wide', inner=nested(grid=(2600, 2600)))
        add('nest-pct', inner=nested(tblw=('pct', 4000)))
        add('nest-pct-word', inner=nested(word=True, tblw=('pct', 4000)))
        add('nest-fixed', inner=nested(word=True, layout='fixed'))
        add('nest-tblw', inner=nested(word=True, tblw=4000))
        deep = tw.table([2000], [2000], tag='i', content={(1, 1): nested(word=True).replace('Li', 'Lj').replace('Ri', 'Rj') + tw.para('')})
        add('nest-deep', inner=deep)
        add('nest-right', inner2=nested(word=True))
        add('nest-both', inner=nested(word=True), inner2=nested(word=True))
        merge = [[(1, 1000, '<w:vMerge w:val="restart"/>'), (1, 1000)], [(1, 1000, '<w:vMerge/>'), (1, 1000)]]
        add('nest-vmerge', inner=tw.table(None, rows=merge, tag='i', content={(1, 1): tw.para('x') + tw.para(WORD)}).replace(
            '<w:tbl><w:tblPr>', '<w:tbl><w:tblPr>', 1).replace('</w:tblPr>', '</w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>', 1))
        add('vmerge-word', content={(1, 1): tw.para('x') + tw.para(WORD)})
        # The calibre document's table.
        inner = tw.table([1343, 1267], [1678, 1629], tblw=('pct', 4000), layout=layout, tag='i',
                         content={(1, 1): tw.para('Three')})
        inner = with_pr(inner, '<w:jc w:val="center"/>' + BORDERS)
        outer = tw.table([3488, 3215], [4788, 4788], tblw=('pct', 3500), layout=layout, inner=inner)
        out['calibre' + lay] = tw.probe(with_pr(outer, '<w:jc w:val="center"/>' + BORDERS))
    # Aligned tables' room in Word 2010's layout (a percentage's base, an autofit table's most).
    jc = lambda tbl, how: with_pr(tbl, f'<w:jc w:val="{how}"/>')
    out['center-pct100'] = tw.probe(jc(tw.table([4680, 4680], [4680, 4680], tblw=('pct', 5000)), 'center'))
    out['right-pct50'] = tw.probe(jc(tw.table([4680, 4680], [4680, 4680], tblw=('pct', 2500)), 'right'))
    out['center-long'] = tw.probe(jc(tw.table([1000, 1000], [('auto', 0), ('auto', 0)], content={(1, 1): tw.para(tw.LONG)}), 'center'))
    return out


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(HERE, '../../..'))
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/nested-width')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    all_ = families()
    for name, body in all_.items():
        tw.write(os.path.join(out_dir, name + '.docx'), body)
        tw.write(os.path.join(out_dir, name + '-c14.docx'), body, mode=14)
    print(f'{2 * len(all_)} fixtures in {out_dir}')
