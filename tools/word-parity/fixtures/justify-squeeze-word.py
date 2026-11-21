#!/usr/bin/env python3
"""Fixtures for what else than its spaces limits how far a justified line runs past its width in Word 2013's
layout (docs/layout-parity-flow.md, the justify-squeeze rows): the line's last word.

    python3 tools/word-parity/fixtures/justify-squeeze-word.py [out dir]   (default tests/word/fixtures/justify-squeeze-word)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/justify-squeeze-word \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/justify-squeeze-word --refresh-word
    python3 tools/word-parity/fixtures/justify-squeeze.py --table tests/word/fixtures/justify-squeeze-word

(--table is justify-squeeze.py's: per probe whether its first line held its last word - word/squeeze.tsv.)

A probe is a paragraph as justify-squeeze's: n spaces between n tokens of digits and a last word of k letters
"n", then " end"; a right indent makes the line narrower than the text up to the last word's end by o, the
probe's overflow - here a fraction f of the last word's width (0, 0.05 ... 1.0), so that the line's spaces
could give up more than o (a quarter of each) for the first probes and less for the last. In the
probes.tsv the column overflow_per_space is o / n (as justify-squeeze's), and word / over the last word's
width and o.

  <font><size>-n<n>-k<k>[-hyph][-m14]   fonts cour (Courier New), geo (Georgia), cal (Calibri), times; k
                                        letters; -hyph w:autoHyphenation; -m14 Word 2010's layout
  last-cour10-n9-k2                     no " end": the probe's line is its paragraph's last
  fine-<font><size>-n<n>-k<k>           the overflow 0.300 ... 0.420 of the last word and the space before
                                        it, in steps of 0.005 (the boundary the coarse run showed)
No dependencies: the package is written as XML.
"""
import importlib.util
import os
import sys

sys.dont_write_bytecode = True

here = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('justify_squeeze', os.path.join(here, 'justify-squeeze.py'))
js = importlib.util.module_from_spec(spec)
spec.loader.exec_module(js)

LINE = js.LINE
# Advances (units per 2048 em): space, digit '0' (all the digits one width but Georgia's - its '0'), 'n'.
FONTS = {'cour': ('Courier New', 1229, 1229, 1229), 'geo': ('Georgia', 494, 1257, 1210),
         'cal': ('Calibri', 463, 1038, 1076), 'times': ('Times New Roman', 512, 1024, 1024)}


def styles(font, size):
    face = FONTS[font][0]
    return js.styles('cour', size).replace('"Courier New"', f'"{face}"')


def probe(font, size, n, k, f, last=False, gap=False):
    """The probe's paragraph and its overflow: n tokens of '0's as wide as whole digits take them with the
    last word to LINE at most, the line narrowed by a right indent to the text's width less f x the word
    (gap: f x the word and the space before it)."""
    _, sp, dg, nn = FONTS[font]
    s, d, w = sp * size / 2048, dg * size / 2048, k * nn * size / 2048
    digits = int((LINE - n * s - w) // d)
    sizes = [digits // n + (1 if i < digits % n else 0) for i in range(n)]
    natural = digits * d + n * s + w
    over = f * (w + s if gap else w)
    indent = round((LINE - natural + over) * 20)
    actual = natural - (LINE - indent / 20)
    text = ' '.join('0' * c for c in sizes) + ' ' + 'n' * k + ('' if last else ' end')
    return f'<w:p>{js.ppr("both", size=size, indent=indent)}{js.r(text)}</w:p>', actual, s, w


def family(out, name, font, size, n, k, mode=15, hyph=False, last=False, steps=None, gap=False):
    rows, paras = [], []
    for i, f in enumerate(steps if steps is not None else [0.05 * j for j in range(21)]):
        para, actual, s, w = probe(font, size, n, k, f, last, gap)
        paras.append(para)
        rows.append(f'{name}\t{i}\t{font}\t{size}\t{n}\t{s:.4f}\t{actual / n:.4f}\tboth\t{mode}\t{int(last)}\tbody'
                    f'\t{k}\t{w:.4f}\t{actual:.4f}')
    js.write(os.path.join(out, name + '.docx'), ''.join(paras), styles(font, size), js.settings(mode, hyph))
    return rows


def generate(out):
    os.makedirs(out, exist_ok=True)
    for f in os.listdir(out):
        if f.endswith('.docx'):
            os.remove(os.path.join(out, f))
    rows = []
    for k in (1, 2, 3, 4):
        rows += family(out, f'cour10-n9-k{k}', 'cour', 10, 9, k)
    for k in (1, 2, 3):
        rows += family(out, f'geo10-n9-k{k}', 'geo', 10, 9, k)
        rows += family(out, f'cal11-n9-k{k}', 'cal', 11, 9, k)
    rows += family(out, 'geo11-n9-k1', 'geo', 11, 9, 1)
    rows += family(out, 'cal10-n9-k1', 'cal', 10, 9, 1)
    rows += family(out, 'times10-n9-k1', 'times', 10, 9, 1)
    rows += family(out, 'geo10-n20-k1', 'geo', 10, 20, 1)
    rows += family(out, 'cour20-n9-k2', 'cour', 20, 9, 2)
    rows += family(out, 'cour10-n9-k2-hyph', 'cour', 10, 9, 2, hyph=True)
    rows += family(out, 'geo10-n9-k1-hyph', 'geo', 10, 9, 1, hyph=True)
    rows += family(out, 'cour10-n9-k2-m14', 'cour', 10, 9, 2, mode=14)
    # (Steps of 0.1: a page holds them all - Word's last line moved to the next page is no probe of its own.)
    rows += family(out, 'last-cour10-n9-k2', 'cour', 10, 9, 2, last=True, steps=[0.1 * j for j in range(11)])
    # Fine: the overflow 0.300 ... 0.420 of the last word and its space, in steps of 0.005.
    fine = [0.300 + 0.005 * j for j in range(25)]
    for n in (5, 7, 9, 14, 20):
        rows += family(out, f'fine-cour10-n{n}-k2', 'cour', 10, n, 2, steps=fine, gap=True)
    for n in (9, 14, 20):
        rows += family(out, f'fine-geo10-n{n}-k2', 'geo', 10, n, 2, steps=fine, gap=True)
    rows += family(out, 'fine-cal11-n9-k1', 'cal', 11, 9, 1, steps=fine, gap=True)
    with open(os.path.join(out, 'probes.tsv'), 'w') as f:
        f.write('doc\tprobe\tfont\tsize\tspaces\tspace\toverflow_per_space\tjc\tmode\tlast\twhere\tletters\tword\tover\n'
                + '\n'.join(rows) + '\n')
    print(len(set(r.split('\t')[0] for r in rows)), 'documents,', len(rows), 'probes')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(here, '../../..'))
    generate(sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/justify-squeeze-word'))
