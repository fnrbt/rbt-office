#!/usr/bin/env python3
"""Locks / evaluates the printer-unit fit candidate for the full<->hash anchor workbooks.

Model (same family as predict-print-count.py, with real Georgia Italic metrics and a
GDI floor(x+0.5) height):

  dots  = floor(size_pt*scale*600/72 + 0.5)
  adv(c)= floor(glyphEm(c)*dots + 0.5) * 72/600                  # pt at page scale
  a     = adv('#')
  c     = digitEm*size_pt*scale/2 + 0.875
          + italic ? floor(floor(usWinAscent*dots/upm + 0.5)/5)*72/600 : 0
  W(w)  = floor(screenPx(w)*47/7 + 0.5) * 72/600 * scale
          screenPx(w) = floor(((256*w + floor(128/7))/256)*7)
  full text fits iff sum(adv(glyph) for glyph in the formatted string, pads included)
                 <= W - c
  else n = max(1, floor((W-c)/a) - leading `_x` pad slots)

`--measure RESULTS.tsv` replaces W(w) with the measured physical cell width and adds
the model's decision margin.  The holdout predictions are locked with the
pre-registered W(w) formula only, before that workbook is exported.

    python3 tools/excel-parity/predict-print-fullhash-anchor.py META.tsv OUT.tsv [--measure RESULTS.tsv]
"""
import collections
import math
import os
import struct
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
FONT_FILES = {
    ('Verdana', False): 'artifacts/fonts/windows/verdana.ttf',
    ('Verdana', True): 'artifacts/fonts/windows/verdanai.ttf',
    ('Georgia', False): 'artifacts/fonts/windows-extra/georgia.ttf',
    ('Georgia', True): 'artifacts/fonts/windows-extra/georgiai.ttf',
    ('Cambria', False): 'artifacts/fonts/windows/cambria.ttc',
    ('Trebuchet MS', False): 'artifacts/fonts/windows-extra/trebuc.ttf',
    ('Trebuchet MS', True): 'artifacts/fonts/windows-extra/trebucit.ttf',
}
CHARS = '$(),.% 0123456789#-'


def u16(d, o):
    return struct.unpack_from('>H', d, o)[0]


def font_dir(path):
    d = open(path, 'rb').read()
    base = 0
    if d[:4] == b'ttcf':
        base = struct.unpack_from('>I', d, 12)[0]          # first face of the collection
    n = u16(d, base + 4)
    t = {d[base + 12 + 16 * i:base + 16 + 16 * i].decode('latin1'):
         struct.unpack_from('>II', d, base + 20 + 16 * i) for i in range(n)}
    return d, t


def font_metrics(path):
    d, t = font_dir(path)
    upm = u16(d, t['head'][0] + 18)
    hhea = t['hhea'][0]
    nhm = u16(d, hhea + 34)
    hm = t['hmtx'][0]
    cmap = t['cmap'][0]
    sub = None
    for i in range(u16(d, cmap + 2)):
        pid, eid = u16(d, cmap + 4 + 8 * i), u16(d, cmap + 6 + 8 * i)
        so = struct.unpack_from('>I', d, cmap + 8 + 8 * i)[0]
        if u16(d, cmap + so) == 4 and (pid, eid) in ((3, 1), (0, 3), (0, 4)):
            sub = cmap + so
            break
    seg = u16(d, sub + 6) // 2
    end = sub + 14
    start = end + 2 * seg + 2
    delta = start + 2 * seg
    ro = delta + 2 * seg

    def lookup(code):
        for i in range(seg):
            s, e = u16(d, start + 2 * i), u16(d, end + 2 * i)
            if s <= code <= e:
                de = struct.unpack_from('>h', d, delta + 2 * i)[0]
                r = u16(d, ro + 2 * i)
                if r == 0:
                    return (code + de) & 0xFFFF
                g = u16(d, ro + 2 * i + r + 2 * (code - s))
                return (g + de) & 0xFFFF if g else 0
        return 0

    glyphs = {}
    for ch in CHARS:
        g = lookup(ord(ch))
        if g:
            glyphs[ch] = u16(d, hm + 4 * min(g, nhm - 1)) / upm
    os2 = t['OS/2'][0]
    return glyphs, upm, u16(d, os2 + 74)


def height_dots(size, scale):
    return math.floor(size * scale * 600.0 / 72.0 + 0.5)


def advance(glyphs, ch, dots):
    return math.floor(glyphs[ch] * dots + 0.5) * 72.0 / 600.0


def italic_shift(win_ascent, upm, dots):
    return math.floor(math.floor(win_ascent * dots / upm + 0.5) / 5.0) * 72.0 / 600.0


def physical_width(w, scale):
    screen_px = int(((256.0 * w + 128 // 7) / 256.0) * 7.0)
    return int(screen_px * 47.0 / 7.0 + 0.5) * 72.0 / 600.0 * scale


ACCT = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)'


def main():
    args = sys.argv[1:]
    meta_path, out_path = args[0], args[1]
    measured = None
    if '--measure' in args:
        measured = args[args.index('--measure') + 1]
    meta = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    head, meta = meta[0], meta[1:]
    measured_w = {}
    if measured:
        m = [l.rstrip('\n').split('\t') for l in open(measured, encoding='utf-8')]
        mh = m[0]
        for p in m[1:]:
            d = dict(zip(mh, p))
            if d['row'] == '1' and d['cellWidth']:
                measured_w[(d['sheet'], int(d['column']))] = float(d['cellWidth'])

    cache = {}
    out = ['sheet\tcolumn\twidth_chars\tface\tsize_pt\titalic\tscale\tformat\tvalue\ttext\t'
           'pred_W_pt\tpred_text_pt\tpred_c_pt\tpred_a_pt\tpred_full\tpred_n\tmargin_pt\tlocked_basis']
    for p in meta:
        d = dict(zip(head, p))
        if d['row'] != '1':
            continue
        face, italic = d['face'], d['italic'] == '1'
        key = (face, italic)
        if key not in cache:
            cache[key] = font_metrics(os.path.join(ROOT, FONT_FILES[key]))
        glyphs, upm, win_ascent = cache[key]
        size = float(d['size_pt'])
        scale = int(d['scale']) / 100.0
        dots = height_dots(size, scale)

        def adv(ch):
            return advance(glyphs, ch, dots)

        a = adv('#')
        c = glyphs['0'] * size * scale / 2.0 + 0.875
        if italic:
            c += italic_shift(win_ascent, upm, dots)
        text = d['text']
        slots = 1 if d['format'] == ACCT else 0
        text_pt = sum(adv(ch) for ch in text)
        if slots:
            text_pt += adv('(') + adv(')')
        w = measured_w.get((d['sheet'], int(d['column']))) if measured else None
        basis = 'measured-W' if w is not None else 'W-model'
        if w is None:
            w = physical_width(float(d['width_chars']), scale)
        room = w - c
        fits = text_pt <= room + 1e-9
        n = 0 if fits else max(1, math.floor(room / a) - slots)
        out.append('\t'.join([
            d['sheet'], d['column'], d['width_chars'], face, d['size_pt'], d['italic'], d['scale'],
            d['format'], d['value'], text, f'{w:.4f}', f'{text_pt:.4f}', f'{c:.4f}', f'{a:.4f}',
            'yes' if fits else 'no', str(n), f'{room - text_pt:.4f}', basis]))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        if not measured:
            f.write('# locked before the corresponding Excel export; model in '
                    'predict-print-fullhash-anchor.py; W from the pre-registered formula only\n')
        f.write('\n'.join(out) + '\n')
    print(f'wrote {len(out) - 1} predictions to {out_path}')


if __name__ == '__main__':
    main()
