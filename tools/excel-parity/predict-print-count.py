#!/usr/bin/env python3
"""Locks the CountTransitions candidate predictions for printed numeric overflow.

Model (printer units, 600 dpi; fitted on the dense train split, not on the holdout):
  dots      = round(size * scale * 600 / 72)
  advance(c)= round(glyphEm(c) * dots) * 72 / 600            # pt at the page scale
  a         = advance('#')
  c         = digitEm * size * scale / 2 + 0.875             # constant content offset
  c italic += printItalicShift (full right-align reserve: floor(round(usWinAscent*dots/upm)/5)*72/600)
  W(w)      = floor(screenPx(w) * 47 / 7 + 0.5) * 72 / 600 * scale,
              screenPx(w) = floor(((256*w + floor(128/7)) / 256) * 7)
  formatted text fits when its per-character rounded-dot width <= W - c; otherwise
  n = max(0, floor((W - c)/a) - leadingPrintHashSlots).

    python3 tools/excel-parity/predict-print-count.py META.tsv OUT.tsv
"""
import struct
import sys

FONT_FILES = {
    ('Aptos Narrow', False): 'windows-extra/aptosnarrow-27970306811.ttf',
    ('Aptos Narrow', True): 'windows-extra/aptosnarrow-28903934312.ttf',
    ('Calibri', False): 'windows/calibri.ttf',
    ('Calibri', True): 'windows/calibrii.ttf',
    ('Arial', False): 'windows/arial.ttf',
    ('Arial', True): 'windows/ariali.ttf',
    ('Verdana', False): 'windows/verdana.ttf',
    ('Verdana', True): 'windows/verdanai.ttf',
    ('Georgia', False): 'windows/GARA.TTF',
    ('Georgia', True): 'windows/GARAIT.TTF',
}
FONT_DIR = 'artifacts/fonts'
CHARS = '$(),.% 0123456789#-'


def u16(d, o):
    return struct.unpack_from('>H', d, o)[0]


def tables(d):
    n = u16(d, 4)
    return {d[12 + 16 * i:16 + 16 * i].decode('latin1'): struct.unpack_from('>II', d, 20 + 16 * i)
            for i in range(n)}


def font_metrics(path):
    d = open(path, 'rb').read()
    t = tables(d)
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


def italic_shift(us_win_ascent, upm, points, scale):
    height_dots = round(points * scale * 600.0 / 72.0)
    ascent_dots = round(us_win_ascent * height_dots / upm)
    return (ascent_dots // 5) * 72.0 / 600.0


def physical_width(w, scale):
    screen_px = int(((256.0 * w + 128 // 7) / 256.0) * 7.0)
    return int(screen_px * 47.0 / 7.0 + 0.5) * 72.0 / 600.0 * scale


def fmt_text(fmt, value):
    if fmt == '0':
        return str(int(value)), 0
    if fmt.startswith('_("$"'):
        return '$1,234,567', 1
    if fmt == '0.00%':
        return f'{value * 100:.2f}%', 0
    if fmt == '#,##0.00':
        return f'{int(value):,}.00', 0
    raise SystemExit(f'unsupported format {fmt!r}')


def main():
    meta_path, out_path = sys.argv[1:3]
    rows = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    head, rows = rows[0], rows[1:]
    cache = {}
    out = ['sheet\tcolumn\tpred_n\tpred_full\tpred_W_pt\tpred_advance_pt\tpred_c_pt\tpred_text_pt']
    for p in rows:
        d = dict(zip(head, p))
        face = d['face']
        italic = d['italic'] == '1'
        key = (face, italic)
        if key not in cache:
            cache[key] = font_metrics(f'{FONT_DIR}/{FONT_FILES[key]}')
        glyphs, upm, win_ascent = cache[key]
        size = float(d['size_pt'])
        scale = int(d['scale']) / 100.0
        dots = round(size * scale * 600.0 / 72.0)

        def adv(ch):
            return round(glyphs[ch] * dots) * 72.0 / 600.0

        a = adv('#')
        c = glyphs['0'] * size * scale / 2.0 + 0.875
        if italic:
            c += italic_shift(win_ascent, upm, size, scale)
        text, slots = fmt_text(d['format'], float(d['value']))
        text_width = sum(adv(ch) for ch in text) + 2 * adv('(') if slots else sum(adv(ch) for ch in text)
        w = physical_width(float(d['width_chars']), scale)
        fits = text_width <= w - c + 1e-9
        n = 0 if fits else max(0, int((w - c) // a) - slots)
        out.append('\t'.join([d['sheet'], d['column'], str(n), 'yes' if fits else 'no',
                              f'{w:.3f}', f'{a:.4f}', f'{c:.4f}', f'{text_width:.3f}']))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('# locked before the holdout Excel export; model in predict-print-count.py\n')
        f.write('\n'.join(out) + '\n')
    print(f'wrote {len(out) - 1} predictions to {out_path}')


if __name__ == '__main__':
    main()
