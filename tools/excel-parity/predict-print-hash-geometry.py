#!/usr/bin/env python3
"""Locks the predictions for make-print-hash-geometry.py's workbook BEFORE the Excel export is opened.

Round 1 (same tool, 8-digit value) established the mechanism and is the training data for the constants here:
  * Generated hash overflow is CENTRED in the cell whatever the cell's horizontal alignment is; the cell
    alignment affects ordinary text/numbers but not the `####` placeholder.
  * Ordinary text follows the cell alignment normally.
  * The fit test uses a content area inset 2 pt on each side (W - 4 pt total); a centred run of n hashes needs
    n*a <= W - 4. Italic generated hashes are centred and then translated left by r_gen, so the fit test becomes
    n*a <= W - 4 - 2*r_gen.
  * Literal '#' text is ordinary text: it follows alignment and the ordinary italic text reserve r_text.
  * a is the per-glyph '#' advance quantised to whole dots at the 600 dpi print device for the effective font
    size (size * print scale); r_gen is approximated as half the face's ordinary right-aligned italic reserve
    (PrinterMetrics' measured coefficients), and scales with the effective size.

The W(w) conversion is the round-1 measured physical cell width: W = 5.6382*w + 0.0533 pt at 100 %, and
W = 4.2249*w + 0.1207 pt at 75 % (the printed page is 0.75x; column width itself is not a function of cell font).
These constants are training data, not a claim about any production rule.

    python3 tools/excel-parity/predict-print-hash-geometry.py META.tsv FONTDIR OUT.tsv
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
}
# PrinterMetrics' measured ordinary right-aligned italic reserve, reserve / font point at 100 %.
R_TEXT_COEFF = {'Aptos Narrow': 0.2007, 'Calibri': 0.19, 'Arial': 0.18}
# Round-1 training: physical printed cell width vs column-width characters.
W_FIT = {100: (5.6382, 0.0533), 75: (4.2249, 0.1207)}
CONTENT_INSET_TOTAL = 4.0  # 2 pt each side, from round-1 count boundaries


def u16(d, o):
    return struct.unpack_from('>H', d, o)[0]


def tables(d):
    n = u16(d, 4)
    return {d[12 + 16 * i:16 + 16 * i].decode('latin1'): struct.unpack_from('>II', d, 20 + 16 * i) for i in range(n)}


def font_metrics(path):
    """Return ({char: advance em}, unitsPerEm) for ASCII digits and '#'."""
    d = open(path, 'rb').read()
    t = tables(d)
    upm = u16(d, t['head'][0] + 18)
    hhea = t['hhea'][0]
    nhm = u16(d, hhea + 34)
    hm = t['hmtx'][0]
    c = t['cmap'][0]
    sub = None
    for i in range(u16(d, c + 2)):
        pid, eid = u16(d, c + 4 + 8 * i), u16(d, c + 6 + 8 * i)
        so = struct.unpack_from('>I', d, c + 8 + 8 * i)[0]
        if u16(d, c + so) == 4 and ((pid, eid) in ((3, 1), (0, 3), (0, 4))):
            sub = c + so
            break
    if sub is None:
        raise SystemExit(f'no format-4 cmap in {path}')
    seg = u16(d, sub + 6) // 2
    end = sub + 14
    start = end + 2 * seg + 2
    delta = start + 2 * seg
    ro = delta + 2 * seg
    glyphs = {}
    wanted = {ord(ch) for ch in '0123456789#'}

    def lookup(code):
        for i in range(seg):
            s, e = u16(d, start + 2 * i), u16(d, end + 2 * i)
            if s <= code <= e:
                de = struct.unpack_from('>h', d, delta + 2 * i)[0]
                r = u16(d, ro + 2 * i)
                if r == 0:
                    return (code + de) & 0xFFFF
                gi = ro + 2 * i + r + 2 * (code - s)
                g = u16(d, gi)
                return (g + de) & 0xFFFF if g else 0
        return 0

    for code in wanted:
        g = lookup(code)
        if g:
            glyphs[chr(code)] = u16(d, hm + 4 * min(g, nhm - 1)) / upm
    return glyphs, upm


def quantised_pt(em, size_pt):
    return round(em * size_pt * 600 / 72) / 600 * 72


def main():
    meta_path, fontdir, out_path = sys.argv[1:4]
    metrics = {}
    rows = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    head, rows = rows[0], rows[1:]
    out = ['sheet\tcolumn\twidth_chars\talign\tface\tsize_pt\titalic\tscale\tkind\t'
           'pred_W_pt\tpred_advance_pt\tpred_n\tpred_slope_xMaxLeft\tpred_slope_xMinLeft\t'
           'pred_slope_rightGap\tpred_r_gen\tpred_number_fits\tnote']
    for p in rows:
        d = dict(zip(head, p))
        face, size = d['face'], int(d['size_pt'])
        italic = d['italic'] == '1'
        key = (face, italic if italic else False)
        if key not in metrics:
            metrics[key] = font_metrics(f'{fontdir}/{FONT_FILES[key]}')[0]
        size_eff = size * int(d['scale']) / 100
        wch = float(d['width_chars'])
        scale = int(d['scale'])
        slope_w, inter = W_FIT[scale]
        w_pt = slope_w * wch + inter
        a = quantised_pt(metrics[key]['#'], size_eff)
        r_gen = 0.5 * R_TEXT_COEFF[face] * size_eff if italic else 0.0
        content = w_pt - CONTENT_INSET_TOTAL
        if d['kind'] == 'overflow':
            n = int((content - 2 * r_gen) // a)
            # Does the 13-digit row-1 number itself fit? (format '0', no grouping)
            digits = metrics[key]['1']
            number_w = sum(quantised_pt(digits, size_eff) for _ in range(13))
            number_fits = 'yes' if number_w <= content else 'no'
            pred_n = str(n)
        else:
            pred_n = 'text'
            number_fits = '-'
        if d['kind'] == 'overflow':
            # generated overflow is centred for every cell alignment
            slope_xmax = slope_xmin = 0.5
            slope_gap = 0.5
            note = 'centred placeholder; literal text would follow cell alignment'
        else:
            align = d['align']
            slope_xmax = slope_xmin = {'right': 1.0, 'center': 0.5, 'left': 0.0}[align]
            slope_gap = {'right': 0.0, 'center': 0.5, 'left': 1.0}[align]
            note = 'literal text follows cell alignment'
        out.append('\t'.join([d['sheet'], d['column'], d['width_chars'], d['align'], face, str(size),
                              d['italic'], d['scale'], d['kind'],
                              f'{w_pt:.3f}', f'{a:.4f}', pred_n,
                              f'{slope_xmax:.2f}', f'{slope_xmin:.2f}', f'{slope_gap:.2f}',
                              f'{r_gen:.4f}', number_fits, note]))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('# locked before the round-2 Excel export; model in predict-print-hash-geometry.py\n')
        f.write('\n'.join(out) + '\n')


if __name__ == '__main__':
    main()
