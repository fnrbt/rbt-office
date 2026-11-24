#!/usr/bin/env python3
"""Compares printed full/hash fit decisions against a fresh direct Excel PDF measurement.

Three models, one row per measured row-1 test column:

  baseline      the runtime's screen-shaped print rule replicated from CellText.fittedRoom /
                fittedHashes on the layout cell width (the rule shipped before the fit cutoff);
  candidate     the printed fit cutoff with actual face metrics on the layout cell width
                (dots = floor(size*scale*600/72 + 0.5); a = round(hashEm*dots)*72/600;
                 c = digitEm*size*scale/2 + 0.875 + italic printed reserve;
                 fits iff per-character rounded-dot run width + `_x` pad advances <= W - c;
                 otherwise n = max(1, floor((W-c)/a) - leading `_x` slots));
  hypothesis    the candidate on the measured PDF cell width instead of the layout width.

The layout width is `floor(floor(((256*w + floor(128/7))/256)*7) * 47/7 + 0.5) * 72/600 * scale`
(Normal Calibri 11: mdw 7, print digit 47). `W_measured` is the PDF fill-rectangle width.

    python3 tools/excel-parity/analyze-print-fit-cutoff.py META RESULTS OUT.tsv

Prints exact / full / overflow / count tallies and every false-hash / false-full row. Faces come from
predict-print-count.py's table with the real Georgia Italic file (windows-extra/georgiai.ttf; the
windows/GARA* files are Garamond).
"""
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = __import__('importlib.util').util.spec_from_file_location(
    'countpredict', os.path.join(HERE, 'predict-print-count.py'))
countpredict = __import__('importlib.util').util.module_from_spec(spec)
spec.loader.exec_module(countpredict)
countpredict.FONT_FILES[('Georgia', False)] = 'windows-extra/georgia.ttf'
countpredict.FONT_FILES[('Georgia', True)] = 'windows-extra/georgiai.ttf'

FACES = {}


def face(name, italic):
    key = (name, italic)
    if key not in FACES:
        path = os.path.join(countpredict.FONT_DIR, countpredict.FONT_FILES[key])
        FACES[key] = countpredict.font_metrics(path)
    return FACES[key]


def gdi_dots(size, scale):
    return math.floor(size * scale * 600.0 / 72.0 + 0.5)


def advance_pt(g, ch, dots):
    return math.floor(g[ch] * dots + 0.5) * 72.0 / 600.0


def italic_shift(g, upm, win_ascent, size, scale):
    dots = gdi_dots(size, scale)
    ascent = math.floor(win_ascent * dots / upm + 0.5)
    return (ascent // 5) * 72.0 / 600.0


def layout_width(width_chars, scale):
    screen_px = int(((256.0 * width_chars + 128 // 7) / 256.0) * 7.0)
    return int(screen_px * 47.0 / 7.0 + 0.5) * 72.0 / 600.0 * scale


def formatted_run(fmt, value):
    """The full formatted text (pads as spaces, `*` fill once), the visible run, and lead/trail `_x` pads."""
    n = float(value)
    if fmt == '0':
        return str(int(n)), str(int(n)), '', ''
    if fmt == '0.00%':
        text = f'{n * 100:.2f}%'
        return text, text, '', ''
    if fmt == '#,##0.00':
        text = f'{int(n):,}.00'
        return text, text, '', ''
    if fmt.startswith('_("$'):
        # `_("$"* #,##0_)` positive: one `_(` and one `_)`; negative `_("$"* (#,##0)`: one `_(` only.
        if n >= 0:
            return ' $ ' + f'{int(n):,}' + ' ', '$' + f'{int(n):,}', '(', ')'
        return ' $ (' + f'{int(-n):,}' + ')', '$(' + f'{int(-n):,}' + ')', '(', ''
    raise SystemExit(f'unsupported format {fmt!r}')


def baseline(g, size, scale, italic, formatted, width_pt):
    # .NET's Math.Round for measurePdf's 600-dpi dot count, then the screen margin `fittedRoom`.
    dots = round(size * scale * 600.0 / 72.0)
    measure = lambda text: sum(round(g[c] * dots) for c in text) * 96.0 / 600.0
    cell_px = width_pt * 96.0 / 72.0
    widest_digit = max(measure(c) for c in '0123456789')
    inset = math.ceil(widest_digit / 4.0)
    px = max(1, int(math.floor(size * 96.0 / 72.0 * scale + 0.5)))
    small = math.floor(round(size * scale * 20.0, 9) + 0.5) < 160
    overhang = (math.floor(px / 2.0) if (italic and small)
                else (math.floor(px * 0.16 + 0.5) if italic else 0.0))
    room = max(0.0, round(cell_px) - (2 * inset + 1) * scale - overhang)
    return measure(formatted) <= room + 1e-9


def candidate(g, upm, win_ascent, size, scale, italic, visible, lead, trail, width_pt):
    dots = gdi_dots(size, scale)
    c = g['0'] * size * scale / 2.0 + 0.875
    if italic:
        c += italic_shift(g, upm, win_ascent, size, scale)
    pads = lead + trail
    run = sum(advance_pt(g, ch, dots) for ch in visible) + sum(advance_pt(g, ch, dots) for ch in pads)
    a = advance_pt(g, '#', dots)
    full = run <= width_pt - c + 1e-9
    n = 0 if full else max(1, int((width_pt - c) // a) - len(lead))
    return full, n, run, c, a


def tsv(path):
    rows = [line.rstrip('\n').split('\t') for line in open(path, encoding='utf-8')]
    return rows[0], rows[1:]


def main():
    meta_path, results_path, out_path = sys.argv[1:4]
    meta_head, meta_rows = tsv(meta_path)
    meta = {(d['sheet'], d['column']): d for d in (dict(zip(meta_head, r)) for r in meta_rows)}
    result_head, result_rows = tsv(results_path)
    truth = {}
    for r in (dict(zip(result_head, x)) for x in result_rows):
        if r['row'] == '1' and r['nHashes'] not in ('', None):
            truth[(r['sheet'], r['column'])] = r
    out = ['sheet\tcolumn\twidth_chars\tface\tsize_pt\titalic\tscale\tformat\tvalue\tprinted\texcel_full\texcel_n\t'
           'W_layout_pt\tW_measured_pt\tbaseline_full\tbaseline_n\tcandidate_full\tcandidate_n\tcandidate_run_pt\tcandidate_c_pt\t'
           'hypothesis_full\thypothesis_n\toutcome']
    models = {'baseline': [0, 0, 0, 0], 'candidate': [0, 0, 0, 0], 'hypothesis': [0, 0, 0, 0]}
    # [decision_exact, full_exact, overflow_exact, count_exact]
    mismatches = []
    for key in sorted(truth):
        r = truth[key]
        d = meta.get(key)
        if d is None:
            continue
        fmt, value = d.get('format', ''), d.get('value', '')
        if value == '' or d.get('role') == 'lit':
            continue
        formatted, visible, lead, trail = formatted_run(fmt, value)
        name, italic = d['face'], d['italic'] == '1'
        size, scale = float(d['size_pt']), int(d['scale']) / 100.0
        g, upm, win_ascent = face(name, italic)
        w_layout = layout_width(float(d['width_chars']), scale)
        w_measured = float(r['cellWidth'])
        excel_full = int(r['nHashes']) == 0
        b_full = baseline(g, size, scale, italic, formatted, w_layout)
        c_full, c_n, run, c, a = candidate(g, upm, win_ascent, size, scale, italic, visible, lead, trail, w_layout)
        h_full, h_n, _, _, _ = candidate(g, upm, win_ascent, size, scale, italic, visible, lead, trail, w_measured)
        verdict = []
        for model, (full, n) in (('baseline', (b_full, 0)), ('candidate', (c_full, c_n)),
                                 ('hypothesis', (h_full, h_n))):
            tally = models[model]
            tally[0] += 1 if full == excel_full else 0
            if excel_full:
                tally[1] += 1 if full else 0
            else:
                tally[2] += 1 if not full else 0
                if model != 'baseline':
                    tally[3] += 1 if (not full and n == int(r['nHashes'])) else 0
                    if not full and n != int(r['nHashes']):
                        verdict.append(f'{model}-count-off-by-{n - int(r["nHashes"])}')
            if full != excel_full:
                verdict.append(f'{model}-false-{"full" if full else "hash"}')
        out.append('\t'.join([
            key[0], key[1], d['width_chars'], name, d['size_pt'], d['italic'], d['scale'], fmt, value,
            r['text'], 'yes' if excel_full else 'no', r['nHashes'],
            f'{w_layout:.3f}', f'{w_measured:.3f}',
            'yes' if b_full else 'no', '-',
            'yes' if c_full else 'no', str(c_n), f'{run:.3f}', f'{c:.4f}',
            'yes' if h_full else 'no', str(h_n), ';'.join(verdict) or 'ok']))
        if verdict:
            mismatches.append((key, verdict))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print(f'wrote {len(out) - 1} rows to {out_path}')
    total = sum(1 for _ in out) - 1
    full_n = sum(1 for line in out[1:] if line.split('\t')[10] == 'yes')
    overflow_n = total - full_n
    for model in ('baseline', 'candidate', 'hypothesis'):
        exact, full_exact, overflow_exact, count_exact = models[model]
        if model == 'baseline':
            print(f'{model:10s} decision {exact}/{total}  full {full_exact}/{full_n}  overflow {overflow_exact}/{overflow_n}  '
                  f'(count not modelled)')
        else:
            print(f'{model:10s} decision {exact}/{total}  full {full_exact}/{full_n}  overflow {overflow_exact}/{overflow_n}  '
                  f'count {count_exact}/{overflow_n}  output {full_exact + count_exact}/{total}')
    for key, verdict in mismatches:
        print(f'  {key[0]} col{key[1]}: {" ".join(verdict)}')


if __name__ == '__main__':
    main()
