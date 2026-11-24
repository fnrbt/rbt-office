#!/usr/bin/env python3
"""Locks the pre-export predictions for make-print-column-edges.py's workbook.

Written before the workbook is sent to Excel.  The hypotheses are structural, so
they can be falsified by the PDF content stream:

  H_width     the grid border centreline spacing (true printed column width)
              equals the product's stored-width conversion `layout_width()`.
  H_expand    each cell's solid fill rectangle is expanded horizontally beyond
              the border centreline by E/2 on both sides, so the fill rectangle
              is E pt wider than the true column and adjacent fill rectangles
              overlap by E.  E is predicted 0.143 pt (the mean measured
              fill-minus-layout width of 41 no-border Georgia 15 upright 90 %
              columns in print-fullhash-anchor-results.tsv).
  H_page      widths are invariant to page origin and page break: the same
              width character gives the same border-to-border width on every
              page, and only the page x origin translates.
  H_fit       the full-number fit decision uses the fill width (previous direct
              anchor PDFs agree with that), so 16.2188-16.3438 at 90 % print the
              full number.  The layout-width model would print hashes there.

Per column: predicted layout width, predicted fill width (layout+E), predicted
border width (layout), predicted adjacent-fill overlap (E), and the two fit
decisions.  Held-out sheets H85_edge and H90_shift are included; they are not
exported until this file exists.

    python3 tools/excel-parity/lock-print-column-edges.py META.tsv OUT.tsv
"""
import importlib.util
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location('countpredict', os.path.join(HERE, 'predict-print-count.py'))
countpredict = importlib.util.module_from_spec(spec)
spec.loader.exec_module(countpredict)
countpredict.FONT_FILES[('Georgia', False)] = 'windows-extra/georgia.ttf'
countpredict.FONT_FILES[('Georgia', True)] = 'windows-extra/georgiai.ttf'
_g, _upm, _wa = countpredict.font_metrics(os.path.join(countpredict.FONT_DIR,
                                                      countpredict.FONT_FILES[('Georgia', False)]))

E_FILL = 0.143
INSET_C_BASE = 0.875


def layout_width(width_chars, scale_pct):
    scale = scale_pct / 100.0
    screen_px = int(((256.0 * width_chars + 128 // 7) / 256.0) * 7.0)
    return int(screen_px * 47.0 / 7.0 + 0.5) * 72.0 / 600.0 * scale


def gdi_dots(size, scale):
    return int(size * scale * 600.0 / 72.0 + 0.5)


def advance(g, ch, dots):
    return int(g[ch] * dots + 0.5) * 72.0 / 600.0


def fit(width_pt, scale_pct):
    scale = scale_pct / 100.0
    size = 15
    dots = gdi_dots(size, scale)
    c = _g['0'] * size * scale / 2.0 + INSET_C_BASE
    run = sum(advance(_g, ch, dots) for ch in '1,234,567.00')
    a = advance(_g, '#', dots)
    full = run <= width_pt - c + 1e-9
    n = 0 if full else max(1, int((width_pt - c) // a))
    return full, n


def main():
    meta_path, out_path = sys.argv[1:3]
    rows = [l.rstrip('\n').split('\t') for l in open(meta_path, encoding='utf-8')]
    head, rows = rows[0], rows[1:]
    out = ['sheet\tcolumn\tblock\tblock_index\twidth_chars\tscale\t'
           'pred_W_layout_pt\tpred_fill_width_pt\tpred_border_width_pt\tpred_fill_overlap_pt\t'
           'pred_fill_decision\tpred_fill_n\tpred_layout_decision\tpred_layout_n\tpred_page_invariant_width\tnote']
    for p in rows:
        d = dict(zip(head, p))
        w, scale = float(d['width_chars']), int(d['scale'])
        W = layout_width(w, scale)
        full, n = fit(W + E_FILL, scale)
        layout_full, layout_n = fit(W, scale)
        out.append('\t'.join([
            d['sheet'], d['column'], d['block'], d['block_index'], d['width_chars'], str(scale),
            f'{W:.3f}', f'{W + E_FILL:.3f}', f'{W:.3f}', f'{E_FILL:.3f}',
            'full' if full else 'hash', str(n),
            'full' if layout_full else 'hash', str(layout_n), 'yes',
            'border centreline = layout W; fill rectangle inflated by E for coverage; '
            'fit follows previous measured fill width']))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('# locked before the column-edges Excel export; model in lock-print-column-edges.py\n')
        f.write('\n'.join(out) + '\n')
    print(f'wrote {len(out) - 2} predictions to {out_path}')


if __name__ == '__main__':
    main()
