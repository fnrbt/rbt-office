#!/usr/bin/env python3
"""Summarises the full<->hash anchor measurement against the locked fit candidate.

Inputs:
  META.tsv            generated column metadata
  EXCEL-RESULTS.tsv   measure-print-fullhash-anchor.py output on the fresh Excel PDFs
  PREDICTIONS.tsv     predict-print-fullhash-anchor.py --measure EXCEL-RESULTS.tsv
  OUT.tsv             per-column table (written)
  [OURS-RESULTS.tsv]  same measurement on our own PDF render (baseline)

The per-sheet summary (also written to OUT.summary.tsv) reports:
  wide row      printed advance (runWidth; lastWordWidth for accounting), model advance,
                right gap to the physical cell edge, and the glyph-composition bias
  hash rows     the c interval implied by every printed hash count at its measured W
  transition    last-hash W, first-full W, the text-advance interval they imply once the
                hash-count c interval is used, and whether the candidate text width is in it
  decisions     Excel vs candidate vs baseline; false hashes and false fulls listed
  newly wrong   columns the baseline gets right and the candidate gets wrong

    python3 tools/excel-parity/analyze-print-fullhash-anchor.py META EXCEL-RESULTS PREDICTIONS OUT [OURS-RESULTS]
"""
import collections
import importlib.util
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location('fhpredict', os.path.join(_here, 'predict-print-fullhash-anchor.py'))
fhpredict = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(fhpredict)


def read_tsv(path):
    raw = [l.rstrip('\n') for l in open(path, encoding='utf-8') if l.strip()]
    raw = [l for l in raw if not l.startswith('#')]
    rows = [l.split('\t') for l in raw]
    head, rows = rows[0], [r for r in rows[1:] if r]
    return [dict(zip(head, r)) for r in rows]


def main():
    meta_path, excel_path, pred_path, out_path = sys.argv[1:5]
    ours_path = sys.argv[5] if len(sys.argv) > 5 else None
    meta = {(d['sheet'], int(d['column'])): d for d in read_tsv(meta_path) if d['row'] == '1'}
    excel = {(d['sheet'], int(d['column'])): d for d in read_tsv(excel_path) if d['row'] == '1'}
    pred = {(d['sheet'], int(d['column'])): d for d in read_tsv(pred_path)}
    ours = {}
    if ours_path and os.path.exists(ours_path):
        ours = {(d['sheet'], int(d['column'])): d for d in read_tsv(ours_path) if d['row'] == '1'}

    cache = {}
    cols = []
    for key in sorted(meta, key=lambda k: (k[0], k[1])):
        m, e, p = meta[key], excel.get(key), pred.get(key)
        if not e or not p:
            continue
        face, italic = m['face'], m['italic'] == '1'
        if (face, italic) not in cache:
            cache[(face, italic)] = fhpredict.font_metrics(os.path.join(
                fhpredict.ROOT, fhpredict.FONT_FILES[(face, italic)]))
        glyphs, upm, _ = cache[(face, italic)]
        scale = int(m['scale']) / 100.0
        dots = fhpredict.height_dots(float(m['size_pt']), scale)
        text_pt = sum(fhpredict.advance(glyphs, ch, dots) for ch in m['text'])
        number_pt = sum(fhpredict.advance(glyphs, ch, dots)
                        for ch in (m['text'].replace('$', '') if m['format'] == fhpredict.ACCT
                                   else m['text']))
        slots = 1 if m['format'] == fhpredict.ACCT else 0
        if slots:
            text_pt += fhpredict.advance(glyphs, '(', dots) + fhpredict.advance(glyphs, ')', dots)
        a_pt = fhpredict.advance(glyphs, '#', dots)
        c_pt = glyphs['0'] * float(m['size_pt']) * scale / 2.0 + 0.875
        if italic:
            c_pt += fhpredict.italic_shift(cache[(face, italic)][2], upm, dots)
        w_model = fhpredict.physical_width(float(m['width_chars']), scale)
        fit_model = text_pt <= w_model - c_pt + 1e-9
        n_model = 0 if fit_model else max(1, int((w_model - c_pt) // a_pt) - slots)
        w_meas = float(e['cellWidth']) if e['cellWidth'] else None
        fit_meas = None if w_meas is None else (text_pt <= w_meas - c_pt + 1e-9)
        n_meas = None if w_meas is None else (
            0 if fit_meas else max(1, int((w_meas - c_pt) // a_pt) - slots))
        locked_mismatch = (p['pred_full'] != ('yes' if fit_model else 'no')
                           or int(p['pred_n']) != n_model)
        excel_hash = int(e['nHashes']) > 0
        cand_hash = not fit_model
        o = ours.get(key)
        cols.append(dict(
            sheet=key[0], column=key[1], width=float(m['width_chars']), kind=m['kind'],
            face=face, size=m['size_pt'], italic=m['italic'], scale=m['scale'],
            fmt=m['format'], text=m['text'], excel_text=e['text_printed'],
            cellWidth=float(e['cellWidth']) if e['cellWidth'] else None,
            runWidth=float(e['runWidth']) if e['runWidth'] else None,
            lastWordWidth=float(e['lastWordWidth']) if e.get('lastWordWidth') else None,
            nWords=int(e.get('nWords') or 0),
            rightGap=float(e['rightGap']) if e['rightGap'] else None,
            run_pt=text_pt, number_pt=number_pt, a_pt=a_pt, c_pt=c_pt, slots=slots,
            pred_W=float(p['pred_W_pt']), margin=w_model - c_pt - text_pt,
            n_excel=int(e['nHashes']), excel_hash=excel_hash, cand_hash=cand_hash,
            n_pred=n_model, n_meas=n_meas,
            meas_hash=None if fit_meas is None else (not fit_meas),
            margin_meas=None if w_meas is None else (w_meas - c_pt - text_pt),
            locked_mismatch=locked_mismatch,
            ours_hash=(int(o['nHashes']) > 0) if o else None,
            ours_n=int(o['nHashes']) if o else None,
            diag=e['diagnostics']))

    out = ['sheet\tcolumn\tkind\twidth_chars\tface\tsize_pt\titalic\tscale\tformat\ttext\t'
           'excel_text\tn_excel\texcel_decision\tcand_decision\tcand_n\tcand_margin_pt\t'
           'cellWidth_pt\trunWidth_pt\tlastWordWidth_pt\tmodel_text_pt\tmodel_c_pt\tmodel_a_pt\t'
           'rightGap_pt\tours_decision\tours_n\tmeasuredW_decision\tmeasuredW_margin_pt\t'
           'locked_file_mismatch\tdiagnostics']
    for c in cols:
        def f(x):
            return '' if x is None else f'{x:.4f}'
        out.append('\t'.join([
            c['sheet'], str(c['column']), c['kind'], f'{c["width"]:g}', c['face'], c['size'],
            c['italic'], c['scale'], c['fmt'], c['text'], c['excel_text'], str(c['n_excel']),
            'hash' if c['excel_hash'] else 'full', 'hash' if c['cand_hash'] else 'full',
            str(c['n_pred']), f'{c["margin"]:.4f}', f(c['cellWidth']), f(c['runWidth']),
            f(c['lastWordWidth']), f'{c["run_pt"]:.4f}', f'{c["c_pt"]:.4f}', f'{c["a_pt"]:.4f}',
            f(c['rightGap']), '' if c['ours_hash'] is None else ('hash' if c['ours_hash'] else 'full'),
            '' if c['ours_n'] is None else str(c['ours_n']),
            '' if c['meas_hash'] is None else ('hash' if c['meas_hash'] else 'full'),
            f(c['margin_meas']) if c['margin_meas'] is not None else '',
            'yes' if c['locked_mismatch'] else 'no', c['diag']]))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write('\n'.join(out) + '\n')

    by_sheet = collections.OrderedDict()
    for c in cols:
        by_sheet.setdefault(c['sheet'], []).append(c)

    summary = ['sheet\tface\tsize_pt\titalic\tscale\tformat\tdense_columns\texcel_full\texcel_hash\t'
               'candidate_wrong\tcand_false_hash\tcand_false_full\tbaseline_wrong\tbase_false_hash\t'
               'base_false_full\tnewly_wrong\tmeasuredW_wrong\tlocked_mismatch\tcount_exact\tcount_checked\t'
               'c_lo_pt\tc_hi_pt\ttrans_lo_pt\ttrans_hi_pt\ttrans_ok\t'
               'wide_text\trunWidth_pt\tlastWordWidth_pt\tmodel_text_pt\tadvance_bias_pt\trightGap_pt']
    print()
    for sheet, items in by_sheet.items():
        dense = [c for c in items if c['kind'] == 'dense']
        fulls = [c for c in dense if not c['excel_hash']]
        hashes = [c for c in dense if c['excel_hash']]
        cand_wrong = [c for c in dense if c['cand_hash'] != c['excel_hash']]
        cand_fh = [c for c in cand_wrong if c['cand_hash']]
        cand_ff = [c for c in cand_wrong if not c['cand_hash']]
        base_wrong = [c for c in dense if c['ours_hash'] is not None and c['ours_hash'] != c['excel_hash']]
        base_fh = [c for c in base_wrong if c['ours_hash']]
        base_ff = [c for c in base_wrong if not c['ours_hash']]
        base_keys = {(c['sheet'], c['column']) for c in base_wrong}
        newly = [c for c in cand_wrong if (c['sheet'], c['column']) not in base_keys]
        cand_wrong_meas = [c for c in dense
                           if c['meas_hash'] is not None and c['meas_hash'] != c['excel_hash']]
        locked_mismatches = [c for c in items if c['locked_mismatch']]

        # c interval from every printed hash count: c in (W-(n+slots+1)a, W-(n+slots)a]
        c_lo, c_hi = -1e9, 1e9
        for c in hashes:
            if c['cellWidth'] is None:
                continue
            c_lo = max(c_lo, c['cellWidth'] - (c['n_excel'] + c['slots'] + 1) * c['a_pt'])
            c_hi = min(c_hi, c['cellWidth'] - (c['n_excel'] + c['slots']) * c['a_pt'])
        # transition: full fits iff text_adv <= W - c  =>  c <= Wf - text_adv and c > Wh - text_adv
        trans_lo = trans_hi = None
        trans_ok = ''
        text_model = items[0]['run_pt'] if items else None
        if hashes and fulls and text_model is not None:
            Wh = max(c['cellWidth'] for c in hashes)
            Wf = min(c['cellWidth'] for c in fulls)
            trans_lo, trans_hi = Wh - text_model, Wf - text_model
            ok = c_lo < trans_hi and c_hi > trans_lo
            trans_ok = 'yes' if ok else 'no'
        count_checked = len(hashes)
        count_exact = sum(1 for c in hashes if c['n_pred'] == c['n_excel'])
        wide = [c for c in items if c['kind'] == 'wide' and c['column'] >= 2][:7]
        wide_text = wide[0]['text'] if wide else ''
        rw = wide[0]['runWidth'] if wide else None
        lw = wide[0]['lastWordWidth'] if wide else None
        bias = (lw if (wide and wide[0]['slots']) else rw) - (wide[0]['number_pt'] if wide else 0.0)
        row = [
            sheet, items[0]['face'], items[0]['size'], items[0]['italic'], items[0]['scale'],
            items[0]['fmt'], str(len(dense)), str(len(fulls)), str(len(hashes)),
            str(len(cand_wrong)), str(len(cand_fh)), str(len(cand_ff)),
            str(len(base_wrong)), str(len(base_fh)), str(len(base_ff)), str(len(newly)),
            str(len(cand_wrong_meas)), str(len(locked_mismatches)),
            str(count_exact), str(count_checked),
            f'{c_lo:.4f}' if c_lo > -1e9 else '', f'{c_hi:.4f}' if c_hi < 1e9 else '',
            f'{trans_lo:.4f}' if trans_lo is not None else '',
            f'{trans_hi:.4f}' if trans_hi is not None else '', trans_ok,
            wide_text,
            f'{rw:.4f}' if rw is not None else '', f'{lw:.4f}' if lw is not None else '',
            f'{items[0]["number_pt"]:.4f}' if items else '',
            f'{bias:.4f}' if wide else '',
            f'{wide[0]["rightGap"]:.4f}' if wide and wide[0]['rightGap'] else '']
        summary.append('\t'.join(row))
        print(f'{sheet}: dense {len(dense)} (full {len(fulls)}, hash {len(hashes)}); '
              f'candidate wrong {len(cand_wrong)} (false-hash {len(cand_fh)}, false-full {len(cand_ff)}); '
              f'baseline wrong {len(base_wrong)}; newly wrong {len(newly)}; '
              f'count {count_exact}/{count_checked}; c='
              f'({"" if c_lo <= -1e9 else f"{c_lo:.3f}"},{"-" if c_hi >= 1e9 else f"{c_hi:.3f}"}]; '
              f'transition c=' +
              ('n/a' if trans_lo is None else f'({trans_lo:.3f},{trans_hi:.3f}] ok={trans_ok}') +
              f'; measured-W wrong {len(cand_wrong_meas)}; locked mismatches {len(locked_mismatches)}')
        for c in cand_wrong:
            print(f'    {"false-hash" if c["cand_hash"] else "false-full"}: col {c["column"]} '
                  f'w={c["width"]:g} W={c["cellWidth"]:.3f} model_text={c["run_pt"]:.3f} '
                  f'margin={c["margin"]:+.3f} excel={c["excel_text"]}')
    with open(out_path + '.summary.tsv', 'w', encoding='utf-8', newline='\n') as fh:
        fh.write('\n'.join(summary) + '\n')


if __name__ == '__main__':
    main()
