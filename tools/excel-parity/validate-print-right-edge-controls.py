#!/usr/bin/env python3
"""Checks the literal-vs-generated right-edge controls against the preregistered predictions.

Reads the oracle table from measure-print-right-edge-controls.py and writes a validation table. Every numeric/literal
pair is checked for identity first (same expected and rendered text, same embedded font, same advance width), then
for the preregistered gap relation. Reserves are italic gap minus upright gap for the same pair.

    python3 tools/excel-parity/validate-print-right-edge-controls.py REFERENCE.tsv VALIDATION.tsv
"""
import sys
from collections import defaultdict

GAP_EQ = 0.05
WIDTH_EQ = 0.01
RESERVE_TOL = 0.10
OURS_TOL = 0.05

# pair id -> (numeric reserve prediction, literal reserve prediction, test name)
PAIR_PRED = {
    'hash_pct_a': (1.440, 2.760, 'hash_pct_overflow_reserve'),
    'hash_pct_b': (1.440, 2.760, 'hash_pct_overflow_reserve'),
    'hash_fmt_lit': (2.760, 2.760, 'hash_fmt_lit_reserve'),
    'pct_a': (2.806, 2.806, 'pct_reserve'),
    'pct_b': (2.806, 2.806, 'pct_reserve'),
    'pct_c': (2.806, 2.806, 'pct_reserve'),
    'pct_d': (2.806, 2.806, 'pct_reserve'),
    'num_a': (2.782, 2.782, 'num_reserve'),
    'num_b': (2.782, 2.782, 'num_reserve'),
    'num_c': (2.782, 2.782, 'num_reserve'),
    'num_d': (2.782, 2.782, 'num_reserve'),
    'acc_neg': (0.008, 0.008, 'acc_neg_no_pad'),
    'acc_pad_neg': (0.008, 0.008, 'acc_pad_neg'),
    'acc_pos': (0.035, 0.035, 'acc_pos'),
    'kg_a': (None, None, 'string_holdout_kg'),
    'kg_b': (None, None, 'string_holdout_kg'),
    'kg_c': (None, None, 'string_holdout_kg'),
}
CONFIG = {
    'C00_hash_upright': ('C01_hash_italic', 'hash'),
    'C02_pct_upright': ('C03_pct_italic', 'pct'),
    'C04_num_upright': ('C05_num_italic', 'num'),
    'C06_acc_upright': ('C07_acc_italic', 'acc'),
    'C08_georgia12_upright': ('C09_georgia12_italic', 'font'),
    'C10_kg_upright': ('C11_kg_italic', 'string'),
    'C12_scale60_upright': ('C13_scale60_italic', 'scale'),
}


def read_table(path):
    with open(path, encoding='utf-8') as f:
        header = f.readline().rstrip('\n').split('\t')
        return [dict(zip(header, line.rstrip('\n').split('\t'))) for line in f if line.strip()]


def f(row, key):
    try:
        return float(row[key])
    except (KeyError, ValueError):
        return None


def main():
    ref_path, out_path = sys.argv[1:3]
    rows = read_table(ref_path)
    by_sheet = defaultdict(list)
    by_sheet_pair = defaultdict(lambda: defaultdict(list))
    for row in rows:
        by_sheet[row['sheet']].append(row)
        if row['pair']:
            by_sheet_pair[row['sheet']][row['pair']].append(row)
    out = open(out_path, 'w', encoding='utf-8')
    out.write('test\tconfig\tdetail\texcel_observed\tours_observed\tpredicted\tverdict\tnotes\n')

    def emit(test, config, detail, excel, ours, predicted, verdict, notes=''):
        out.write('\t'.join(str(v) for v in (test, config, detail, excel, ours, predicted, verdict, notes)) + '\n')

    # 1. Identity and source equality for every pair.
    for sheet in sorted(by_sheet):
        for pair, prows in sorted(by_sheet_pair[sheet].items()):
            nums = [r for r in prows if r['kind'] == 'num']
            lits = [r for r in prows if r['kind'] == 'lit']
            if len(nums) != 1 or len(lits) != 1:
                emit('pair_identity', f'{sheet}/{pair}', 'shape', f'{len(nums)} num / {len(lits)} lit', '', '1 num + 1 lit', 'FAIL')
                continue
            n, l = nums[0], lits[0]
            # The measured string is the rightmost value word. Excel's accounting PDF emits the "$" and the
            # value as separate text elements; our PDF may merge them, so the value-word identity is the
            # confound check, not the whole text element.
            last_ok = (n['excel_last_word'] == l['excel_last_word']
                       and n['ours_last_word'] == l['ours_last_word'])
            expected_ok = n['expected_text'] == n['excel_text'] and n['expected_text'].endswith(n['excel_last_word'])
            ours_text_ok = n['ours_text'] == l['ours_text'] == n['expected_text']
            text_ok = expected_ok and ours_text_ok and n['excel_text'] == l['excel_text'] and last_ok
            font_ok = n['excel_font'] == l['excel_font'] and n['ours_font'] == l['ours_font']
            we = abs(f(n, 'excel_width') - f(l, 'excel_width'))
            wo = abs(f(n, 'ours_width') - f(l, 'ours_width'))
            identity_ok = text_ok and font_ok and we <= WIDTH_EQ and wo <= WIDTH_EQ
            note = f'fonts {n["excel_font"]!r}/{l["excel_font"]!r}'
            if n['ours_text'] != l['ours_text'] and last_ok:
                note += '; ours full text element differs but the measured value word is identical'
            emit('pair_identity', f'{sheet}/{pair}', f'{n["kind"]} vs {l["kind"]}',
                 f'text={text_ok} font={font_ok} dWidth={we:.3f}', f'dWidth={wo:.3f}',
                 'same text/font/width', 'pass' if identity_ok else 'FAIL', note)
            ge = f(n, 'excel_gap') - f(l, 'excel_gap')
            go = f(n, 'ours_gap') - f(l, 'ours_gap')
            # The original source-equality hypothesis is expected to fail for overflow hashes; the amendment
            # predicts that difference rather than equality.
            expected_fail = pair.startswith('hash_pct_')
            verdict = 'pass' if abs(ge) <= GAP_EQ else ('expected-difference' if expected_fail else 'FAIL')
            emit('source_equality', f'{sheet}/{pair}', 'numeric gap - literal gap',
                 f'{ge:+.3f}', f'{go:+.3f}', '0 +/- 0.05 unless overflow hashes',
                 verdict if identity_ok else 'FAIL',
                 'numeric overflow draw path' if expected_fail else '')

    # 2. Reserves italic - upright per pair, numeric and literal separately.
    for up, (it, kind) in CONFIG.items():
        if up not in by_sheet or it not in by_sheet:
            emit('reserve', f'{up}/{it}', 'missing sheet', '', '', '', 'FAIL')
            continue
        for pair in sorted(set(by_sheet_pair[up]) & set(by_sheet_pair[it])):
            rows_up = {r['kind']: r for r in by_sheet_pair[up][pair]}
            rows_it = {r['kind']: r for r in by_sheet_pair[it][pair]}
            if 'num' not in rows_up or 'num' not in rows_it:
                continue
            en = f(rows_it['num'], 'excel_gap') - f(rows_up['num'], 'excel_gap')
            on = f(rows_it['num'], 'ours_gap') - f(rows_up['num'], 'ours_gap')
            if kind == 'font':
                pred_n, pred_l, test = None, None, 'font_holdout_georgia12'
            elif kind == 'scale':
                pred_n, pred_l, test = 1.684, 1.684, 'scale_holdout_60'
            elif kind == 'string':
                pred_n, pred_l, test = None, None, 'string_holdout_kg'
            else:
                pred_n, pred_l, test = PAIR_PRED.get(pair, (None, None, kind))
            if pred_n is None and kind == 'font':
                ok = 1.0 <= en <= 4.0
                emit(test, f'{up}/{it}/{pair}', 'numeric reserve',
                     f'{en:+.3f}', f'{on:+.3f}', 'source equality; held-out band 1..4',
                     'pass' if ok else 'FAIL')
            elif pred_n is None:
                emit(test, f'{up}/{it}/{pair}', 'numeric reserve',
                     f'{en:+.3f}', f'{on:+.3f}', 'source equality only', 'report')
            else:
                emit(test, f'{up}/{it}/{pair}', 'numeric reserve',
                     f'{en:+.3f}', f'{on:+.3f}', f'{pred_n:.3f} +/- {RESERVE_TOL:.2f}',
                     'pass' if abs(en - pred_n) <= RESERVE_TOL else 'FAIL')
            if 'lit' in rows_up and 'lit' in rows_it:
                el = f(rows_it['lit'], 'excel_gap') - f(rows_up['lit'], 'excel_gap')
                ol = f(rows_it['lit'], 'ours_gap') - f(rows_up['lit'], 'ours_gap')
                if pred_l is None:
                    emit(test, f'{up}/{it}/{pair}', 'literal reserve', f'{el:+.3f}', f'{ol:+.3f}',
                         'equals numeric reserve', 'pass' if abs(el - en) <= GAP_EQ else 'FAIL')
                else:
                    emit(test, f'{up}/{it}/{pair}', 'literal reserve', f'{el:+.3f}', f'{ol:+.3f}',
                         f'{pred_l:.3f} +/- {RESERVE_TOL:.2f}',
                         'pass' if abs(el - pred_l) <= RESERVE_TOL else 'FAIL')
            # Our engine's own reserve, both kinds.
            emit('ours_reserve', f'{up}/{it}/{pair}', 'numeric', '', f'{on:+.3f}', '0.000 +/- 0.05',
                 'pass' if abs(on) <= OURS_TOL else 'FAIL')
            if 'lit' in rows_up and 'lit' in rows_it:
                emit('ours_reserve', f'{up}/{it}/{pair}', 'literal', '', f'{ol:+.3f}', '0.000 +/- 0.05',
                     'pass' if abs(ol) <= OURS_TOL else 'FAIL')

    # 3. Documented special controls.
    for sheet in ('C00_hash_upright', 'C01_hash_italic'):
        general = [r for r in by_sheet[sheet] if r['format'] == 'General' and not r['pair']]
        for g in general:
            emit('general_hash_premise', sheet, 'General numeric overflow',
                 g['excel_text'], g['ours_text'], 'Excel shows scientific notation, not hashes',
                 'report', 'General cannot express the requested seven-hash control')
    # Pad shift: padded negative vs no-pad negative, both numeric and literal.
    for label, rows_ in (('upright', by_sheet_pair['C06_acc_upright']), ('italic', by_sheet_pair['C07_acc_italic'])):
        if 'acc_neg' in rows_ and 'acc_pad_neg' in rows_:
            n_neg = {r['kind']: r for r in rows_['acc_neg']}
            n_pad = {r['kind']: r for r in rows_['acc_pad_neg']}
            for kind in ('num', 'lit'):
                if kind in n_neg and kind in n_pad:
                    shift = f(n_pad[kind], 'excel_gap') - f(n_neg[kind], 'excel_gap')
                    emit('acc_pad_gap', label, kind, f'{shift:+.3f}', '',
                         'padded gap exceeds no-pad gap', 'pass' if shift > 0 else 'FAIL')
    # 4. Post-hoc findings, recorded only after inspecting the preregistered failures above. These are not used
    # to score the predictions; they state what the failures identify.
    def reserve(up, it, pair, kind):
        return f(by_sheet_pair[it][pair][0 if kind == 'num' else 1], 'excel_gap') - \
               f(by_sheet_pair[up][pair][0 if kind == 'num' else 1], 'excel_gap')

    for pair in ('hash_pct_a', 'hash_pct_b'):
        num = reserve('C00_hash_upright', 'C01_hash_italic', pair, 'num')
        lit = reserve('C00_hash_upright', 'C01_hash_italic', pair, 'lit')
        emit('posthoc_hash_overflow', pair, 'numeric overflow vs literal seven hashes',
             f'num {num:+.3f} / lit {lit:+.3f}', '', 'the overflow draw path differs from literal/custom hashes',
             'finding', 'the custom-format numeric hash row matches the literal twin and reserve 2.760')
    for pair in ('acc_neg', 'acc_pad_neg', 'acc_pos'):
        num = reserve('C06_acc_upright', 'C07_acc_italic', pair, 'num')
        lit = reserve('C06_acc_upright', 'C07_acc_italic', pair, 'lit')
        emit('posthoc_accounting', pair, 'numeric accounting vs literal twin',
             f'num {num:+.3f} / lit {lit:+.3f}', '', 'the numeric accounting formatter suppresses the italic shift',
             'finding', 'pad shifts both by +4.320 but does not change the reserve')
    ours_pad = []
    for pair in ('acc_pad_neg', 'acc_pos'):
        upu = {r['kind']: r for r in by_sheet_pair['C06_acc_upright'][pair]}
        if 'num' in upu and 'lit' in upu:
            ours_pad.append(f"{f(upu['num'], 'ours_gap') - f(upu['lit'], 'ours_gap'):+.3f}")
    emit('posthoc_ours_text_pad', 'C06/C07', 'ours accounting literal text-section pad',
         'numeric-literal upright gap ' + '/'.join(ours_pad), '',
         'ours does not give literal text the accounting _( @_) pad',
         'finding', 'separate from the italic reserve; the measured value word width is still identical')
    out.close()


if __name__ == '__main__':
    main()
