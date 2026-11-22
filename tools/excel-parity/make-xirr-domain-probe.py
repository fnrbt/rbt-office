#!/usr/bin/env python3
"""Generate synthetic XIRR inputs for eval-cells-on-vm.sh.

    python3 tools/excel-parity/make-xirr-domain-probe.py /tmp/xirr-domain-cells.tsv
    tools/excel-parity/eval-cells-on-vm.sh /tmp/xirr-domain-cells.tsv /tmp/xirr-domain-output.tsv

Each row's two requested cells are the direct XIRR and its IFERROR wrapper.
Input order defines row numbers; results include Excel's exact binary encoding.
"""
import sys

# name, cashflows, day offsets from 2020-01-01
SERIES = [
    ('positive', (-1000, 1100), (0, 365)),
    ('negative', (-1000, 900), (0, 365)),
    ('negative-half', (-1000, 500), (0, 365)),
    ('zero', (-1000, 1000), (0, 365)),
    ('positive-reversed', (1000, -1100), (0, 365)),
    ('negative-reversed', (1000, -900), (0, 365)),
    ('positive-small', (-1, 1.1), (0, 365)),
    ('negative-half-year', (-1000, 950), (0, 180)),
    ('positive-large', (-1000000000, 1100000000), (0, 365)),
    ('half-year', (-1000, 1050), (0, 180)),
    ('two-years', (-1000, 1210), (0, 730)),
    ('near-negative-one', (-1000, 10), (0, 365)),
    ('multiple-roots-reversed', (100, -230, 132), (0, 365, 730)),
    ('three-positive', (-100, 50, 100), (0, 365, 730)),
    ('three-positive-reversed', (100, -50, -100), (0, 365, 730)),
    ('large-positive-root', (-1000, 4000), (0, 365)),
    ('multiple-roots', (-100, 230, -132), (0, 365, 730)),
    ('wide-multiple-roots', (-100, 525, -500), (0, 365, 730)),
    ('zero-and-half-root', (-100, 250, -150), (0, 365, 730)),
    ('opposite-signed-roots', (-100, 300, -150), (0, 365, 730)),
    ('no-root', (-100, 100, -100), (0, 365, 730)),
    ('three-positive-root', (-1000, 100, 1000), (0, 180, 365)),
    ('three-negative-root', (-1000, 100, 800), (0, 180, 365)),
    ('near-zero-positive', (-1000, 1000.000001), (0, 365)),
    ('near-zero-negative', (-1000, 999.999999), (0, 365)),
]
GUESSES = ('default', '-2', '-1', '-0.999', '-0.5', '-0.1', '-0.01',
           '-0.000000001', '-0.000000000001', '-0.0000000000000001', '0',
           '0.0000000000000001', '0.000000001', '0.1', '0.11', '0.15', '0.19', '0.2', '0.25', '0.5',
           '0.75', '0.9', '0.99', '1', '1.0001', '1.01', '2', '4', '10')

with open(sys.argv[1], 'w', encoding='utf-8', newline='\n') as out:
    for row, (name, flows, days, guess) in enumerate(
            ((name, flows, days, guess) for name, flows, days in SERIES for guess in GUESSES), 1):
        values = ','.join(map(str, flows))
        dates = ','.join(str(43831 + day) for day in days)
        args = f'{{{values}}},{{{dates}}}' + ('' if guess == 'default' else f',{guess}')
        formula = f'XIRR({args})'
        out.write(f'Domain\tA{row}\t{name}:{guess}\n')
        out.write(f'Domain\tB{row}\t={formula}\n')
        out.write(f'Domain\tC{row}\t=IFERROR({formula},"error")\n')
