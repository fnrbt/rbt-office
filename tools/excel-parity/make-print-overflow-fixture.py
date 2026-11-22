#!/usr/bin/env python3
"""Create a public direct-PDF numeric overflow threshold fixture.

Usage: python3 tools/excel-parity/make-print-overflow-fixture.py OUT.xlsx
Widths B:V sweep 8.00 through 13.00 in quarter-character steps. Each
format/font has an occupied-neighbor row and 21 isolated-cell rows. The
last format is a held-out percentage with decimal places.
"""
import importlib.util
from pathlib import Path
import sys

spec = importlib.util.spec_from_file_location('pages', Path(__file__).with_name('make-print-pages-fixture.py'))
pages = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pages)

pages.FONTS += ['<sz val="10"/><name val="Arial"/><family val="2"/>']
formats = ['&quot;$&quot;#,##0', '_(&quot;$&quot;* #,##0_);_(&quot;$&quot;* (#,##0);_(&quot;$&quot;* &quot;-&quot;_);_(@_)', '0.00%',
           '&quot;$&quot;#,##0_);\\(&quot;$&quot;#,##0\\)',
           '_(&quot;$&quot; #,##0_);_(&quot;$&quot; (#,##0);_(&quot;$&quot; &quot;-&quot;_);_(@_)']
pages.FIGURE_NUMFMTS = '<numFmts count="5">' + ''.join(
    f'<numFmt numFmtId="{164+i}" formatCode="{f}"/>' for i, f in enumerate(formats)) + '</numFmts>'
pages.FIGURE_XFS = [f'<xf numFmtId="{fmt}" fontId="{font}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>'
                    for font in (0, 2) for fmt in (164, 165, 0, 166)]
pages.FIGURE_XFS.append('<xf numFmtId="167" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>')
pages.FIGURE_XFS += [f'<xf numFmtId="{fmt}" fontId="{font}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>'
                     for font, fmt in ((0, 167), (1, 164), (0, 168), (1, 168))]

sheets = []
for font in range(2):
    for fmt, label in enumerate(('Currency', 'Accounting', 'General', 'PercentHoldout')):
        cells = [(1, 1, f'{label} {"Calibri11" if font == 0 else "Arial10"}', 0)]
        cells += [(2, c, f'{8 + (c-2)/4:.2f}', 0) for c in range(2, 23)]
        style = 8 + font * 4 + fmt
        value = 12345.6789 if fmt == 3 else 1234567
        cells += [(3, 1, 'occupied', 0)] + [(3, c, value, style) for c in range(2, 23)]
        for c in range(2, 23):
            cells += [(c + 2, 1, f'free {8 + (c-2)/4:.2f}', 0), (c + 2, c, value, style)]
        sheets.append(pages.sheet(f'{label[:12]}{font}', cells,
                      cols={1: 18, **{c: 8 + (c-2)/4 for c in range(2, 23)}},
                      setup='<pageSetup paperSize="1" orientation="landscape" scale="100"/>',
                      options='<printOptions gridLines="1"/>'))
cells = [(1, 1, 'Bold currency Calibri11', 0)] + [(2, c, f'{8 + (c-2)/4:.2f}', 0) for c in range(2, 23)]
cells += [(3, 1, 'occupied', 0)] + [(3, c, 1234567, 16) for c in range(2, 23)]
for c in range(2, 23):
    cells += [(c + 2, 1, f'free {8 + (c-2)/4:.2f}', 0), (c + 2, c, 1234567, 16)]
sheets.append(pages.sheet('BoldCurrency', cells,
              cols={1: 18, **{c: 8 + (c-2)/4 for c in range(2, 23)}},
              setup='<pageSetup paperSize="1" orientation="landscape" scale="100"/>',
              options='<printOptions gridLines="1"/>'))
# Factorial controls: identical cells and widths, varying font weight, reserve, and print scale.
widths = (9.25, 9.50, 9.625, 9.75, 9.875, 10.00, 10.125, 10.25, 10.50)
for scale in (100, 75):
    for bold in (False, True):
        for mode in ('plain', 'reserve', 'accountNoFill'):
            style = {(False, 'plain'): 8, (False, 'reserve'): 17, (False, 'accountNoFill'): 19,
                     (True, 'plain'): 18, (True, 'reserve'): 16, (True, 'accountNoFill'): 20}[bold, mode]
            label = f'{"bold" if bold else "regular"} {mode} {scale}'
            cells = [(1, 1, label, 0)] + [(2, c, f'{w:g}', 0) for c, w in enumerate(widths, 2)]
            cells += [(3, 1, 'occupied', 0)] + [(3, c, 1234567, style) for c in range(2, 2 + len(widths))]
            for c, w in enumerate(widths, 2):
                cells += [(c + 2, 1, f'free {w:g}', 0), (c + 2, c, 1234567, style)]
            sheets.append(pages.sheet(f'F{scale}{"B" if bold else "R"}{mode}', cells,
                          cols={1: 15, **{c: w for c, w in enumerate(widths, 2)}},
                          setup=f'<pageSetup paperSize="1" orientation="landscape" scale="{scale}"/>',
                          options='<printOptions gridLines="1"/>'))
pages.sets = lambda: {'bestfit-figures-overflow': sheets}
pages.build(sys.argv[1], 'bestfit-figures-overflow')
