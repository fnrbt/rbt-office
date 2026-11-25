#!/usr/bin/env python3
"""Probe cells for ERF/ERFC, the BESSEL functions and CONVERT (FnEngineering.fs), as eval-cells-on-vm.sh
takes them.

    python3 tools/excel-parity/make-engineering-probe.py > engineering-cells.txt

ERF and ERFC over [-6, 6] and beyond (where they saturate or underflow), the four BESSEL functions over
orders 0 to 20 and x from tiny to large (negative x, fractional and negative orders); CONVERT's every unit
to and from its category's base unit, every unit with the prefixes k, m, da and ki/Ki (and all prefixes
for a few), unit names that only look like a prefixed unit, temperatures, units of other categories and
names Excel does not know. engineering-excel.tsv holds Excel
2019's values; function-parity-engine.fsx compares the engine with them.
"""

cells = []


def cell(formula):
    cells.append(formula)


# ERF, ERFC.
xs = [round(-6 + 0.1 * k, 1) for k in range(121)]
xs += [6.5, 7, 9, 10, 26, 26.5, 27, 27.3, 28, 30, 100, -7, -10, -27, -100, 1e-10, 1e-300, -1e-10, 0.4769362762044699,
       0.46875, 0.4687, 0.4688, 4, 5.9, 3.999, 2.5e-8, 1e-5, 0.01, 0.15, 0.65, 0.66, 0.67, 1.2, 1.45, 1.6, 2.3]
for x in xs:
    cell('=ERF(%r)' % x)
    cell('=ERFC(%r)' % x)
for x in xs[::3]:
    cell('=ERF.PRECISE(%r)' % x)
    cell('=ERFC.PRECISE(%r)' % x)
for a, b in [(0, 1), (1, 0), (0.5, 2), (-1, 1), (2, 2), (-3, -2), (1, 30), (-30, 30), (0.1, 0.10001), (5, 6),
             (3, 4), (0.3, 0.4)]:
    cell('=ERF(%r,%r)' % (a, b))
for args in ('"a"', 'TRUE', '"1"', '#N/A', '1,"b"', '1,TRUE', 'Z1', '1,Z1', '-1,'):
    cell('=ERF(%s)' % args)
    cell('=ERFC(%s)' % args.split(',')[0] if args else '=ERFC()')
    cell('=ERF.PRECISE(%s)' % args.split(',')[0] if args else '=ERF.PRECISE()')

# BESSEL.
BX = [0, 0.001, 0.1, 0.5, 1, 1.5, 2, 2.5, 3, 5, 7.5, 10, 15, 20, 30, 50, 100, -1, -2.5, -10, 1e-8, 0.25, 4.5, 8]
BN = ['0', '1', '2', '3', '5', '10', '20', '1.9', '-1', '50']
for fn in ('BESSELI', 'BESSELJ', 'BESSELK', 'BESSELY'):
    for x in BX:
        for n in BN:
            cell('=%s(%r,%s)' % (fn, x, n))
    for args in ('500,0', '1000,1', '700,0', '710,0', '713,2', '1E+5,0', '1E+8,1', '1E+9,0', '134217728,0',
                 '"a",1', '1,"a"', 'TRUE,1', '"2",1', '1,', '2,Z1', '#N/A,1', '1,#DIV/0!', '1,-0.5', '1,0.999',
                 '40,30', '1,100', '0.1,30', '150,3', '25,24'):
        cell('=%s(%s)' % (fn, args))

# CONVERT: the units, by category with each category's base unit first.
UNITS = {
    'mass': ['g', 'sg', 'lbm', 'u', 'ozm', 'grain', 'cwt', 'shweight', 'uk_cwt', 'lcwt', 'hweight', 'stone', 'ton',
             'uk_ton', 'LTON', 'brton', 'pweight'],
    'distance': ['m', 'mi', 'Nmi', 'in', 'ft', 'yd', 'ang', 'ell', 'ly', 'parsec', 'pc', 'Picapt', 'Pica', 'pica',
                 'survey_mi'],
    'time': ['sec', 'yr', 'day', 'd', 'hr', 'mn', 'min', 's'],
    'pressure': ['Pa', 'p', 'atm', 'at', 'mmHg', 'psi', 'Torr'],
    'force': ['N', 'dyn', 'dy', 'lbf', 'pond'],
    'energy': ['J', 'e', 'c', 'cal', 'eV', 'ev', 'HPh', 'hh', 'Wh', 'wh', 'flb', 'BTU', 'btu'],
    'power': ['W', 'HP', 'h', 'PS', 'w'],
    'magnetism': ['T', 'ga'],
    'temperature': ['K', 'C', 'cel', 'F', 'fah', 'kel', 'Rank', 'Reau'],
    'volume': ['m3', 'tsp', 'tspm', 'tbs', 'oz', 'cup', 'pt', 'us_pt', 'uk_pt', 'qt', 'uk_qt', 'gal', 'uk_gal', 'l',
               'L', 'lt', 'ang3', 'ang^3', 'barrel', 'bushel', 'ft3', 'ft^3', 'in3', 'in^3', 'ly3', 'ly^3', 'm^3',
               'mi3', 'mi^3', 'yd3', 'yd^3', 'Nmi3', 'Nmi^3', 'Picapt3', 'Picapt^3', 'Pica3', 'Pica^3', 'GRT',
               'regton', 'MTON'],
    'area': ['m2', 'uk_acre', 'us_acre', 'ang2', 'ang^2', 'ar', 'ft2', 'ft^2', 'ha', 'in2', 'in^2', 'ly2', 'ly^2',
             'm^2', 'Morgen', 'mi2', 'mi^2', 'Nmi2', 'Nmi^2', 'Picapt2', 'Pica2', 'Pica^2', 'Picapt^2', 'yd2',
             'yd^2'],
    'information': ['bit', 'byte'],
    'speed': ['m/s', 'admkn', 'kn', 'm/h', 'm/hr', 'm/sec', 'mph'],
}
NOT_UNITS = ['kph', 'ozfl', 'Ang', 'psia', 'tonne', 't', 'acre', 'bar', 'Btu', 'Gal', 'gram', 'M', 'G', 'Sec', 'SEC',
             'mile', 'foot', 'hp', 'Hp', 'j', 'n', 'pa', 'k', 'B', 'b', 'yr2', 'm4', 'm^4', 'sec2', 'm2 ', ' m',
             'ft/s', 'km/h/h', '', 'lbm2', 'ozm3', 'g2', 'g^2', 'in^2^2', 'm3^', 'mm^', 'Lt', 'LT']
PREFIXES = ['Y', 'Z', 'E', 'P', 'T', 'G', 'M', 'k', 'h', 'da', 'e', 'd', 'c', 'm', 'u', 'n', 'p', 'f', 'a', 'z',
            'y', 'Yi', 'Zi', 'Ei', 'Pi', 'Ti', 'Gi', 'Mi', 'ki', 'Ki', 'K', 'D', 'H', 'µ', 'μ']

for cat, units in UNITS.items():
    base = units[0]
    for u in units:
        cell('=CONVERT(1,"%s","%s")' % (u, base))
        cell('=CONVERT(1,"%s","%s")' % (base, u))
        cell('=CONVERT(2.5,"%s","%s")' % (u, u))
        for p in ('k', 'm', 'da', 'ki', 'Ki', 'G', 'u'):
            cell('=CONVERT(1,"%s%s","%s")' % (p, u, u))
    # Products of the factors: every pair of the category one way.
    for i, u in enumerate(units):
        for v in units[i + 1:i + 6]:
            cell('=CONVERT(3.7,"%s","%s")' % (u, v))
for u in ('g', 'm', 'bit', 'byte', 'm2', 'm3', 'l', 'J', 'sec', 'Pa', 'ang^2', 'mmHg', 'K', 'm/h'):
    for p in PREFIXES:
        cell('=CONVERT(1,"%s%s","%s")' % (p, u, u))
for name in ('dam', 'cc', 'ms', 'mm2', 'cm3', 'km^2', 'Ym3', 'uu', 'ue', 'mpc', 'kpc', 'Mly', 'km/h', 'km/hr',
             'km/s', 'Mbyte', 'Mibyte', 'kibit', 'KiB', 'Mibit', 'mmm', 'kkm', 'dadam', 'hha', 'kcal', 'mcal', 'MeV',
             'keV', 'kWh', 'kW', 'mW', 'kN', 'kPa', 'hPa', 'MPa', 'katm', 'kdyn', 'kpond', 'mT', 'kga', 'mK', 'kkel',
             'kC', 'kF', 'mL', 'ml', 'dl', 'cl', 'kl', 'mlt', 'cang', 'nang', 'kar', 'har', 'dar', 'mm/s', 'mm/sec',
             'kmi', 'kft', 'kNmi', 'kHP', 'ksg', 'klbm', 'mmi', 'ma', 'mn2', 'kmn', 'ksec', 'kday', 'kyr', 'khr',
             'kmin', 'Ppc', 'dapc', 'Tly', 'kly2', 'kly^3', 'km3', 'km^3', 'mm^3', 'Gm2', 'dam2', 'dam3', 'Ehbit',
             'Ybyte', 'Yibyte', 'Zibit', 'yb', 'dabyte', 'cbit', 'mbit', 'kbyte'):
    cell('=CONVERT(1,"%s","%s")' % (name, 'm'))
    cell('=CONVERT(1,"%s","%s")' % (name, 'J'))
    cell('=CONVERT(1,"%s","%s")' % (name, 'bit'))
    cell('=CONVERT(1,"%s","%s")' % (name, 'g'))
for name in NOT_UNITS:
    cell('=CONVERT(1,"%s","m")' % name)
    cell('=CONVERT(1,"m","%s")' % name)
    cell('=CONVERT(1,"%s","%s")' % (name, name))
temps = ['K', 'C', 'F', 'Rank', 'Reau', 'kel', 'cel', 'fah', 'mK', 'kK']
for a in temps:
    for b in temps:
        for t in ('-40', '0', '37.5', '100', '-273.15', '1000', '-500'):
            cell('=CONVERT(%s,"%s","%s")' % (t, a, b))
# Units of different categories, and argument handling.
cats = list(UNITS)
for i, a in enumerate(cats):
    for b in cats[i + 1:]:
        cell('=CONVERT(1,"%s","%s")' % (UNITS[a][1], UNITS[b][1]))
for args in ('"a","m","ft"', 'TRUE,"m","ft"', '"2","m","ft"', '#N/A,"m","ft"', '1,#N/A,"ft"', '1,"m",#DIV/0!',
             '1,1,"m"', '1,"m",1', 'Z1,"m","ft"', '1,Z1,"m"', '-1,"m","ft"', '9E+307,"km","m"',
             '9E+307,"ly","mm"', '"1e3","m","km"', '1,"in","ft"', '1,"km","mi"', '1,"kg","lbm"',
             '100,"m/s","mph"', '1,"Pica","pt"', '1,"Pica","Picapt"', '1,"ft2","m2"', '1,"m","in"',
             '6,"C","F"', '1,"lbm","kg"', '1,"BTU","J"', '1,"gal","l"', '1,"mi2","ha"', '1,"yr","day"'):
    cell('=CONVERT(%s)' % args)

for i, f in enumerate(cells):
    print('S\tA%d\t%s' % (i + 1, f))
