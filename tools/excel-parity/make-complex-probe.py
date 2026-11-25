#!/usr/bin/env python3
"""Probe cells for the complex-number functions (FnComplex.fs), as eval-cells-on-vm.sh takes them.

    python3 tools/excel-parity/make-complex-probe.py > complex-cells.txt

Excel's reading of complex text ("3+4i", "4j", "-i", exponents, spaces, mixed suffixes), its writing of a
result (15 significant digits, dropped zero parts and unit coefficients, the suffix kept), COMPLEX's
arguments, and every IM function over the four quadrants, zero, pure real and pure imaginary values,
both suffixes and values that overflow. complex-excel.tsv holds Excel
2019's values; function-parity-engine.fsx compares the engine with them.
"""

cells = []


def cell(formula):
    cells.append(formula)


def q(s):
    return '"' + s.replace('"', '""') + '"'


TEXTS = ['3+4i', '3-4i', '-3+4i', '-3-4i', '4i', '-4i', '+4i', 'i', '-i', '+i', '3', '-3', '+3', '0', '-0',
         '0i', '0+0i', '-0-0i', '3+i', '3-i', '3+1i', '3+0i', '0+4i', '3.5+4.25i', '1e3+2e-1i', '1E3+2E+2i',
         '1.5e-3i', '.5+.5i', '5.+5.i', '3+4j', '4j', 'j', '-j', '3+4I', '3+4J', '3 +4i', ' 3+4i', '3+4i ',
         '3+ 4i', 'i3', '3i+4', '3+4ii', '3++4i', '3+-4i', '--3', '3e+4i', '3e4i', '1e309', '1e308+1e308i',
         'abc', '', 'i+3', '3+4k', '3+4', '+', '-', 'e', '1e', '1e+i', '0x10', '3,4', '(3+4i)', '3+4i+5',
         '3.4.5', '1,000', '$3', '3%', '1/2', 'TRUE', '4i3', '1+i1', '2+3.i', '1e-400+1i', '00012+0004i',
         '1.23456789012345678+9.87654321098765432i', '+-1', '-+1i', 'ii', '1i2i', '3.0+4.0i', '7.000i',
         '1E+3', '1e3', '-1.5E-3-2.5E+3j', '1d3', '٣+4i', '3 +4i', '0j', '0.0', '1+0j',
         'I', 'J', '2E', '2E+', '2E+i', '1.e3i', '-.5j']

for t in TEXTS:
    cell('=IMREAL(%s)' % q(t))
    cell('=IMAGINARY(%s)' % q(t))
    cell('=IMSUM(%s)' % q(t))
    cell('=IMCONJUGATE(%s)' % q(t))

for t in ('CHAR(9)&"3+4i"', '"3"&CHAR(9)&"+4i"', '"3+4i"&CHAR(10)'):
    cell('=IMREAL(%s)' % t)
    cell('=IMSUM(%s)' % t)

# Values that are not text.
for v in ('5', '-2.5', '0', 'TRUE', 'FALSE', '#N/A', '1/0', '1e300', 'Z1', '""', '1E-300*1E-10', '123456789012345678'):
    for fn in ('IMREAL', 'IMAGINARY', 'IMSUM', 'IMABS', 'IMCONJUGATE', 'IMSQRT'):
        cell('=%s(%s)' % (fn, v))

# COMPLEX's writing.
NUMS = ['0', '1', '-1', '0.1', '1/3', '-2/3', '2/3*1E+20', '1E-20', '123456789012345678', '1.23456789012345E+15',
        '1E+15', '99999999999999.9', '999999999999999', '9999999999999999', '0.000123456789012345', '1E-300',
        '-1E+300', 'PI()', '0.5', '1E+21', '1E+100', '12345.6789', '-0.000001', '0.0001', '1.5E-5', '2^53',
        '1-2^-52', '1E-15', '1.2345678901234567E-5', '100', '1E+16', '-1E-16', '0.1+0.2']
for a in NUMS:
    cell('=COMPLEX(%s,1)' % a)
    cell('=COMPLEX(1,%s)' % a)
    cell('=COMPLEX(%s,%s)' % (a, a))
    cell('=COMPLEX(0,%s,"j")' % a)
    cell('=COMPLEX(%s,0)' % a)
for s in ('"i"', '"j"', '""', '"I"', '"J"', '"k"', '1', 'TRUE', '"ij"', ' "i "', 'Z1', '#N/A'):
    cell('=COMPLEX(3,4,%s)' % s)
    cell('=COMPLEX(0,1,%s)' % s)
for args in ('"3","4"', '"a",1', '1,"b"', 'TRUE,1', '1,FALSE', '#DIV/0!,1', '1,#N/A', 'Z1,Z2', '3,', ',4',
             '3.5,-1', '-0.0,-0.0', '1,-1,"j"', 'POWER(10,309),1'):
    cell('=COMPLEX(%s)' % args)

# Every one-argument IM function over the quadrants, zero, pure values, both suffixes.
VALUES = ['0', '1', '-1', 'i', '-i', '3+4i', '-3+4i', '-3-4i', '3-4i', '2.5', '-2.5', '0.5i', '-7i',
          '1e-5+1e-5i', '100+200i', '1.2345678+9.87654321i', '3+4j', '-2j', '0.25-0.75j', '710', '-710',
          '1+800i', '800+1i', '1e300+1e300i', '1e-300', '-1e-300i', '1e200+1e200i', '5e-324']
UNARY = ['IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH',
         'IMEXP', 'IMLN', 'IMLOG10', 'IMLOG2', 'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMSQRT', 'IMTAN']
for fn in UNARY:
    for v in VALUES:
        cell('=%s(%s)' % (fn, q(v)))

POWERS = ['2', '3', '0.5', '-1', '0', '1', '1.5', '-2.5', '10', '1/3', '-0.5', '100', '2000', '"2"', '"x"', 'TRUE']
for v in VALUES[:20]:
    for p in POWERS:
        cell('=IMPOWER(%s,%s)' % (q(v), p))

PAIR = ['0', '1', '-1', 'i', '3+4i', '-3-4i', '2.5', '0.5i', '1-1i', '1e-200', '1e200+1e200i', '3+4j']
for x in PAIR:
    for y in PAIR:
        cell('=IMDIV(%s,%s)' % (q(x), q(y)))
        cell('=IMSUB(%s,%s)' % (q(x), q(y)))
        cell('=IMPRODUCT(%s,%s)' % (q(x), q(y)))
        cell('=IMSUM(%s,%s)' % (q(x), q(y)))

# The aggregates over arrays, ranges, blanks and mixed suffixes.
for args in ('"1+i","2"', '{"1+i","2";"3i","-4"}', '"1+i","2j"', '"j","j"', '"i","j"', '"3","4j"', '"3j","4"',
             '"1+i",""', '"1+i",Z1:Z3', 'Z1:Z3', '"1+i",TRUE', '"1+i",5', '"1+i",{1,2,3}', '"1+i",#N/A',
             '"1+i","x"', '"0.1","0.2"', '"1e308","1e308"', '"1+2i","3+4i","5+6i","7+8i"', '{"1+i";"2"},"3"',
             '"1+i",{"2","x"}', '"1+i",{"2",""}', '"1+i",,"2"', '"j","0"', '"0","j"', '"2j","0j"', '"1e200i","1e200i"'):
    cell('=IMSUM(%s)' % args)
    cell('=IMPRODUCT(%s)' % args)
for args in ('"1+i","2j"', '"j","0"', '"3j","4"', '"3","4j"', '"0","0"', '"1+i",""', '"","1+i"',
             '"1+i",5', '5,"2j"', '"1+i",TRUE'):
    cell('=IMSUB(%s)' % args)
    cell('=IMDIV(%s)' % args)
cell('=IMPOWER("0",-1)')
cell('=IMPOWER("0",0)')
cell('=IMPOWER("0",0.5)')
cell('=IMPOWER("0j",2)')

# Where the written number turns to an exponent, and its 15-digit rounding.
for k in range(-26, 27):
    for m in ('1', '1.5', '1.23456789012345', '9.99999999999999', '9.999999999999999', '-1', '-4.56'):
        cell('=COMPLEX(%sE%+d,0)' % (m, k))
for x in ('0.1+0.7', '1/7*1E+20', '1/7*1E-5', '1/7*1E-6', '1/7*1E-7', '1/7*1E-8', '1/7*1E-9', '2/3*1E+15',
          '2/3*1E+14', '123456789012345.5', '0.5E-15', '5E-16', '4.9999999999999E-16'):
    cell('=COMPLEX(%s,%s)' % (x, x))
# More text forms.
for t in ['3+--4i', '3-+4i', '3--4i', '+3+4i', '-3--4i', '3+4.5e1i', '3+4e1.5i', '1e-308', '2.2250738585072014E-308',
          '2.2250738585072E-308', '2.225073858507E-308', '9.99999999999999E+307', '9.999999999999999E+307',
          '3 + 4 i', '3+4 i', ' i', '- i', '3 4i', '3i ', '1 000', '1,000+2,000i', '$3+4i', '3%+4%i', '1/2+i',
          '12:00', '1e1000i', '0e0', '-0i', '3+0.0000i', '1E+15+1E+15i', '  ', ' ', '3+4i+', '+i+i', '1+e',
          '1+ei', '1e5e5', '1.5.i', '3+.i', '3+.5i', '3+-.5i', '3e-2-4e-2j', '0.000000000000000000001i',
          '123456789012345678901234567890', '1e-307', '-1e-307i', '1e15', '3+4i ', 'i ', ' j', '3 ',
          '1+1E+308i', '1+9E+307i', '2E+307+2E+307i', 'TRUE+i', 'pi', '3,5+4i', '(3)', '3+(4)i']:
    cell('=IMREAL(%s)' % q(t))
    cell('=IMAGINARY(%s)' % q(t))
    cell('=IMSUM(%s)' % q(t))
# Overflowing and underflowing moduli and results.
for t in ['1e300', '1e300i', '1e200', '1e154', '1e155+1e155i', '1e-200', '1e-160+1e-160i', '3e307+3e307i',
          '-1e300', '1e-300i', '1e150+1e150i', '1e-155']:
    for fn in ('IMABS', 'IMSQRT', 'IMLN', 'IMARGUMENT', 'IMLOG10', 'IMCONJUGATE', 'IMEXP', 'IMSIN'):
        cell('=%s(%s)' % (fn, q(t)))
    cell('=IMPOWER(%s,2)' % q(t))
    cell('=IMPOWER(%s,0.5)' % q(t))
    cell('=IMPRODUCT(%s,%s)' % (q(t), q(t)))
    cell('=IMDIV("1",%s)' % q(t))
    cell('=IMDIV(%s,"1e-200")' % q(t))
    cell('=IMSUM(%s,%s)' % (q(t), q(t)))

for i, f in enumerate(cells):
    print('S\tA%d\t%s' % (i + 1, f))
