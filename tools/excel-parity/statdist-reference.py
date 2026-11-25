#!/usr/bin/env python3
"""High-precision reference values for the statdist probe cells.

    python3 tools/excel-parity/statdist-reference.py CELLS-DIR OUT.tsv

For every probe formula of make-statdist-probe.py that is a distribution
function of plain numbers (no references, text, logicals or omitted arguments)
it computes the true value in 50-digit decimal arithmetic (Python's decimal:
log-gamma by Stirling's series, the incomplete gamma and beta functions by
their series and continued fractions, inverses by Newton's method from the
double-precision answer) and writes sheet, address, value (17 significant
digits) - or nothing for a cell it does not cover or whose arguments are
outside the domain. Each argument is taken as the double Excel parses it to.
statdist-compare.fsx reads the file to tell, of every disagreement, whether
Excel or the engine is the nearer to the truth.
"""
import math
import os
import re
import sys
from decimal import Decimal as D, getcontext, localcontext

getcontext().prec = 50
EPS = D('1e-45')
PI = D('3.141592653589793238462643383279502884197169399375105820974944592307816406286')
LN2PI_HALF = (2 * PI).ln() / 2
BERN = [D(1) / 6, D(-1) / 30, D(1) / 42, D(-1) / 30, D(5) / 66, D(-691) / 2730, D(7) / 6, D(-3617) / 510,
        D(43867) / 798, D(-174611) / 330, D(854513) / 138, D(-236364091) / 2730]
LIMIT = 2_000_000


class Skip(Exception):
    pass


def lngamma(z):
    shift = D(0)
    while z < 40:
        shift -= z.ln()
        z += 1
    s = (z - D('0.5')) * z.ln() - z + LN2PI_HALF
    zz = z * z
    p = z
    for k, b in enumerate(BERN, 1):
        s += b / (2 * k * (2 * k - 1) * p)
        p *= zz
    return s + shift


def gamma_pq(a, x):
    """(P(a, x), Q(a, x))."""
    if x <= 0:
        return D(0), D(1)
    lead = (a * x.ln() - x - lngamma(a)).exp()
    if x < a + 1:
        s = t = 1 / a
        n = 0
        while True:
            n += 1
            t = t * x / (a + n)
            s += t
            if t < s * EPS:
                break
            if n > LIMIT:
                raise Skip
        p = lead * s
        return p, 1 - p
    tiny = D('1e-300')
    b = x + 1 - a
    c = 1 / tiny
    d = 1 / b
    h = d
    i = 0
    while True:
        i += 1
        an = -i * (i - a)
        b += 2
        d = an * d + b
        d = d if abs(d) > tiny else tiny
        c = b + an / c
        c = c if abs(c) > tiny else tiny
        d = 1 / d
        de = d * c
        h *= de
        if abs(de - 1) < EPS:
            break
        if i > LIMIT:
            raise Skip
    q = lead * h
    return 1 - q, q


def beta_ij(a, b, x):
    """(I_x(a, b), 1 - I_x(a, b))."""
    if x <= 0:
        return D(0), D(1)
    if x >= 1:
        return D(1), D(0)
    if x > (a + 1) / (a + b + 2):
        j, i = beta_ij(b, a, 1 - x)
        return i, j
    lead = (a * x.ln() + b * (1 - x).ln() - (lngamma(a) + lngamma(b) - lngamma(a + b))).exp()
    tiny = D('1e-300')
    qab, qap, qam = a + b, a + 1, a - 1
    c = D(1)
    d = 1 - qab * x / qap
    d = 1 / (d if abs(d) > tiny else tiny)
    h = d
    m = 0
    while True:
        m += 1
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 + aa * d
        d = 1 / (d if abs(d) > tiny else tiny)
        c = 1 + aa / c
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 + aa * d
        d = 1 / (d if abs(d) > tiny else tiny)
        c = 1 + aa / c
        de = d * c
        h *= de
        if abs(de - 1) < EPS:
            break
        if m > LIMIT:
            raise Skip
    i = lead * h / a
    return i, 1 - i


def gamma_pdf(a, x):
    return ((a - 1) * x.ln() - x - lngamma(a)).exp()


def beta_pdf(a, b, x):
    return ((a - 1) * x.ln() + (b - 1) * (1 - x).ln() - (lngamma(a) + lngamma(b) - lngamma(a + b))).exp()


def t_upper(n, t):
    if t == 0:
        return D('0.5')
    x = n / (n + t * t)
    i, _ = beta_ij(n / 2, D('0.5'), x)
    return i / 2 if t > 0 else 1 - i / 2


def t_pdf(n, t):
    return (lngamma((n + 1) / 2) - lngamma(n / 2) - (n * PI).ln() / 2 - (n + 1) / 2 * (1 + t * t / n).ln()).exp()


def f_tails(m, n, x):
    if x <= 0:
        return D(0), D(1)
    return beta_ij(m / 2, n / 2, m * x / (m * x + n))


def f_pdf(m, n, x):
    return (m / 2 * (m / n).ln() + (m / 2 - 1) * x.ln() - (m + n) / 2 * (1 + m * x / n).ln()
            - (lngamma(m / 2) + lngamma(n / 2) - lngamma((m + n) / 2))).exp()


def norm_cdf(z):
    if z == 0:
        return D('0.5')
    _, q = gamma_pq(D('0.5'), z * z / 2)
    return q / 2 if z < 0 else 1 - q / 2


def binom_pmf(k, n, p):
    if p == 0:
        return D(1) if k == 0 else D(0)
    if p == 1:
        return D(1) if k == n else D(0)
    return (lngamma(n + 1) - lngamma(k + 1) - lngamma(n - k + 1) + k * p.ln() + (n - k) * (1 - p).ln()).exp()


def newton(f, df, x, target, lo=None, hi=None):
    """Refine x with f(x) = target."""
    x = D(x)
    for _ in range(8):
        fx = f(x) - target
        d = df(x)
        if d == 0:
            break
        nx = x - fx / d
        if lo is not None and nx <= lo:
            nx = (x + lo) / 2
        if hi is not None and nx >= hi:
            nx = (x + hi) / 2
        if abs(nx - x) <= abs(x) * D('1e-40'):
            x = nx
            break
        x = nx
    return x


def trunc(x):
    return D(int(x))


def dist(name, a, start):
    """The true value of NAME(a...); `start` the double answer, for inverses."""
    g = lambda i: a[i]
    start = D(start)
    if name == 'T.DIST':
        n = trunc(g(1))
        if g(2) != 0:
            return t_upper(n, -g(0)) if g(0) < 0 else 1 - t_upper(n, g(0))
        return t_pdf(n, g(0))
    if name == 'T.DIST.2T':
        return min(D(1), 2 * t_upper(trunc(g(1)), g(0)))
    if name == 'T.DIST.RT':
        return t_upper(trunc(g(1)), g(0))
    if name == 'TDIST':
        return min(D(1), trunc(g(2)) * t_upper(trunc(g(1)), g(0)))
    if name in ('T.INV', 'T.INV.2T', 'TINV'):
        n = trunc(g(1))
        p = g(0)
        if name == 'T.INV':
            target_upper = 1 - p
        else:
            if p > 1:
                raise Skip
            target_upper = p / 2
        return newton(lambda t: t_upper(n, t), lambda t: -t_pdf(n, t), start, target_upper)
    if name == 'CONFIDENCE.T':
        n = trunc(g(2))
        t = newton(lambda t: t_upper(n - 1, t), lambda t: -t_pdf(n - 1, t), start / g(1) * n.sqrt(), g(0) / 2)
        return t * g(1) / n.sqrt()
    if name == 'F.DIST':
        m, n = trunc(g(1)), trunc(g(2))
        return f_tails(m, n, g(0))[0] if g(3) != 0 else f_pdf(m, n, g(0))
    if name in ('F.DIST.RT', 'FDIST'):
        return f_tails(trunc(g(1)), trunc(g(2)), g(0))[1]
    if name in ('F.INV', 'F.INV.RT', 'FINV'):
        m, n = trunc(g(1)), trunc(g(2))
        target = g(0) if name == 'F.INV' else 1 - g(0)
        return newton(lambda x: f_tails(m, n, x)[0], lambda x: f_pdf(m, n, x), start, target, lo=D(0))
    if name == 'CHISQ.DIST':
        k = trunc(g(1)) / 2
        return gamma_pq(k, g(0) / 2)[0] if g(2) != 0 else gamma_pdf(k, g(0) / 2) / 2
    if name in ('CHISQ.DIST.RT', 'CHIDIST'):
        return gamma_pq(trunc(g(1)) / 2, g(0) / 2)[1]
    if name in ('CHISQ.INV', 'CHISQ.INV.RT', 'CHIINV'):
        k = trunc(g(1)) / 2
        if name == 'CHISQ.INV':
            return 2 * newton(lambda x: gamma_pq(k, x)[0], lambda x: gamma_pdf(k, x), start / 2, g(0), lo=D(0))
        return 2 * newton(lambda x: gamma_pq(k, x)[1], lambda x: -gamma_pdf(k, x), start / 2, g(0), lo=D(0))
    if name in ('BETA.DIST', 'BETADIST'):
        lower_i = 4 if name == 'BETA.DIST' else 3
        lo = a[lower_i] if len(a) > lower_i else D(0)
        hi = a[lower_i + 1] if len(a) > lower_i + 1 else D(1)
        t = (g(0) - lo) / (hi - lo)
        if name == 'BETADIST' or g(3) != 0:
            return beta_ij(g(1), g(2), t)[0]
        return beta_pdf(g(1), g(2), t) / (hi - lo)
    if name in ('BETA.INV', 'BETAINV'):
        lo = a[3] if len(a) > 3 else D(0)
        hi = a[4] if len(a) > 4 else D(1)
        x = newton(lambda x: beta_ij(g(1), g(2), x)[0], lambda x: beta_pdf(g(1), g(2), x), (start - lo) / (hi - lo), g(0), lo=D(0), hi=D(1))
        return lo + x * (hi - lo)
    if name in ('GAMMA.DIST', 'GAMMADIST'):
        al, be = g(1), g(2)
        return gamma_pq(al, g(0) / be)[0] if g(3) != 0 else gamma_pdf(al, g(0) / be) / be
    if name in ('GAMMA.INV', 'GAMMAINV'):
        al, be = g(1), g(2)
        return be * newton(lambda x: gamma_pq(al, x)[0], lambda x: gamma_pdf(al, x), start / be, g(0), lo=D(0))
    if name in ('GAMMALN', 'GAMMALN.PRECISE'):
        return lngamma(g(0))
    if name == 'GAMMA':
        x = g(0)
        if x > 0:
            return lngamma(x).exp()
        # reflection: Gamma(x) = pi / (sin(pi x) Gamma(1 - x))
        with localcontext() as ctx:
            ctx.prec = 60
            r = x - D(math.floor(x))
            s = dsin(PI * r) * (-1 if math.floor(x) % 2 else 1)
            return PI / (s * lngamma(1 - x).exp())
    if name in ('BINOM.DIST', 'BINOMDIST'):
        s, n, p = trunc(g(0)), trunc(g(1)), g(2)
        if n > 20000:
            raise Skip
        if g(3) != 0:
            return sum((binom_pmf(D(k), n, p) for k in range(int(s) + 1)), D(0))
        return binom_pmf(s, n, p)
    if name in ('POISSON', 'POISSON.DIST'):
        k, m = trunc(g(0)), g(1)
        if m == 0:
            return D(1) if g(2) != 0 or k == 0 else D(0)
        if g(2) != 0:
            return gamma_pq(k + 1, m)[1]
        return (k * m.ln() - m - lngamma(k + 1)).exp()
    if name in ('EXPON.DIST', 'EXPONDIST'):
        x, l = g(0), g(1)
        return 1 - (-l * x).exp() if g(2) != 0 else l * (-l * x).exp()
    if name in ('WEIBULL', 'WEIBULL.DIST'):
        x, al, be = g(0), g(1), g(2)
        if x == 0:
            raise Skip
        pw = ((x / be).ln() * al).exp()
        if g(3) != 0:
            return 1 - (-pw).exp()
        return al / be * ((x / be).ln() * (al - 1)).exp() * (-pw).exp()
    if name in ('LOGNORM.DIST', 'LOGNORMDIST'):
        x, mu, sd = g(0), g(1), g(2)
        z = (x.ln() - mu) / sd
        if name == 'LOGNORMDIST' or g(3) != 0:
            return norm_cdf(z)
        return (-z * z / 2).exp() / ((2 * PI).sqrt() * sd * x)
    if name in ('LOGNORM.INV', 'LOGINV'):
        p, mu, sd = g(0), g(1), g(2)
        z0 = (D(start).ln() - mu) / sd
        z = newton(norm_cdf, lambda z: (-z * z / 2).exp() / (2 * PI).sqrt(), z0, p)
        return (mu + sd * z).exp()
    if name in ('HYPGEOM.DIST', 'HYPGEOMDIST'):
        x, n, r, tot = (trunc(v) for v in a[:4])
        def pmf(k):
            return (lnchoose(r, k) + lnchoose(tot - r, n - k) - lnchoose(tot, n)).exp()
        if name == 'HYPGEOM.DIST' and a[4] != 0:
            low = max(0, int(n - (tot - r)))
            return sum((pmf(D(k)) for k in range(low, int(x) + 1)), D(0))
        return pmf(x)
    if name in ('NEGBINOM.DIST', 'NEGBINOMDIST'):
        f, s, p = trunc(g(0)), trunc(g(1)), g(2)
        def pmf(k):
            return (lnchoose(k + s - 1, k) + s * p.ln() + k * (1 - p).ln()).exp()
        if name == 'NEGBINOM.DIST' and g(3) != 0:
            if f > 20000:
                raise Skip
            return sum((pmf(D(k)) for k in range(int(f) + 1)), D(0))
        return pmf(f)
    raise Skip


def lnchoose(n, k):
    return lngamma(n + 1) - lngamma(k + 1) - lngamma(n - k + 1)


def dsin(x):
    """sin by its series (x reduced to [-pi, pi])."""
    x = x % (2 * PI)
    if x > PI:
        x -= 2 * PI
    s = t = x
    k = 1
    while abs(t) > D('1e-55'):
        t = -t * x * x / ((2 * k) * (2 * k + 1))
        s += t
        k += 1
    return s


NUMBER = re.compile(r'^-?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$')


def parse(formula):
    m = re.match(r'^=([A-Z][A-Z0-9.]*)\((.*)\)$', formula)
    if not m:
        return None
    args = m.group(2).split(',')
    values = []
    for t in args:
        if t in ('TRUE', 'FALSE'):
            values.append(D(1) if t == 'TRUE' else D(0))
        elif NUMBER.match(t):
            values.append(D(float(t)))
        else:
            return None
    return m.group(1), values


def main():
    cells_dir, out_path = sys.argv[1], sys.argv[2]
    excel_path = sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.path.dirname(__file__), 'statdist-excel-results.tsv')
    excel = {}
    for line in open(excel_path, encoding='utf-8'):
        p = line.rstrip('\n').split('\t')
        if len(p) >= 4:
            excel[(p[0], p[1])] = p
    out = open(out_path, 'w', encoding='utf-8', newline='\n')
    files = sorted(f for f in os.listdir(cells_dir) if f.startswith('statdist-cells-'))
    done = 0
    for name in files:
        for line in open(os.path.join(cells_dir, name), encoding='utf-8'):
            sheet, addr, formula = line.rstrip('\n').split('\t')
            if not formula.startswith('='):
                continue
            parsed = parse(formula)
            r = excel.get((sheet, addr))
            if parsed is None or r is None or r[2] != 'n':
                continue
            fn, args = parsed
            start = float(r[3])
            try:
                v = dist(fn, args, start)
            except (Skip, ArithmeticError, ValueError, IndexError):
                continue
            out.write(f'{sheet}\t{addr}\t{float(v)!r}\t{v:.20e}\n')
            done += 1
    print(f'{done} reference values')


if __name__ == '__main__':
    main()
