# usage: python3 -I tokdiff.py a.zip b.zip part [max]  -- XML tokens (tags / text) of the part that are only in one of the two
import sys, zipfile, re, collections
a, b, part = sys.argv[1:4]; mx = int(sys.argv[4]) if len(sys.argv) > 4 else 12
def toks(z, p):
    s = zipfile.ZipFile(z).read(p).decode("utf8", "replace")
    return re.findall(r"<[^>]*>|[^<]+", s)
ta, tb = toks(a, part), toks(b, part)
ca, cb = collections.Counter(ta), collections.Counter(tb)
gone = list((ca - cb).elements()); new = list((cb - ca).elements())
print(f"{part}: {len(ta)} -> {len(tb)} tokens; {len(gone)} only in the original, {len(new)} only in the saved")
def show(label, l):
    c = collections.Counter(re.sub(r'"[^"]*"', '"…"', t)[:70] for t in l)
    print(f"  {label} by shape:", "; ".join(f"{n}x {k}" for k, n in c.most_common(mx)))
    for t in l[:mx]: print("     ", t[:230])
show("gone", gone); show("new", new)
