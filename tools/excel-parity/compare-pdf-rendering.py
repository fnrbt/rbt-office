#!/usr/bin/env python3
"""Compare two PDFs page by page as rasterized pixels.

A source-preserving PDF text fix changes the ToUnicode map (what pdftotext reads
back), not the glyphs a page paints.  This renders both PDFs with Poppler's
pdftoppm into the same temporary directory and compares each page's PPM bytes:
identical output means every pixel is unchanged.  The page count must match.

    python3 tools/excel-parity/compare-pdf-rendering.py BASELINE.pdf CANDIDATE.pdf [--dpi 150]

Exit status 0 only when every page is pixel-identical.
"""
import argparse
import hashlib
import os
import subprocess
import sys
import tempfile


def render(pdf, prefix, dpi):
    subprocess.run(["pdftoppm", "-r", str(dpi), pdf, prefix], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    return sorted(
        os.path.join(root, name)
        for root, _, names in os.walk(os.path.dirname(prefix))
        for name in names
        if name.startswith(os.path.basename(prefix) + "-") and name.endswith(".ppm"))


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("baseline")
    parser.add_argument("candidate")
    parser.add_argument("--dpi", type=int, default=150)
    args = parser.parse_args()

    with tempfile.TemporaryDirectory(prefix="pdf-render-compare-") as scratch:
        left = render(args.baseline, os.path.join(scratch, "baseline"), args.dpi)
        right = render(args.candidate, os.path.join(scratch, "candidate"), args.dpi)
        if len(left) != len(right):
            print(f"page count differs: baseline {len(left)}, candidate {len(right)}", file=sys.stderr)
            return 1
        differing = 0
        for index, (a, b) in enumerate(zip(left, right), start=1):
            with open(a, "rb") as f:
                data_a = f.read()
            with open(b, "rb") as f:
                data_b = f.read()
            hash_a = hashlib.sha256(data_a).hexdigest()
            hash_b = hashlib.sha256(data_b).hexdigest()
            same = data_a == data_b
            print(f"page {index}: {'identical' if same else 'DIFFERS'}  {hash_a[:16]}  {hash_b[:16]}")
            if not same:
                differing += 1
        print(f"{len(left) - differing}/{len(left)} pages pixel-identical at {args.dpi} dpi")
        return 0 if differing == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
