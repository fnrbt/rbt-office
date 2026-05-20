#!/usr/bin/env python3
"""Visual guard for table-cell-auto-pattern-shading.docx.

The text layout guard for this repro cannot see the remaining pattern-density
drift. This samples interior cell colors from 144dpi Word/ours rasters and
checks that the diagonal stripe cell has Word-like average coverage while the
percentage fills remain close.
"""

import argparse
import sys
from pathlib import Path

try:
    from PIL import Image, ImageStat
except ImportError:
    sys.exit("assert-table-auto-pattern-shading-visual: pillow not installed")


REGIONS = {
    "pct5": (170, 225, 400, 35),
    "pct10": (640, 225, 400, 35),
    "stripe": (675, 315, 350, 25),
}


def mean_rgb(path: Path, region):
    image = Image.open(path).convert("RGB")
    x, y, width, height = region
    return tuple(ImageStat.Stat(image.crop((x, y, x + width, y + height))).mean)


def max_delta(left, right):
    return max(abs(a - b) for a, b in zip(left, right))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("word_png", type=Path)
    parser.add_argument("ours_png", type=Path)
    parser.add_argument("--tolerance", type=float, default=8.0)
    args = parser.parse_args()

    failures = []
    for name, region in REGIONS.items():
        word = mean_rgb(args.word_png, region)
        ours = mean_rgb(args.ours_png, region)
        delta = max_delta(word, ours)
        print(
            f"{name}: word=({word[0]:.1f},{word[1]:.1f},{word[2]:.1f}) "
            f"ours=({ours[0]:.1f},{ours[1]:.1f},{ours[2]:.1f}) "
            f"max-delta={delta:.1f}"
        )
        if delta > args.tolerance:
            failures.append(f"{name} max channel delta {delta:.1f} > {args.tolerance:.1f}")

    if failures:
        print("table auto pattern shading visual assertion failed:", "; ".join(failures), file=sys.stderr)
        return 1

    print("table auto pattern shading visual assertion ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
