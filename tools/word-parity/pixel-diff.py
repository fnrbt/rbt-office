#!/usr/bin/env python3
"""Compare two rendered PNGs pixel-by-pixel.

Use when you need a number instead of an opinion. Designed for "is our
renderer close to Word's reference?" — emit total pixels, how many differ
beyond a per-channel threshold, and a percentage. Optionally write a diff
image with the differing pixels highlighted.

Usage:
  pixel-diff.py REF.png OUR.png
  pixel-diff.py REF.png OUR.png --threshold 24 --diff-out diff.png
  pixel-diff.py REF.png OUR.png --region 100,150,800,500
  pixel-diff.py REF.png OUR.png --resize-to-ref     # resize OUR to REF's size first
"""

import argparse
import sys
from pathlib import Path

try:
    from PIL import Image, ImageChops
except ImportError:
    sys.exit("pixel-diff: pillow not installed (apt: python3-pil)")


def parse_region(text: str) -> tuple[int, int, int, int]:
    parts = [p.strip() for p in text.split(",")]
    if len(parts) != 4:
        raise argparse.ArgumentTypeError("region must be 'x,y,w,h'")
    x, y, w, h = (int(p) for p in parts)
    return x, y, w, h


def load_and_prepare(path: Path, region, target_size):
    img = Image.open(path).convert("RGB")
    if target_size is not None and img.size != target_size:
        img = img.resize(target_size, Image.LANCZOS)
    if region is not None:
        x, y, w, h = region
        img = img.crop((x, y, x + w, y + h))
    return img


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("ref", type=Path, help="Reference PNG (the truth, e.g. Word render)")
    p.add_argument("ours", type=Path, help="Our PNG to compare")
    p.add_argument("--threshold", type=int, default=16,
                   help="Max per-channel diff to ignore. Default 16 (~6%% of 255).")
    p.add_argument("--region", type=parse_region, default=None,
                   help="Compare only this region, as 'x,y,w,h' in REF pixels.")
    p.add_argument("--resize-to-ref", action="store_true",
                   help="Resize OURS to REF's dimensions before comparing.")
    p.add_argument("--diff-out", type=Path, default=None,
                   help="Write a diff PNG: red = differs above threshold, faded = original.")
    args = p.parse_args()

    if not args.ref.exists():
        sys.exit(f"pixel-diff: ref not found: {args.ref}")
    if not args.ours.exists():
        sys.exit(f"pixel-diff: ours not found: {args.ours}")

    ref = Image.open(args.ref).convert("RGB")
    ours = Image.open(args.ours).convert("RGB")
    target = ref.size if args.resize_to_ref else None
    ref_img = load_and_prepare(args.ref, args.region, target)
    our_img = load_and_prepare(args.ours, args.region, target)

    if ref_img.size != our_img.size:
        print(f"SIZE MISMATCH: ref={ref_img.size} ours={our_img.size}")
        print("  (use --resize-to-ref to scale OURS to match)")
        return 2

    w, h = ref_img.size
    total = w * h

    diff = ImageChops.difference(ref_img, our_img)
    # For each pixel, count if max(R, G, B) diff > threshold.
    pixels = diff.load()
    bad = 0
    max_channel_diff = 0
    sum_channel_diff = 0
    for y in range(h):
        for x in range(w):
            r, g, b = pixels[x, y]
            m = max(r, g, b)
            if m > max_channel_diff:
                max_channel_diff = m
            sum_channel_diff += m
            if m > args.threshold:
                bad += 1
    pct = 100.0 * bad / total
    mean_channel_diff = sum_channel_diff / total

    print(f"size:               {w}x{h} ({total} pixels)")
    print(f"threshold:          {args.threshold} per channel")
    print(f"pixels differing:   {bad} ({pct:.2f}%)")
    print(f"max channel diff:   {max_channel_diff}")
    print(f"mean channel diff:  {mean_channel_diff:.2f}")

    if args.diff_out is not None:
        # Faded reference with red overlay where pixels differ.
        out = Image.new("RGB", (w, h), (255, 255, 255))
        ref_load = ref_img.load()
        out_load = out.load()
        for y in range(h):
            for x in range(w):
                r, g, b = ref_load[x, y]
                # fade reference toward white
                rr = int(r * 0.3 + 255 * 0.7)
                gg = int(g * 0.3 + 255 * 0.7)
                bb = int(b * 0.3 + 255 * 0.7)
                d = pixels[x, y]
                if max(d) > args.threshold:
                    out_load[x, y] = (255, 0, 0)
                else:
                    out_load[x, y] = (rr, gg, bb)
        args.diff_out.parent.mkdir(parents=True, exist_ok=True)
        out.save(args.diff_out)
        print(f"diff written to:    {args.diff_out}")

    # Exit code reflects severity bands so it's CI-usable.
    if pct >= 5.0:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
