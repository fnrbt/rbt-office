#!/usr/bin/env python3
"""Compare two rendered PNGs.

Reports two numbers, both as percentages of total pixels:

  raw       — fraction of pixels whose max-channel diff > threshold.
              Picks up anti-aliasing wobble; in a page that is mostly
              white space this stays low even when the text is in the
              wrong place.
  ink-XOR   — binarize both images (max-channel < ink_max → "ink"),
              then count pixels where one image has ink and the other
              does not (relative to the *union* of ink pixels). This
              is the "is the text actually in the same place?" number
              — when line breaks differ, this jumps even when raw
              looks flat.

When the layouts agree, both numbers are small and close. When the
layouts disagree structurally, ink-XOR is much larger than raw and
the gap itself is the signal.

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


def is_ink(rgb, ink_max: int) -> bool:
    r, g, b = rgb
    # A pixel counts as "ink" if any channel is dark enough. Anti-aliased
    # edges around glyphs sit in the gray zone; ink_max of ~200 keeps the
    # text body + heavier AA pixels in, drops pure background.
    return min(r, g, b) < ink_max


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("ref", type=Path, help="Reference PNG (the truth, e.g. Word render)")
    p.add_argument("ours", type=Path, help="Our PNG to compare")
    p.add_argument("--threshold", type=int, default=16,
                   help="Raw: max per-channel diff to ignore. Default 16.")
    p.add_argument("--ink-max", type=int, default=200,
                   help="Ink: min-channel below this counts as 'ink'. Default 200.")
    p.add_argument("--region", type=parse_region, default=None,
                   help="Compare only this region, as 'x,y,w,h' in REF pixels.")
    p.add_argument("--resize-to-ref", action="store_true",
                   help="Resize OURS to REF's dimensions before comparing.")
    p.add_argument("--diff-out", type=Path, default=None,
                   help="Write a diff PNG: red = differs above threshold, faded = original.")
    p.add_argument("--ink-diff-out", type=Path, default=None,
                   help="Write an ink-XOR PNG: blue = ref-only ink, red = ours-only ink, "
                        "purple = both ink.")
    args = p.parse_args()

    if not args.ref.exists():
        sys.exit(f"pixel-diff: ref not found: {args.ref}")
    if not args.ours.exists():
        sys.exit(f"pixel-diff: ours not found: {args.ours}")

    ref = Image.open(args.ref).convert("RGB")
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
    diff_px = diff.load()
    ref_px = ref_img.load()
    our_px = our_img.load()

    raw_bad = 0
    max_channel_diff = 0
    sum_channel_diff = 0
    ref_ink = 0
    our_ink = 0
    both_ink = 0
    ref_only_ink = 0
    our_only_ink = 0
    for y in range(h):
        for x in range(w):
            dr, dg, db = diff_px[x, y]
            m = max(dr, dg, db)
            if m > max_channel_diff:
                max_channel_diff = m
            sum_channel_diff += m
            if m > args.threshold:
                raw_bad += 1
            r_ink = is_ink(ref_px[x, y], args.ink_max)
            o_ink = is_ink(our_px[x, y], args.ink_max)
            if r_ink:
                ref_ink += 1
            if o_ink:
                our_ink += 1
            if r_ink and o_ink:
                both_ink += 1
            elif r_ink:
                ref_only_ink += 1
            elif o_ink:
                our_only_ink += 1

    raw_pct = 100.0 * raw_bad / total
    mean_channel_diff = sum_channel_diff / total
    union_ink = ref_ink + our_ink - both_ink
    xor_ink = ref_only_ink + our_only_ink
    ink_xor_pct_of_union = (100.0 * xor_ink / union_ink) if union_ink else 0.0
    ink_xor_pct_of_total = 100.0 * xor_ink / total
    ref_ink_pct = 100.0 * ref_ink / total
    our_ink_pct = 100.0 * our_ink / total

    print(f"size:                  {w}x{h} ({total} pixels)")
    print(f"raw threshold:         {args.threshold} per channel")
    print(f"raw pixels differing:  {raw_bad} ({raw_pct:.2f}%)")
    print(f"raw max channel diff:  {max_channel_diff}")
    print(f"raw mean channel diff: {mean_channel_diff:.2f}")
    print(f"ink-max:               {args.ink_max}")
    print(f"ref ink coverage:      {ref_ink} ({ref_ink_pct:.2f}% of total)")
    print(f"ours ink coverage:     {our_ink} ({our_ink_pct:.2f}% of total)")
    print(f"ink XOR (one but not the other):")
    print(f"   ref-only:           {ref_only_ink}")
    print(f"   ours-only:          {our_only_ink}")
    print(f"   total:              {xor_ink}")
    print(f"   % of union ink:     {ink_xor_pct_of_union:.2f}%")
    print(f"   % of total pixels:  {ink_xor_pct_of_total:.2f}%")

    if args.diff_out is not None:
        out = Image.new("RGB", (w, h), (255, 255, 255))
        out_load = out.load()
        for y in range(h):
            for x in range(w):
                r, g, b = ref_px[x, y]
                rr = int(r * 0.3 + 255 * 0.7)
                gg = int(g * 0.3 + 255 * 0.7)
                bb = int(b * 0.3 + 255 * 0.7)
                if max(diff_px[x, y]) > args.threshold:
                    out_load[x, y] = (255, 0, 0)
                else:
                    out_load[x, y] = (rr, gg, bb)
        args.diff_out.parent.mkdir(parents=True, exist_ok=True)
        out.save(args.diff_out)
        print(f"raw diff:              {args.diff_out}")

    if args.ink_diff_out is not None:
        out = Image.new("RGB", (w, h), (255, 255, 255))
        out_load = out.load()
        for y in range(h):
            for x in range(w):
                r_ink = is_ink(ref_px[x, y], args.ink_max)
                o_ink = is_ink(our_px[x, y], args.ink_max)
                if r_ink and o_ink:
                    out_load[x, y] = (170, 50, 200)   # purple — both
                elif r_ink:
                    out_load[x, y] = (20, 70, 220)    # blue — ref only
                elif o_ink:
                    out_load[x, y] = (230, 30, 30)    # red — ours only
        args.ink_diff_out.parent.mkdir(parents=True, exist_ok=True)
        out.save(args.ink_diff_out)
        print(f"ink-XOR diff:          {args.ink_diff_out}")

    # Exit non-zero when either metric is clearly bad.
    if raw_pct >= 5.0 or ink_xor_pct_of_union >= 30.0:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
