#!/usr/bin/env python3
import argparse
import csv
import math
import os
from pathlib import Path
import re
import subprocess
import sys

from PIL import Image, ImageChops, ImageEnhance, ImageStat


ROOT = Path(os.environ.get("OFFICE_ROOT", Path.cwd())).resolve()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compare Microsoft Excel PDF output with RedBat spreadsheet PDF output."
    )
    parser.add_argument("xlsx", nargs="+", help="XLSX file(s) to compare")
    parser.add_argument("--out", default="artifacts/pdf-compare-excel", help="output directory")
    parser.add_argument("--dpi", type=int, default=96, help="rasterization DPI")
    parser.add_argument(
        "--crop",
        action="append",
        default=[],
        metavar="NAME:X,Y,W,H",
        help="absolute crop in raster pixels; may be repeated",
    )
    parser.add_argument(
        "--build",
        action="store_true",
        help="allow dotnet run to build the spreadsheet PDF runner",
    )
    parser.add_argument(
        "--reuse-excel",
        action="store_true",
        help="reuse an existing excel.pdf instead of calling Excel through wordvm",
    )
    parser.add_argument(
        "--max-body-mae",
        type=float,
        default=None,
        help=(
            "Fail (exit 2) if any 'body' crop's MAE exceeds this threshold."
            " Use to gate CI on Excel-vs-ours print-PDF parity."
        ),
    )
    parser.add_argument(
        "--max-full-mae",
        type=float,
        default=None,
        help="Fail (exit 2) if any 'full' crop's MAE exceeds this threshold.",
    )
    args = parser.parse_args()

    out_root = (ROOT / args.out).resolve()
    out_root.mkdir(parents=True, exist_ok=True)
    custom_crops = [parse_crop(value) for value in args.crop]
    summary_path = out_root / "summary.tsv"
    rows = []

    for input_arg in args.xlsx:
        xlsx = (ROOT / input_arg).resolve()
        if not xlsx.exists():
            raise FileNotFoundError(xlsx)
        doc_rows = compare_doc(xlsx, out_root, args.dpi, custom_crops, args)
        rows.extend(doc_rows)

    write_summary(summary_path, rows)
    print(f"summary: {summary_path}")
    print_top(rows)

    # Enforce thresholds when set. A row's `mae` is a float string; rows for
    # page-count-mismatch errors carry an empty `mae` and are skipped here
    # (they fail loudly via their `detail` column instead).
    breaches = []
    for row in rows:
        try:
            mae = float(row["mae"])
        except (TypeError, ValueError, KeyError):
            continue
        crop = row.get("crop", "")
        limit = None
        if crop == "body" and args.max_body_mae is not None:
            limit = args.max_body_mae
        elif crop == "full" and args.max_full_mae is not None:
            limit = args.max_full_mae
        if limit is not None and mae > limit:
            breaches.append((row, limit))
    if breaches:
        print(f"\nFAIL: {len(breaches)} crop(s) exceeded MAE threshold:")
        for row, limit in breaches:
            print(
                f"  {row['doc']} page={row['page']} crop={row['crop']}"
                f" mae={row['mae']} > {limit}"
            )
        return 2
    return 0


def compare_doc(xlsx: Path, out_root: Path, dpi: int, custom_crops, args):
    rel = xlsx.relative_to(ROOT) if is_relative_to(xlsx, ROOT) else xlsx
    stem = sanitize(str(rel.with_suffix("")))
    doc_out = out_root / stem
    doc_out.mkdir(parents=True, exist_ok=True)
    excel_pdf = doc_out / "excel.pdf"
    ours_pdf = doc_out / "ours.pdf"

    print(f"compare print pdf: {rel}", flush=True)
    if not (args.reuse_excel and excel_pdf.exists()):
        run(["tools/excel-parity/excel-print-pdf.sh", str(xlsx), str(excel_pdf)])

    runner = [
        "dotnet",
        "run",
        "--project",
        "tools/spreadsheet-pdf-runner/SpreadsheetPdfRunner.fsproj",
    ]
    if not args.build:
        runner.append("--no-build")
    runner.extend(["--", str(xlsx), str(ours_pdf)])
    run(runner)

    excel_pages = rasterize(excel_pdf, doc_out / "excel", dpi)
    ours_pages = rasterize(ours_pdf, doc_out / "ours", dpi)
    page_count = min(len(excel_pages), len(ours_pages))
    rows = []
    if len(excel_pages) != len(ours_pages):
        rows.append(
            {
                "doc": str(rel),
                "page": "",
                "crop": "page-count",
                "width": "",
                "height": "",
                "mae": "",
                "rms": "",
                "max": "",
                "bbox": "",
                "detail": f"excel={len(excel_pages)} ours={len(ours_pages)}",
            }
        )

    for page_index in range(page_count):
        excel_img = Image.open(excel_pages[page_index]).convert("RGB")
        ours_img = Image.open(ours_pages[page_index]).convert("RGB")
        if excel_img.size != ours_img.size:
            rows.append(
                {
                    "doc": str(rel),
                    "page": page_index + 1,
                    "crop": "dimensions",
                    "width": "",
                    "height": "",
                    "mae": "",
                    "rms": "",
                    "max": "",
                    "bbox": "",
                    "detail": f"excel={excel_img.size} ours={ours_img.size}",
                }
            )
        width = min(excel_img.width, ours_img.width)
        height = min(excel_img.height, ours_img.height)
        excel_img = excel_img.crop((0, 0, width, height))
        ours_img = ours_img.crop((0, 0, width, height))
        crops = custom_crops or default_crops(width, height)
        for crop in crops:
            row = compare_crop(str(rel), doc_out, page_index + 1, crop, excel_img, ours_img)
            rows.append(row)

    write_summary(doc_out / "summary.tsv", rows)
    return rows


def compare_crop(doc, doc_out: Path, page_num: int, crop, excel_img, ours_img):
    name, box = crop_box(crop, excel_img.width, excel_img.height)
    excel_crop = excel_img.crop(box)
    ours_crop = ours_img.crop(box)
    diff = ImageChops.difference(excel_crop, ours_crop)
    stat = ImageStat.Stat(diff)
    mae = sum(stat.mean) / len(stat.mean)
    rms = math.sqrt(sum(value * value for value in stat.rms) / len(stat.rms))
    max_delta = max((max(channel) for channel in diff.getextrema()), default=0)
    bbox = diff.getbbox()

    prefix = doc_out / f"page{page_num:02d}-{name}"
    excel_crop.save(str(prefix) + "-excel.png")
    ours_crop.save(str(prefix) + "-ours.png")
    ImageEnhance.Contrast(diff).enhance(4.0).save(str(prefix) + "-diff.png")

    return {
        "doc": doc,
        "page": page_num,
        "crop": name,
        "width": excel_crop.width,
        "height": excel_crop.height,
        "mae": f"{mae:.4f}",
        "rms": f"{rms:.4f}",
        "max": max_delta,
        "bbox": "" if bbox is None else ",".join(str(v) for v in bbox),
        "detail": "",
    }


def default_crops(width: int, height: int):
    top_h = min(height, max(1, round(height * 0.22)))
    bottom_h = top_h
    body_top = round(height * 0.18)
    body_bottom = round(height * 0.82)
    return [
        ("full", (0, 0, width, height)),
        ("top", (0, 0, width, top_h)),
        ("body", (0, body_top, width, max(body_top + 1, body_bottom))),
        ("bottom", (0, max(0, height - bottom_h), width, height)),
    ]


def parse_crop(value: str):
    match = re.fullmatch(r"([^:]+):(\d+),(\d+),(\d+),(\d+)", value)
    if not match:
        raise ValueError(f"invalid --crop {value!r}; expected NAME:X,Y,W,H")
    name = sanitize(match.group(1))
    x, y, w, h = (int(match.group(i)) for i in range(2, 6))
    return name, (x, y, x + w, y + h)


def crop_box(crop, width: int, height: int):
    name, (x1, y1, x2, y2) = crop
    x1 = clamp(x1, 0, width)
    y1 = clamp(y1, 0, height)
    x2 = clamp(x2, x1 + 1, width)
    y2 = clamp(y2, y1 + 1, height)
    return name, (x1, y1, x2, y2)


def rasterize(pdf: Path, prefix: Path, dpi: int):
    for old in prefix.parent.glob(prefix.name + "-*.png"):
        old.unlink()
    run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(prefix)])
    pages = sorted(prefix.parent.glob(prefix.name + "-*.png"), key=page_sort_key)
    if not pages:
        raise RuntimeError(f"pdftoppm produced no pages for {pdf}")
    return pages


def write_summary(path: Path, rows):
    fieldnames = ["doc", "page", "crop", "width", "height", "mae", "rms", "max", "bbox", "detail"]
    with path.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter="\t", extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def print_top(rows):
    scored = []
    for row in rows:
        try:
            scored.append((float(row["mae"]), row))
        except (ValueError, KeyError):
            pass
    for _, row in sorted(scored, key=lambda item: item[0], reverse=True)[:12]:
        print(
            f"{row['doc']} page={row['page']} crop={row['crop']} "
            f"mae={row['mae']} rms={row['rms']} bbox={row['bbox']}"
        )


def run(command):
    result = subprocess.run(command, cwd=ROOT)
    if result.returncode != 0:
        raise RuntimeError(f"command failed ({result.returncode}): {' '.join(command)}")


def sanitize(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", value)


def page_sort_key(path: Path):
    match = re.search(r"-(\d+)\.png$", path.name)
    return int(match.group(1)) if match else path.name


def clamp(value: int, lo: int, hi: int) -> int:
    return max(lo, min(value, hi))


def is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(1)
