#!/usr/bin/env python3
"""Measure simple visual geometry in Word parity render PNGs.

The goal is not pixel-perfect image diffing. This script extracts stable
landmarks such as the yellow table fill and large white cell interiors so a
layout fix can be checked against Word without eyeballing mismatched crops.
"""

from __future__ import annotations

import argparse
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from PIL import Image


@dataclass(frozen=True)
class Box:
    x0: int
    y0: int
    x1: int
    y1: int

    @property
    def width(self) -> int:
        return self.x1 - self.x0 + 1

    @property
    def height(self) -> int:
        return self.y1 - self.y0 + 1

    @property
    def area(self) -> int:
        return self.width * self.height

    def logical(self, scale: float) -> tuple[float, float, float, float]:
        return (
            self.x0 / scale,
            self.y0 / scale,
            self.width / scale,
            self.height / scale,
        )


def bbox(points: Iterable[tuple[int, int]]) -> Box | None:
    xs: list[int] = []
    ys: list[int] = []
    for x, y in points:
        xs.append(x)
        ys.append(y)
    if not xs:
        return None
    return Box(min(xs), min(ys), max(xs), max(ys))


def yellow_mask(pixels, width: int, height: int) -> list[list[bool]]:
    rows: list[list[bool]] = []
    for y in range(height):
        row: list[bool] = []
        for x in range(width):
            r, g, b = pixels[x, y][:3]
            row.append(r >= 220 and 135 <= g <= 215 and b <= 80)
        rows.append(row)
    return rows


def white_mask_in_box(pixels, box: Box) -> list[list[bool]]:
    rows: list[list[bool]] = []
    for y in range(box.y0, box.y1 + 1):
        row: list[bool] = []
        for x in range(box.x0, box.x1 + 1):
            r, g, b = pixels[x, y][:3]
            row.append(r >= 242 and g >= 242 and b >= 242)
        rows.append(row)
    return rows


def dark_mask_in_box(pixels, box: Box) -> list[list[bool]]:
    rows: list[list[bool]] = []
    for y in range(box.y0, box.y1 + 1):
        row: list[bool] = []
        for x in range(box.x0, box.x1 + 1):
            r, g, b = pixels[x, y][:3]
            row.append(r <= 90 and g <= 90 and b <= 90)
        rows.append(row)
    return rows


def component_boxes(mask: list[list[bool]], origin_x: int = 0, origin_y: int = 0) -> list[Box]:
    if not mask:
        return []
    height = len(mask)
    width = len(mask[0])
    seen = [[False for _ in range(width)] for _ in range(height)]
    components: list[Box] = []
    for sy in range(height):
        for sx in range(width):
            if seen[sy][sx] or not mask[sy][sx]:
                continue
            q: deque[tuple[int, int]] = deque([(sx, sy)])
            seen[sy][sx] = True
            points: list[tuple[int, int]] = []
            while q:
                x, y = q.popleft()
                points.append((origin_x + x, origin_y + y))
                for nx, ny in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
                    if nx < 0 or ny < 0 or nx >= width or ny >= height:
                        continue
                    if seen[ny][nx] or not mask[ny][nx]:
                        continue
                    seen[ny][nx] = True
                    q.append((nx, ny))
            box = bbox(points)
            if box is not None:
                components.append(box)
    return components


def parse_scale(value: str) -> tuple[str, float]:
    label, raw_scale = value.split("=", 1)
    return label, float(raw_scale)


def parse_image(value: str) -> tuple[str, Path]:
    if "=" in value:
        label, raw_path = value.split("=", 1)
        return label, Path(raw_path)
    path = Path(value)
    return path.stem, path


def fmt_box(box: Box, scale: float) -> str:
    x, y, width, height = box.logical(scale)
    return (
        f"native=({box.x0},{box.y0}) {box.width}x{box.height} "
        f"logical=({x:.2f},{y:.2f}) {width:.2f}x{height:.2f}"
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("images", nargs="+", help="PNG paths or label=PNG paths")
    parser.add_argument(
        "--scale",
        action="append",
        default=[],
        help="Coordinate scale as label=scale. Use 1.5 for 144dpi Word/PDF when comparing to 96dpi browser.",
    )
    parser.add_argument("--min-white-area", type=int, default=800)
    args = parser.parse_args()

    scales = dict(parse_scale(item) for item in args.scale)

    for image_arg in args.images:
        label, path = parse_image(image_arg)
        scale = scales.get(label, 1.0)
        image = Image.open(path).convert("RGB")
        pixels = image.load()
        width, height = image.size
        mask = yellow_mask(pixels, width, height)
        yellow = bbox(
            (x, y)
            for y, row in enumerate(mask)
            for x, value in enumerate(row)
            if value
        )

        print(f"{label}: {path}")
        print(f"  image: native={width}x{height} logical={width / scale:.2f}x{height / scale:.2f}")
        if yellow is None:
            print("  yellow: none")
            continue
        print(f"  yellow: {fmt_box(yellow, scale)}")

        whites = component_boxes(white_mask_in_box(pixels, yellow), yellow.x0, yellow.y0)
        whites = [
            box
            for box in whites
            if box.area >= args.min_white_area
            and box.width >= 16
            and box.height >= 8
        ]
        whites.sort(key=lambda box: (box.y0, box.x0))
        for index, box in enumerate(whites[:12], start=1):
            print(f"  white[{index}]: {fmt_box(box, scale)}")

        darks = component_boxes(dark_mask_in_box(pixels, yellow), yellow.x0, yellow.y0)
        horizontal = [
            box
            for box in darks
            if box.width >= 32 and box.height <= 4
        ]
        vertical = [
            box
            for box in darks
            if box.height >= 16 and box.width <= 4
        ]
        horizontal.sort(key=lambda box: (box.y0, box.x0))
        vertical.sort(key=lambda box: (box.x0, box.y0))
        for index, box in enumerate(horizontal[:16], start=1):
            print(f"  hline[{index}]: {fmt_box(box, scale)}")
        for index, box in enumerate(vertical[:16], start=1):
            print(f"  vline[{index}]: {fmt_box(box, scale)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
