#!/usr/bin/env bash
# Build a page-aligned 3-way comparison (MS Word | Ours | docx-editor) for a
# fixture. Unlike a naive montage, each panel is cropped to its PAGE SHEET and
# scaled to the SAME pixels-per-inch (same zoom) and top-left page-aligned (same
# pan), so a given document position lands at the same pixel in all three. A
# 1-inch blue grid is overlaid for cross-panel vertical comparison.
#
# Inputs (in --dir, default artifacts/compare):
#   <name>-ours.png         (from capture-ours.mjs)
#   <name>-docxeditor.png   (from capture-docx-editor.mjs)
#   docs/images/word/<name>-word.png   (committed MS Word reference)
# Output: <dir>/norm/<name>-aligned.png
#
# Usage: tools/docx-editor-compare/align-3way.sh <name> [--dir <dir>]
set -euo pipefail

name="${1:?usage: align-3way.sh <name> [--dir <dir>]}"; shift || true
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
office_root="$(cd "$here/../.." && pwd)"
dir="$office_root/artifacts/compare"
[ "${1:-}" = "--dir" ] && { dir="$2"; shift 2; }

word="$office_root/docs/images/word/${name}-word.png"
ours="$dir/${name}-ours.png"
de="$dir/${name}-docxeditor.png"
out_dir="$dir/norm"; mkdir -p "$out_dir"
for f in "$word" "$ours" "$de"; do [ -f "$f" ] || { echo "missing: $f" >&2; exit 1; }; done

PW=820; PH=1061   # US Letter (8.5x11) at ~96.5 dpi
GRID=""; for y in $(seq 96 96 $PH); do GRID="$GRID line 0,$y $PW,$y"; done

# Crop each render to its white page sheet: Word PNG already IS the page; ours is
# a white page on the #eaecf0 canvas; docx-editor is a white page on #f8f9fa.
cp "$word" "$out_dir/${name}-word-page.png"
convert "$ours" -bordercolor '#eaecf0' -border 2 -fuzz 4% -trim +repage "$out_dir/${name}-ours-page.png"
convert "$de"   -bordercolor '#f8f9fa' -border 2 -fuzz 3% -trim +repage "$out_dir/${name}-docxeditor-page.png"

for s in word:"MS Word" ours:"Ours" docxeditor:"docx-editor"; do
  key="${s%%:*}"; lbl="${s##*:}"
  convert "$out_dir/${name}-${key}-page.png" -resize ${PW}x${PH}! \
    -stroke '#33bbff' -strokewidth 1 -fill none -draw "$GRID" \
    -bordercolor '#888' -border 1 -background '#222' -fill white -gravity North -splice 0x30 -annotate +0+6 "$lbl" \
    "$out_dir/_${name}_${key}.png"
done
montage "$out_dir/_${name}_word.png" "$out_dir/_${name}_ours.png" "$out_dir/_${name}_docxeditor.png" \
  -tile 3x1 -geometry +8+0 -background '#bbb' "$out_dir/${name}-aligned.png"
rm -f "$out_dir/_${name}_"*.png "$out_dir/${name}"-*-page.png
echo "built $out_dir/${name}-aligned.png"
