#!/usr/bin/env bash
# Render a set of repro fixtures three ways (ours + docx-editor) and build
# page-aligned comparisons. Resilient: a failed fixture is logged and skipped.
# Usage: tools/docx-editor-compare/run-batch.sh [name ...]   (defaults to the set below)
set -uo pipefail
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
office_root="$(cd "$here/../.." && pwd)"
cd "$office_root"
out="artifacts/compare"; mkdir -p "$out/norm"

names=("$@")
if [ ${#names[@]} -eq 0 ]; then
  names=(
    float-image-wrap-square float-image-wrap-top-bottom float-image-paragraph-posv
    two-column-section
    table-gridspan-bottom-fit table-grid-style-borders table-vmerge-valign-rowspan
    nested-table-cell-layout table-cell-auto-pattern-shading
    endnote-placement running-header-body-offset
    heading-cambria-font-wrap widow-control-last-line
  )
fi

for n in "${names[@]}"; do
  src="tests/repros/word/${n}.docx"
  if [ ! -f "$src" ]; then echo "SKIP $n (no docx)"; continue; fi
  echo "=== $n ==="
  if timeout 130 node "$here/capture-ours.mjs" "$src" "$out/${n}-ours.png" >/dev/null 2>"$out/${n}-ours.err"; then
    echo "  ours ok"
  else echo "  ours FAILED ($(tail -1 "$out/${n}-ours.err" 2>/dev/null))"; continue; fi
  if timeout 130 node "$here/capture-docx-editor.mjs" "$src" "$out/${n}-docxeditor.png" >/dev/null 2>"$out/${n}-de.err"; then
    echo "  docx-editor ok"
  else
    # retry single-page for multi-page docs that bust the screenshot cap
    if timeout 130 node "$here/capture-docx-editor.mjs" "$src" "$out/${n}-docxeditor.png" --page 0 >/dev/null 2>"$out/${n}-de.err"; then
      echo "  docx-editor ok (page 0)"
    else echo "  docx-editor FAILED ($(tail -1 "$out/${n}-de.err" 2>/dev/null))"; continue; fi
  fi
  if "$here/align-3way.sh" "$n" >/dev/null 2>&1; then echo "  aligned -> $out/norm/${n}-aligned.png"; else echo "  align FAILED"; fi
done
echo "=== batch done ==="
