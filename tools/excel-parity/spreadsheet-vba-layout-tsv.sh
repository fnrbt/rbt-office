#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 2 ]]; then
  cat >&2 <<'USAGE'
usage: spreadsheet-vba-layout-tsv.sh input.xlsx output.tsv [sheetIndex]

Dump our spreadsheet engine's per-cell layout (address, geometry, format,
value, display text) as a TSV in the same schema as
`extract-excel-layout.vbs`. Pair with that script's Excel-side TSV and feed
both to `compare-vba-layout.py` (or read both manually) to diff our
rendering against Excel COM's reported cell values.
USAGE
  exit 2
fi

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
repo=$(cd -- "$script_dir/../.." && pwd)

dotnet run --project "$repo/tools/spreadsheet-vba-layout" -- "$@"
