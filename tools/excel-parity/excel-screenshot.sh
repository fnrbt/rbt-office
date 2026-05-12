#!/usr/bin/env bash
set -euo pipefail

# Usage: tools/excel-parity/excel-screenshot.sh <input.xlsx> <output.png>
#
# Runs `screenshot-excel.ps1` on the Word/Excel VM under an interactive
# scheduled task so Excel's main window paints with column/row headers and
# gridlines visible, then copies the PNG back.

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
exec "$script_dir/run-excel-vbs-on-vm.sh" "$script_dir/screenshot-excel.ps1" "$@"
