#!/usr/bin/env bash
set -euo pipefail

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
exec "$script_dir/run-excel-vbs-on-vm.sh" "$script_dir/print-excel-pdf.vbs" "$@"
