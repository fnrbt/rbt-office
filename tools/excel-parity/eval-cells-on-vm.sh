#!/usr/bin/env bash
# eval-cells-on-vm.sh CELLS.txt OUT.txt: types the cells of CELLS.txt (UTF-8; sheet, address, input,
# tab-separated: see eval-cells.vbs) into a new workbook in Excel on the Windows VM, recalculates it and
# writes what Excel computes to OUT.txt (UTF-8). Excel runs in the desktop session; another Excel
# there is left alone (run-vbs-in-session.ps1).
set -euo pipefail
if [[ $# -ne 2 ]]; then echo "usage: eval-cells-on-vm.sh CELLS.txt OUT.txt" >&2; exit 2; fi
here=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
vm=${WORD_VM:-winvm}
vm_dir=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}
run_id=cells-$(date -u +%Y%m%dT%H%M%SZ)-$$
remote="${vm_dir}/${run_id}"
work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT
{ printf '\xff\xfe'; iconv -f UTF-8 -t UTF-16LE "$1"; } > "$work/cells.tsv"
ssh "$vm" "powershell -NoProfile -Command \"New-Item -ItemType Directory -Force -Path '$remote' | Out-Null\"" 2>/dev/null
scp -q "$here/eval-cells.vbs" "$here/run-vbs-in-session.ps1" "$work/cells.tsv" "${vm}:${remote}/" 2>/dev/null
win=${remote//\//\\}
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote}/run-vbs-in-session.ps1\" -ScriptPath \"${win}\\eval-cells.vbs\" -Arguments \"${win}\\cells.tsv ${win}\\out.tsv\" -TimeoutSeconds 300" 2>/dev/null
scp -q "${vm}:${remote}/out.tsv" "$work/out.tsv" 2>/dev/null
iconv -f UTF-16 -t UTF-8 "$work/out.tsv" | tr -d '\r' > "$2"
