#!/usr/bin/env bash
# eval-file-on-vm.sh BOOK.xlsx OUT.txt [full|none|calc]: opens BOOK.xlsx in Excel on the Windows VM and
# writes every formula cell's value (and the circular references Excel reports, and its iteration
# setting) to OUT.txt (UTF-8): see eval-file.vbs. "full" (the default) recalculates fully first.
set -euo pipefail
if [[ $# -lt 2 ]]; then echo "usage: eval-file-on-vm.sh BOOK.xlsx OUT.txt [full|none|calc]" >&2; exit 2; fi
here=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
vm=${WORD_VM:-winvm}
vm_dir=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}
mode=${3:-full}
run_id=file-$(date -u +%Y%m%dT%H%M%SZ)-$$
remote="${vm_dir}/${run_id}"
work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT
name=$(basename -- "$1")
ssh "$vm" "powershell -NoProfile -Command \"New-Item -ItemType Directory -Force -Path '$remote' | Out-Null\"" 2>/dev/null
scp -q "$here/eval-file.vbs" "$here/run-vbs-in-session.ps1" "$1" "${vm}:${remote}/" 2>/dev/null
win=${remote//\//\\}
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote}/run-vbs-in-session.ps1\" -ScriptPath \"${win}\\eval-file.vbs\" -Arguments \"${win}\\${name} ${win}\\out.tsv ${mode}\" -TimeoutSeconds 600" 2>/dev/null
scp -q "${vm}:${remote}/out.tsv" "$work/out.tsv" 2>/dev/null
iconv -f UTF-16 -t UTF-8 "$work/out.tsv" | tr -d '\r' > "$2"
