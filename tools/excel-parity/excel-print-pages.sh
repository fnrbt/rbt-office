#!/usr/bin/env bash
# usage: excel-print-pages.sh BOOK.xlsx OUTDIR
# Runs probe-print-pages.vbs on the Windows VM (as run-excel-vbs-on-vm.sh does) and copies back OUTDIR/<book>.tsv
# (every visible sheet's page setup and page breaks as Excel 2019 has them) and OUTDIR/<book>-<n>.pdf (the sheet's
# own PDF export). The probe quits its own Excel; no other Excel on the VM is touched (-KeepExcel). WORD_VM / WORD_VM_DIR as for run-excel-vbs-on-vm.sh; TIMEOUT (seconds, default 600).
set -euo pipefail
script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
book=$1
outdir=$2
vm=${WORD_VM:-winvm}
vm_dir=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}
remote="${vm_dir}/pages-$(date -u +%Y%m%dT%H%M%SZ)-$$"
name=$(basename "$book")
stem=${name%.*}
mkdir -p "$outdir"
ssh "$vm" "powershell -NoProfile -Command \"New-Item -ItemType Directory -Force -Path '$remote' | Out-Null\"" 2>/dev/null
scp -q "$script_dir/probe-print-pages.vbs" "$script_dir/run-excel-vbs-task.ps1" "$book" "${vm}:${remote}/" 2>/dev/null
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote}/run-excel-vbs-task.ps1\" -ScriptPath \"${remote}/probe-print-pages.vbs\" -InputPath \"${remote}/${name}\" -OutputPath \"${remote}/${stem}.tsv\" -TimeoutSeconds ${TIMEOUT:-600} -KeepExcel" 2>/dev/null >/dev/null
scp -q "${vm}:${remote}/${stem}.tsv" "${vm}:${remote}/${stem}-*.pdf" "$outdir/" 2>/dev/null || true
if [[ -f "$outdir/$stem.tsv" ]] && file -b --mime-encoding "$outdir/$stem.tsv" | grep -q utf-16; then
  iconv -f UTF-16 -t UTF-8 "$outdir/$stem.tsv" | tr -d '\r' > "$outdir/$stem.tsv.tmp" && mv "$outdir/$stem.tsv.tmp" "$outdir/$stem.tsv"
fi
ls "$outdir/$stem".tsv
