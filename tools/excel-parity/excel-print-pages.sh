#!/usr/bin/env bash
# usage: excel-print-pages.sh BOOK.xlsx OUTDIR
# Exports each visible sheet directly to PDF before any preview or sheet activation. A reopened workbook supplies
# setup and preview-only page breaks; Page Break Preview can refit bestFit columns, so those breaks need not describe
# the PDFs. The probe quits its own Excel; no other Excel on the VM is touched.
# WORD_VM / WORD_VM_DIR as for run-excel-vbs-on-vm.sh; TIMEOUT (seconds, default 600).
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
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote}/run-excel-vbs-task.ps1\" -ScriptPath \"${remote}/probe-print-pages.vbs\" -InputPath \"${remote}/${name}\" -OutputPath \"${remote}/${stem}.tsv\" -TimeoutSeconds ${TIMEOUT:-600}" 2>/dev/null >/dev/null
scp -q "${vm}:${remote}/${stem}.tsv" "${vm}:${remote}/${stem}-*.pdf" "$outdir/" 2>/dev/null || true
if [[ -f "$outdir/$stem.tsv" ]] && file -b --mime-encoding "$outdir/$stem.tsv" | grep -q utf-16; then
  iconv -f UTF-16 -t UTF-8 "$outdir/$stem.tsv" | tr -d '\r' > "$outdir/$stem.tsv.tmp" && mv "$outdir/$stem.tsv.tmp" "$outdir/$stem.tsv"
fi
# The PDF belongs to the fresh direct-export pass, never to the preview break pass.
command -v pdfinfo >/dev/null || { echo "pdfinfo is required to record direct page counts" >&2; exit 1; }
counts=$(mktemp)
trap 'rm -f "$counts"' EXIT
while IFS=$'\t' read -r sheet index key value; do
  if [[ "$key" == pdf && "$value" == *.pdf ]]; then
    pages=$(pdfinfo "$outdir/$value" | awk '/^Pages:/ {print $2}')
    [[ -n "$pages" ]] || { echo "Cannot read page count: $value" >&2; exit 1; }
    printf '%s\t%s\tdirect.pages\t%s\n' "$sheet" "$index" "$pages"
  fi
done < "$outdir/$stem.tsv" > "$counts"
cat "$counts" >> "$outdir/$stem.tsv"
ls "$outdir/$stem".tsv
