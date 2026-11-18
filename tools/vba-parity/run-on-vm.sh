#!/usr/bin/env bash
# tools/vba-parity/run-on-vm.sh DOCS OUT.tsv: runs every case module (.bas) of DOCS (made by parity.fsx build) in
# Microsoft Excel on the Windows VM (winvm, the parity tooling's machine) and copies the result back.
set -euo pipefail
docs=$1; output=$2
vm=${WORD_VM:-winvm}
vm_dir=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}
here=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
run_id=vba-$(date -u +%Y%m%dT%H%M%SZ)-$$
remote="${vm_dir}/${run_id}"
ssh "$vm" "powershell -NoProfile -Command \"New-Item -ItemType Directory -Force -Path '$remote/docs' | Out-Null\""
scp -q "$here/run-cases.vbs" "$here/../excel-parity/run-excel-vbs-task.ps1" "${vm}:${remote}/"
scp -q "$docs"/*.bas "${vm}:${remote}/docs/"
status=0
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote}/run-excel-vbs-task.ps1\" -ScriptPath \"${remote}/run-cases.vbs\" -InputPath \"${remote}/docs\" -OutputPath \"${remote}/excel.tsv\" -TimeoutSeconds ${VBA_PARITY_TIMEOUT:-900}" || status=$?
# (Where it got to, also when it timed out.)
scp -q "${vm}:${remote}/excel.tsv.progress.txt" "$output.progress.txt" 2>/dev/null || true
if [[ $status -ne 0 ]]; then echo "the VM run failed ($status); progress: $output.progress.txt" >&2; cat "$output.progress.txt" >&2 2>/dev/null || true; exit $status; fi
mkdir -p "$(dirname "$output")"
scp -q "${vm}:${remote}/excel.tsv" "$output.utf16"
iconv -f UTF-16 -t UTF-8 "$output.utf16" | tr -d '\r' > "$output"
rm -f "$output.utf16"
echo "$output"
