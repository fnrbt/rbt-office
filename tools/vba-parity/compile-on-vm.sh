#!/usr/bin/env bash
# tools/vba-parity/compile-on-vm.sh CASES OUT.tsv: compiles each case of CASES (tests/vba-parity/compile-cases)
# in Microsoft Excel on the Windows VM (compile-cases.ps1: VBE's Compile VBAProject, the compile error's
# dialog read) and copies the result back: case <TAB> ok | Excel's compile error.
set -euo pipefail
cases=$1; output=$2
vm=${WORD_VM:-winvm}
vm_dir=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}
here=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
run_id=compile-$(date -u +%Y%m%dT%H%M%SZ)-$$
remote="${vm_dir}/${run_id}"
ssh "$vm" "powershell -NoProfile -Command \"New-Item -ItemType Directory -Force -Path '$remote/cases' | Out-Null\""
scp -q "$here/compile-cases.ps1" "$here/../excel-parity/run-excel-vbs-task.ps1" "${vm}:${remote}/"
scp -q "$cases"/* "${vm}:${remote}/cases/"
# (A fresh folder and output each run: the task runner takes an output that is there as done.)
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote}/run-excel-vbs-task.ps1\" -Engine powershell -ScriptPath \"${remote}/compile-cases.ps1\" -InputPath \"${remote}/cases\" -OutputPath \"${remote}/compile.tsv\" -TimeoutSeconds ${VBA_PARITY_TIMEOUT:-900}" >/dev/null
mkdir -p "$(dirname "$output")"
scp -q "${vm}:${remote}/compile.tsv" "$output"
echo "$output"
