#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 3 ]]; then
  cat >&2 <<'USAGE'
usage: run-excel-vbs-on-vm.sh script.vbs input.xlsx output-file [extra cscript args...]

Copies the VBS script and XLSX to the Windows VM (the same machine the Word
parity tooling uses), runs the script with cscript, and copies the output
file back to the requested local path.

Defaults:
  WORD_VM=winvm
  WORD_VM_DIR='C:/Users/olivia/RbtExcelParity'
  WORD_VM_RUN_MODE=task
USAGE
  exit 2
fi

script_path=$1
input_xlsx=$2
local_output=$3
shift 3
script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)

if [[ ! -f "$script_path" ]]; then
  echo "script not found: $script_path" >&2
  exit 2
fi

if [[ ! -f "$input_xlsx" ]]; then
  echo "input XLSX not found: $input_xlsx" >&2
  exit 2
fi

vm=${WORD_VM:-winvm}
vm_dir=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}
run_mode=${WORD_VM_RUN_MODE:-task}
engine=${WORD_VM_ENGINE:-}
if [[ -z "$engine" ]]; then
  case "$script_path" in
    *.ps1) engine=powershell ;;
    *)     engine=cscript ;;
  esac
fi
run_id=$(date -u +%Y%m%dT%H%M%SZ)-$$
remote_dir="${vm_dir}/${run_id}"

script_name=$(basename "$script_path")
xlsx_name=$(basename "$input_xlsx")
output_name=$(basename "$local_output")
local_output_dir=$(dirname "$local_output")

mkdir -p "$local_output_dir"

ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -Command \"New-Item -ItemType Directory -Force -Path '$remote_dir' | Out-Null\""
scp "$script_path" "${vm}:${remote_dir}/${script_name}"
scp "$input_xlsx" "${vm}:${remote_dir}/${xlsx_name}"

remote_script="${remote_dir}/${script_name}"
remote_xlsx="${remote_dir}/${xlsx_name}"
remote_output="${remote_dir}/${output_name}"

case "$run_mode" in
  direct)
    if [[ "$engine" == "powershell" ]]; then
      ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"$remote_script\" -InputPath \"$remote_xlsx\" -OutputPath \"$remote_output\""
    else
      ssh "$vm" "cscript //nologo \"$remote_script\" \"$remote_xlsx\" \"$remote_output\" $*"
    fi
    ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -Command \"if (-not (Test-Path -LiteralPath '$remote_output')) { Write-Error 'expected output was not created: $remote_output'; exit 1 }\""
    ;;
  task)
    task_runner="$script_dir/run-excel-vbs-task.ps1"
    scp "$task_runner" "${vm}:${remote_dir}/run-excel-vbs-task.ps1"
    ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote_dir}/run-excel-vbs-task.ps1\" -ScriptPath \"$remote_script\" -InputPath \"$remote_xlsx\" -OutputPath \"$remote_output\" -Engine \"$engine\""
    ;;
  *)
    echo "unknown WORD_VM_RUN_MODE: $run_mode" >&2
    exit 2
    ;;
esac

scp "${vm}:${remote_output}" "$local_output"

if command -v file >/dev/null 2>&1 && command -v iconv >/dev/null 2>&1; then
  encoding=$(file -b --mime-encoding "$local_output" || true)
  case "$encoding" in
    utf-16le|utf-16)
      tmp_output="${local_output}.utf8.tmp"
      iconv -f UTF-16 -t UTF-8 "$local_output" > "$tmp_output"
      mv "$tmp_output" "$local_output"
      ;;
  esac
fi

echo "$local_output"
