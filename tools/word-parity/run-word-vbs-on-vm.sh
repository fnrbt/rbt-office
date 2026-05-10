#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 3 ]]; then
  cat >&2 <<'USAGE'
usage: run-word-vbs-on-vm.sh script.vbs input.docx output-file [extra cscript args...]

Copies the VBS script and DOCX to the Windows Word VM, runs the script with
cscript, and copies the output file back to the requested local path.

Defaults:
  WORD_VM=wordvm
  WORD_VM_DIR='C:/Users/worduser/RedBatWordParity'
  WORD_VM_RUN_MODE=task
USAGE
  exit 2
fi

script_path=$1
input_docx=$2
local_output=$3
shift 3
script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)

if [[ ! -f "$script_path" ]]; then
  echo "script not found: $script_path" >&2
  exit 2
fi

if [[ ! -f "$input_docx" ]]; then
  echo "input DOCX not found: $input_docx" >&2
  exit 2
fi

vm=${WORD_VM:-wordvm}
vm_dir=${WORD_VM_DIR:-C:/Users/worduser/RedBatWordParity}
run_mode=${WORD_VM_RUN_MODE:-task}
run_id=$(date -u +%Y%m%dT%H%M%SZ)-$$
remote_dir="${vm_dir}/${run_id}"

script_name=$(basename "$script_path")
docx_name=$(basename "$input_docx")
output_name=$(basename "$local_output")
local_output_dir=$(dirname "$local_output")

mkdir -p "$local_output_dir"

ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -Command \"New-Item -ItemType Directory -Force -Path '$remote_dir' | Out-Null\""
scp "$script_path" "${vm}:${remote_dir}/${script_name}"
scp "$input_docx" "${vm}:${remote_dir}/${docx_name}"

remote_script="${remote_dir}/${script_name}"
remote_docx="${remote_dir}/${docx_name}"
remote_output="${remote_dir}/${output_name}"

case "$run_mode" in
  direct)
    ssh "$vm" "cscript //nologo \"$remote_script\" \"$remote_docx\" \"$remote_output\" $*"
    ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -Command \"if (-not (Test-Path -LiteralPath '$remote_output')) { Write-Error 'expected output was not created: $remote_output'; exit 1 }\""
    ;;
  task)
    task_runner="$script_dir/run-word-vbs-task.ps1"
    scp "$task_runner" "${vm}:${remote_dir}/run-word-vbs-task.ps1"
    ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote_dir}/run-word-vbs-task.ps1\" -ScriptPath \"$remote_script\" -InputPath \"$remote_docx\" -OutputPath \"$remote_output\""
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
