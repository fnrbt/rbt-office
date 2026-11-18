#!/usr/bin/env bash
# tools/word-parity/word-macro-on-vm.sh <input.docm> <macro> <out-dir>
#
# Runs a macro of a macro-enabled document in Microsoft Word on the Windows VM (run-word-macro.vbs, through
# an interactive-user scheduled task as run-word-vbs-on-vm.sh does) and copies back <out-dir>/<macro>.tsv
# (paragraphs, formatting, tables, counts) and <out-dir>/<macro>.pdf (the pages).
set -euo pipefail
if [[ $# -ne 3 ]]; then
  echo "usage: word-macro-on-vm.sh <input.docm> <macro> <out-dir>" >&2
  exit 2
fi
input=$1
macro=$2
out_dir=$3
script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
vm=${WORD_VM:-winvm}
vm_dir=${WORD_VM_DIR:-C:/Users/olivia/RbtWordParity}
remote_dir="${vm_dir}/macro-$(date -u +%Y%m%dT%H%M%SZ)-$$"
mkdir -p "$out_dir"
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -Command \"New-Item -ItemType Directory -Force -Path '$remote_dir' | Out-Null\""
scp -q "$script_dir/run-word-macro.vbs" "$script_dir/run-word-vbs-task.ps1" "$input" "${vm}:${remote_dir}/"
name=$(basename "$input")
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote_dir}/run-word-vbs-task.ps1\" -ScriptPath \"${remote_dir}/run-word-macro.vbs\" -InputPath \"${remote_dir}/${name}\" -OutputPath \"${remote_dir}/${macro}.tsv\""
scp -q "${vm}:${remote_dir}/${macro}.tsv" "${vm}:${remote_dir}/${macro}.pdf" "$out_dir/" || true
if [[ -f "$out_dir/$macro.tsv" ]] && file -b --mime-encoding "$out_dir/$macro.tsv" | grep -q utf-16; then
  iconv -f UTF-16 -t UTF-8 "$out_dir/$macro.tsv" > "$out_dir/$macro.tsv.tmp" && mv "$out_dir/$macro.tsv.tmp" "$out_dir/$macro.tsv"
fi
ls "$out_dir/$macro".*
