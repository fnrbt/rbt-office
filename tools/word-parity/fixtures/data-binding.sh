#!/usr/bin/env bash
# Asks Word (on $WORD_VM, default winvm) what it shows and saves for the documents of a folder:
#   tools/word-parity/fixtures/data-binding.sh <folder>    -> <folder>/word/<name>.tsv, .saved.docx, .edited.docx
# (tools/word-parity/content-controls.vbs; see fixtures/data-binding.py.)
set -euo pipefail
folder=$(realpath "$1")
here=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)
vm=${WORD_VM:-winvm}
remote=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}/content-controls
ssh "$vm" "powershell -NoProfile -Command \"Remove-Item -Recurse -Force -ErrorAction SilentlyContinue '$remote'; New-Item -ItemType Directory -Force -Path '$remote/docs' | Out-Null\""
scp -q "$here/content-controls.vbs" "$vm:$remote/"
scp -q "$folder"/*.docx "$vm:$remote/docs/"
ssh "$vm" "cscript //nologo $remote/content-controls.vbs $remote/docs $remote/out"
mkdir -p "$folder/word"
scp -q "$vm:$remote/out/*" "$folder/word/"
