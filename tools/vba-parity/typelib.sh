#!/usr/bin/env bash
# tools/vba-parity/typelib.sh: Excel's and Word's type libraries from the Windows VM (winvm), as the F#
# modules the macro hosts compile against (early binding: Rbt.Vb's TypeCheck) and take constants from:
#   src/Spreadsheet/ExcelTypeLibrary.fs   src/Word/WordTypeLibrary.fs
# The TSV format is tools/vba-parity/typelib.ps1's.
set -euo pipefail
vm=${WORD_VM:-winvm}
remote=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}/typelib
office='C:\Program Files\Microsoft Office\Root\Office16'
here=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
root=$(cd -- "$here/../.." && pwd)
work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT
ssh "$vm" "powershell -NoProfile -Command \"New-Item -ItemType Directory -Force -Path '$remote' | Out-Null\""
scp -q "$here/typelib.ps1" "$vm:$remote/"
generate() { # name file module target
    ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File $remote/typelib.ps1 -Path \"$office\\$2\" -OutputPath $remote/$1.tsv" >/dev/null
    scp -q "$vm:$remote/$1.tsv" "$work/$1.tsv"
    if grep -q '"""' "$work/$1.tsv"; then echo "$1.tsv has a triple quote" >&2; exit 1; fi
    {
        echo "/// $5's type library (made by tools/vba-parity/typelib.sh from $2 on the Windows VM; do not edit):"
        echo "/// the interfaces, their members, classes and enums, as Rbt.Vb's TypeCheck reads them. The macro host"
        echo "/// compiles against it (early binding, as VBA's compiler) and takes its constants from it."
        echo "module $3"
        echo
        echo "let private text = \"\"\""
        cat "$work/$1.tsv"
        echo "\"\"\""
        echo
        echo "/// Read once, when a project is first compiled."
        echo "let library = lazy (Rbt.Vb.TypeCheck.Library.parse text)"
    } > "$root/$4"
    echo "$4: $(grep -c '^member' "$work/$1.tsv") members, $(grep -c '^enum' "$work/$1.tsv") constants"
}
generate excel EXCEL.EXE Rbt.Spreadsheet.ExcelTypeLibrary src/Spreadsheet/ExcelTypeLibrary.fs Excel
generate word MSWORD.OLB Rbt.Word.WordTypeLibrary src/Word/WordTypeLibrary.fs Word
