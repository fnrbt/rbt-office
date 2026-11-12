#!/usr/bin/env bash
# Validate our vbaProject.bin reader against REAL Excel output: mint a
# macro-enabled .xlsm on the Excel VM (gen-macro-xlsm.vbs), pull it back, and
# check VbaProject.readMacroSource recovers the module — and that the recovered
# source runs on our engine.
#
# Requires the parity VM (ssh winvm). Run from anywhere.
set -euo pipefail
repo=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/../.." && pwd)
cd "$repo"

out=${1:-/tmp/real-macro.xlsm}
dummy=tests/test-documents-excel/pandas-test_decimal.xlsx   # unused input the runner requires

echo "==> minting $out on the Excel VM"
tools/excel-parity/run-excel-vbs-on-vm.sh tools/vba/gen-macro-xlsm.vbs "$dummy" "$out"

echo "==> building Office.Formats"
dotnet build src/Office.Formats/OfficeFormat.fsproj -c Debug >/dev/null
dotnet build src/Spreadsheet/RedBat.Spreadsheet.fsproj -c Debug >/dev/null

dll="$repo/src/Office.Formats/bin/Debug/netstandard2.0/OfficeFormat.dll"
rb="$repo/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"
vb="$repo/src/Spreadsheet/bin/Debug/netstandard2.0/FsVb.dll"

cat > /tmp/validate-real.fsx <<EOF
#r "$dll"
#r "$vb"
#r "$rb"
open RedBat.Spreadsheet.Model
module V = RedBat.Spreadsheet.Vba
let bytes = System.IO.File.ReadAllBytes "$out"
match VbaProject.extractFromXlsx bytes with
| None -> eprintfn "FAIL: no vbaProject.bin / no module recovered"; exit 1
| Some src ->
    printfn "recovered source from REAL Excel .xlsm:\n%s\n---" src
    if not (src.Contains "Demo") then (eprintfn "FAIL: 'Demo' sub not found"; exit 1)
    match V.runOnModel (createEmpty 10 10) (src + "\nDemo()") with
    | Result.Ok m ->
        let d r c = displayValue (getCellValue { Row=r; Col=c } m)
        printfn "ran recovered macro: A1=%s A2=%s A3=%s" (d 0 0) (d 1 0) (d 2 0)
        if d 0 0 = "42" && d 1 0 = "84" then printfn "PASS: real Excel macro read + ran"
        else (eprintfn "FAIL: unexpected values"; exit 1)
    | Result.Error e -> eprintfn "FAIL run: %s" e; exit 1
EOF
dotnet fsi /tmp/validate-real.fsx
