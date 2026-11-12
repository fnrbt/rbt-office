#!/usr/bin/env bash
# The writers change XML that was read through XmlPatch alone, and the readers read XML through the one
# scanner, XmlDsl (its lexer, `tryScan`, the outline and element-by-element reads) (docs/save-architecture.md):
# no searching, cutting or matching of XML text. This counts the lines of each writer and reader that do text
# operations (Substring, IndexOf, LastIndexOf, Regex, StartsWith, EndsWith, Split) and fails when a count is
# above the one recorded below - the counts may only go down. Values that are text but not XML (VML styles,
# field codes, colours, cell references, paths) are read by ValueText, ModelText, CellRef, ColorText. Lower a
# count here when a change lowers it.
set -u
cd "$(dirname "$0")/.."
declare -A LIMIT=(
    [src/Office.Formats/src/WordPackageWriterModel.fs]=0
    [src/Office.Formats/src/WordPackageWriterSplice.fs]=0
    [src/Office.Formats/src/WordPackageWriterAssembly.fs]=0
    [src/Office.Formats/src/XlsxWriter.fs]=0
    [src/Office.Formats/src/DocxReader.fs]=0
    [src/Office.Formats/src/XlsxReader.fs]=0
)
pattern='\.Substring\(|\.IndexOf\(|\.LastIndexOf\(|Regex|\.StartsWith|\.EndsWith|\.Split\('
failed=0
for file in "${!LIMIT[@]}"; do
    count=$(grep -cE "$pattern" "$file")
    limit=${LIMIT[$file]}
    if [ "$count" -gt "$limit" ]; then
        echo "FAIL $file: $count lines with text operations on XML, at most $limit (use XmlPatch / XmlDsl)"
        failed=1
    elif [ "$count" -lt "$limit" ]; then
        echo "ok   $file: $count (limit $limit: lower it in tests/xml-text-ratchet.sh)"
    else
        echo "ok   $file: $count"
    fi
done
if [ $failed -ne 0 ]; then echo "XML TEXT RATCHET: FAILED"; exit 1; fi
echo "XML TEXT RATCHET: PASSED"
