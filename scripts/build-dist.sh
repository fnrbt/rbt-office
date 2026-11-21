#!/usr/bin/env bash
# bash scripts/build-dist.sh [dist]   (pnpm run build:fable)
#
# The browser dist: one Fable compile of src/Dist/Rbt.Dist.fsproj, which references every browser
# project, so dist has one compiled copy of every module and one fable-library
# (scripts/layout-dist.mjs puts them at dist's public paths: dist/word/WordHost.js,
# dist/spreadsheet/Spreadsheet.js, dist/office-formats/src/Main.js, dist/fable_modules/...).
set -euo pipefail
repository="$(cd "$(dirname "$0")/.." && pwd)"
dist="$(realpath -m "${1:-$repository/dist}")"
# (Not under obj/: Fable leaves out package sources it copies to an outDir there.)
out="$(mktemp -d "${TMPDIR:-/tmp}/rbt-office-fable.XXXXXX")"
trap 'rm -rf "$out"' EXIT
cd "$repository"

dotnet tool restore
rm -rf "$dist"
dotnet fable src/Dist --outDir "$out" --lang JavaScript --noCache
node scripts/layout-dist.mjs "$out" "$dist"
dotnet fsi scripts/stage-word-pdf-data.fsx "$dist/word-pdf"
cp -r fonts "$dist/fonts"

# What must load only on use is not in the start-up graphs: the function modules and the VBA engine;
# the XLSX package writer (ExportXlsx); in Word the DOCX reader (opening a file: WordPackageOpen) and
# writer (saving: WordPackageWriterAssembly); in either, office-formats' Main (which loads all of
# office-formats) and the other editor's file formats.
node scripts/check-function-modules.mjs "$dist/spreadsheet"
node scripts/check-start-up-modules.mjs "$dist/spreadsheet" Spreadsheet.js,Toolbar.js,XlsxImporter.js \
  'MacroProject\.js$,MacroHost\.js$,Rbt\.Vb\.,^rbt-vb/,XlsxPackageWriter\.js$,DocxReader\.js$,WordPackage,WordEditorPackage,office-formats/src/Main\.js$'
node scripts/check-start-up-modules.mjs "$dist/word" Word.js,WordShell.js,WordHost.js \
  'WordMacroProject\.js$,WordMacroHost\.js$,HyphenPatternsEnUs\.js$,HyphenWordsEnUs\.js$,Rbt\.Vb\.,^rbt-vb/,DocxReader\.js$,WordPackageOpen\.js$,WordPackageWriter,XlsxReader\.js$,XlsxPackageWriter\.js$,office-formats/src/Main\.js$'
# The start-up graphs' sizes and what they load on use, for the record (`--top N --min`: the biggest
# modules, minified and gzipped, with the import chain to each).
node scripts/start-up-report.mjs "$dist/spreadsheet" Spreadsheet.js,Toolbar.js,XlsxImporter.js --top 0
node scripts/start-up-report.mjs "$dist/word" Word.js,WordShell.js,WordHost.js --top 0
