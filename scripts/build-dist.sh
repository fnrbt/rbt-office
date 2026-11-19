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

# What must load only on use (the function modules, the VBA engine) is not in the start-up graphs.
node scripts/check-function-modules.mjs "$dist/spreadsheet"
node scripts/check-start-up-modules.mjs "$dist/spreadsheet" Spreadsheet.js,Toolbar.js,XlsxImporter.js 'MacroProject\.js$,MacroHost\.js$,Rbt\.Vb\.,^rbt-vb/'
node scripts/check-start-up-modules.mjs "$dist/word" Word.js,WordShell.js,WordHost.js 'WordMacroProject\.js$,WordMacroHost\.js$,Rbt\.Vb\.,^rbt-vb/'
