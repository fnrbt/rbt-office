#!/usr/bin/env bash
# tools/word-parity/macro-parity.sh <out-dir> <module.bas>...
#
# For each module: a .docm of Word's blank document with the module (make-docm.fsx), its Main run in
# Microsoft Word on the VM (word-macro-on-vm.sh: <out>/<name>/word/Main.tsv + .pdf) and in ours
# (run-macro-native.fsx: <out>/<name>/ours/Main.tsv + .docx), and the two compared
# (compare-macro-dumps.py: <out>/<name>/compare.txt). Needs the Debug build of src/Word.
set -uo pipefail
cd "$(dirname "$0")/../.."
out=$1
shift
for bas in "$@"; do
  name=$(basename "$bas" .bas)
  dir="$out/$name"
  mkdir -p "$dir/word" "$dir/ours"
  dotnet fsi --quiet tools/word-parity/make-docm.fsx tests/macros/word-blank.docx "$dir/$name.docm" "$bas" >/dev/null
  if [[ "${SKIP_WORD:-0}" != 1 ]]; then
    tools/word-parity/word-macro-on-vm.sh "$dir/$name.docm" Main "$dir/word" >/dev/null 2>&1 || echo "$name: Word run failed"
  fi
  dotnet fsi --quiet tools/word-parity/run-macro-native.fsx "$dir/$name.docm" Main "$dir/ours/Main.tsv" "$dir/ours/$name.docx" >/dev/null 2>"$dir/ours/stderr.txt"
  if [[ -f "$dir/word/Main.tsv" ]]; then
    python3 tools/word-parity/compare-macro-dumps.py "$dir/word/Main.tsv" "$dir/ours/Main.tsv" > "$dir/compare.txt"
    echo "$name: $(tail -1 "$dir/compare.txt")"
  else
    echo "$name: no Word result"
  fi
done
