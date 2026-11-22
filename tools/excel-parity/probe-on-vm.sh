#!/usr/bin/env bash
# probe-on-vm.sh PROBE.vbs OUT [FILE.xlsx ...]: runs a probe that writes one output file (its first argument) in
# Excel's desktop session on the Windows VM (run-vbs-in-session.ps1), the workbooks given copied beside it and passed
# as its further arguments, and brings the output back as UTF-8. A probe that writes pictures gets a folder `png`
# beside it as the argument after the workbooks when PNGDIR is set, copied back to $PNGDIR. BOOKFIRST=1 passes the
# workbooks before the output (chart-probe.vbs WORKBOOK OUT [PNGDIR]).
set -euo pipefail
if [[ $# -lt 2 ]]; then echo "usage: probe-on-vm.sh PROBE.vbs OUT [FILE.xlsx ...]" >&2; exit 2; fi
here=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
vm=${WORD_VM:-winvm}
vm_dir=${WORD_VM_DIR:-C:/Users/olivia/RbtExcelParity}
probe=$1; out=$2; shift 2
run_id=probe-$(date -u +%Y%m%dT%H%M%SZ)-$$
remote="${vm_dir}/${run_id}"
win=${remote//\//\\}
ssh "$vm" "powershell -NoProfile -Command \"New-Item -ItemType Directory -Force -Path '$remote/png' | Out-Null\"" 2>/dev/null
scp -q "$probe" "$here/run-vbs-in-session.ps1" "$@" "${vm}:${remote}/" 2>/dev/null
books=""
for file in "$@"; do books="$books ${win}\\$(basename -- "$file")"; done
# (BOOKFIRST=1: a probe that takes its workbook before its output, as chart-probe.vbs.)
if [[ -n "${BOOKFIRST:-}" ]]; then args="${books# } ${win}\\out.tsv"; else args="${win}\\out.tsv${books}"; fi
if [[ -n "${PNGDIR:-}" ]]; then args="$args ${win}\\png"; fi
ssh "$vm" "powershell -NoProfile -ExecutionPolicy Bypass -File \"${remote}/run-vbs-in-session.ps1\" -ScriptPath \"${win}\\$(basename -- "$probe")\" -Arguments \"${args}\" -TimeoutSeconds ${TIMEOUT:-3600}" 2>/dev/null || true
work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT
scp -q "${vm}:${remote}/out.tsv" "$work/out.tsv" 2>/dev/null
if [[ "$(file -b --mime-encoding "$work/out.tsv")" == utf-16* ]]; then iconv -f UTF-16 -t UTF-8 "$work/out.tsv" | tr -d '\r' > "$out"; else tr -d '\r' < "$work/out.tsv" > "$out"; fi
if [[ -n "${PNGDIR:-}" ]]; then mkdir -p "$PNGDIR"; scp -q "${vm}:${remote}/png/*" "$PNGDIR/" 2>/dev/null || true; fi
echo "$out: $(wc -l < "$out") lines"
