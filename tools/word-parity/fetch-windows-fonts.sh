#!/usr/bin/env bash
set -euo pipefail

vm="${WORD_VM_HOST:-wordvm}"
out_dir="${RB_WORD_FONT_DIR:-artifacts/fonts/windows}"

fonts=(
  arial.ttf
  arialbd.ttf
  arialbi.ttf
  ariali.ttf
  ARIALN.TTF
  ARIALNB.TTF
  ARIALNBI.TTF
  ARIALNI.TTF
  ARIALUNI.TTF
  calibri.ttf
  calibrib.ttf
  calibrii.ttf
  calibril.ttf
  calibrili.ttf
  calibriz.ttf
  cambria.ttc
  cambriab.ttf
  cambriai.ttf
  cambriaz.ttf
  times.ttf
  timesbd.ttf
  timesbi.ttf
  timesi.ttf
)

mkdir -p "$out_dir"

for font in "${fonts[@]}"; do
  scp -q "${vm}:C:/Windows/Fonts/${font}" "${out_dir}/${font}"
done

chmod u+rw,go-rwx "$out_dir"/* || true
printf 'windows fonts copied to %s\n' "$out_dir"
