#!/usr/bin/env bash
set -euo pipefail

vm="${WORD_VM_HOST:-winvm}"
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
  verdana.ttf
  verdanab.ttf
  verdanai.ttf
  verdanaz.ttf
  # Monotype Garamond (Office) + Segoe UI — used by the dashboard fixtures.
  GARA.TTF
  GARABD.TTF
  GARAIT.TTF
  segoeui.ttf
  segoeuib.ttf
  segoeuii.ttf
  segoeuiz.ttf
  # CJK font-link fallbacks Excel uses when a Calibri/Arial cell hits a
  # codepoint outside the Latin set. Matches the GDI font-link chain
  # that Office uses on a stock Win10/11 install: Yu Gothic for
  # Japanese, Microsoft YaHei for Simplified Chinese, Microsoft
  # JhengHei for Traditional, Malgun Gothic for Korean. msgothic kept
  # as a legacy fallback for older fixtures.
  YuGothR.ttc
  YuGothB.ttc
  msgothic.ttc
  msyh.ttc
  msyhbd.ttc
  msjh.ttc
  msjhbd.ttc
  malgun.ttf
  malgunbd.ttf
)

mkdir -p "$out_dir"

for font in "${fonts[@]}"; do
  # Tolerate missing fonts — older Windows images don't always ship
  # every CJK face — but keep the standard ones load-bearing.
  if ! scp -q "${vm}:C:/Windows/Fonts/${font}" "${out_dir}/${font}"; then
    printf '  warning: failed to copy %s\n' "$font" >&2
  fi
done

chmod u+rw,go-rwx "$out_dir"/* || true
printf 'windows fonts copied to %s\n' "$out_dir"
