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

# More faces the corpus documents name (Consolas, Courier New, Tahoma, Georgia,
# Trebuchet MS, Lucida Console, Arial Black, the symbol fonts), for the page
# comparison only (tools/word-parity/page-parity.mjs registers them as web
# fonts). Kept apart so that the native font lookup and the browser checks,
# which read $out_dir, see the same faces as before.
extra_dir="${RB_WORD_EXTRA_FONT_DIR:-$(dirname "$out_dir")/windows-extra}"
extra=(ariblk.ttf consola.ttf consolab.ttf consolai.ttf consolaz.ttf cour.ttf courbd.ttf courbi.ttf couri.ttf
  georgia.ttf georgiab.ttf georgiai.ttf georgiaz.ttf lucon.ttf tahoma.ttf tahomabd.ttf trebuc.ttf trebucbd.ttf
  trebucbi.ttf trebucit.ttf symbol.ttf wingding.ttf seguisym.ttf seguibl.ttf seguibli.ttf
  TCM_____.TTF TCB_____.TTF TCMI____.TTF TCBI____.TTF)
mkdir -p "$extra_dir"
for font in "${extra[@]}"; do
  scp -q "${vm}:C:/Windows/Fonts/${font}" "${extra_dir}/${font}" || printf '  warning: failed to copy %s\n' "$font" >&2
done
chmod u+rw,go-rwx "$extra_dir"/* || true
printf 'extra windows fonts copied to %s\n' "$extra_dir"

# Office's cloud fonts a corpus document names (free-test-100kb and -500kb: Open Sans; Excel models: Aptos Narrow), which Word
# downloads into the user's font cache when it opens the document - not in C:/Windows/Fonts.
cloud="C:/Users/${WORD_VM_USER:-olivia}/AppData/Local/Microsoft/FontCache/4/CloudFonts"
for family in "Open Sans" "Aptos Narrow"; do
  slug="$(printf '%s' "$family" | tr -d ' ' | tr '[:upper:]' '[:lower:]')"
  # (The cache names its files by number: one per face.)
  for file in $(ssh "$vm" "powershell -NoProfile -Command \"Get-ChildItem '${cloud}/${family}' -Name\"" 2>/dev/null | tr -d '\r'); do
    scp -q "${vm}:${cloud}/${family}/${file}" "${extra_dir}/${slug}-${file}" || printf '  warning: failed to copy %s/%s\n' "$family" "$file" >&2
  done
done
chmod u+rw,go-rwx "$extra_dir"/* || true
printf 'cloud fonts copied to %s\n' "$extra_dir"
