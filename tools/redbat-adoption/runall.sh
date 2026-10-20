#!/bin/bash
# usage: runall.sh <tag> [suite...]
A=/tmp/redbat-adoption
tag=$1; shift
suites=${@:-word-a word-b word-c word-d word-hf word-prog word-new word-load sheet-a sheet-b sheet-reg sheet-new10 ui-1 ui-2 ui-3 preview chat files shots fidelity word-new12 sheet-new12 word-new13 sheet-new13 word-new14 sheet-new14 word-new16 sheet-new16 word-new18 sheet-new18}
tp=0; tf=0
for s in $suites; do
  ( cd $A; timeout 900 node $s.mjs > $A/run$tag-$s.log 2>&1 )
  p=$(grep -c "^PASS" $A/run$tag-$s.log); f=$(grep -c "^FAIL" $A/run$tag-$s.log)
  echo "$s: $p pass, $f fail, exit-tail: $(tail -1 $A/run$tag-$s.log | cut -c1-150)"
  tp=$((tp+p)); tf=$((tf+f))
done
echo "TOTAL $tp pass $tf fail"
