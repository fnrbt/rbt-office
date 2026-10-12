#!/bin/bash
cd /tmp/redbat-adoption
node static.mjs /tmp/redbat-adoption/build 16140 > static14.log 2>&1 &
SP=$!
sleep 1
PORT=16140 node perf-open.mjs /tmp/redbat-adoption/long-generated.docx 3 2>&1 | grep -v 404 > perf-open-now14.txt
PORT=16140 node perf-scroll.mjs 2>&1 | grep -v 404 > perf-scroll-now14.txt
PORT=16140 node perf-mass.mjs 2>&1 | grep -v 404 > perf-mass-now14.txt
PORT=16140 node perf-save.mjs 2>&1 | grep -v 404 > perf-save-now14.txt
kill $SP
node perf-open.mjs /tmp/redbat-adoption/long-generated.docx 3 2>&1 | grep -v 404 > perf-open-now16.txt
node perf-scroll.mjs 2>&1 | grep -v 404 > perf-scroll-now16.txt
node perf-mass.mjs 2>&1 | grep -v 404 > perf-mass-now16.txt
node perf-save.mjs 2>&1 | grep -v 404 > perf-save-now16.txt
node fid16-extra.mjs sheet > run-fid16-extra-sheet.log 2>&1; cp fid16-extra-table.json fid16-extra-sheet-table.json
node fid16-extra.mjs word > run-fid16-extra-word.log 2>&1
cp fid16-extra-table.json fid16-extra-word-table.json
echo DONE > batch2.done
