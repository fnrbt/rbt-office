#!/bin/bash
cd /tmp/redbat-adoption
node perf-first-edit.mjs 2>&1 | grep -v 404 > perf-first-edit-run16.txt
node perf-sheet.mjs 2>&1 | grep -v 404 > perf-sheet-run16.txt
node perf-open.mjs /tmp/redbat-adoption/long-generated.docx 3 2>&1 | grep -v 404 > perf-open-run16.txt
node perf-scroll.mjs 2>&1 | grep -v 404 > perf-scroll-run16.txt
node perf-save.mjs 2>&1 | grep -v 404 > perf-save-run16.txt
node perf-mass.mjs 2>&1 | grep -v 404 > perf-mass-run16.txt
echo DONE > perf-all.done
