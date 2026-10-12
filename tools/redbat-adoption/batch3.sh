#!/bin/bash
cd /tmp/redbat-adoption
node fid14-word.mjs > run-fid16-word.log 2>&1; cp fid14-word-table-all.json fid16-word-table-all.json
node fid14-sheet.mjs > run-fid16-sheet.log 2>&1; cp fid14-sheet-table.json fid16-sheet-table.json
node fid16-extra.mjs word fieldbold > run-fid16-extra-fieldbold.log 2>&1
node explore16-word.mjs > run-explore16-word.log 2>&1
node explore16-sheet.mjs > run-explore16-sheet.log 2>&1
python3 -I validate16.py out/*.docx out/*.xlsx > validate16-saved.txt 2>/dev/null
echo DONE > batch3.done
