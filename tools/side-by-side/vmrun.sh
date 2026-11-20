#!/usr/bin/env bash
# Run a command line in olivia's interactive desktop session on winvm (via a one-off /IT task); waits for it.
# usage: vmrun.sh 'C:\sbs\x.cmd' [timeout-seconds]
set -euo pipefail
cmd=$1; t=${2:-120}; name=RbtSbs-$RANDOM$RANDOM
ssh winvm "schtasks /Create /TN $name /TR \"cmd /c $cmd > C:\\sbs\\$name.log 2>&1 & echo done> C:\\sbs\\$name.done\" /SC ONCE /ST 23:59 /F /IT >nul && schtasks /Run /TN $name >nul" 2>/dev/null
for ((i=0;i<t;i++)); do ssh winvm "if exist C:\\sbs\\$name.done (type C:\\sbs\\$name.log & del C:\\sbs\\$name.done C:\\sbs\\$name.log)" 2>/dev/null | grep . && break; sleep 1; done || true
ssh winvm "schtasks /Delete /TN $name /F >nul" 2>/dev/null || true
