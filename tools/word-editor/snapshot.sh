#!/usr/bin/env bash
# Launch the word-editor, snapshot its window, kill it.
#
# Usage: snapshot.sh <output.png> [docx-path]
set -u

OUT="${1:-/tmp/word-editor.png}"
DOC="${2:-}"

PROJ="$(cd "$(dirname "$0")" && pwd)"
cd "$PROJ/../.."  # repo root so the relative `tests/...` path works

WAIT_SECS="${WORD_EDITOR_SNAPSHOT_WAIT:-2.5}"

# Ensure no stale word-editor processes survive from earlier runs.
pkill -9 -f "/word-editor/bin/" 2>/dev/null
sleep 0.2

EXE="tools/word-editor/bin/Debug/net10.0/WordEditor"
(
    if [ -n "$DOC" ]; then
        "$EXE" "$DOC" >/tmp/word-editor.log 2>&1
    else
        "$EXE" >/tmp/word-editor.log 2>&1
    fi
) &
PID=$!

cleanup() {
    kill "$PID" 2>/dev/null
    pkill -9 -f "/word-editor/bin/" 2>/dev/null
    wait 2>/dev/null
}
trap cleanup EXIT

sleep "$WAIT_SECS"

# Find the X11 window by title prefix. Tree lists newest first; take that.
WIN_ID="$(xwininfo -root -tree 2>/dev/null | awk '/"RedBat Word/ {print $1; exit}')"

if [ -n "$WIN_ID" ]; then
    import -window "$WIN_ID" "$OUT" 2>&1
else
    import -window root "$OUT" 2>&1
    echo "warning: word-editor window not found; captured root" >&2
fi

ls -la "$OUT"
