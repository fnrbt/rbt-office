#!/usr/bin/env bash
# Hold an *active* RDP session on the Word VM.
#
# Word's layout/print automation runs through an interactive scheduled task and
# needs a live desktop. When the VM's interactive session is merely "Disc"
# (disconnected) Word hangs and every extraction stalls until the 180s task
# timeout. Connecting over RDP flips the session to "Active"; this script holds
# that connection open (headless, via Xvfb) and periodically wiggles the mouse
# so Windows never idle-locks the session.
#
# Commands:
#   ensure   start the connection if the session is not already Active (idempotent)
#   start    (re)start the connection
#   stop     tear down the connection (and the wiggle + Xvfb it owns)
#   status   exit 0 if the VM session is Active, non-zero otherwise
#
# Config (env, with defaults):
#   WORD_VM                 ssh alias for the VM            (default: winvm)
#   WORD_VM_RDP_HOST        RDP host/IP   (default: resolved from `ssh -G $WORD_VM`)
#   WORD_VM_RDP_USER        RDP username                    (default: olivia)
#   WORD_VM_RDP_PASSWORD    RDP password                    (default: pass)
#   WORD_VM_RDP_DISPLAY     Xvfb display for the client     (default: :99)
#   WORD_VM_RDP_GEOMETRY    screen geometry                 (default: 1280x800)
#   WORD_VM_RDP_WIGGLE_SECS seconds between mouse wiggles   (default: 30)
set -uo pipefail

vm=${WORD_VM:-winvm}
rdp_user=${WORD_VM_RDP_USER:-olivia}
rdp_pass=${WORD_VM_RDP_PASSWORD:-pass}
display=${WORD_VM_RDP_DISPLAY:-:99}
geometry=${WORD_VM_RDP_GEOMETRY:-1280x800}
wiggle_secs=${WORD_VM_RDP_WIGGLE_SECS:-30}

rdp_host=${WORD_VM_RDP_HOST:-}
if [[ -z "$rdp_host" ]]; then
  rdp_host=$(ssh -G "$vm" 2>/dev/null | awk '/^hostname /{print $2; exit}')
  rdp_host=${rdp_host:-$vm}
fi

state_dir=${WORD_VM_RDP_STATE_DIR:-/tmp/rb-word-rdp}
mkdir -p "$state_dir"
xvfb_pidf="$state_dir/xvfb.pid"
rdp_pidf="$state_dir/rdp.pid"
wiggle_pidf="$state_dir/wiggle.pid"
xvfb_ownedf="$state_dir/xvfb.owned"
rdp_log="$state_dir/rdp.log"

log() { echo "[vm-rdp] $*" >&2; }

session_active() {
  # Returns 0 if the RDP user's Windows session reports STATE=Active.
  ssh "$vm" 'query user' 2>/dev/null \
    | awk -v u="$rdp_user" 'tolower($1)==tolower(u) && $0 ~ /Active/ {found=1} END{exit found?0:1}'
}

alive() { [[ -f "$1" ]] && kill -0 "$(cat "$1" 2>/dev/null)" 2>/dev/null; }

start_xvfb() {
  if DISPLAY="$display" xdpyinfo >/dev/null 2>&1; then
    return 0  # display already up (someone else owns it)
  fi
  setsid Xvfb "$display" -screen 0 "${geometry}x24" >"$state_dir/xvfb.log" 2>&1 </dev/null &
  echo $! >"$xvfb_pidf"
  : >"$xvfb_ownedf"
  for _ in $(seq 1 20); do
    DISPLAY="$display" xdpyinfo >/dev/null 2>&1 && return 0
    sleep 0.25
  done
  log "Xvfb on $display did not come up"; return 1
}

start_rdp() {
  DISPLAY="$display" setsid xfreerdp3 \
    /v:"$rdp_host" /u:"$rdp_user" /p:"$rdp_pass" \
    /cert:ignore /gdi:sw /size:"$geometry" \
    +auto-reconnect /auto-reconnect-max-retries:0 \
    -grab-keyboard /log-level:ERROR \
    >"$rdp_log" 2>&1 </dev/null &
  echo $! >"$rdp_pidf"
}

start_wiggle() {
  # Nudge the pointer inside the client window so Windows never idle-locks.
  setsid bash -c '
    while :; do
      DISPLAY="'"$display"'" xdotool mousemove 40 40 >/dev/null 2>&1
      sleep 1
      DISPLAY="'"$display"'" xdotool mousemove 120 120 >/dev/null 2>&1
      sleep "'"$wiggle_secs"'"
    done' >/dev/null 2>&1 </dev/null &
  echo $! >"$wiggle_pidf"
}

cmd_start() {
  command -v xfreerdp3 >/dev/null 2>&1 || { log "xfreerdp3 not installed"; return 1; }
  command -v Xvfb >/dev/null 2>&1 || { log "Xvfb not installed"; return 1; }
  cmd_stop_quiet
  start_xvfb || return 1
  start_rdp
  start_wiggle
  for _ in $(seq 1 45); do
    if session_active; then
      log "session Active for $rdp_user on $rdp_host (display $display)"
      return 0
    fi
    alive "$rdp_pidf" || { log "xfreerdp exited early; see $rdp_log"; tail -3 "$rdp_log" >&2; return 1; }
    sleep 1
  done
  log "connected but session not reported Active; see $rdp_log"; return 1
}

cmd_stop_quiet() {
  alive "$wiggle_pidf" && kill "$(cat "$wiggle_pidf")" 2>/dev/null
  alive "$rdp_pidf" && kill "$(cat "$rdp_pidf")" 2>/dev/null
  if [[ -f "$xvfb_ownedf" ]] && alive "$xvfb_pidf"; then
    kill "$(cat "$xvfb_pidf")" 2>/dev/null
  fi
  rm -f "$wiggle_pidf" "$rdp_pidf" "$xvfb_pidf" "$xvfb_ownedf"
}

# Foreground supervisor: hold the session Active, reconnecting if the client
# drops, until killed. Launch this with the harness `run_in_background` (or
# nohup) so it survives across commands -- a plain `start` that backgrounds and
# returns is killed when its spawning shell/cgroup exits.
cmd_run() {
  command -v xfreerdp3 >/dev/null 2>&1 || { log "xfreerdp3 not installed"; exit 1; }
  command -v Xvfb >/dev/null 2>&1 || { log "Xvfb not installed"; exit 1; }
  trap 'cmd_stop_quiet; exit 0' INT TERM
  start_xvfb || exit 1
  start_wiggle
  log "supervising RDP session for $rdp_user@$rdp_host (display $display)"
  while :; do
    if ! alive "$rdp_pidf"; then
      start_rdp
      for _ in $(seq 1 45); do session_active && break; sleep 1; done
      session_active && log "session Active" || log "waiting for session; see $rdp_log"
    fi
    sleep 5
  done
}

case "${1:-ensure}" in
  ensure)
    # Idempotent: only (re)connect if the session is not already Active. The
    # connection lives for the lifetime of the calling command's process group.
    if session_active; then exit 0; fi
    cmd_start ;;
  start) cmd_start ;;
  run) cmd_run ;;
  stop) cmd_stop_quiet; log "stopped" ;;
  status)
    if session_active; then log "Active"; exit 0; else log "not Active"; exit 1; fi ;;
  *) echo "usage: $0 {ensure|start|run|stop|status}" >&2; exit 2 ;;
esac
