#!/usr/bin/env bash
# build/loop.sh -- tight feedback loop for the Excel agent probe macros.
#
#   ./build/loop.sh S4                 # build probe .xlam, run S4, report
#   ./build/loop.sh S4 --form          # include frmAgent
#   ./build/loop.sh S4 --timeout 60
#
# Decides in one second-resolution step whether the probe
#   finished  -> RESULT: OK / FAILED   (with breadcrumbs)
#   stalled   -> RESULT: HANG          (with breadcrumbs + VBE capture)
# and never leaves a stuck Excel behind.
set -u

MACRO="${1:-PbGuard}"; shift || true
FORM=""
TIMEOUT=45
while [ $# -gt 0 ]; do
  case "$1" in
    --form)    FORM="-WithForm" ;;
    --timeout) TIMEOUT="$2"; shift ;;
    *) echo "unknown arg: $1" >&2; exit 2 ;;
  esac
  shift
done

WIN_SRC='\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent'
PS() { ( cd /mnt/c && powershell.exe -NoProfile -ExecutionPolicy Bypass "$@" ); }
kill_excel() {
  PS -Command 'Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force' >/dev/null 2>&1
  sleep 0.6
}
trim() { tr -d '\r' | tail -1; }

kill_excel

WTEMP=$(PS -Command 'Write-Output $env:TEMP' | trim)
WSLTEMP=$(wslpath -u "$WTEMP")
# Unique per run: Windows keeps a handle on the previous run's file, so deleting
# it from WSL fails and stale breadcrumbs would be mistaken for fresh ones.
STAMP=$(date +%H%M%S)
LOG_WIN="$WTEMP\\probe_$STAMP.log"
LOG_WSL=$(wslpath -u "$LOG_WIN")

XLAM=$(PS -File "$WIN_SRC\\build\\probe.ps1" $FORM 2>/dev/null | trim)
if [ -z "$XLAM" ] || [ ! -f "$(wslpath -u "$XLAM" 2>/dev/null)" ]; then
  echo "BUILD FAILED (no xlam produced)"; exit 1
fi

PS -File "$WIN_SRC\\build\\runner.ps1" -Xlam "$XLAM" -Macro "$MACRO" -LogPath "$LOG_WIN" \
   > /tmp/loop_run.out 2>&1 &
BPID=$!

for _ in $(seq 1 "$TIMEOUT"); do
  kill -0 "$BPID" 2>/dev/null || break
  sleep 1
done

HANG=0
if kill -0 "$BPID" 2>/dev/null; then
  HANG=1
  XPID=$(PS -Command '(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1).Id' | trim)
  if [ -n "$XPID" ]; then
    PS -File "$WIN_SRC\\build\\vbepoke.ps1" -ProcId "$XPID" -Out "probe_hang.png" 2>/dev/null | head -1
  fi
fi

kill_excel
wait "$BPID" 2>/dev/null

echo "=============================================================="
echo "macro : $MACRO"
if [ "$HANG" -eq 1 ]; then
  echo "RESULT: HANG (no return within ${TIMEOUT}s)"
else
  grep -a 'RESULT:' /tmp/loop_run.out || echo "RESULT: (runner produced no verdict)"
fi
echo "--- breadcrumbs ----------------------------------------------"
if [ -f "$LOG_WSL" ]; then cat "$LOG_WSL"; else echo "(no breadcrumb file written)"; fi
echo "=============================================================="
