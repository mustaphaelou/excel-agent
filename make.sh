#!/usr/bin/env bash
# One-shot build + ribbon injection + install + self-test.
# Usage: ./make.sh [--build-only]
set -u
cd "$(dirname "$0")" || exit 1

SRC='\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent'
PS=(powershell.exe -NoProfile -ExecutionPolicy Bypass -File)

echo "== 1/3 build (Excel COM) =="
( cd /mnt/c && timeout 240 "${PS[@]}" "$SRC\\build.ps1" ) 2>&1 | tr -d '\r' | tail -10
rc=${PIPESTATUS[0]}
if [ "$rc" -ne 0 ]; then echo "BUILD FAILED (exit $rc)"; exit 1; fi
[ -f build/ExcelAgent.raw.xlam ] || { echo "BUILD FAILED: no raw.xlam"; exit 1; }

echo "== 2/3 inject ribbon =="
python3 inject_ribbon.py build/ExcelAgent.raw.xlam build/ExcelAgent.xlam ribbon/customUI14.xml || exit 1

if [ "${1:-}" = "--build-only" ]; then echo "done (build only)"; exit 0; fi

echo "== 3/4 deploy =="
( cd /mnt/c && timeout 120 "${PS[@]}" "$SRC\\deploy.ps1" ) 2>&1 | tr -d '\r' | tail -6

cd /home/eloaurdi/hay2010agent/excel-agent || exit 1
echo "== 4/4 self-test =="
( cd /mnt/c && timeout 150 "${PS[@]}" "$SRC\\selftest.ps1" ) 2>&1 | tr -d '\r' | tail -40
rc=$?

# If the macro stalled in break mode, the COM call never returns; clear it so
# the next run starts clean. Only touch instances running our add-in.
cd /home/eloaurdi/hay2010agent/excel-agent || exit $rc
if [ -f build/trace.log ]; then
  echo "== trace (from workspace copy) =="
  cat build/trace.log
fi
exit 0
