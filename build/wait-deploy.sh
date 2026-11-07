#!/usr/bin/env bash
# Wait for Excel to exit, then deploy the rebuilt add-in.
cd /home/eloaurdi/hay2010agent/excel-agent || exit 1
echo "waiting for Excel to close..."
for i in $(seq 1 720); do
  running=$(tasklist.exe 2>/dev/null | grep -ci "excel.exe")
  if [ "$running" -eq 0 ]; then
    echo "Excel closed after ~$((i*5))s - deploying"
    for attempt in 1 2 3 4 5; do
      if (cd /mnt/c && timeout 120 powershell.exe -NoProfile -ExecutionPolicy Bypass \
            -File '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent\deploy.ps1' 2>&1 | tr -d '\r' | tail -3); then
        echo "deploy finished (attempt $attempt)"
        exit 0
      fi
      sleep 3
    done
    echo "deploy kept failing"
    exit 1
  fi
  sleep 5
done
echo "gave up waiting (10 min)"
exit 1
