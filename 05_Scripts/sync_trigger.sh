#!/bin/bash
# sync_trigger.sh — รันโดย LaunchAgent com.opendurian.sync-trigger
# เมื่อ Claude เขียน .sync_trigger file → LaunchAgent ตรวจจับ → รัน script นี้
TRIGGER="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts/.sync_trigger"
LOG="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts/meta_pull.log"
PYTHON="/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
RUNNER="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts/opd_runner.py"

# ลบ trigger file ก่อน (ป้องกัน re-trigger ซ้ำ)
rm -f "$TRIGGER"

# รัน opd_runner.py
echo "[TRIGGER] sync_trigger.sh fired at $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG"
"$PYTHON" "$RUNNER"
