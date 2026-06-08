#!/bin/zsh
# meta_hourly.sh — รัน meta_pull.py ทุกชั่วโมง (ผ่าน LaunchAgent)

PYTHON="/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
SCRIPTS="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts"
LOG="$SCRIPTS/meta_pull.log"

echo "========================================" >> "$LOG"
echo "  Meta Pull: $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG"
echo "========================================" >> "$LOG"

cd "$SCRIPTS"
$PYTHON -u "$SCRIPTS/meta_pull.py" >> "$LOG" 2>&1
STATUS=$?

if [ $STATUS -eq 0 ]; then
  echo "  ✅ meta_pull เสร็จสิ้น" >> "$LOG"
else
  echo "  ❌ meta_pull ล้มเหลว (exit $STATUS)" >> "$LOG"
fi

echo "◀ $(date '+%Y-%m-%d %H:%M:%S') จบ" >> "$LOG"
echo "========================================" >> "$LOG"
