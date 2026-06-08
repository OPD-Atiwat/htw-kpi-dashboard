#!/bin/bash
# ════════════════════════════════════════════════════════
#  OPD Daily Update — รันทุกวันจันทร์–ศุกร์ 10:00 น.
#  ดาวน์โหลดข้อมูลจาก Google Sheets → อัปเดต index.html
# ════════════════════════════════════════════════════════

SCRIPT_DIR="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard"
LOG_FILE="$SCRIPT_DIR/05_Scripts/update.log"
PYTHON=$(which python3)

echo "======================================" >> "$LOG_FILE"
echo "$(date '+%Y-%m-%d %H:%M:%S') — เริ่ม update" >> "$LOG_FILE"

# 1. OPD Daily (MK13 orders + GOALS_DATA + Affiliate)
echo "--- [1/2] opd_pull.py ---" >> "$LOG_FILE"
cd "$SCRIPT_DIR" && "$PYTHON" opd_pull.py >> "$LOG_FILE" 2>&1
OPD_STATUS=$?

# 2. Meta Ads API
echo "--- [2/3] meta_pull.py ---" >> "$LOG_FILE"
cd "$SCRIPT_DIR" && "$PYTHON" meta_pull.py >> "$LOG_FILE" 2>&1
META_STATUS=$?

# 3. KMS Content Sheet
echo "--- [3/3] kms_pull.py ---" >> "$LOG_FILE"
cd "$SCRIPT_DIR" && "$PYTHON" kms_pull.py >> "$LOG_FILE" 2>&1
KMS_STATUS=$?

# สรุปผล
echo "$(date '+%Y-%m-%d %H:%M:%S') — สรุป:" >> "$LOG_FILE"
[ $OPD_STATUS -eq 0 ]  && echo "  ✅ opd_pull.py"  >> "$LOG_FILE" || echo "  ❌ opd_pull.py"  >> "$LOG_FILE"
[ $META_STATUS -eq 0 ] && echo "  ✅ meta_pull.py" >> "$LOG_FILE" || echo "  ❌ meta_pull.py" >> "$LOG_FILE"
[ $KMS_STATUS -eq 0 ]  && echo "  ✅ kms_pull.py"  >> "$LOG_FILE" || echo "  ❌ kms_pull.py"  >> "$LOG_FILE"

# 4. Git push → GitHub Pages (อัปเดต Dashboard อัตโนมัติ)
echo "--- [4/4] git push ---" >> "$LOG_FILE"
cd "$SCRIPT_DIR" && \
  git add index.html opd_daily_data.json && \
  git commit -m "auto update: $(date '+%Y-%m-%d %H:%M')" >> "$LOG_FILE" 2>&1 && \
  git push >> "$LOG_FILE" 2>&1
GIT_STATUS=$?
[ $GIT_STATUS -eq 0 ] && echo "  ✅ git push สำเร็จ" >> "$LOG_FILE" || echo "  ❌ git push ล้มเหลว" >> "$LOG_FILE"
echo "======================================" >> "$LOG_FILE"
