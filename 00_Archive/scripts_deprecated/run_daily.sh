#!/bin/zsh
# run_daily.sh — OpenDurian How-to Dashboard Auto Sync
# รันทุกชั่วโมง ผ่าน LaunchAgent com.opendurian.dailyrun
# Source of truth: LOCAL → push to GitHub
#
# Run order (สำคัญ):
#   meta_pull.py  → เรียก mk13_sync ข้างใน → RAW_DATA + CREATOR_META
#   opd_pull.py   → AFFILIATE_DATA + GOALS_DATA (ไม่เขียน OPD_DAILY แล้ว)
#   kms_pull.py   → TikTok rows เข้า RAW_DATA + KMS_CAL_DATA
# ─────────────────────────────────────────────────────

# PATH ใน LaunchAgent จำกัด — ระบุ path ตรงๆ
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
PYTHON="/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
if [ ! -f "$PYTHON" ]; then
  PYTHON=$(which python3 2>/dev/null || which python 2>/dev/null)
fi
if [ -z "$PYTHON" ]; then echo "ERROR: python3 not found" >> "$LOG"; exit 1; fi
DIR="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard"
SCRIPTS="$DIR/05_Scripts"
LOG="$SCRIPTS/meta_pull.log"

echo "========================================" >> "$LOG"
echo "  Run: $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG"
echo "========================================" >> "$LOG"

# 1. Meta Ads API → (mk13_sync inside) → RAW_DATA + CREATOR_META
#    mk13_sync ทำ: OPD_DAILY + OPD_PROD_DATA + ADSM44_PCTADS
echo "[1/3] Meta Ads + MK13 Sync..." >> "$LOG"
$PYTHON "$SCRIPTS/meta_pull.py" >> "$LOG" 2>&1
META_STATUS=$?

# 2. OPD Pull → AFFILIATE_DATA + GOALS_DATA (skip OPD_DAILY)
echo "[2/3] OPD Pull (Affiliate + Goals)..." >> "$LOG"
$PYTHON "$SCRIPTS/opd_pull.py" >> "$LOG" 2>&1
OPD_STATUS=$?

# 3. KMS Sheet → TikTok rows + KMS_CAL_DATA
echo "[3/3] KMS Sheet Pull..." >> "$LOG"
$PYTHON "$SCRIPTS/kms_pull.py" >> "$LOG" 2>&1
KMS_STATUS=$?

# สรุปผล
echo "--- สรุป ---" >> "$LOG"
[ $META_STATUS -eq 0 ] && echo "  ✅ meta_pull.py (+ mk13_sync)"  >> "$LOG" || echo "  ❌ meta_pull.py  (exit $META_STATUS)" >> "$LOG"
[ $OPD_STATUS  -eq 0 ] && echo "  ✅ opd_pull.py"                 >> "$LOG" || echo "  ❌ opd_pull.py   (exit $OPD_STATUS)"  >> "$LOG"
[ $KMS_STATUS  -eq 0 ] && echo "  ✅ kms_pull.py"                 >> "$LOG" || echo "  ❌ kms_pull.py   (exit $KMS_STATUS)"  >> "$LOG"

# Git push → GitHub
echo "--- Git push ---" >> "$LOG"
cd "$DIR"
find .git -name "*.lock" -delete >> "$LOG" 2>&1
git add index.html
if git diff --cached --quiet; then
  echo "  ℹ️  ไม่มีการเปลี่ยนแปลง ข้าม push" >> "$LOG"
else
  git commit -m "auto: $(date '+%Y-%m-%d %H:%M')" >> "$LOG" 2>&1
  git pull --rebase --autostash origin main >> "$LOG" 2>&1
  PULL_STATUS=$?
  if [ $PULL_STATUS -ne 0 ]; then
    echo "  ⚠️  pull rebase ล้มเหลว (exit $PULL_STATUS) — force push local" >> "$LOG"
    git push --force-with-lease >> "$LOG" 2>&1
  else
    git push >> "$LOG" 2>&1
  fi
  [ $? -eq 0 ] && echo "  ✅ Git push สำเร็จ" >> "$LOG" || echo "  ❌ Git push ล้มเหลว" >> "$LOG"
fi

echo "◀ $(date '+%Y-%m-%d %H:%M:%S') เสร็จสิ้น" >> "$LOG"
echo "========================================" >> "$LOG"
