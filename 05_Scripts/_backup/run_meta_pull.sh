#!/bin/bash
# Meta Ads Auto Pull — รันโดย cron ทุกชั่วโมง
# ─────────────────────────────────────────────

SCRIPT_DIR="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts"
LOG_FILE="$SCRIPT_DIR/meta_pull.log"
PYTHON="/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"

echo "──────────────────────────────────" >> "$LOG_FILE"
echo "▶ $(date '+%Y-%m-%d %H:%M:%S') เริ่มรัน" >> "$LOG_FILE"

cd "$SCRIPT_DIR"
$PYTHON meta_pull.py >> "$LOG_FILE" 2>&1

EXIT_CODE=$?
if [ $EXIT_CODE -eq 0 ]; then
    echo "✅ Pull สำเร็จ (exit 0)" >> "$LOG_FILE"

    # ── Git push ──────────────────────────────────
    REPO_DIR="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard"
    cd "$REPO_DIR"

    # เช็คว่ามีการเปลี่ยนแปลงจริง
    if git diff --quiet; then
        echo "ℹ️  ไม่มีการเปลี่ยนแปลง ข้าม git push" >> "$LOG_FILE"
    else
        git add index.html >> "$LOG_FILE" 2>&1
        git commit -m "auto: Meta Ads update $(date '+%Y-%m-%d %H:%M')" >> "$LOG_FILE" 2>&1
        git push >> "$LOG_FILE" 2>&1

        GIT_CODE=$?
        if [ $GIT_CODE -eq 0 ]; then
            echo "✅ Git push สำเร็จ" >> "$LOG_FILE"
        else
            echo "❌ Git push error (exit $GIT_CODE)" >> "$LOG_FILE"
        fi
    fi
    # ─────────────────────────────────────────────

else
    echo "❌ Pull Error (exit $EXIT_CODE) — ข้าม git push" >> "$LOG_FILE"
fi

echo "◀ $(date '+%Y-%m-%d %H:%M:%S') เสร็จสิ้น" >> "$LOG_FILE"
