#!/bin/bash
# ────────────────────────────────────────────────────────────
#  How-to Dashboard — Sync from GS + Push to GitHub
#  Double-click ได้เลย ไม่ต้องเปิด Terminal
# ────────────────────────────────────────────────────────────

PROJ="/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard"
SCRIPT="$PROJ/05_Scripts/mk13_sync.py"

echo "============================================"
echo "  How-to Dashboard Sync"
echo "  $(date '+%Y-%m-%d %H:%M:%S')"
echo "============================================"

# ── 1. Pull from Google Sheets → update local HTML ──
echo ""
echo "📥 Step 1: ดึงข้อมูลจาก Google Sheets..."
python3 "$SCRIPT"
if [ $? -ne 0 ]; then
  echo ""
  echo "❌ Sync ล้มเหลว — ตรวจ internet / Sheet permissions"
  echo "กด Enter เพื่อปิด..."
  read
  exit 1
fi

# ── 2. Git push ──
echo ""
echo "🚀 Step 2: Push ขึ้น GitHub..."
cd "$PROJ"
rm -f .git/HEAD.lock .git/index.lock 2>/dev/null
git add index.html
git commit -m "auto: MK13 sync $(date '+%Y-%m-%d %H:%M')"
git push origin main

if [ $? -eq 0 ]; then
  echo ""
  echo "✅ เสร็จแล้ว! Dashboard อัปเดตแล้ว"
else
  echo ""
  echo "⚠️  Push ไม่สำเร็จ — แต่ local HTML อัปเดตแล้ว"
  echo "   ลอง git push เองจาก Terminal"
fi

echo ""
echo "กด Enter เพื่อปิด..."
read
