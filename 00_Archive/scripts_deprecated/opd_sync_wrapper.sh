#!/bin/zsh
# wrapper สำหรับ LaunchAgent — ไม่มี colon/space ใน path ของตัวเอง
# ใช้ cd เข้า dir ก่อนแล้วเรียก ./run_daily.sh (relative path ไม่มี colon)
export PATH="/Library/Frameworks/Python.framework/Versions/3.13/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
cd "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts"
/bin/zsh ./run_daily.sh
