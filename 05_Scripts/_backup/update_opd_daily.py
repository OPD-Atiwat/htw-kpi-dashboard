#!/usr/bin/env python3
"""
════════════════════════════════════════════════════════════════
  OPD Daily Auto-Updater  v1.0
  อ่าน CSV จาก Google Drive Desktop → อัปเดต OPD_DAILY ใน index.html
  รันโดย Scheduled Task ทุกวัน 10:00 น.
════════════════════════════════════════════════════════════════
"""

import csv
import json
import os
import re
import sys
import glob
from datetime import datetime
from pathlib import Path

# ════════════════════════════════════════════════════════════════
#  CONFIG — แก้ค่าเหล่านี้ให้ตรงกับเครื่องของคุณ
# ════════════════════════════════════════════════════════════════

# Path ของ Google Drive Desktop (แก้ตามเครื่อง)
# Mac: "/Users/[username]/Library/CloudStorage/GoogleDrive-[email]/My Drive"
# Mac (เก่า): "/Users/[username]/Google Drive/My Drive"
# Windows: "G:/My Drive"  หรือ  "C:/Users/[username]/Google Drive/My Drive"
GOOGLE_DRIVE_PATHS = [
    # ลำดับ: ค้นหาจากบนลงล่าง ใช้ path แรกที่เจอ
    os.path.expanduser("~/Library/CloudStorage/GoogleDrive-how-to.cowork@opendurian.com/My Drive"),
    os.path.expanduser("~/Library/CloudStorage/GoogleDrive-jorjae@opendurian.com/My Drive"),
    os.path.expanduser("~/Google Drive/My Drive"),
    os.path.expanduser("~/GoogleDrive/My Drive"),
    # เพิ่ม path อื่นถ้าจำเป็น
]

# ชื่อไฟล์ CSV ที่ Apps Script สร้าง
CSV_FILENAME = "OPD_Daily_Export.csv"

# Dashboard files
DASHBOARD_DIR = Path(__file__).parent.parent
DASHBOARD_FILES = [
    DASHBOARD_DIR / "01_Dashboard" / "index.html",
    DASHBOARD_DIR / "index.html",
]

# ════════════════════════════════════════════════════════════════
#  COLUMN MAP — ชื่อ column ใน CSV → ชื่อ channel ใน OPD_DAILY
#  key = ชื่อ column ใน CSV (case-insensitive)
#  value = ชื่อ channel ที่ใช้ใน dashboard
# ════════════════════════════════════════════════════════════════
COLUMN_MAP = {
    # Revenue columns
    "date"          : "__date__",          # special: วันที่
    "วันที่"          : "__date__",
    "tiktok"        : "TikTok",
    "tiktok_brand"  : "TikTok",
    "tiktok brand"  : "TikTok",
    "tiktok live"   : "TikTok Live",
    "tiktokive"     : "TikTok Live",
    "tiktok_live"   : "TikTok Live",
    "tiktok affi"   : "TikTok Affi",
    "tiktok_affi"   : "TikTok Affi",
    "tiktok affiliate": "TikTok Affi",
    "shopee"        : "Shopee",
    "facebook"      : "Facebook",
    "meta"          : "Facebook",
    "หน้าร้าน"       : "หน้าร้าน",
    "hanaraan"      : "หน้าร้าน",
    "store"         : "หน้าร้าน",
    "shopify"       : "Shopify",
    "instagram"     : "Instagram",
    "ig"            : "Instagram",
    "lazada"        : "Lazada",
    "line"          : "LINE",
    "line shopping" : "LINE",
    "web"           : "Web",
    "website"       : "Web",
    "bookfair"      : "Bookfair",
    "book fair"     : "Bookfair",
    "youtube"       : "YouTube",
    "yt"            : "YouTube",

    # Quantity columns (ชื่อ column _q หรือ qty)
    "tiktok_qty"      : "TikTok_q",
    "tiktok qty"      : "TikTok_q",
    "tiktok live_qty" : "TikTok Live_q",
    "tiktok affi_qty" : "TikTok Affi_q",
    "shopee_qty"      : "Shopee_q",
    "facebook_qty"    : "Facebook_q",
    "หน้าร้าน_qty"    : "หน้าร้าน_q",
    "shopify_qty"     : "Shopify_q",
    "instagram_qty"   : "Instagram_q",
    "lazada_qty"      : "Lazada_q",
    "line_qty"        : "LINE_q",
    "web_qty"         : "Web_q",
    "bookfair_qty"    : "Bookfair_q",
    "youtube_qty"     : "YouTube_q",
}

ALL_CHANNELS = [
    "TikTok", "TikTok Live", "TikTok Affi", "Shopee", "Facebook",
    "หน้าร้าน", "Shopify", "Instagram", "Lazada", "LINE", "Web",
    "Bookfair", "YouTube"
]

# ════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════

def find_csv_file():
    """ค้นหาไฟล์ CSV ใน Google Drive Desktop paths"""
    for drive_path in GOOGLE_DRIVE_PATHS:
        csv_path = os.path.join(drive_path, CSV_FILENAME)
        if os.path.exists(csv_path):
            print(f"✅ พบไฟล์: {csv_path}")
            return csv_path

    # ถ้าไม่เจอ ลองค้นหาแบบ wildcard
    for drive_path in GOOGLE_DRIVE_PATHS:
        if os.path.isdir(os.path.expanduser("~/Library/CloudStorage")):
            pattern = os.path.expanduser(f"~/Library/CloudStorage/GoogleDrive-*/{CSV_FILENAME}")
            matches = glob.glob(pattern)
            if matches:
                print(f"✅ พบไฟล์ (auto-detect): {matches[0]}")
                return matches[0]

    return None


def parse_date(val):
    """แปลง date string ให้เป็น YYYY-MM-DD"""
    if not val:
        return None
    val = str(val).strip()
    # ลอง format ต่างๆ
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(val, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    # ถ้าเป็น serial number (Excel date)
    try:
        n = float(val)
        if 40000 < n < 60000:
            from datetime import timedelta
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=n)).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    return None


def parse_number(val):
    """แปลงตัวเลข (handle comma, thai baht sign)"""
    if val is None or val == "":
        return 0
    val = str(val).replace(",", "").replace("฿", "").replace(" ", "").strip()
    try:
        return int(float(val))
    except ValueError:
        return 0


def map_header(header):
    """Map ชื่อ column CSV → ชื่อ channel dashboard"""
    h = header.strip().lower()
    if h in COLUMN_MAP:
        return COLUMN_MAP[h]
    # ลองหา partial match
    for key, val in COLUMN_MAP.items():
        if key in h or h in key:
            return val
    return None


def read_csv(csv_path):
    """อ่าน CSV และแปลงเป็น OPD_DAILY format"""
    rows = []
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames
        if not headers:
            print("❌ ไม่พบ header ใน CSV")
            return None

        print(f"📋 Columns ใน CSV: {headers}")

        # สร้าง header mapping
        col_map = {}
        for h in headers:
            mapped = map_header(h)
            if mapped:
                col_map[h] = mapped
                print(f"   {h} → {mapped}")
            else:
                print(f"   {h} → (ไม่ได้ map)")

        # ตรวจว่ามี date column ไหม
        date_cols = [h for h, m in col_map.items() if m == "__date__"]
        if not date_cols:
            # ลองใช้ column แรก
            first_col = headers[0]
            col_map[first_col] = "__date__"
            print(f"   ⚠️ ไม่พบ date column → ใช้ column แรก: {first_col}")
            date_cols = [first_col]

        date_col = date_cols[0]

        for row in reader:
            date_str = parse_date(row.get(date_col, ""))
            if not date_str:
                continue  # ข้ามแถวที่ไม่มีวันที่

            entry = {"d": date_str}
            for csv_col, dashboard_col in col_map.items():
                if dashboard_col == "__date__":
                    continue
                val = parse_number(row.get(csv_col, 0))
                if val != 0:
                    entry[dashboard_col] = val

            rows.append(entry)

    # เรียงตามวันที่
    rows.sort(key=lambda r: r["d"])
    print(f"✅ อ่านได้ {len(rows)} แถว ({rows[0]['d'] if rows else 'N/A'} → {rows[-1]['d'] if rows else 'N/A'})")
    return rows


def update_html(html_path, new_data):
    """แทนที่ OPD_DAILY ใน index.html"""
    if not html_path.exists():
        print(f"⚠️ ไม่พบไฟล์: {html_path}")
        return False

    content = html_path.read_text(encoding="utf-8")

    # สร้าง OPD_DAILY JS object ใหม่
    opd_obj = {
        "channels": ALL_CHANNELS,
        "data": new_data
    }
    opd_json = json.dumps(opd_obj, ensure_ascii=False, separators=(",", ":"))

    # Replace ด้วย regex
    pattern = r'var OPD_DAILY\s*=\s*\{.*?\};'
    replacement = f'var OPD_DAILY = {opd_json};'
    new_content, count = re.subn(pattern, replacement, content, count=1, flags=re.DOTALL)

    if count == 0:
        print(f"❌ ไม่พบ var OPD_DAILY ใน {html_path.name}")
        return False

    html_path.write_text(new_content, encoding="utf-8")
    print(f"✅ อัปเดต {html_path.name} สำเร็จ")
    return True


# ════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print(f"  OPD Daily Auto-Updater — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # 1. หาไฟล์ CSV
    csv_path = find_csv_file()
    if not csv_path:
        print("❌ ไม่พบไฟล์ OPD_Daily_Export.csv ใน Google Drive Desktop")
        print("   ตรวจสอบ:")
        print("   1. ติดตั้ง Google Drive Desktop app แล้วหรือยัง")
        print("   2. Google Apps Script export ทำงานแล้วหรือยัง")
        print("   3. แก้ GOOGLE_DRIVE_PATHS ใน script ให้ถูกต้อง")
        sys.exit(1)

    # 2. อ่านและ parse CSV
    new_data = read_csv(csv_path)
    if not new_data:
        print("❌ ไม่สามารถ parse CSV ได้")
        sys.exit(1)

    # 3. อัปเดต HTML files
    success_count = 0
    for html_path in DASHBOARD_FILES:
        if update_html(html_path, new_data):
            success_count += 1

    print("-" * 60)
    if success_count > 0:
        print(f"🎉 อัปเดตสำเร็จ {success_count}/{len(DASHBOARD_FILES)} ไฟล์")
        print(f"   แถวล่าสุด: {new_data[-1]['d']} — {sum(v for k,v in new_data[-1].items() if k != 'd' and not k.endswith('_q')):,} บาท")
    else:
        print("❌ ไม่มีไฟล์ที่อัปเดตสำเร็จ")
        sys.exit(1)

    print("=" * 60)


if __name__ == "__main__":
    main()
