#!/usr/bin/env python3
"""
MK13 + ADSM44 Sync — OpenDurian How-to
────────────────────────────────────────
อ่านข้อมูลจาก Google Sheets (CSV export) แล้วอัปเดต
  • OPD_DAILY      ← MK13 raw orders
  • ADSM44_PCTADS  ← ADSM44 % Ads by product/month
ใน index.html (local) โดยไม่ push เอง — meta_pull.py จะ push รวมกัน

รัน standalone: python mk13_sync.py
Import:         import mk13_sync; mk13_sync.sync()
"""

import csv
import io
import json
import re
import ssl
import sys
from datetime import datetime
from collections import defaultdict
import urllib.request
import urllib.error

# macOS SSL fix
ssl._create_default_https_context = ssl._create_unverified_context

# ============================================================
# CONFIG
# ============================================================
SHEET_ID      = "1qYdwXuCHDHeHN6a8vU_RVFBYT5MuYq-8K5AMK3cP4DM"
MK13_GID      = "964123706"
ADSM44_GID    = "1860322120"
DASHBOARD_PATH = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"
# ============================================================

MONTH_LABEL = {
    "2025-01":"Jan 25","2025-02":"Feb 25","2025-03":"Mar 25",
    "2025-04":"Apr 25","2025-05":"May 25","2025-06":"Jun 25",
    "2025-07":"Jul 25","2025-08":"Aug 25","2025-09":"Sep 25",
    "2025-10":"Oct 25","2025-11":"Nov 25","2025-12":"Dec 25",
    "2026-01":"Jan 26","2026-02":"Feb 26","2026-03":"Mar 26",
    "2026-04":"Apr 26","2026-05":"May 26","2026-06":"Jun 26",
    "2026-07":"Jul 26","2026-08":"Aug 26","2026-09":"Sep 26",
    "2026-10":"Oct 26","2026-11":"Nov 26","2026-12":"Dec 26",
}


# ─── CSV fetch ──────────────────────────────────────────────

def fetch_csv(gid):
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
           f"/export?format=csv&gid={gid}")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.read().decode("utf-8-sig")  # utf-8-sig กัน BOM
    except urllib.error.HTTPError as e:
        raise Exception(f"HTTP {e.code}: Sheet ยังไม่ public — ต้องทำ File → Share → Publish to web")


# ─── Helpers ────────────────────────────────────────────────

def find_col(header, candidates):
    for name in candidates:
        for i, h in enumerate(header):
            if name.lower() in h.lower():
                return i
    return -1


def parse_date(raw):
    """แปลงวันที่หลายรูปแบบ → YYYY-MM-DD"""
    if not raw:
        return ""
    raw = str(raw).strip()
    # Google Sheets date serial (float)
    try:
        n = float(raw)
        # Excel/Sheets serial: days since 1899-12-30
        from datetime import timedelta
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=n)).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y",
                "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def parse_month_label(raw):
    """แปลง month value → 'May 26' format"""
    if not raw:
        return None
    raw = str(raw).strip()
    # Already in 'May 26' format
    if re.match(r'^[A-Za-z]{3}\s+\d{2}$', raw):
        return raw
    # YYYY-MM format
    if re.match(r'^\d{4}-\d{2}$', raw):
        return MONTH_LABEL.get(raw)
    # Date serial
    date_str = parse_date(raw)
    if date_str:
        return MONTH_LABEL.get(date_str[:7])
    return None


def get_val(row, idx):
    if idx < 0 or idx >= len(row):
        return 0.0
    try:
        return float(str(row[idx]).replace(",", "").strip() or 0)
    except (ValueError, TypeError):
        return 0.0


# ─── MK13 ───────────────────────────────────────────────────

def map_channel(ch, mt):
    ch = (ch or "").strip()
    mt = (mt or "").strip()
    if ch == "TikTok":
        if mt == "Affiliate":                   return "TikTok Affi"
        if mt in ("Live", "TikTokLive"):        return "TikTok Live"
        return "TikTok"
    if ch == "Shopee":
        if mt in ("Live", "ShopeeLive"):        return "Shopee Live"
        return "Shopee"
    if ch == "Facebook":
        if mt in ("Salepage", "Shopify"):       return "Shopify"
        return "Facebook"
    if ch == "Instagram":                       return "Instagram"
    if ch in ("LINE", "Line"):                  return "LINE"
    if ch == "YouTube":                         return "YouTube"
    if ch == "หน้าร้าน":                        return "หน้าร้าน"
    if ch == "Bookfair":                        return "Bookfair"
    return None


def read_mk13():
    """อ่าน MK13 → OPD_DAILY (by date+channel) + OPD_PROD_DATA (by product+date+channel)"""
    print("📥 ดึง MK13 จาก Google Sheets...")
    text = fetch_csv(MK13_GID)
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        raise Exception("MK13: ได้ CSV ว่างเปล่า")

    header = [h.strip() for h in rows[0]]
    date_col    = find_col(header, ["วันที่", "Date", "date"])
    ch_col      = find_col(header, ["Sale Channel"])
    method_col  = find_col(header, ["Sale Method"])
    amt_col     = find_col(header, ["ราคาแยกรายการ"])
    free_col    = find_col(header, ["แถม?", "แถม"])
    product_col = find_col(header, ["ชื่อสินค้า"])

    if date_col < 0: raise Exception("MK13: ไม่เจอ column 'วันที่'")
    if ch_col   < 0: raise Exception("MK13: ไม่เจอ column 'Sale Channel'")
    if amt_col  < 0: raise Exception("MK13: ไม่เจอ column 'ราคาแยกรายการ'")

    print(f"   Columns: วันที่={date_col}, Channel={ch_col}, "
          f"Method={method_col}, ราคา={amt_col}, แถม={free_col}, สินค้า={product_col}")

    by_date    = {}   # OPD_DAILY: {date: {channel: amt}}
    prod_data  = {}   # OPD_PROD_DATA: {product: {date: {channel: amt}}}
    ch_set     = set()

    for row in rows[1:]:
        if not row or len(row) <= max(date_col, ch_col, amt_col):
            continue
        # ข้าม แถม / discount rows
        if free_col >= 0 and free_col < len(row):
            if str(row[free_col]).strip().upper() in ("TRUE", "YES", "1", "✓"):
                continue
        # วันที่
        d = parse_date(row[date_col])
        if not d:
            continue
        # ช่องทาง
        ch_raw = row[ch_col].strip() if ch_col < len(row) else ""
        mt_raw = row[method_col].strip() if 0 <= method_col < len(row) else ""
        channel = map_channel(ch_raw, mt_raw)
        if not channel:
            continue
        # ราคา
        try:
            amt = float(str(row[amt_col]).replace(",", "").replace("฿", "").strip())
        except (ValueError, TypeError):
            continue
        if amt <= 0:
            continue

        # ── OPD_DAILY aggregation ──
        ch_set.add(channel)
        if d not in by_date:
            by_date[d] = {"d": d}
        by_date[d][channel] = by_date[d].get(channel, 0) + amt

        # ── OPD_PROD_DATA aggregation ──
        if product_col >= 0 and product_col < len(row):
            prod = row[product_col].strip()
            # ข้าม rows ที่ไม่ใช่ชื่อหนังสือ (discount codes, vouchers)
            if prod and not prod.startswith("*") and not prod.startswith("ส่วนลด"):
                if prod not in prod_data:
                    prod_data[prod] = {}
                if d not in prod_data[prod]:
                    prod_data[prod][d] = {}
                prod_data[prod][d][channel] = prod_data[prod][d].get(channel, 0) + amt

    channels  = sorted(ch_set)
    data_rows = sorted(by_date.values(), key=lambda r: r["d"])
    print(f"   ✅ {len(data_rows)} วัน | {len(prod_data)} สินค้า | channels: {channels}")
    return {"channels": channels, "data": data_rows}, prod_data


# ─── ADSM44 ─────────────────────────────────────────────────

def read_adsm44():
    print("📥 ดึง ADSM44 จาก Google Sheets...")
    text = fetch_csv(ADSM44_GID)
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        raise Exception("ADSM44: ได้ CSV ว่างเปล่า")

    header = [h.strip() for h in rows[0]]
    print(f"   Header: {' | '.join(h for h in header if h)}")

    product_col  = find_col(header, ["Book name", "Book Name", "Product", "product", "เล่ม", "หนังสือ"])
    month_col    = find_col(header, ["Month", "month", "เดือน", "Date", "วันที่"])
    # Ads spend columns — ชื่อจริงใน Sheet: "Ads TT", "Ads FB", "Ads Shopee", "MKT"
    tt_ads_col   = find_col(header, ["Ads TT", "Ads Tiktok Ads", "TT Ads", "TikTok Ads"])
    tt_aff_col   = find_col(header, ["Ads TT Aff", "Ads Tiktok Aff", "TT Aff Ads"])
    fb_msg_col   = find_col(header, ["Ads FB", "Ads FB MSG", "FB Ads", "Facebook Ads"])
    fb_sp_col    = find_col(header, ["Ads FB Salepage", "FB Salepage Ads"])
    sp_ads_col   = find_col(header, ["Ads Shopee", "Shopee Ads"])
    total_ad_col = find_col(header, ["MKT", "Ads Cost", "Total Ads", "_spend"])
    # Revenue columns — ชื่อจริงใน Sheet: "Revenue TT", "Revenue FB", "Revenue Shopee", "Revenue MMS"
    sale_tt_col     = find_col(header, ["Revenue TT", "Sale Tiktok", "TT Rev", "TikTok Revenue"])
    sale_tt_aff_col = find_col(header, ["Revenue TT Aff", "Sale Tiktok Aff", "TT Aff Rev"])
    sale_fb_msg_col = find_col(header, ["Revenue FB", "Sale FB MSG", "FB Rev", "Facebook Revenue"])
    sale_fb_sp_col  = find_col(header, ["Revenue FB Salepage", "Sale FB Salepage", "FB Salepage Rev"])
    sale_sp_col     = find_col(header, ["Revenue Shopee", "Sale Shopee", "Shopee Rev"])
    sale_total_col  = find_col(header, ["Revenue MMS", "Total revenue", "Sale Total", "Revenue Total"])

    print(f"   Columns → product:{product_col} | Ads TT:{tt_ads_col} FB:{fb_msg_col} SP:{sp_ads_col} | Rev TT:{sale_tt_col} FB:{sale_fb_msg_col} SP:{sale_sp_col} | total_rev:{sale_total_col}")

    if product_col < 0:
        print("⚠️  ADSM44: ไม่เจอ column Product/Book name — ข้าม")
        return {}

    cur_month = MONTH_LABEL.get(datetime.now().strftime("%Y-%m"))
    result = {}

    for row in rows[1:]:
        if not row:
            continue
        product = str(row[product_col]).strip() if product_col < len(row) else ""
        if not product:
            continue
        # ลบ prefix [หนังสือ] / [นิยาย] / [...] เพื่อให้ key ตรงกับ GOALS_DATA
        product = re.sub(r'^\[[^\]]+\]\s*', '', product).strip()

        # หาเดือน
        month_label = None
        if month_col >= 0 and month_col < len(row) and row[month_col].strip():
            month_label = parse_month_label(row[month_col])
        if not month_label:
            month_label = cur_month
        if not month_label:
            continue

        if month_label not in result:
            result[month_label] = {}

        tt_spend  = get_val(row, tt_ads_col) + get_val(row, tt_aff_col)
        fb_spend  = get_val(row, fb_msg_col) + get_val(row, fb_sp_col)
        sp_spend  = get_val(row, sp_ads_col)
        tot_spend = (get_val(row, total_ad_col) if total_ad_col >= 0
                     else tt_spend + fb_spend + sp_spend)

        tt_rev  = get_val(row, sale_tt_col)     + get_val(row, sale_tt_aff_col)
        fb_rev  = get_val(row, sale_fb_msg_col) + get_val(row, sale_fb_sp_col)
        sp_rev  = get_val(row, sale_sp_col)
        tot_rev = (get_val(row, sale_total_col) if sale_total_col >= 0
                   else tt_rev + fb_rev + sp_rev)

        # %Ads cost ratio = spend/sale × 100 (per channel)
        def pct(spend, sale):
            if spend == 0 and sale == 0: return None
            return round(spend / sale * 100, 2) if sale > 0 else None

        result[month_label][product] = {
            "TikTok":    pct(tt_spend, tt_rev),
            "Facebook":  pct(fb_spend, fb_rev),
            "Shopee":    pct(sp_spend, sp_rev),
            "_tt_spend": tt_spend,
            "_fb_spend": fb_spend,
            "_sp_spend": sp_spend,
            "_spend":    tot_spend,
            "_tt_sale":  tt_rev,
            "_fb_sale":  fb_rev,
            "_sp_sale":  sp_rev,
            "_rev":      tot_rev,
        }

    print(f"   ✅ {len(result)} เดือน | {list(result.keys())}")
    return result


# ─── Replace variable in HTML (bracket-counting, O(n)) ──────

def replace_var(html, var_name, new_value):
    candidates = [
        f"const {var_name} = ",
        f"var {var_name} = ",
        f"let {var_name} = ",
        f"const {var_name}=",
        f"var {var_name}=",
    ]
    value_start = -1
    for prefix in candidates:
        idx = html.find(prefix)
        if idx >= 0:
            value_start = idx + len(prefix)
            break
    if value_start < 0:
        print(f"   ⚠️  ไม่เจอตัวแปร '{var_name}'")
        return html

    # หา start bracket
    ch = html[value_start]
    if ch not in ('{', '['):
        print(f"   ⚠️  '{var_name}' ไม่ได้เริ่มด้วย {{ หรือ [")
        return html

    open_b  = ch
    close_b = '}' if ch == '{' else ']'
    depth   = 0
    i       = value_start
    in_str  = False
    escape  = False
    str_ch  = None

    while i < len(html):
        c = html[i]
        if escape:
            escape = False
        elif c == '\\' and in_str:
            escape = True
        elif in_str:
            if c == str_ch:
                in_str = False
        elif c in ('"', "'", '`'):
            in_str = True
            str_ch = c
        elif c == open_b:
            depth += 1
        elif c == close_b:
            depth -= 1
            if depth == 0:
                break
        i += 1

    if depth != 0:
        print(f"   ⚠️  '{var_name}' bracket ไม่สมดุล")
        return html

    result = html[:value_start] + new_value + html[i+1:]
    print(f"   ✅ แทนที่ {var_name} สำเร็จ ({len(new_value):,} chars)")
    return result


# ─── Update HTML ────────────────────────────────────────────

def update_html(opd_data, prod_data, adsm44_data, html_path=DASHBOARD_PATH):
    print(f"\n💾 อัปเดต {html_path}...")
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    if len(content) < 100_000:
        raise Exception(f"❌ Safety: ไฟล์ {len(content)} chars — เล็กเกินไป หยุด")

    opd_json      = json.dumps(opd_data,   ensure_ascii=False, separators=(",", ":"))
    prod_json     = json.dumps(prod_data,  ensure_ascii=False, separators=(",", ":"))
    adsm44_json   = json.dumps(adsm44_data, ensure_ascii=False, separators=(",", ":"))

    content = replace_var(content, "OPD_DAILY",      opd_json)
    content = replace_var(content, "OPD_PROD_DATA",  prod_json)
    content = replace_var(content, "ADSM44_PCTADS",  adsm44_json)

    # อัปเดต timestamp การ sync ล่าสุด
    now_ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    content = replace_var(content, "META_PULL_TS", f'"{now_ts}"')

    # อัปเดต version tag
    today = datetime.now().strftime("%Y%m%d")
    content = re.sub(r'v\d{8}[a-z]*', f'v{today}as', content)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"   ✅ บันทึกแล้ว ({len(content):,} chars)")


# ─── Main ───────────────────────────────────────────────────

def sync(html_path=DASHBOARD_PATH):
    print("=" * 60)
    print("  MK13 + ADSM44 Sync — OpenDurian How-to")
    print("=" * 60)
    opd_data, prod_data = read_mk13()
    adsm44_data = read_adsm44()
    update_html(opd_data, prod_data, adsm44_data, html_path)
    print("\n✅ mk13_sync เสร็จสิ้น")


if __name__ == "__main__":
    sync()
