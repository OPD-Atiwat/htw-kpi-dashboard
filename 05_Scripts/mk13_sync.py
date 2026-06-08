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


def fetch_csv_gviz(gid, date_from):
    """ดึง CSV เฉพาะ rows ที่ date >= date_from ผ่าน gviz/tq — ไม่ถูก truncate"""
    # column F = วันที่ (index 5, 0-based → gviz ใช้ A=col0, F=col5)
    tq = f"SELECT * WHERE F >= date '{date_from}'"
    import urllib.parse
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
           f"/gviz/tq?tqx=out:csv&gid={gid}&tq={urllib.parse.quote(tq)}")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.read().decode("utf-8-sig")
    except Exception as e:
        print(f"   ⚠️  gviz fetch failed: {e}")
        return None


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
    mt_l = mt.lower()
    if ch == "TikTok":
        if mt in ("Affiliate", "TikTokAffi"):           return "TikTok Affi"
        if "live" in mt_l or mt == "TikTokLive":        return "TikTok Live"
        return "TikTok"
    if ch == "Shopee":
        if "live" in mt_l or mt == "ShopeeLive":        return "Shopee Live"
        return "Shopee"
    if ch == "Facebook":
        if mt in ("Salepage", "Shopify", "FBSalepage"):  return "Shopify"
        return "Facebook"
    if ch == "Instagram":                       return "Instagram"
    if ch in ("LINE", "Line"):                  return "LINE"
    if ch == "YouTube":                         return "YouTube"
    if ch in ("Lazada", "lazada"):              return "Lazada"
    if ch in ("Web", "web", "Website"):         return "Web"
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

    print(f"   ได้ {len(rows)} rows (รวม header) จาก CSV export")

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

    # ── Phase 1: process main CSV rows ──────────────────────────
    covered_dates = set()  # dates ที่มีข้อมูลจาก main CSV แล้ว
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
        covered_dates.add(d)
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

    # ── Phase 2: เสริม gviz rows สำหรับเดือนปัจจุบัน (แก้ CSV truncation) ──
    today = datetime.today()
    month_start = today.replace(day=1).strftime("%Y-%m-%d")
    print(f"   🔄 ดึง gviz rows เสริมตั้งแต่ {month_start} เพื่อแก้ truncation...")
    gviz_text = fetch_csv_gviz(MK13_GID, month_start)
    if gviz_text:
        gviz_reader = csv.reader(io.StringIO(gviz_text))
        gviz_all = list(gviz_reader)
        added = 0
        skipped_dates = set()
        for row in gviz_all[1:]:  # ข้าม header
            if not row or len(row) <= max(date_col, ch_col, amt_col):
                continue
            if free_col >= 0 and free_col < len(row):
                if str(row[free_col]).strip().upper() in ("TRUE", "YES", "1", "✓"):
                    continue
            d = parse_date(row[date_col])
            if not d:
                continue
            # ข้ามวันที่มีข้อมูลจาก main CSV แล้ว (เพื่อกัน double-count)
            if d in covered_dates:
                skipped_dates.add(d)
                continue
            ch_raw  = row[ch_col].strip() if ch_col < len(row) else ""
            mt_raw  = row[method_col].strip() if 0 <= method_col < len(row) else ""
            channel = map_channel(ch_raw, mt_raw)
            if not channel:
                continue
            try:
                amt = float(str(row[amt_col]).replace(",", "").replace("฿", "").strip())
            except (ValueError, TypeError):
                continue
            if amt <= 0:
                continue
            ch_set.add(channel)
            if d not in by_date:
                by_date[d] = {"d": d}
            by_date[d][channel] = by_date[d].get(channel, 0) + amt
            added += 1
            if product_col >= 0 and product_col < len(row):
                prod = row[product_col].strip()
                if prod and not prod.startswith("*") and not prod.startswith("ส่วนลด"):
                    if prod not in prod_data:
                        prod_data[prod] = {}
                    if d not in prod_data[prod]:
                        prod_data[prod][d] = {}
                    prod_data[prod][d][channel] = prod_data[prod][d].get(channel, 0) + amt
        new_dates = sorted(set(by_date.keys()) - covered_dates)
        print(f"   ✅ gviz เพิ่ม {added} rows | วันใหม่: {new_dates} | skip (already covered): {sorted(skipped_dates)}")
    else:
        print(f"   ⚠️  gviz ล้มเหลว — ใช้ main CSV เท่านั้น ({len(covered_dates)} วัน)")

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
    # Pre-computed % Ads columns (ถ้าSheet คำนวณไว้แล้ว ใช้ตรงๆ เลย — แม่นกว่า recompute)
    pct_total_col   = find_col(header, ["% Ads"])          # col 24 overall %Ads
    pct_fb_col      = find_col(header, ["% Ads FB MSG"])   # col 25 FB %Ads
    pct_tt_ads_col  = find_col(header, ["% Ads Tiktok Ads"])  # col 27
    pct_tt_aff_col  = find_col(header, ["% Ads Tiktok Aff"])  # col 28
    pct_sp_col      = find_col(header, ["% Ads Shopee"])   # col 30

    print(f"   Columns → product:{product_col} | Ads TT:{tt_ads_col} FB:{fb_msg_col} SP:{sp_ads_col} | Rev TT:{sale_tt_col} FB:{sale_fb_msg_col} SP:{sale_sp_col} | total_rev:{sale_total_col} | pct_total:{pct_total_col}")

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

        tt_ads_spend = get_val(row, tt_ads_col)
        tt_aff_spend = get_val(row, tt_aff_col)
        tt_spend  = tt_ads_spend + tt_aff_spend
        fb_spend  = get_val(row, fb_msg_col) + get_val(row, fb_sp_col)
        sp_spend  = get_val(row, sp_ads_col)
        tot_spend = (get_val(row, total_ad_col) if total_ad_col >= 0
                     else tt_spend + fb_spend + sp_spend)

        tt_ads_rev = get_val(row, sale_tt_col)          # TikTok Shop revenue
        tt_aff_rev = get_val(row, sale_tt_aff_col)      # TikTok Affi revenue
        tt_rev     = tt_ads_rev + tt_aff_rev            # รวม (backward compat)
        fb_rev  = get_val(row, sale_fb_msg_col) + get_val(row, sale_fb_sp_col)
        sp_rev  = get_val(row, sale_sp_col)
        tot_rev = (get_val(row, sale_total_col) if sale_total_col >= 0
                   else tt_rev + fb_rev + sp_rev)

        # ── MERGE per-channel: ถ้ามีหลาย row ต่อเล่ม ใช้ max() per field ──
        # Logic: max() ทำงานถูกต้องทุก pattern:
        #   • 1 row/product       → ไม่ถูก trigger เลย ✅
        #   • sub-rows only       → max(0, X) = X เก็บค่าสูงสุดรายช่อง ✅
        #   • total+sub-rows      → max(total, partial) = total ไม่ double-count ✅
        if product in result[month_label]:
            ex = result[month_label][product]
            tt_ads_spend = max(tt_ads_spend, ex["_tt_ads_spend"])
            tt_aff_spend = max(tt_aff_spend, ex["_tt_aff_spend"])
            tt_spend     = max(tt_spend,     ex["_tt_spend"])
            fb_spend     = max(fb_spend,     ex["_fb_spend"])
            sp_spend     = max(sp_spend,     ex["_sp_spend"])
            tt_ads_rev   = max(tt_ads_rev,   ex["_tt_ads_sale"])
            tt_aff_rev   = max(tt_aff_rev,   ex["_tt_aff_sale"])
            fb_rev       = max(fb_rev,       ex["_fb_sale"])
            sp_rev       = max(sp_rev,       ex["_sp_sale"])

        # _spend / _rev: ใช้ MKT / Revenue MMS column จาก sheet เป็นหลัก
        # (ครอบคลุม YouTube, LINE, หน้าร้าน ฯลฯ ที่ per-channel sum ไม่รวม)
        tt_rev    = tt_ads_rev + tt_aff_rev
        _mkt_val  = get_val(row, total_ad_col)  if total_ad_col  >= 0 else 0
        _rev_val  = get_val(row, sale_total_col) if sale_total_col >= 0 else 0
        # merge-safe: ถ้ามีหลาย rows ใช้ max กับค่าก่อนหน้า
        if product in result[month_label]:
            ex = result[month_label][product]
            _mkt_val = max(_mkt_val, ex.get("_spend", 0))
            _rev_val = max(_rev_val, ex.get("_rev",   0))
        tot_spend = _mkt_val  if _mkt_val  > 0 else (tt_spend + fb_spend + sp_spend)
        tot_rev   = _rev_val  if _rev_val  > 0 else (tt_rev   + fb_rev   + sp_rev)

        # Pre-computed % Ads — อ่านตรงจาก Sheet (แม่นกว่า recompute จาก spend/sale)
        # Sheet เก็บเป็น decimal 0-1 → ×100 ถ้า <1, ไม่งั้นใช้ตรง
        def _read_pct(col):
            if col < 0: return None
            v = get_val(row, col)
            if v == 0: return None
            return round(v * 100, 2) if v < 1 else round(v, 2)

        _pct_total  = _read_pct(pct_total_col)
        _pct_fb     = _read_pct(pct_fb_col)
        _pct_tt_ads = _read_pct(pct_tt_ads_col)
        _pct_tt_aff = _read_pct(pct_tt_aff_col)
        _pct_sp     = _read_pct(pct_sp_col)

        # %Ads cost ratio = spend/sale × 100 (per channel)
        def pct(spend, sale):
            if spend == 0 and sale == 0: return None
            return round(spend / sale * 100, 2) if sale > 0 else None

        result[month_label][product] = {
            "TikTok":         pct(tt_ads_spend, tt_ads_rev),  # TT Shop เท่านั้น
            "TikTokAff":      pct(tt_aff_spend, tt_aff_rev),  # TT Affi เท่านั้น
            "Facebook":       pct(fb_spend,     fb_rev),
            "Shopee":         pct(sp_spend,     sp_rev),
            "_tt_spend":      tt_spend,           # TikTok Shop + Affi รวม
            "_tt_ads_spend":  tt_ads_spend,        # TikTok Shop เท่านั้น
            "_tt_aff_spend":  tt_aff_spend,        # TikTok Affi commission
            "_fb_spend":      fb_spend,
            "_sp_spend":      sp_spend,
            "_spend":         tot_spend,           # คำนวณจาก TT+FB+SP เสมอ
            "_tt_sale":       tt_rev,              # TT รวม (backward compat)
            "_tt_ads_sale":   tt_ads_rev,          # TikTok Shop revenue
            "_tt_aff_sale":   tt_aff_rev,          # TikTok Affi revenue
            "_fb_sale":       fb_rev,
            "_sp_sale":       sp_rev,
            "_rev":           tot_rev,             # คำนวณจาก TT+FB+SP เสมอ
            # Pre-computed % จาก Sheet (ใช้แทน recompute เมื่อมีค่า)
            "_pct_total":    _pct_total,   # % Ads รวม (col 24)
            "_pct_fb":       _pct_fb,      # % Ads FB MSG (col 25)
            "_pct_tt_ads":   _pct_tt_ads,  # % Ads Tiktok Ads (col 27)
            "_pct_tt_aff":   _pct_tt_aff,  # % Ads Tiktok Aff (col 28)
            "_pct_sp":       _pct_sp,      # % Ads Shopee (col 30)
        }

    # Summary log
    for mo, prods in result.items():
        with_spend = sum(1 for d in prods.values() if (d.get("_spend") or 0) > 0)
        print(f"   {mo}: {len(prods)} products, {with_spend} with spend>0")
    print(f"   ✅ {len(result)} เดือน | {list(result.keys())}")
    return result


# ─── Replace variable in HTML (bracket-counting, O(n)) ──────

def replace_var(html, var_name, new_value):
    candidates = [
        f"const {var_name} = ",
        f"var {var_name} = ",
        f"let {var_name} = ",
        f"const {var_name} =",
        f"var {var_name} =",
        f"let {var_name} =",
        f"const {var_name}=",
        f"var {var_name}=",
        f"let {var_name}=",
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

    # หา start bracket หรือ string quote
    ch = html[value_start]

    # ── กรณีที่ค่าเป็น string ที่ขึ้นต้นด้วย " ' ` ──────────────────
    if ch in ('"', "'", '`'):
        quote = ch
        i = value_start + 1
        while i < len(html):
            c = html[i]
            if c == '\\':          # escape → ข้ามตัวถัดไป
                i += 2
                continue
            if c == quote:         # ปิด string
                break
            i += 1
        result = html[:value_start] + new_value + html[i+1:]
        print(f"   ✅ แทนที่ {var_name} สำเร็จ ({len(new_value):,} chars)")
        return result

    # ── กรณีที่ค่าเป็น object/array ────────────────────────────────
    if ch not in ('{', '['):
        print(f"   ⚠️  '{var_name}' ไม่ได้เริ่มด้วย {{ หรือ [ หรือ quote")
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
