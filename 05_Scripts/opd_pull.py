#!/usr/bin/env python3
"""
OPD Daily Puller — OpenDurian How-to
ดึงข้อมูล Affiliate Creator จาก OPD Daily Google Sheet แล้วอัปเดต Dashboard

Flow:
  1. Download MK13 tab → filter TikTok Affi orders (แถม?=FALSE)
     → group by Creator + Month → revenue, orders, top products
  2. Download %Ads tab → group by Product + Month
     → Ads Cost, TikTok Aff revenue, % ROAS
  3. อัปเดต AFFILIATE_DATA ใน dashboard HTML
"""

import csv
import json
import re
import os
import io
import requests
from datetime import datetime
from collections import defaultdict

# ============================================================
# CONFIG — แก้ตรงนี้
# ============================================================
OPD_SHEET_ID         = "1qYdwXuCHDHeHN6a8vU_RVFBYT5MuYq-8K5AMK3cP4DM"
MK13_GID             = "964123706"
ADS_PCT_GID          = "1860322120"
PRODUCT_MONITOR_GID  = "2052764970"   # "Goal by Product" tab

DASHBOARD_PATH = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"
# ============================================================

SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}"

# คอลัมน์ใน MK13 (0-indexed)
MK13_COL = {
    "date":         5,
    "product":      13,   # ชื่อสินค้า (ACC) — ชื่อมาตรฐาน
    "is_gift":      16,   # แถม? — ต้องเป็น FALSE / เว้นว่าง
    "sale_channel": 21,   # Sale Channel ("TikTok", "Facebook" ฯลฯ)
    "sale_method":  22,   # Sale Method ("Affiliate") ← ตัวกรองหลัก
    "creator":      23,   # Creator Name (affiliate username เช่น unsexnn)
    "revenue":      27,   # ยอดเงินรับ
}

# คอลัมน์ใน %Ads (0-indexed)
ADS_COL = {
    "product":           0,
    "ads_cost":          2,
    "ads_fb_msg":        3,   # Ads FB MSG
    "ads_fb_salepage":   4,   # Ads FB Salepage
    "ads_tiktok_ads":    5,
    "ads_tiktok_aff":    6,
    "ads_shopee":        8,   # Ads Shopee
    "sale_total":        13,
    "sale_tiktok":       16,
    "pct_ads":           24,  # % Ads หรือ ROAS field
    "date":              35,
}

# MONTH_LABELS ถูกแทนด้วย dynamic function get_month_label() ด้านล่าง


def download_csv(url, timeout=120, retries=3):
    """
    ดาวน์โหลด CSV พร้อม retry logic
    - timeout=120  รองรับไฟล์ขนาดใหญ่ (MK13 = 15,000+ rows)
    - retries=3    ลองใหม่ถ้า timeout / connection error
    """
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(url, timeout=timeout, stream=True)
            r.raise_for_status()
            # อ่านแบบ streaming ป้องกัน memory spike
            chunks = []
            for chunk in r.iter_content(chunk_size=1024 * 256):
                chunks.append(chunk)
            raw = b"".join(chunks)
            text = raw.decode("utf-8-sig")
            return list(csv.reader(io.StringIO(text)))
        except requests.exceptions.Timeout:
            print(f"  ⚠️  Timeout (ครั้งที่ {attempt}/{retries})", flush=True)
            if attempt == retries:
                print(f"  ❌ ดาวน์โหลดไม่ได้หลังลอง {retries} ครั้ง")
                return []
        except Exception as e:
            print(f"  ⚠️  ดาวน์โหลดไม่ได้: {e}")
            return []
    return []


def clean(s):
    return s.strip().strip('"').strip()


def flt(s):
    try:
        return float(re.sub(r'[^\d.\-]', '', s))
    except:
        return 0.0


def parse_date_mk13(raw):
    """
    แปลงวันที่จาก MK13 เช่น '2026-04-01', '1/4/2026', '01/04/26'
    → คืนค่า (year, month) tuple หรือ None
    """
    raw = clean(raw)
    if not raw:
        return None
    # ลอง YYYY-MM-DD ก่อน
    m = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', raw)
    if m:
        return int(m.group(1)), int(m.group(2))
    # ลอง D/M/YYYY หรือ DD/MM/YYYY
    m = re.match(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})', raw)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        return y, mo
    return None


def parse_date_ads(raw):
    """แปลงวันที่จาก %Ads tab — รูปแบบเดียวกับ MK13"""
    return parse_date_mk13(raw)


def parse_full_date(raw):
    """
    แปลงวันที่จาก MK13 → คืนค่า 'YYYY-MM-DD' string หรือ None
    รองรับ: '2026-04-01', '1/4/2026', '01/04/26', '2026/04/01'
    """
    raw = clean(raw)
    if not raw:
        return None
    # YYYY-MM-DD หรือ YYYY/MM/DD
    m = re.match(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', raw)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{y:04d}-{mo:02d}-{d:02d}"
    # D/M/YYYY หรือ DD/MM/YYYY หรือ D-M-YY
    m = re.match(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})', raw)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        return f"{y:04d}-{mo:02d}-{d:02d}"
    return None


# ── Channel name normalization (MK13 col21 → OPD_DAILY key) ──────────────────
CHANNEL_NORMALIZE = {
    "tiktok":        "TikTok",
    "tiktok live":   "TikTok Live",
    "tiktok affi":   "TikTok Affi",
    "tiktok affi.":  "TikTok Affi",
    "shopee":        "Shopee",
    "shopee live":   "Shopee",        # รวมเข้า Shopee
    "facebook":      "Facebook",
    "fb":            "Facebook",
    "instagram":     "Instagram",
    "ig":            "Instagram",
    "lazada":        "Lazada",
    "line":          "LINE",
    "web":           "Web",
    "shopify":       "Shopify",
    "bookfair":      "Bookfair",
    "book fair":     "Bookfair",
    "หน้าร้าน":      "หน้าร้าน",
    "youtube":       "YouTube",
    "yt":            "YouTube",
}

def normalize_channel(sale_channel, sale_method):
    """
    แปลง col21 + col22 → OPD_DAILY channel key
    TikTok + Affiliate → 'TikTok Affi'
    """
    ch  = sale_channel.strip()
    chL = ch.lower()
    mL  = sale_method.strip().lower()

    # ถ้า col21 = "TikTok" (ไม่มี affi) และ col22 = "Affiliate" → TikTok Affi
    if "affiliate" in mL and chL == "tiktok":
        return "TikTok Affi"

    # Lookup ใน CHANNEL_NORMALIZE
    if chL in CHANNEL_NORMALIZE:
        return CHANNEL_NORMALIZE[chL]

    # คืนค่าเดิมถ้าไม่รู้จัก (capitalized)
    return ch if ch else None


def get_month_label(ym_tuple):
    """Auto-generate label จาก (year, month) tuple — รองรับทุกปีโดยไม่ hardcode"""
    if ym_tuple is None:
        return None
    try:
        from datetime import datetime as _dt
        dt = _dt(ym_tuple[0], ym_tuple[1], 1)
        return dt.strftime("%b") + dt.strftime("%y")  # e.g. "May26"
    except Exception:
        return None


def parse_mk13(rows):
    """
    อ่าน rows จาก MK13 tab
    → กรองเฉพาะ Sale Channel = 'TikTok Affi' และ แถม? ≠ 'TRUE'
    → คืนค่า tuple:
        creators      = { creator: { month: {revenue, orders, products:{name:rev}} } }
        daily_creators = { date_str: { creator: {revenue, orders} } }

    Dedup: col0 = Order ID — ถ้า 1 order มีหลาย items (rows) แต่ col27 = ยอดรวม order
    → นับ revenue ครั้งเดียวต่อ (order_id, date, creator)
    """
    if not rows:
        return {}, {}

    data_start = 1

    creators = defaultdict(lambda: defaultdict(lambda: {
        "revenue": 0.0,
        "orders": 0,
        "products": defaultdict(float)
    }))

    # daily aggregate: date_str → creator → {revenue, orders}
    daily_creators = defaultdict(lambda: defaultdict(lambda: {
        "revenue": 0.0,
        "orders": 0
    }))

    seen_orders = set()   # (order_id, date_str, creator) สำหรับ dedup
    skipped        = 0
    deduped        = 0
    col0_samples   = []   # เก็บตัวอย่าง col0 สำหรับ debug
    col0_empty     = 0

    for row in rows[data_start:]:
        max_col = max(MK13_COL.values())
        if len(row) <= max_col:
            skipped += 1
            continue

        order_id    = clean(row[0]) if len(row) > 0 else ""
        sale_method = clean(row[MK13_COL["sale_method"]])
        is_gift     = clean(row[MK13_COL["is_gift"]]).upper()
        creator     = clean(row[MK13_COL["creator"]])
        rev_raw     = clean(row[MK13_COL["revenue"]])
        date_raw    = clean(row[MK13_COL["date"]])
        prod_raw    = clean(row[MK13_COL["product"]])

        if "affiliate" not in sale_method.lower():
            continue
        if is_gift in ("TRUE", "1", "YES", "Y"):
            continue
        if not creator or creator in ("-", "#N/A", ""):
            continue

        ym = parse_date_mk13(date_raw)
        month = get_month_label(ym)
        if not month:
            continue

        date_str = parse_full_date(date_raw)  # 'YYYY-MM-DD' หรือ None

        # เก็บตัวอย่าง col0 สำหรับ debug (10 รายการแรก)
        if len(col0_samples) < 10:
            col0_samples.append(repr(order_id) if order_id else "(empty)")
        if not order_id:
            col0_empty += 1

        revenue = flt(rev_raw)
        product = prod_raw if prod_raw else "Unknown"

        # Dedup: ถ้ามี order_id → นับ revenue ครั้งเดียวต่อ order+date+creator
        if order_id and date_str:
            key = (order_id, date_str, creator)
            if key in seen_orders:
                deduped += 1
                # ยังนับ product เพื่อรู้ว่า order นี้มีอะไรบ้าง แต่ไม่เพิ่ม revenue
                creators[creator][month]["products"][product]  # just touch it (no add)
                continue
            seen_orders.add(key)

        # Monthly aggregate
        d = creators[creator][month]
        d["revenue"] += revenue
        d["orders"]  += 1
        d["products"][product] += revenue

        # Daily aggregate
        if date_str:
            dd = daily_creators[date_str][creator]
            dd["revenue"] += revenue
            dd["orders"]  += 1

    if skipped > 0:
        print(f"  (ข้าม {skipped} rows ที่คอลัมน์ไม่ครบ)")

    # ── Diagnostic report ────────────────────────────────────
    total_aff = sum(
        d["orders"] for cr in creators.values() for d in cr.values()
    )
    total_rev = sum(
        d["revenue"] for cr in creators.values() for d in cr.values()
    )
    print(f"  [parse_mk13] orders(deduped): {total_aff}  revenue: ฿{total_rev:,.0f}")
    print(f"  [parse_mk13] deduped rows: {deduped}  col0_empty: {col0_empty}")
    print(f"  [parse_mk13] col0 samples: {col0_samples[:5]}")

    return creators, daily_creators


def parse_ads_pct(rows):
    """
    อ่าน rows จาก %Ads tab
    → คืนค่า dict: { month: { product: {ads_cost, ads_tiktok_aff, sale_tiktok, pct_ads} } }
    """
    if not rows:
        return {}

    data_start = 1
    products = defaultdict(lambda: defaultdict(lambda: {
        "ads_cost": 0.0,
        "ads_fb": 0.0,
        "ads_tiktok_ads": 0.0,
        "ads_tiktok_aff": 0.0,
        "ads_shopee": 0.0,
        "sale_total": 0.0,
        "sale_tiktok": 0.0,
    }))

    max_col = max(ADS_COL.values())
    for row in rows[data_start:]:
        if len(row) <= max_col:
            continue

        product  = clean(row[ADS_COL["product"]])
        date_raw = clean(row[ADS_COL["date"]])

        if not product or product in ("-", "#N/A", ""):
            continue

        ym = parse_date_ads(date_raw)
        month = get_month_label(ym)
        if not month:
            continue

        d = products[month][product]
        d["ads_cost"]        += flt(row[ADS_COL["ads_cost"]])
        d["ads_fb"]          += flt(row[ADS_COL["ads_fb_msg"]]) + flt(row[ADS_COL["ads_fb_salepage"]])
        d["ads_tiktok_ads"]  += flt(row[ADS_COL["ads_tiktok_ads"]])
        d["ads_tiktok_aff"]  += flt(row[ADS_COL["ads_tiktok_aff"]])
        d["ads_shopee"]      += flt(row[ADS_COL["ads_shopee"]])
        d["sale_total"]      += flt(row[ADS_COL["sale_total"]])
        d["sale_tiktok"]     += flt(row[ADS_COL["sale_tiktok"]])

    return products


def build_affiliate_data(creators, ads_products, daily_creators=None):
    """
    Combine creator data + ads data → AFFILIATE_DATA JS structure
    """
    # รวบรวม months ที่มีข้อมูล
    all_months = set()
    for cr_months in creators.values():
        all_months.update(cr_months.keys())
    for month in ads_products.keys():
        all_months.add(month)

    # เรียงเดือนตามลำดับเวลา (dynamic — ไม่ hardcode เพื่อให้เดือนใหม่ทุกเดือนโผล่เองอัตโนมัติ)
    def _mkey(m):
        try:
            return datetime.strptime(m, "%b%y")
        except Exception:
            return datetime(2000, 1, 1)
    months_sorted = sorted(all_months, key=_mkey)

    # สร้าง creator summary
    creator_data = {}
    for creator, months_dict in creators.items():
        creator_data[creator] = {}
        total_rev = 0.0
        total_ord = 0
        for month in months_sorted:
            if month not in months_dict:
                continue
            d = months_dict[month]
            rev = round(d["revenue"], 2)
            ord_ = d["orders"]
            total_rev += rev
            total_ord += ord_

            # top 5 products by revenue
            top_prods = sorted(d["products"].items(), key=lambda x: x[1], reverse=True)[:5]
            creator_data[creator][month] = {
                "revenue": rev,
                "orders": ord_,
                "top_products": [{"name": p, "revenue": round(r, 2)} for p, r in top_prods]
            }
        creator_data[creator]["__total"] = {
            "revenue": round(total_rev, 2),
            "orders": total_ord
        }

    # เรียง creators ตาม total revenue มากสุด
    creator_data_sorted = dict(
        sorted(creator_data.items(),
               key=lambda x: x[1].get("__total", {}).get("revenue", 0),
               reverse=True)
    )

    # สร้าง products_ads summary
    products_ads = {}
    for month, prod_dict in ads_products.items():
        products_ads[month] = {}
        for product, d in prod_dict.items():
            ads_aff = d["ads_tiktok_aff"]
            sale_tk = d["sale_tiktok"]
            roas = round(sale_tk / ads_aff, 4) if ads_aff > 0 else 0
            products_ads[month][product] = {
                "ads_cost":       round(d["ads_cost"], 2),
                "ads_tiktok_ads": round(d["ads_tiktok_ads"], 2),
                "ads_tiktok_aff": round(ads_aff, 2),
                "sale_total":     round(d["sale_total"], 2),
                "sale_tiktok":    round(sale_tk, 2),
                "roas_tiktok_aff": roas,
            }

    # สร้าง daily field: { date_str: { creator: {revenue, orders} } }
    daily_data = {}
    if daily_creators:
        for date_str, cr_dict in sorted(daily_creators.items()):
            daily_data[date_str] = {
                cr: {"revenue": round(v["revenue"], 2), "orders": v["orders"]}
                for cr, v in cr_dict.items()
                if v["revenue"] > 0
            }

    return {
        "months": months_sorted,
        "creators": creator_data_sorted,
        "products_ads": products_ads,
        "daily": daily_data,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def parse_opd_daily(rows):
    """
    อ่าน MK13 rows → group by date + channel
    → คืนค่า dict: { "YYYY-MM-DD": { channel: revenue, channel+"_q": qty, ... } }
    ตัดของแถม (is_gift=TRUE) ออก

    Dedup logic: col0 = Order ID (Shopify order name เช่น #1234)
    แต่ละ order อาจมีหลาย rows (หลายเล่ม) แต่ col27 = ยอดรวมทั้ง order
    → นับ revenue ครั้งเดียวต่อ (order_id, date, channel) เพื่อไม่ให้นับซ้ำ
    """
    if not rows:
        return {}

    daily        = defaultdict(lambda: defaultdict(float))
    seen         = set()   # (order_id, date_str, channel) ที่นับไปแล้ว

    max_col    = max(MK13_COL.values())
    skipped    = 0
    deduped    = 0
    col0_empty = 0
    col0_samp  = []   # debug samples
    ch_rev     = defaultdict(float)   # revenue totals per channel (for report)

    for row in rows[1:]:
        if len(row) <= max_col:
            skipped += 1
            continue

        order_id   = clean(row[0]) if len(row) > 0 else ""   # col0 = Shopify Order ID
        date_raw   = clean(row[MK13_COL["date"]])
        sale_ch    = clean(row[MK13_COL["sale_channel"]])
        sale_meth  = clean(row[MK13_COL["sale_method"]])
        is_gift    = clean(row[MK13_COL["is_gift"]]).upper()
        rev_raw    = clean(row[MK13_COL["revenue"]])

        # ตัดของแถม
        if is_gift in ("TRUE", "1", "YES", "Y"):
            continue

        date_str = parse_full_date(date_raw)
        if not date_str:
            continue

        channel = normalize_channel(sale_ch, sale_meth)
        if not channel:
            continue

        rev = flt(rev_raw)

        # เก็บตัวอย่าง col0 สำหรับ debug
        if len(col0_samp) < 10:
            col0_samp.append(repr(order_id) if order_id else "(empty)")
        if not order_id:
            col0_empty += 1

        # Dedup: ถ้ามี order_id → นับ revenue แค่ครั้งเดียวต่อ order+channel+date
        if order_id:
            key = (order_id, date_str, channel)
            if key in seen:
                # row นี้เป็น item ที่ 2+ ใน order เดิม — นับ qty เพิ่มแต่ไม่นับ revenue ซ้ำ
                daily[date_str][channel + "_q"] += 1.0
                deduped += 1
                continue
            seen.add(key)

        daily[date_str][channel]          += rev
        daily[date_str][channel + "_q"]   += 1.0
        ch_rev[channel]                   += rev

    if skipped > 0:
        print(f"  (OPD_DAILY: ข้าม {skipped} rows ที่คอลัมน์ไม่ครบ)")

    # ── Diagnostic report ────────────────────────────────────
    total_rev = sum(ch_rev.values())
    print(f"  [OPD_DAILY] deduped rows: {deduped}  col0_empty: {col0_empty}")
    print(f"  [OPD_DAILY] col0 samples: {col0_samp[:5]}")
    print(f"  [OPD_DAILY] revenue by channel (total ฿{total_rev:,.0f}):")
    for ch in sorted(ch_rev, key=lambda c: -ch_rev[c]):
        print(f"    {ch:<20} ฿{ch_rev[ch]:>12,.0f}")

    return dict(daily)


def build_opd_daily_data(daily_dict):
    """
    แปลง daily_dict → OPD_DAILY JS structure
    { "channels": [...], "data": [ {d:"YYYY-MM-DD", ch:rev, ch_q:qty, ...}, ... ] }
    """
    # รวบรวม channels ที่ปรากฏ (ไม่รวม _q)
    channel_set = set()
    for day_data in daily_dict.values():
        for k in day_data.keys():
            if not k.endswith("_q"):
                channel_set.add(k)

    # เรียงลำดับ channels ตาม preferred order
    PREFERRED_ORDER = [
        "TikTok","TikTok Live","TikTok Affi",
        "Shopee","Facebook","หน้าร้าน",
        "Shopify","Instagram","Lazada","LINE","Web","Bookfair","YouTube"
    ]
    channels = [c for c in PREFERRED_ORDER if c in channel_set]
    channels += sorted([c for c in channel_set if c not in PREFERRED_ORDER])

    # เรียง dates
    sorted_dates = sorted(daily_dict.keys())

    data_rows = []
    for date_str in sorted_dates:
        day = daily_dict[date_str]
        row = {"d": date_str}
        for ch in channels:
            rev = day.get(ch, 0)
            qty = int(day.get(ch + "_q", 0))
            if rev > 0 or qty > 0:
                row[ch]          = round(rev)
                row[ch + "_q"]   = qty
        data_rows.append(row)

    # ── Bookfair hardcode overlay (งานสัปดาห์หนังสือ BIBT 2026: 26 มี.ค. – 6 เม.ย.) ──
    # ยอด Bookfair ถูกบันทึกเป็น หน้าร้าน ใน MK13 → ดึงออกมาแยก channel
    BOOKFAIR_OVERLAY = {
        "2026-03-26": 36621, "2026-03-27": 31998, "2026-03-28": 74804,
        "2026-03-29": 65451, "2026-03-30": 37310, "2026-03-31": 34158,
        "2026-04-01": 32044, "2026-04-02": 31280, "2026-04-03": 28885,
        "2026-04-04": 59935, "2026-04-05": 58089, "2026-04-06": 45985,
    }
    for row in data_rows:
        bf = BOOKFAIR_OVERLAY.get(row.get("d", ""), 0)
        if bf > 0:
            row["Bookfair"] = bf
            if "หน้าร้าน" in row:
                row["หน้าร้าน"] = max(0, row["หน้าร้าน"] - bf)
    if any(BOOKFAIR_OVERLAY.get(r.get("d",""),0) > 0 for r in data_rows):
        if "Bookfair" not in channels:
            channels.insert(channels.index("หน้าร้าน") + 1, "Bookfair")

    return {
        "channels":   channels,
        "data":       data_rows,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def update_dashboard_opd_daily(opd_data, html_path):
    """
    อัปเดต OPD_DAILY ใน dashboard HTML
    แทนที่ var OPD_DAILY = {...}; ด้วยข้อมูลใหม่
    """
    if not os.path.exists(html_path):
        print(f"⚠️  ไม่พบ dashboard: {html_path}")
        return

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    new_js = json.dumps(opd_data, ensure_ascii=False, separators=(",", ":"))

    match = re.search(r'var OPD_DAILY\s*=\s*(\{.*?\});', content, re.DOTALL)
    if match:
        new_content = content[:match.start(1)] + new_js + content[match.end(1):]
        print(f"✅ อัปเดต OPD_DAILY แล้ว ({len(opd_data['data'])} วัน, {len(opd_data['channels'])} channels)")
    else:
        print("⚠️  ไม่พบ var OPD_DAILY ใน dashboard — ข้าม")
        return

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_content)


def update_dashboard_affiliate(affiliate_data, html_path):
    """
    อัปเดต AFFILIATE_DATA ใน dashboard HTML
    ถ้ายังไม่มีให้เพิ่มหลัง CREATOR_META_APR26
    """
    if not os.path.exists(html_path):
        print(f"⚠️  ไม่พบ dashboard: {html_path}")
        return

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    new_js = json.dumps(affiliate_data, ensure_ascii=False, separators=(",", ":"))

    # ถ้ามี AFFILIATE_DATA อยู่แล้ว → แทนที่
    match = re.search(r'var AFFILIATE_DATA\s*=\s*(\{.*?\});', content, re.DOTALL)
    if match:
        new_content = content[:match.start(1)] + new_js + content[match.end(1):]
        print(f"✅ อัปเดต AFFILIATE_DATA แล้ว")
    else:
        # เพิ่มใหม่หลัง CREATOR_META_APR26 line
        marker = re.search(r'(const CREATOR_META_APR26\s*=.*?;)', content, re.DOTALL)
        if marker:
            insert_pos = marker.end()
            insert_str = f'\n\nvar AFFILIATE_DATA = {new_js};'
            new_content = content[:insert_pos] + insert_str + content[insert_pos:]
            print(f"✅ เพิ่ม AFFILIATE_DATA ใหม่ใน dashboard แล้ว")
        else:
            # fallback: เพิ่มก่อน </script> tag แรก
            script_end = content.find('</script>')
            if script_end == -1:
                print("⚠️  ไม่พบตำแหน่งใส่ AFFILIATE_DATA")
                return
            insert_str = f'\nvar AFFILIATE_DATA = {new_js};\n'
            new_content = content[:script_end] + insert_str + content[script_end:]
            print(f"✅ เพิ่ม AFFILIATE_DATA (fallback) แล้ว")

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    n_creators = len([k for k in affiliate_data["creators"] if not k.startswith("__")])
    print(f"   Creators: {n_creators}, Months: {affiliate_data['months']}")


# ─── Product Monitor → GOALS_DATA ──────────────────────────

# Month keys ที่ dashboard รู้จัก (ตรงกับ Month_Key ใน CSV)
_VALID_MONTH_KEYS = {"Jan 26","Feb 26","Mar 26","Apr 26","May 26","Jun 26",
                     "Jul 26","Aug 26","Sep 26","Oct 26","Nov 26","Dec 26"}

def parse_product_monitor(rows, ads_lookup=None):
    """
    Parse Product Monitor CSV (GID=2052764970) → GOALS_DATA format

    Column mapping (0-indexed, เหมือนกันทุกเดือน):
      col 6  = Status
      col 7  = Product Name
      col 8  = Month_Key  (e.g. "Apr 26")
      col 18 = Total Actual (Online+BF+หน้าร้าน) — ใช้เป็น actual
      col 20 = Online Goal — ใช้เป็น goal (excl. Consign, BF)
      col 21 = Online Actual (excl. หน้าร้าน) — ใช้คำนวณ Other channel
      col 15 = หน้าร้าน Actual (Mar/Apr: มีค่า, Jan/Feb: ว่าง)
      col 41 = Actual TikTok (Brand+Affi+Live combined)  ┐
      col 43 = Actual Shopee (all Shopee)                ├ Jan ว่าง, Feb-Apr มีค่า
      col 45 = Actual Facebook                           │
      col 47 = Shopify                                   ┘
      col 49 = % Ads Online (รวม)
      col 50 = % Ads TikTok (pct_ads.TikTok)
      col 51 = % Ads Shopee (pct_ads.Shopee)
      col 52 = % Ads Facebook (pct_ads.Facebook)
      "Other" = col21 - (col41+col43+col45+col47) = Instagram+Lazada+LINE+Web+YouTube รวม

    ตัว filter product rows:
      - col 8 ต้องเป็น month key ที่รู้จัก
      - col 7 ต้องไม่ว่าง และไม่ใช่ "Total"
    """
    goals = {}

    for row in rows:
        if len(row) < 22:
            continue

        product = clean(row[7])
        month   = clean(row[8])

        if month not in _VALID_MONTH_KEYS:
            continue
        if not product or product == "Total":
            continue

        status     = clean(row[6]) or "Critical"
        goal       = flt(row[20]) if len(row) > 20 else 0.0  # Online Goal (col20)
        actual     = flt(row[18]) if len(row) > 18 else 0.0  # Total Actual incl. หน้าร้าน
        online_act = flt(row[21]) if len(row) > 21 else 0.0  # Online Actual (excl. หน้าร้าน)
        achieve    = round(actual / goal * 100, 1) if goal > 0 else 0.0

        tt  = flt(row[41]) if len(row) > 41 else 0.0
        sp  = flt(row[43]) if len(row) > 43 else 0.0
        fb  = flt(row[45]) if len(row) > 45 else 0.0
        sfy = flt(row[47]) if len(row) > 47 else 0.0
        hr  = flt(row[15]) if len(row) > 15 else 0.0   # หน้าร้าน

        # "Other" = ช่องทางย่อยที่ไม่มีคอลัมน์แยก (Instagram, Lazada, LINE, Web, YouTube)
        # = Online Actual − (TikTok + Shopee + FB + Shopify)
        dig_sum = tt + sp + fb + sfy
        other = round(online_act - dig_sum, 0) if online_act > dig_sum else 0.0

        channels = {}
        if tt    > 0: channels["TikTok"]   = int(tt)
        if sp    > 0: channels["Shopee"]   = int(sp)
        if fb    > 0: channels["Facebook"] = int(fb)
        if sfy   > 0: channels["Shopify"]  = int(sfy)
        if hr    > 0: channels["หน้าร้าน"] = int(hr)
        if other > 0: channels["Other"]    = int(other)

        # pct_ads — คำนวณจาก ADSM44 lookup (ถูกต้อง) หรือ fallback ค่า pre-computed (ผิด)
        # month key ใน ads_lookup ใช้รูปแบบ "Apr26" (ไม่มีช่องว่าง)
        # month key ใน Product Monitor ใช้รูปแบบ "Apr 26" (มีช่องว่าง)
        ads_month_key = month.replace(" ", "")  # "Apr 26" → "Apr26"
        pct_tt = pct_sp = pct_fb = 0.0
        if ads_lookup and ads_month_key in ads_lookup and product in ads_lookup[ads_month_key]:
            d_ads = ads_lookup[ads_month_key][product]
            sale = d_ads.get("sale_total", 0.0)
            if sale > 0:
                pct_tt = round((d_ads.get("ads_tiktok_ads", 0) + d_ads.get("ads_tiktok_aff", 0)) / sale * 100, 2)
                pct_sp = round(d_ads.get("ads_shopee", 0) / sale * 100, 2)
                pct_fb = round(d_ads.get("ads_fb", 0) / sale * 100, 2)
        else:
            # fallback: ค่า pre-computed จาก Product Monitor (อาจผิดสำหรับ TikTok Aff)
            pct_tt = round(flt(row[50]), 2) if len(row) > 50 else 0.0
            pct_sp = round(flt(row[51]), 2) if len(row) > 51 else 0.0
            pct_fb = round(flt(row[52]), 2) if len(row) > 52 else 0.0
        pct_ads = {"TikTok": pct_tt, "Shopee": pct_sp, "Facebook": pct_fb}

        if month not in goals:
            goals[month] = {}

        # dedup: ถ้า product ซ้ำ ให้ keep entry ที่ actual สูงกว่า
        existing = goals[month].get(product)
        if existing and int(actual) <= existing.get("actual", 0):
            continue

        goals[month][product] = {
            "product": product,
            "status":  status,
            "goal":    int(goal),
            "actual":  int(actual),
            "achieve": achieve,
            "channels": channels,
            "pct_ads": pct_ads,
        }

    # แปลง dict → list (backward compat)
    return {month: list(prods.values()) for month, prods in goals.items()}


def update_dashboard_goals(goals_data, html_path):
    """Inject GOALS_DATA ใหม่ลง index.html — ใช้ bracket-counting แทน regex"""
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    # หา var GOALS_DATA = { ... }; โดยนับ bracket depth (ไม่ใช้ non-greedy regex)
    prefix = "var GOALS_DATA = "
    idx = content.find(prefix)
    if idx == -1:
        print("   ⚠️  ไม่พบ var GOALS_DATA ใน HTML — ข้าม")
        return

    start = idx + len(prefix)  # ตำแหน่งที่ { เริ่มต้น
    if content[start] != "{":
        print("   ⚠️  GOALS_DATA format ไม่ถูกต้อง — ข้าม")
        return

    # นับ bracket depth จนหา closing }
    depth = 0
    i = start
    in_str = False
    str_char = None
    end = -1
    while i < len(content):
        c = content[i]
        if in_str:
            if c == "\\":
                i += 1  # skip escaped char
            elif c == str_char:
                in_str = False
        else:
            if c in ('"', "'", "`"):
                in_str = True
                str_char = c
            elif c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    end = i + 1  # position หลัง closing }
                    break
        i += 1

    if end == -1:
        print("   ⚠️  หา closing } ของ GOALS_DATA ไม่เจอ — ข้าม")
        return

    # ป้องกัน overwrite goal และ actual ด้วย 0
    try:
        existing = json.loads(content[start:end])
        for month, prods in list(goals_data.items()):
            exist_map = {p["product"]: p for p in existing.get(month, [])
                         if p.get("actual", 0) > 0 or p.get("goal", 0) > 0}
            if exist_map:
                merged = []
                for p in prods:
                    ex = exist_map.get(p["product"])
                    if ex:
                        # คง goal ถ้า new=0 แต่ existing>0
                        if p.get("goal", 0) == 0 and ex.get("goal", 0) > 0:
                            p["goal"] = ex["goal"]
                        # คง actual ถ้า new=0 หรือ new ลดจาก existing >50% (ผิดปกติ)
                        new_actual = p.get("actual", 0)
                        ex_actual  = ex.get("actual", 0)
                        if new_actual == 0 and ex_actual > 0:
                            p["actual"] = ex_actual
                        elif ex_actual > 0 and 0 < new_actual < ex_actual * 0.5:
                            print(f"   ⚠️  PROTECT actual {p['product'][:25]} {month}: "
                                  f"{ex_actual:,.0f} → {new_actual:,.0f} (ลด {100-new_actual/ex_actual*100:.0f}%) — คง existing")
                            p["actual"] = ex_actual
                        if p["goal"] > 0:
                            p["achieve"] = round(p["actual"] / p["goal"] * 100, 1)
                    merged.append(p)
                goals_data[month] = merged
    except Exception:
        pass

    new_js = json.dumps(goals_data, ensure_ascii=False, separators=(",", ":"))
    new_content = content[:start] + new_js + content[end:]

    if new_content == content:
        print("   ⚠️  GOALS_DATA ไม่มีการเปลี่ยนแปลง — ข้าม")
        return

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    total_products = sum(len(v) for v in goals_data.values())
    months_str = ", ".join(sorted(goals_data.keys()))
    print(f"   ✅ อัปเดต GOALS_DATA แล้ว: {total_products} products ({months_str})")
    for month, prods in sorted(goals_data.items()):
        total_actual = sum(p["actual"] for p in prods)
        print(f"      {month}: {len(prods)} products, total actual = ฿{total_actual:,.0f}")


# ─── main ───────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  OPD Daily Puller — OpenDurian How-to (Affiliate)")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    # ── MK13: Affiliate Creator Orders ──────────────────────
    mk13_url = SHEET_CSV_URL.format(sid=OPD_SHEET_ID, gid=MK13_GID)
    print(f"\n[MK13] กำลังดาวน์โหลด... (ไฟล์ใหญ่ อาจใช้เวลา 1-2 นาที)", flush=True)
    mk13_rows = download_csv(mk13_url, timeout=180, retries=3)
    print(f"  ได้ {len(mk13_rows)} rows (รวม header)")

    # gviz supplement: export?format=csv ถูก truncate ตัดวันล่าสุดของเดือนปัจจุบันทิ้ง
    # → ดึง gviz (de-truncate) เฉพาะเดือนปัจจุบันมาเสริม (col F = วันที่)
    # parse_mk13 dedup ด้วย (order_id, date, creator) อยู่แล้ว → append ซ้ำได้ ไม่ double-count
    try:
        import urllib.parse as _up
        _mfrom = datetime.now().replace(day=1).strftime("%Y-%m-%d")
        _tq = _up.quote("SELECT * WHERE F >= date '%s'" % _mfrom)
        _gv_url = "https://docs.google.com/spreadsheets/d/%s/gviz/tq?tqx=out:csv&gid=%s&tq=%s" % (OPD_SHEET_ID, MK13_GID, _tq)
        _gv = download_csv(_gv_url, timeout=120, retries=2)
        if _gv and len(_gv) > 1:
            mk13_rows.extend(_gv[1:])   # skip gviz header
            print(f"  + gviz เสริมเดือนปัจจุบัน ({_mfrom}+): +{len(_gv)-1} rows")
    except Exception as _gve:
        print(f"  gviz supplement fail (ข้าม): {_gve}")

    creators = {}
    daily_creators = {}
    if mk13_rows:
        creators, daily_creators = parse_mk13(mk13_rows)
        total_creators = len(creators)
        total_orders   = sum(
            d["orders"]
            for cr in creators.values()
            for d in cr.values()
        )
        print(f"  TikTok Affi creators: {total_creators} คน")
        print(f"  Total orders (non-gift): {total_orders}")
        for cr, months in sorted(creators.items()):
            rev_total = sum(d["revenue"] for d in months.values())
            ord_total = sum(d["orders"]  for d in months.values())
            print(f"    {cr}: ฿{rev_total:,.0f} / {ord_total} orders")

    # ── %Ads: Product-level TikTok Aff Spend ────────────────
    ads_url = SHEET_CSV_URL.format(sid=OPD_SHEET_ID, gid=ADS_PCT_GID)
    print(f"\n[%Ads] กำลังดาวน์โหลด...")
    ads_rows = download_csv(ads_url)
    print(f"  ได้ {len(ads_rows)} rows (รวม header)")

    ads_products = {}
    if ads_rows:
        ads_products = parse_ads_pct(ads_rows)
        for month, prods in sorted(ads_products.items()):
            total_aff = sum(d["ads_tiktok_aff"] for d in prods.values())
            print(f"  {month}: {len(prods)} products, TikTok Aff Spend ฿{total_aff:,.0f}")

    # ── OPD_DAILY: Channel revenue by date ──────────────────
    if mk13_rows:
        print(f"\n[OPD_DAILY] กำลัง parse ยอดรายวัน...")
        daily_dict = parse_opd_daily(mk13_rows)
        print(f"  {len(daily_dict)} วันที่มีข้อมูล")
        opd_data = build_opd_daily_data(daily_dict)

        # Backup
        daily_backup = os.path.join(os.path.dirname(DASHBOARD_PATH), "opd_daily_data.json")
        with open(daily_backup, "w", encoding="utf-8") as f:
            json.dump(opd_data, f, ensure_ascii=False, indent=2)
        print(f"  💾 Backup: opd_daily_data.json")

        # OPD_DAILY จัดการโดย mk13_sync.py แล้ว — ข้ามเพื่อไม่ทับข้อมูล
        print("  ℹ️  OPD_DAILY: ข้าม (mk13_sync.py รับผิดชอบ)")
    else:
        print("  ⚠️  ไม่มีข้อมูล MK13 — ข้าม OPD_DAILY")

    if not creators and not ads_products:
        print("\nไม่มีข้อมูล Affiliate")
        return

    # ── Build + Inject Affiliate ─────────────────────────────
    affiliate_data = build_affiliate_data(creators, ads_products, daily_creators)

    # บันทึก JSON ไว้เป็น backup
    backup_path = os.path.join(os.path.dirname(DASHBOARD_PATH), "affiliate_data.json")
    with open(backup_path, "w", encoding="utf-8") as f:
        json.dump(affiliate_data, f, ensure_ascii=False, indent=2)
    print(f"\n💾 Backup บันทึกแล้ว: affiliate_data.json")

    update_dashboard_affiliate(affiliate_data, DASHBOARD_PATH)

    # ── Product Monitor → GOALS_DATA ────────────────────────
    pm_url = SHEET_CSV_URL.format(sid=OPD_SHEET_ID, gid=PRODUCT_MONITOR_GID)
    print(f"\n[Product Monitor] กำลังดาวน์โหลด...")
    pm_rows = download_csv(pm_url)
    print(f"  ได้ {len(pm_rows)} rows")

    if pm_rows:
        goals_data = parse_product_monitor(pm_rows, ads_lookup=ads_products)
        if goals_data:
            update_dashboard_goals(goals_data, DASHBOARD_PATH)
        else:
            print("  ⚠️  parse ไม่ได้ product rows — ตรวจสอบ column mapping")
    else:
        print("  ⚠️  ดาวน์โหลดไม่ได้ — ข้าม GOALS_DATA")


if __name__ == "__main__":
    main()
