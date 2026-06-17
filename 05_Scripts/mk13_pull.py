#!/usr/bin/env python3
"""
mk13_pull.py — Top-up OPD_DAILY จาก MMS API (api/report/marketing_team, reportsID=13)
────────────────────────────────────────────────────────────────────────
จุดประสงค์: เติม "เฉพาะวันใหม่ที่ Google Sheet ยังไม่มี" (เช่น เมื่อ Sheet อัปเดตช้า)
            โดย **ไม่แตะ/ไม่ทับ** วันที่ Sheet มีอยู่แล้ว และไม่แตะ Bookfair

ความปลอดภัย (กันข้อมูลพัง):
  • เติมเฉพาะ date > max(วันที่มีใน OPD_DAILY) → ไม่มีทางทับของ Sheet
  • เมื่อ Sheet ตามมาทันวันถัดไป mk13_sync (Sheet) จะเขียนทับ top-up นี้เอง
  • cross-check แล้ว: เม.ย./พ.ค./มิ.ย. API == Sheet ทุกช่อง (ดู note ใน HOWTO_UNIVERSE)

รันหลัง meta_pull/mk13_sync ใน opd_runner.py (top-up tail)
"""
import re, json, datetime, urllib.request

API   = "https://mms-admin.opendurian.com/api/report/marketing_team"
HTML  = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"
M07   = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts/m07_pull.py"
TEAM  = 35          # marketingTeamID = OpenDurian HOW-TO
BACKDAYS = 7        # มองหาวันที่ขาดย้อนหลังไม่เกิน N วัน

def get_token():
    return re.search(r'TOKEN\s*=\s*"([^"]+)"', open(M07, encoding="utf-8").read()).group(1)

def fetch_day(day, token):
    rows, page = [], 1
    while True:
        body = {"reportsID": 13, "marketingTeamID": TEAM, "startDate": day, "endDate": day,
                "panal": 1, "page": page, "row_per_pages": 500,
                "sorting": "transactions_time", "sorting_option": False}
        req = urllib.request.Request(API, data=json.dumps(body).encode(),
              headers={"Content-Type": "application/json", "Authorization": token})
        d = json.load(urllib.request.urlopen(req, timeout=90))
        m = d.get("metric", [])
        rows += m
        if len(m) < 500:
            break
        page += 1
    return rows

def map_channel(r):
    """ตรง logic กับ mk13_sync.map_channel + ใช้ is_tiktok_affiliate (เกณฑ์เดียวกับ Sheet)"""
    if r.get("is_tiktok_affiliate"):
        return "TikTok Affi"
    ch = (r.get("sale_channel") or "").strip()
    mt = (r.get("sale_method") or "").strip()
    ml = mt.lower()
    if ch == "TikTok":
        if mt == "Affiliate":   return "TikTok Affi"
        if "live" in ml:        return "TikTok Live"
        return "TikTok"
    if ch == "Shopee":
        return "Shopee Live" if "live" in ml else "Shopee"
    if ch == "Facebook":
        return "Shopify" if mt in ("Salepage", "Shopify", "FBSalepage") else "Facebook"
    return {"Instagram": "Instagram", "LINE": "LINE", "Line": "LINE", "YouTube": "YouTube",
            "Lazada": "Lazada", "Web": "Web", "Website": "Web", "หน้าร้าน": "หน้าร้าน"}.get(ch)

def read_var(html, name):
    m = re.search(r"var %s = (\{.*?\});" % name, html)
    return json.loads(m.group(1)) if m else None

def write_var(html, name, obj):
    new = "var %s = %s;" % (name, json.dumps(obj, ensure_ascii=False, separators=(",", ":")))
    return re.subn(r"var %s = \{.*?\};" % name, new, html, count=1)

def main():
    token = get_token()
    with open(HTML, encoding="utf-8") as f:
        html = f.read()

    opd  = read_var(html, "OPD_DAILY")
    pdat = read_var(html, "OPD_PROD_DATA") or {}
    pqty = read_var(html, "OPD_PROD_QTY") or {}
    if not opd or "data" not in opd:
        print("  ⚠️  ไม่พบ OPD_DAILY — ข้าม"); return

    have = set(r["d"] for r in opd["data"])
    maxd = max(have) if have else "2026-01-01"
    today = datetime.date.today()

    # วันที่ต้อง top-up = (maxd, today] ที่ยังไม่มี — ไม่เกิน BACKDAYS
    targets = []
    d = today
    for _ in range(BACKDAYS + 1):
        ds = d.strftime("%Y-%m-%d")
        if ds > maxd and ds not in have:
            targets.append(ds)
        d -= datetime.timedelta(days=1)
    targets = sorted(targets)

    if not targets:
        print(f"  ℹ️  ไม่มีวันใหม่ให้ top-up (Sheet ล่าสุด {maxd}) — ข้าม"); return
    print(f"  Sheet ล่าสุด {maxd} → top-up จาก API: {targets}")

    ch_set = set(opd.get("channels", []))
    added_rows = 0
    for ds in targets:
        try:
            rows = fetch_day(ds, token)
        except Exception as e:
            print(f"    {ds} fetch fail: {e}"); continue
        day = {"d": ds}
        for r in rows:
            if r.get("is_giveaway"):
                continue
            ch = map_channel(r)
            if not ch:
                continue
            amt = float(r.get("amount") or 0)
            if amt <= 0:
                continue
            qty = int(r.get("quantity") or 0)
            ch_set.add(ch)
            day[ch] = day.get(ch, 0) + amt
            day[ch + "_q"] = day.get(ch + "_q", 0) + (qty if qty else 1)
            # per-product (แยกเล่ม) — ข้ามส่วนลด/voucher
            prod = (r.get("product_name") or "").strip()
            if prod and not prod.startswith("*") and not prod.startswith("ส่วนลด"):
                pdat.setdefault(prod, {}).setdefault(ds, {})
                pdat[prod][ds][ch] = pdat[prod][ds].get(ch, 0) + amt
                pqty.setdefault(prod, {}).setdefault(ds, {})
                pqty[prod][ds][ch] = pqty[prod][ds].get(ch, 0) + (qty if qty else 1)
        # ปัดเศษเงินให้เหมือน Sheet (int)
        for k in list(day):
            if k != "d" and not k.endswith("_q"):
                day[k] = round(day[k])
        opd["data"].append(day)
        added_rows += 1
        tot = sum(v for k, v in day.items() if k != "d" and not k.endswith("_q"))
        print(f"    + {ds}: รวม ฿{tot:,.0f} ({len([k for k in day if not k.endswith('_q') and k!='d'])} ช่อง)")

    if added_rows == 0:
        print("  ℹ️  ไม่มีข้อมูลใหม่จาก API"); return

    opd["data"].sort(key=lambda r: r["d"])
    opd["channels"] = sorted(ch_set)

    for name, obj in [("OPD_DAILY", opd), ("OPD_PROD_DATA", pdat), ("OPD_PROD_QTY", pqty)]:
        html, n = write_var(html, name, obj)
        if n != 1:
            print(f"  ⚠️  เขียน {name} ไม่สำเร็จ — ยกเลิกทั้งหมด"); return

    with open(HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✅ top-up {added_rows} วัน เข้า OPD_DAILY + แยกเล่ม (channel+product)")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"  ⚠️  mk13_pull ล้มเหลว: {e}")
