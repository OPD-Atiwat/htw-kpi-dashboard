#!/usr/bin/env python3
"""
m44_pull.py — ดึง M44 (Ads cost แยกเล่ม) จาก MMS API → ADSM_DAILY + ADSM_PROD_DAILY
────────────────────────────────────────────────────────────────────────
แหล่งความจริง: api/report/management reportsID=44 (ค่าแอด/ยอดขาย/%ads แยกเล่ม แยกช่อง)
  • ADSM_DAILY[]            : รวมทุกเล่ม/วัน  {d, ac, st, pa, fb, tt, sp}  (fb/tt/sp = ค่าแอดบาท)
  • ADSM_PROD_DAILY{book:[]}: ราย book/วัน    {d, ac, st, pa, fb, tt, sp, affi, *_ac, *_st}
self-backfill ตั้งแต่ DAILY_START + refresh 2 วันล่าสุด (เหมือน m07_pull)
รันใน opd_runner.py
"""
import re, json, datetime, urllib.request

API   = "https://mms-admin.opendurian.com/api/report/management"
HTML  = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"
M07   = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts/m07_pull.py"
TEAM  = 35
PRODUCT_GROUP = [286, 287, 288, 290, 291, 292, 319, 362, 487, 488, 546, 547, 548]
DAILY_START = datetime.date(2026, 1, 1)

def get_token():
    return re.search(r'TOKEN\s*=\s*"([^"]+)"', open(M07, encoding="utf-8").read()).group(1)

def fetch(day, token):
    body = {"marketing_team_id": TEAM, "managementID": 1, "reportsID": 44,
            "startDate": day, "endDate": day, "productGroup": PRODUCT_GROUP}
    req = urllib.request.Request(API, data=json.dumps(body).encode(),
          headers={"Content-Type": "application/json", "Authorization": token})
    return json.load(urllib.request.urlopen(req, timeout=120))

def g(row, k):
    try: return float(row.get(k) or 0)
    except (TypeError, ValueError): return 0.0

def pct(cost, sale):
    return round(cost / sale * 100, 2) if sale else 0.0

def fb_cost(r):  return g(r, "adscost_facebook_msg") + g(r, "adscost_facebook_salepage")
def fb_sale(r):  return g(r, "saletotal_facebook_msg") + g(r, "saletotal_facebook_sale_page")
def tt_cost(r):  return g(r, "adscost_tiktok_ads")
def tt_sale(r):  return g(r, "saletotal_tiktok")
def affi_cost(r):return g(r, "adscost_tiktok_aff")
def affi_sale(r):return g(r, "saletotal_tiktok_affiliate")
def sp_cost(r):  return g(r, "adscost_shopee")
def sp_sale(r):  return g(r, "saletotal_shopee")

def daily_row(total):
    """ADSM_DAILY (รวม) — fb/tt/sp = ค่าแอดบาท"""
    ac = g(total, "adscost_total"); st = g(total, "saletotal_total")
    return {"ac": round(ac, 2), "st": round(st, 2), "pa": pct(ac, st),
            "fb": round(fb_cost(total), 2), "tt": round(tt_cost(total), 2), "sp": round(sp_cost(total), 2)}

def prod_row(r):
    """ราย book — fb/tt/sp/affi = %ads ; *_ac/_st = cost/sale"""
    ac = g(r, "adscost_total"); st = g(r, "saletotal_total")
    fa, fs = fb_cost(r), fb_sale(r)
    ta, ts = tt_cost(r), tt_sale(r)
    aa, as_ = affi_cost(r), affi_sale(r)
    sa, ss = sp_cost(r), sp_sale(r)
    return {"ac": round(ac, 2), "st": round(st, 2), "pa": pct(ac, st),
            "fb": pct(fa, fs), "tt": pct(ta, ts), "sp": pct(sa, ss), "affi": pct(aa, as_),
            "fb_ac": round(fa, 2), "fb_st": round(fs, 2),
            "tt_ac": round(ta, 2), "tt_st": round(ts, 2),
            "sp_ac": round(sa, 2), "sp_st": round(ss, 2),
            "affi_ac": round(aa, 2), "affi_st": round(as_, 2)}

def book_key(name):
    """ตัด prefix [หนังสือ]/[xxx] นำหน้า ให้ตรง key เดิมใน ADSM_PROD_DAILY"""
    return re.sub(r'^\[[^\]]*\]\s*', '', (name or "").strip())

def read_var(html, name):
    m = re.search(r"var %s = (\[.*?\]|\{.*?\});" % name, html)
    return json.loads(m.group(1)) if m else None

def write_var(html, name, obj):
    new = "var %s = %s;" % (name, json.dumps(obj, ensure_ascii=False, separators=(",", ":")))
    return re.subn(r"var %s = (?:\[.*?\]|\{.*?\});" % name, new, html, count=1)

def main():
    token = get_token()
    with open(HTML, encoding="utf-8") as f:
        html = f.read()

    adsm = read_var(html, "ADSM_DAILY")
    prod = read_var(html, "ADSM_PROD_DAILY")
    if adsm is None or prod is None:
        print("  ⚠️  ไม่พบ ADSM_DAILY/ADSM_PROD_DAILY — ข้าม"); return

    by_date = {r["d"]: r for r in adsm}
    today = datetime.date.today()

    # วันที่ต้องดึง: ขาด + refresh 2 วันล่าสุด
    want, d = [], DAILY_START
    while d <= today:
        ds = d.strftime("%Y-%m-%d")
        if ds not in by_date or (today - d).days <= 1:
            want.append(ds)
        d += datetime.timedelta(days=1)
    print(f"  M44: ต้องดึง {len(want)} วัน (มีแล้ว {len(by_date)})")

    prod_by = {k: {r["d"]: r for r in v} for k, v in prod.items()}
    ok = 0
    for ds in want:
        try:
            data = fetch(ds, token)
        except Exception as e:
            print(f"    {ds} fail: {e}"); continue
        tot = (data.get("list_total_by_product") or [{}])[0]
        if not tot or g(tot, "adscost_total") == 0 and g(tot, "saletotal_total") == 0:
            continue
        dr = daily_row(tot); dr["d"] = ds
        by_date[ds] = dr
        for it in data.get("list_item", []):
            k = book_key(it.get("product_name"))
            if not k: continue
            pr = prod_row(it); pr["d"] = ds
            if g(it, "adscost_total") == 0 and g(it, "saletotal_total") == 0:
                continue
            prod_by.setdefault(k, {})[ds] = pr
        ok += 1

    if ok == 0:
        print("  ℹ️  ไม่มีข้อมูลใหม่จาก M44"); return

    adsm_out = sorted(by_date.values(), key=lambda r: r["d"])
    prod_out = {k: sorted(v.values(), key=lambda r: r["d"]) for k, v in prod_by.items()}

    for name, obj in [("ADSM_DAILY", adsm_out), ("ADSM_PROD_DAILY", prod_out)]:
        html, n = write_var(html, name, obj)
        if n != 1:
            print(f"  ⚠️  เขียน {name} ไม่สำเร็จ — ยกเลิก"); return

    with open(HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✅ M44 อัปเดต {ok} วัน | ADSM_DAILY {len(adsm_out)} วัน | {len(prod_out)} เล่ม")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"  ⚠️  m44_pull ล้มเหลว: {e}")
