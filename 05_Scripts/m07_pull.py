#!/usr/bin/env python3
"""
m07_pull.py — ดึง Actual Margin + FC จาก MMS M07 (API ตรง, ไม่พึ่งเบราว์เซอร์)
อัปเดต var M07_MARGIN ใน index.html → การ์ด Goal Margin + แถบ sticky
รัน: python3 05_Scripts/m07_pull.py  (เรียกจาก opd_runner.py)
"""
import re, json, datetime, requests

API   = "https://mms-admin.opendurian.com/api/report/management"
# Bearer token (long-lived, exp ~2029) — ถ้าหมดอายุ/เปลี่ยน login ให้เอา token ใหม่จาก MMS มาแทน
TOKEN = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ0b2tlbl90eXBlIjoiYWNjZXNzIiwiZXhwIjoxODYxODQxMDQ4LCJqdGkiOiIxYWY4MmIyNGE1OWQ0MmZkOWZmMTA5MTg3MTgxNjNhZiIsInVzZXJfaWQiOjU0NH0.5I6iHu5WMbtzSxoE_gNoZCtBGxzEuJ0QlLew6737bbk"
HTML  = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"

HOWTO_TEAM = 35          # listMapping id ของ OpenDurian HOW-TO
GOALS = {"online": 431000, "consign": 429005, "total": 860005}   # เป้า margin เดือนปัจจุบัน (Jul 26 / default)
# เป้า Profit ก่อนผลิต ราย month — อิงชีต Base Profit (Brand OpenDurian HOW-TO)
GOALS_BY_MONTH = {
    "Jan 26": {"online": -72897, "consign": 552897, "total": 480000},
    "Feb 26": {"online": 29858,  "consign": 570000, "total": 599858},
    "Mar 26": {"online": 438563, "consign": 458938, "total": 897500},
    "Apr 26": {"online": 468000, "consign": 630000, "total": 1308000},
    "May 26": {"online": 249400, "consign": 624050, "total": 873450},
    "Jun 26": {"online": 174000, "consign": 560000, "total": 734000},
    "Jul 26": {"online": 431000, "consign": 429005, "total": 860005},
}

def fetch_margin(is_consign_group, start, end):
    """is_consign_group: 1=ไม่รวม Consignment (Online), 2=รวม Consignment (Total)"""
    body = {"listMapping": [HOWTO_TEAM], "managementID": 1, "reportsID": 7,
            "startDate": start, "endDate": end, "panal": 5, "week_factor": 95,
            "event_id": 1, "is_consignment_group": is_consign_group}
    r = requests.post(API, headers={"Content-Type": "application/json", "Authorization": TOKEN},
                      json=body, timeout=60)
    r.raise_for_status()
    d = r.json()["metrics"]["data"]
    return {
        "actual": round(d.get("margin_fix_cost_production", 0)),       # Profit ก่อน Fix Cost ผลิต
        "fc":     round(d.get("margin_not_del_fix_cost_production", 0)),  # ประเมินทั้งเดือน
        "sale":   round(d.get("sale_total", 0)),
        # ต้นทุน (ขอเพิ่ม): MKT / VAR / FIX + %
        "mkt":     round(d.get("mkt_cost_total", 0)),
        "mkt_pct": round(d.get("mkt_cost_percentage", 0), 2),
        "var":     round(d.get("var_cost_total", 0)),
        "var_pct": round(d.get("var_cost_total_percentage", 0), 2),
        "fix":     round(d.get("fixed_cost_total_not_production", 0)),
        "fix_pct": round(d.get("fix_cost_total_percentage_not_production", 0), 2),
    }

DAILY_START = datetime.date(2026, 3, 1)   # backfill ตั้งแต่ 1 มี.ค. 26

def read_var(html, name):
    m = re.search(r"var %s = (\{.*?\});" % name, html)
    return json.loads(m.group(1)) if m else None

def write_var(html, name, obj):
    new = "var %s = %s;" % (name, json.dumps(obj, ensure_ascii=False, separators=(",", ":")))
    return re.subn(r"var %s = \{.*?\};" % name, new, html, count=1)

def daily_entry(day):
    """margin (ไม่หักเงินเดือน) + ต้นทุนของวันเดียว — view ไม่รวม Consignment (online)"""
    d = fetch_margin(1, day, day)
    return {"m": d["actual"], "sale": d["sale"],
            "mkt": d["mkt"], "mkt_pct": d["mkt_pct"],
            "var": d["var"], "var_pct": d["var_pct"],
            "fix": d["fix"], "fix_pct": d["fix_pct"]}

def main():
    today = datetime.date.today()
    # MTD actual/cost ยึด "วันปิดล่าสุด" (เมื่อวาน) — ไม่รวมวันนี้ครึ่งวัน เพื่อให้ตรง MMS
    cutoff = today - datetime.timedelta(days=1)
    if cutoff.month != today.month:        # วันที่ 1 ของเดือน → ยังไม่มีวันปิด ใช้ today
        cutoff = today
    start = cutoff.replace(day=1).strftime("%Y-%m-%d")
    end   = cutoff.strftime("%Y-%m-%d")
    print(f"M07 pull (MTD ถึงวันปิดล่าสุด) {start} → {end}")

    # prev month ช่วงเดียวกัน (1 → วันเดียวกันกับ cutoff)
    import calendar
    pmy, pmm = (cutoff.year, cutoff.month-1) if cutoff.month > 1 else (cutoff.year-1, 12)
    pnd = min(cutoff.day, calendar.monthrange(pmy, pmm)[1])
    pstart = f"{pmy}-{pmm:02d}-01"; pend = f"{pmy}-{pmm:02d}-{pnd:02d}"
    _THm = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."]

    # ── 1) M07_MARGIN (MTD + ต้นทุน 2 view + เทียบเดือนก่อน) ──
    online = fetch_margin(1, start, end)
    total  = fetch_margin(2, start, end)
    online_p = fetch_margin(1, pstart, pend)
    total_p  = fetch_margin(2, pstart, pend)
    def costobj(d):
        return {"mkt": d["mkt"], "mkt_pct": d["mkt_pct"], "var": d["var"],
                "var_pct": d["var_pct"], "fix": d["fix"], "fix_pct": d["fix_pct"]}
    M = {
        "online":  {"goal": GOALS["online"],  "actual": online["actual"], "fc": online["fc"]},
        "consign": {"goal": GOALS["consign"], "actual": total["actual"]-online["actual"], "fc": total["fc"]-online["fc"]},
        "total":   {"goal": GOALS["total"],   "actual": total["actual"],  "fc": total["fc"]},
        # ต้นทุน 2 view (ไม่รวม/รวม Consignment) + เดือนก่อนช่วงเดียวกัน
        "cost": {"excl": costobj(online), "incl": costobj(total),
                 "prev_excl": costobj(online_p), "prev_incl": costobj(total_p),
                 "prev_label": _THm[pmm-1] + " 1-" + str(pnd)},
        "updated": end,
    }
    print("  MTD Total:", M["total"])
    print("  cost excl:", M["cost"]["excl"], "| incl:", M["cost"]["incl"])

    with open(HTML, encoding="utf-8") as f:
        html = f.read()

    new_html, n = write_var(html, "M07_MARGIN", M)
    if n != 1:
        print("  ⚠️  ไม่พบ var M07_MARGIN — ข้าม"); return
    html = new_html

    # ── 1b) M07_MARGIN_BY_MONTH (ทุกเดือน full-month — dynamic, ไม่ hardcode) ──
    # PERF: cache เดือนที่จบแล้ว (data ไม่เปลี่ยน) — ดึง API แค่เดือนปัจจุบัน (+เดือนก่อนช่วงต้นเดือน
    # เผื่อ data มาช้า) ตัดจาก ~24 API calls/รอบ เหลือ 2-4 → M07 ไม่ block sync
    ENG = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    bym = read_var(html, "M07_MARGIN_BY_MONTH") or {}   # เริ่มจาก cache เดิม
    _m = datetime.date(2026, 1, 1)
    _curfirst = cutoff.replace(day=1)
    _prevfirst = datetime.date(cutoff.year-1, 12, 1) if cutoff.month == 1 else datetime.date(cutoff.year, cutoff.month-1, 1)
    _fetched = 0
    while _m <= _curfirst:
        y, mo = _m.year, _m.month
        key = f"{ENG[mo-1]} {str(y)[2:]}"
        g = GOALS_BY_MONTH.get(key, GOALS)
        _is_cur  = (y == cutoff.year and mo == cutoff.month)
        _is_prev = (_m == _prevfirst and cutoff.day <= 5)   # refresh prev เฉพาะต้นเดือน (data late)
        _nxt = datetime.date(y+1, 1, 1) if mo == 12 else datetime.date(y, mo+1, 1)
        # เดือนจบแล้ว + มีใน cache → ไม่ดึง API ซ้ำ (แค่ refresh goal เผื่อ GOALS_BY_MONTH เปลี่ยน)
        if key in bym and not _is_cur and not _is_prev:
            try:
                bym[key]["online"]["goal"]  = g["online"]
                bym[key]["consign"]["goal"] = g["consign"]
                bym[key]["total"]["goal"]   = g["total"]
            except Exception:
                pass
            _m = _nxt; continue
        ms = f"{y}-{mo:02d}-01"; last = calendar.monthrange(y, mo)[1]
        # เดือนปัจจุบัน → MTD ถึง cutoff; เดือนจบแล้ว → ทั้งเดือน
        me = end if _is_cur else f"{y}-{mo:02d}-{last:02d}"
        try:
            on = fetch_margin(1, ms, me); to = fetch_margin(2, ms, me); _fetched += 1
            bym[key] = {
                "online":  {"goal": g["online"],  "actual": on["actual"]},
                "consign": {"goal": g["consign"], "actual": to["actual"] - on["actual"]},
                "total":   {"goal": g["total"],   "actual": to["actual"]},
            }
        except Exception as e:
            print(f"  BYM {ms} fail: {e}")
        _m = _nxt
    print(f"  BYM: ดึง API {_fetched} เดือน (cache ที่เหลือ)")
    new_html_b, nb = write_var(html, "M07_MARGIN_BY_MONTH", bym)
    if nb == 1:
        html = new_html_b
        print(f"  ✅ M07_MARGIN_BY_MONTH {len(bym)} เดือน: " + ", ".join(f"{k}={bym[k]['total']['actual']:,}" for k in bym))
    else:
        print("  ⚠️  ไม่พบ var M07_MARGIN_BY_MONTH — ข้าม (เพิ่ม var M07_MARGIN_BY_MONTH = {}; ใน index.html)")

    # ── 2) M07_DAILY (รายวัน — self-backfill ตั้งแต่ 1 มี.ค.) ──
    daily = read_var(html, "M07_DAILY") or {}
    want = []
    d = DAILY_START
    while d <= today:
        ds = d.strftime("%Y-%m-%d")
        # เติมวันที่ขาด + refresh 2 วันล่าสุดเสมอ (ยอดวันนี้/เมื่อวานยังขยับ)
        if ds not in daily or (today - d).days <= 6:   # refresh 6 วันล่าสุด (วันเก่า M07 ยัง allocate → กัน stale ไม่ตรง M07)
            want.append(ds)
        d += datetime.timedelta(days=1)
    print(f"  M07_DAILY: ต้องดึง {len(want)} วัน (มีแล้ว {len(daily)})")
    for ds in want:
        try:
            daily[ds] = daily_entry(ds)
        except Exception as e:
            print(f"    {ds} fail: {e}")
    new_html2, n2 = write_var(html, "M07_DAILY", daily)
    if n2 == 1:
        html = new_html2
        print(f"  ✅ M07_DAILY {len(daily)} วัน")
    else:
        print("  ⚠️  ไม่พบ var M07_DAILY — ข้าม (เพิ่ม var M07_DAILY = {}; ใน index.html)")

    with open(HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print("  ✅ อัปเดต index.html แล้ว")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"  ⚠️  M07 pull ล้มเหลว: {e}")
