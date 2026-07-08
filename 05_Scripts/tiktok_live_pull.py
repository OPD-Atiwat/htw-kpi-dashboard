#!/usr/bin/env python3
"""
tiktok_live_pull.py — ดึง TikTok Live product catalog (Google Sheet pub CSV) ทุกวัน
  → var TIKTOK_LIVE = {catalog:[...], live_days:{"YYYY-MM":[วันที่มีไลฟ์]}, updated}
- catalog: Product/Status(Focus/Maintain/Dead Stock/Clearance)/Product ID/ราคา (ปกติ/Live/LiveNew/Flash)
- live_days: วันที่มียอด TikTok Live ใน OPD_DAILY (มีไลฟ์วันไหน) แยกเดือน
chain ต่อท้าย kms_pull (runner ไม่เรียกตรง). อัปเดตทุก sync cycle
"""
import re, json, csv, io, datetime, urllib.request

CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTn35kR7NfFXOO762nWkhHui4ZWG2cj6l3Q0IP0NjcumtYrEr6ypRpWsw4ICGgoLROe2tlc8Af-Xjlw/pub?output=csv"
# ตาราง Live/Rerun รายชั่วโมง (workbook อีกอัน, tab "Live" gid=943460085)
SCHED_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQI5FqrcI2E7MeJmgVs783XiDQtTrYnlKhmVCKbWWGu4_lI-dre8Obd0lAF6zKI37aatwCzwwI7C-Pt/pub?gid=943460085&single=true&output=csv"
_MON = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
HTML = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"

def fetch_schedule():
    """tab Live: row0=2h, row1=1h slots(col1-24), rows2+=date + Live/Rerun ต่อชั่วโมง
       → {'YYYY-MM-DD': {live_h, rerun_h, live_hrs:[], rerun_hrs:[]}}"""
    import datetime as _dt
    req = urllib.request.Request(SCHED_URL, headers={"User-Agent": "Mozilla/5.0"})
    txt = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "replace")
    rows = list(csv.reader(io.StringIO(txt)))
    yr = _dt.date.today().year
    out = {}
    for r in rows[2:]:
        if not r or not r[0].strip():
            continue
        m = re.search(r"(\d{1,2})\s+([A-Za-z]{3})", r[0])   # "Wed 1 Jul"
        if not m:
            continue
        day = int(m.group(1)); mon = _MON.get(m.group(2).lower())
        if not mon:
            continue
        ds = f"{yr:04d}-{mon:02d}-{day:02d}"
        live, rerun = [], []
        for c in range(1, 25):
            v = (r[c].strip().lower() if len(r) > c else "")
            if v == "live":   live.append(c-1)
            elif v == "rerun": rerun.append(c-1)
        if live or rerun:
            out[ds] = {"live_h": len(live), "rerun_h": len(rerun), "live_hrs": live, "rerun_hrs": rerun}
    return out

def fetch_catalog():
    req = urllib.request.Request(CSV_URL, headers={"User-Agent": "Mozilla/5.0"})
    txt = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "replace")
    rows = list(csv.reader(io.StringIO(txt)))   # csv handles quoted multiline
    out = []
    for r in rows[2:]:                            # row0=section, row1=header
        if len(r) < 6 or not r[1].strip() or not r[5].strip():
            continue
        def n(x):
            try: return float(str(x).replace(",", "")) if str(x).strip() else None
            except Exception: return None
        out.append({
            "product": r[1].strip(),
            "status":  r[3].strip() if len(r) > 3 else "",
            "noted":   r[4].strip() if len(r) > 4 else "",
            "pid":     r[5].strip(),
            "normal":  n(r[6]) if len(r) > 6 else None,
            "live":    n(r[7]) if len(r) > 7 else None,
            "live_new":n(r[8]) if len(r) > 8 else None,
            "flash":   n(r[9]) if len(r) > 9 else None,
        })
    return out

def read_var(html, name):
    m = re.search(r'var %s\s*=\s*(\{.*?\});' % name, html, re.DOTALL)
    try: return json.loads(m.group(1)) if m else None
    except Exception: return None

def live_days_by_month(html):
    """วันที่มียอด 'TikTok Live' ใน OPD_DAILY → {'YYYY-MM':[วัน,...]} (มีไลฟ์วันไหน)"""
    od = read_var(html, "OPD_DAILY")
    res = {}
    if not od or "data" not in od:
        return res
    for row in od.get("data", []):
        d = row.get("d") or ""
        if len(d) < 10:
            continue
        if (row.get("TikTok Live") or 0) > 0:
            res.setdefault(d[:7], []).append({"d": d, "rev": round(row.get("TikTok Live") or 0)})
    for k in res:
        res[k].sort(key=lambda x: x["d"])
    return res

def main():
    print("=== TikTok Live catalog + live-days ===")
    try:
        catalog = fetch_catalog()
        print(f"  catalog: {len(catalog)} products")
    except Exception as e:
        print(f"  ⚠️  ดึง CSV ไม่ได้ ({e}) — คงของเดิม"); return

    html = open(HTML, encoding="utf-8").read()
    ld = live_days_by_month(html)
    print(f"  live months: {', '.join('%s(%d วัน)'%(m,len(v)) for m,v in sorted(ld.items()))}")

    try:
        sched = fetch_schedule()
        print(f"  schedule: {len(sched)} วัน (Live {sum(1 for v in sched.values() if v['live_h']>0)} · Rerun {sum(1 for v in sched.values() if v['rerun_h']>0)})")
    except Exception as _se:
        print(f"  schedule fail (ข้าม): {_se}"); sched = {}
    obj = {"catalog": catalog, "live_days": ld, "schedule": sched,
           "updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}
    new_js = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    if re.search(r'var TIKTOK_LIVE\s*=', html):
        html = re.sub(r'var TIKTOK_LIVE\s*=\s*\{.*?\};', "var TIKTOK_LIVE = %s;" % new_js, html, count=1, flags=re.DOTALL)
    else:
        # สร้างใหม่ก่อน </script> แรก (fallback) หรือหลัง OPD_DAILY
        m = re.search(r'(var OPD_DAILY\s*=\s*\{.*?\};)', html, re.DOTALL)
        if m:
            html = html[:m.end()] + " var TIKTOK_LIVE = %s;" % new_js + html[m.end():]
        else:
            print("  ⚠️  หาที่วาง TIKTOK_LIVE ไม่ได้"); return
    open(HTML, "w", encoding="utf-8").write(html)
    print(f"  ✅ TIKTOK_LIVE updated ({len(catalog)} products, {len(ld)} months)")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"  ⚠️  tiktok_live_pull ล้มเหลว: {e}")
