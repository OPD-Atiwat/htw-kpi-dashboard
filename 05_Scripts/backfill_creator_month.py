#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Backfill CREATOR_META_<MONTH> ลง dashboard สำหรับเดือนย้อนหลัง (เช่น พ.ค. 26)
ดึง Meta Ads เต็มเดือนผ่าน API + ใช้ logic เดียวกับ meta_pull.py (ไทยล้วน + แชร์ 50/50)
แล้ว "ฝังเฉพาะตัวแปรเดือนนั้น" ลง index.html — ไม่แตะ CREATOR_META เดือนอื่น

วิธีรัน (บน Mac, Terminal):
    cd "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts"
    python3 backfill_creator_month.py 2026-05
หลังรันเสร็จ → publish ด้วย .sync_trigger หรือ Sync & Push ได้เลย
"""
import sys, os, re, json, calendar
from collections import defaultdict
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import meta_pull as mp  # มี __main__ guard → import ไม่รัน pull

MONTH_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
              7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
TEAM = ["หมิว", "ริว", "มิ้น", "สต็อป"]   # ตรงกับ meta_pull.py + MAIN_CR ใน dashboard

def main():
    if len(sys.argv) < 2:
        print("ใช้: python3 backfill_creator_month.py YYYY-MM   (เช่น 2026-05)"); sys.exit(1)
    yy, mm = map(int, sys.argv[1].split("-"))
    mcode = f"{MONTH_ABBR[mm]}{str(yy)[2:]}"        # 2026-05 -> May26
    last  = calendar.monthrange(yy, mm)[1]

    # override ช่วงวันที่ของ meta_pull เป็นเต็มเดือนเป้าหมาย
    mp.DATE_FROM = f"{yy:04d}-{mm:02d}-01"
    mp.DATE_TO   = f"{yy:04d}-{mm:02d}-{last:02d}"
    print(f"📅 ดึง Meta เต็มเดือน {mcode}: {mp.DATE_FROM} → {mp.DATE_TO}")

    raw  = mp.pull_ad_insights()
    rows = mp.transform_rows(raw)
    rows = [r for r in rows if r.get("month") == mcode]
    if not rows:
        print(f"❌ ไม่มี records เดือน {mcode}"); sys.exit(1)

    def _crs(r): return r.get("creators") or [r["creator"]]
    def _w(r, cr): return r.get("share", 1.0) if cr in _crs(r) else 0.0
    month_rows = [r for r in rows if any(c in TEAM for c in _crs(r))]

    summary, daily, products = {}, {}, {}
    for cr in TEAM:
        crows = [(r, _w(r, cr)) for r in month_rows if _w(r, cr) > 0]
        if not crows: continue
        sp  = sum(r["spend_thb"]   * w for r, w in crows)
        rv  = sum(r["revenue_thb"] * w for r, w in crows)
        pur = sum(r["purchases"]   * w for r, w in crows)
        imp = sum(r["impressions"] * w for r, w in crows)
        rch = sum(r["reach"]       * w for r, w in crows)
        clk = sum(r["clicks"]      * w for r, w in crows)
        msg = sum(r.get("messages",0)* w for r, w in crows)
        summary[cr] = {"spend":round(sp,2),"revenue":round(rv,2),
            "roas":round(rv/sp,4) if sp>0 else 0,"purchases":round(pur),
            "impressions":round(imp),"reach":round(rch),
            "ctr":round(clk/imp*100,4) if imp>0 else 0,
            "cpm":round(sp/imp*1000,4) if imp>0 else 0,
            "messages":round(msg),"link_clicks":round(clk),
            "cpa":round(sp/pur,2) if pur>0 else 0}
        bd = defaultdict(lambda:{"s":0,"r":0,"p":0,"i":0,"m":0})
        for r, w in crows:
            d = bd[r["ad_date"]]
            d["s"]+=r["spend_thb"]*w; d["r"]+=r["revenue_thb"]*w
            d["p"]+=r["purchases"]*w; d["i"]+=r["impressions"]*w; d["m"]+=r.get("messages",0)*w
        daily[cr] = [{"d":d,"s":round(v["s"],2),"r":round(v["r"],2),
            "p":round(v["p"]),"i":round(v["i"]),"m":round(v["m"])}
            for d,v in sorted(bd.items())]
        bp = defaultdict(lambda:{"spend":0,"revenue":0,"purchases":0})
        for r, w in crows:
            p = r.get("product") or "Unknown"
            bp[p]["spend"]+=r["spend_thb"]*w; bp[p]["revenue"]+=r["revenue_thb"]*w
            bp[p]["purchases"]+=r["purchases"]*w
        products[cr] = sorted([{"name":p,"spend":round(v["spend"],2),
            "revenue":round(v["revenue"],2),
            "roas":round(v["revenue"]/v["spend"],4) if v["spend"]>0 else 0,
            "purchases":round(v["purchases"])} for p,v in bp.items()],
            key=lambda x:-x["revenue"])

    var_name = f"CREATOR_META_{mcode.upper()}"
    obj = json.dumps({"summary":summary,"daily":daily,"products":products},
                     ensure_ascii=False, separators=(",",":"))
    line = f"const {var_name} = {obj};"

    with open(mp.DASHBOARD_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    if var_name in html:   # มีอยู่แล้ว → แทนเฉพาะตัวนี้
        html = re.sub(rf'const {var_name} = \{{.*?\}};', line, html, count=1, flags=re.DOTALL)
        action = "แทนที่"
    else:                  # ยังไม่มี → แทรกต่อท้าย CREATOR_META_JUN26 (ตัวเดียวที่มี)
        anchor = re.search(r'const CREATOR_META_[A-Z0-9]+ = \{.*?\};', html, flags=re.DOTALL)
        if not anchor:
            print("❌ ไม่พบ anchor CREATOR_META_* ใน HTML"); sys.exit(1)
        html = html[:anchor.end()] + "\n" + line + html[anchor.end():]
        action = "แทรกใหม่"

    with open(mp.DASHBOARD_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    sp_tot = sum(v["spend"] for v in summary.values())
    rv_tot = sum(v["revenue"] for v in summary.values())
    print(f"✅ {action} {var_name}")
    for cr in TEAM:
        if cr in summary:
            print(f"   {cr}: rev ฿{summary[cr]['revenue']:,.0f} | spend ฿{summary[cr]['spend']:,.0f}")
    print(f"   รวมทีม: rev ฿{rv_tot:,.0f} | spend ฿{sp_tot:,.0f}")
    print("👉 publish: touch 05_Scripts/.sync_trigger หรือ Sync & Push")

if __name__ == "__main__":
    main()
