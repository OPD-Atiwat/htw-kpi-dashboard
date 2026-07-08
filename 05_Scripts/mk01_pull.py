#!/usr/bin/env python3
"""
mk01_pull.py — ดึง MK01 (%Var Cost ราย channel) จาก MMS API → var MK01_VAR_CH (live)
────────────────────────────────────────────────────────────────────────
แหล่งความจริง: api/report/marketing_team reportsID=1 (MK01) · marketingTeamID=35 (HOW-TO)
  body: {reportsID:1, marketingTeamID:35, startDate, endDate, panal:1,
         productGroupID:[...], type_sales:1}
  response: metrics.data.summary[].report → percentage_var_cost_number (%var), sale_total
mapping Product Group → dashboard channel:
  "OpenDurian HOW-TO"            = Meta  (Facebook/Instagram/LINE/Shopify)
  "TikTok X OpenDurian HOW-TO"   = TikTok (TikTok/TikTok Affi/TikTok Live)
  "Shopee X OpenDurian HOW-TO"   = Shopee (Shopee/Shopee Live)
ยึด MTD (1 → วันปิดล่าสุด) เพื่อ %var นิ่ง · เขียน MK01_VAR_CH merge กับของเดิม (ไม่ทับ channel ที่ไม่ได้ดึง)
เลิกใช้ค่า static — ดึงสดทุกรอบ. รันใน opd_runner.py (หลัง m44)
"""
import re, json, datetime, urllib.request

API  = "https://mms-admin.opendurian.com/api/report/marketing_team"
HTML = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"
M07  = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts/m07_pull.py"
TEAM = 35
PRODUCT_GROUP = [286, 287, 288, 290, 291, 292, 362, 487, 488, 546, 547, 548]

# product_group name (จาก MK01) → dashboard channels ที่ใช้ %var เดียวกัน
GROUP_CHANNELS = {
    "OpenDurian HOW-TO":          ["Facebook", "Instagram", "LINE", "Shopify"],   # Meta
    "TikTok X OpenDurian HOW-TO": ["TikTok", "TikTok Affi", "TikTok Live"],
    "Shopee X OpenDurian HOW-TO": ["Shopee", "Shopee Live"],
}

def get_token():
    return re.search(r'TOKEN\s*=\s*"([^"]+)"', open(M07, encoding="utf-8").read()).group(1)

def fetch(start, end, token):
    body = {"reportsID": 1, "marketingTeamID": TEAM, "startDate": start, "endDate": end,
            "panal": 1, "productGroupID": PRODUCT_GROUP, "type_sales": 1}
    req = urllib.request.Request(API, data=json.dumps(body).encode(),
          headers={"Content-Type": "application/json", "Authorization": token})
    return json.load(urllib.request.urlopen(req, timeout=120))

def main():
    print("=== MK01 %Var ราย channel (live) ===")
    today = datetime.date.today()
    cutoff = today - datetime.timedelta(days=1)          # ยึดวันปิดล่าสุด (ไม่เอาวันนี้ที่ยังไม่นิ่ง)
    if cutoff.month != today.month:                       # วันที่ 1 → ใช้ today
        cutoff = today
    start = cutoff.replace(day=1).strftime("%Y-%m-%d")
    end   = cutoff.strftime("%Y-%m-%d")
    print(f"  MTD {start} → {end}")

    j = fetch(start, end, get_token())
    summ = (((j or {}).get("metrics") or {}).get("data") or {}).get("summary") or []
    varch = {}
    for row in summ:
        pg = row.get("product_group")
        chans = GROUP_CHANNELS.get(pg)
        if not chans:
            continue
        rep = row.get("report") or {}
        pv  = rep.get("percentage_var_cost_number")
        sale = rep.get("sale_total") or 0
        if pv is None or sale <= 0:                       # ไม่มีข้อมูล → ข้าม (คงค่าเดิม ไม่ทับด้วย 0)
            print(f"  ข้าม {pg}: pv={pv} sale={sale}")
            continue
        for ch in chans:
            varch[ch] = round(float(pv), 2)
        print(f"  {pg}: %var={round(float(pv),2)} (sale ฿{sale:,.0f}) → {chans}")

    if not varch:
        print("  ⚠️  ไม่ได้ %var เลย — ข้าม (คงค่าเดิม)"); return

    html = open(HTML, encoding="utf-8").read()
    m = re.search(r'var MK01_VAR_CH\s*=\s*(\{[^}]*\})\s*;', html)
    cur = {}
    if m:
        try: cur = json.loads(m.group(1))
        except Exception: cur = {}
    cur.update(varch)                                     # merge — ไม่ drop channel ที่ไม่ได้ดึง
    new = "var MK01_VAR_CH = %s;" % json.dumps(cur, ensure_ascii=False, separators=(",", ":"))
    html2, n = re.subn(r'var MK01_VAR_CH\s*=\s*\{[^}]*\}\s*;', new, html, count=1)
    if n != 1:
        print("  ⚠️  ไม่พบ var MK01_VAR_CH ใน index.html — ข้าม"); return
    open(HTML, "w", encoding="utf-8").write(html2)
    print(f"  ✅ MK01_VAR_CH updated: {cur}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"  ⚠️  MK01 pull ล้มเหลว: {e}")
