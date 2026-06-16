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
GOALS = {"online": 174000, "consign": 560000, "total": 734000}   # เป้า margin (กำหนดเอง)

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
    }

def main():
    today = datetime.date.today()
    start = today.replace(day=1).strftime("%Y-%m-%d")
    end   = today.strftime("%Y-%m-%d")
    print(f"M07 pull {start} → {end}")

    online = fetch_margin(1, start, end)
    total  = fetch_margin(2, start, end)
    consign = {"actual": total["actual"] - online["actual"],
               "fc":     total["fc"]     - online["fc"]}

    M = {
        "online":  {"goal": GOALS["online"],  "actual": online["actual"],  "fc": online["fc"]},
        "consign": {"goal": GOALS["consign"], "actual": consign["actual"], "fc": consign["fc"]},
        "total":   {"goal": GOALS["total"],   "actual": total["actual"],   "fc": total["fc"]},
        "updated": end,
    }
    print("  Online :", M["online"])
    print("  Consign:", M["consign"])
    print("  Total  :", M["total"])

    new_var = "var M07_MARGIN = " + json.dumps(M, ensure_ascii=False, separators=(",", ":")) + ";"
    with open(HTML, encoding="utf-8") as f:
        html = f.read()
    new_html, n = re.subn(r"var M07_MARGIN = \{.*?\};", new_var, html, count=1)
    if n != 1:
        print("  ⚠️  ไม่พบ var M07_MARGIN ใน index.html — ข้าม")
        return
    if new_html != html:
        with open(HTML, "w", encoding="utf-8") as f:
            f.write(new_html)
        print("  ✅ อัปเดต M07_MARGIN ใน index.html แล้ว")
    else:
        print("  (ไม่มีการเปลี่ยนแปลง)")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"  ⚠️  M07 pull ล้มเหลว: {e}")
