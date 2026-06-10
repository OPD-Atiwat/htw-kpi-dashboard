"""
backfill_raw.py — เติม RAW_DATA ของเดือนที่ระบุ (Meta) ให้ครบทุกวัน
ใช้เมื่อกราฟ Daily Trend ของเดือนที่ปิดแล้ว "เข้าไม่ครบ" (ข้อมูลขาดปลายเดือน)

⚠️ แตะเฉพาะ RAW_DATA เท่านั้น — ไม่ยุ่งกับ AO_DATA / META_AD_PERF (เดือนปัจจุบันปลอดภัย)

รัน:
  python3 05_Scripts/backfill_raw.py 2026-05      # เติมทั้งเดือน พ.ค.
  python3 05_Scripts/backfill_raw.py 2026-05-29 2026-05-31   # เฉพาะช่วง
"""
import sys, re, json, calendar
import meta_pull  # reuse fetch + transform (มี __main__ guard ไม่รัน main ตอน import)

DASHBOARD = meta_pull.DASHBOARD_PATH

# ── ช่วงวันที่ ──
if len(sys.argv) >= 3:
    DATE_FROM, DATE_TO = sys.argv[1], sys.argv[2]
elif len(sys.argv) == 2:
    ym = sys.argv[1]                      # YYYY-MM
    y, m = map(int, ym.split("-")[:2])
    last = calendar.monthrange(y, m)[1]
    DATE_FROM, DATE_TO = f"{ym}-01", f"{ym}-{last:02d}"
else:
    print("ใช้: python3 backfill_raw.py 2026-05   (หรือ 2026-05-29 2026-05-31)")
    sys.exit(1)

print(f"=== Backfill RAW_DATA: {DATE_FROM} → {DATE_TO} ===")

# override ช่วงวันที่ของ meta_pull แล้วเรียก fetch/transform เดิม
meta_pull.DATE_FROM = DATE_FROM
meta_pull.DATE_TO   = DATE_TO
raw  = meta_pull.pull_ad_insights()
rows = meta_pull.transform_rows(raw)
meta_rows = [r for r in rows if r.get("platform") == "Meta"]
if not meta_rows:
    print("⚠️  ไม่ได้ records เลย (token หมดอายุ? / ไม่มี spend ช่วงนี้?) — ยกเลิก")
    sys.exit(1)

# เดือนเป้าหมาย (จาก records) — อาจครอบหลายเดือนถ้าช่วงคร่อมเดือน
tgt_months = sorted(set(r.get("month") for r in meta_rows if r.get("month")))
new_dates  = sorted(set(r.get("ad_date") for r in meta_rows if r.get("ad_date")))
print(f"   ได้ {len(meta_rows)} records · เดือน {tgt_months} · {new_dates[0]} → {new_dates[-1]}")

# ── merge เฉพาะ RAW_DATA ──
with open(DASHBOARD, "r", encoding="utf-8") as f:
    content = f.read()
m = re.search(r'const RAW_DATA = (\[.*?\]);', content, re.DOTALL)
if not m:
    print("❌ ไม่พบ const RAW_DATA"); sys.exit(1)
existing = json.loads(m.group(1))

# เก็บทุกอย่างไว้ ยกเว้น Meta ของเดือนเป้าหมาย (จะแทนด้วยชุดใหม่ที่ครบกว่า)
kept = [r for r in existing
        if not (r.get("platform") == "Meta" and r.get("month") in tgt_months)]
merged = kept + meta_rows
merged.sort(key=lambda r: (r.get("month",""), r.get("ad_date",""), r.get("creator","")))

new_json = json.dumps(merged, ensure_ascii=False, separators=(",", ":"))
new_content = re.sub(r'const RAW_DATA = \[.*?\];',
                     f'const RAW_DATA = {new_json};', content, flags=re.DOTALL)
if new_content == content:
    print("⚠️  ไม่มีการเปลี่ยนแปลง"); sys.exit(0)
with open(DASHBOARD, "w", encoding="utf-8") as f:
    f.write(new_content)
print(f"✅ เติม RAW_DATA แล้ว ({len(merged)} records รวม) — กด Cmd+Shift+R")
