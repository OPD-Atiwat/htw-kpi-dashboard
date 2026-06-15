# 🌌 HOWTO UNIVERSE — สมุดกลาง

> ศูนย์กลางความรู้ของ OpenDurian How-to Dashboard · อ่านไฟล์นี้ก่อนเริ่มงานทุกครั้ง
> เจ้าของ: Digital Marketing Supervisor (ทีม How-to) · อัปเดตล่าสุด: 15 มิ.ย. 2026

---

## 📚 สารบัญเอกสาร (อ่านต่อตามต้องการ)
| ไฟล์ | เนื้อหา |
|---|---|
| `PROJECT_CONTEXT.md` | โครงสร้างธุรกิจ/โปรเจกต์ภาพรวม |
| `HTW_DATA_MEMORY.md` | GOALS + KMS + ข้อมูลอ้างอิงรายเล่ม |
| `DASHBOARD_DEV_KNOWLEDGE.md` | เทคนิคพัฒนา index.html (architecture, data blocks, gotchas, bug log) |
| `index.html` | **Dashboard = แหล่งข้อมูลเดียวที่เชื่อถือได้** |

---

## 🧱 โครงสร้างข้อมูลใน index.html (data blocks)
`OPD_DAILY` (ยอดรายวัน+`_q`=จำนวนเล่ม) · `OPD_PROD_QTY`/`OPD_PROD_DATA` (รายเล่ม) · `GOALS_DATA` (เป้ารายเล่ม — ชื่อเต็มเป็น truth) · `ADSM_DAILY`/`ADSM44_PCTADS` (ค่าแอด, ตัวหาร %Ads) · `PROFIT_DATA` · `CONSIGN_DATA` · `MONTHLY_FINANCE_GOALS` (เป้ารวม Consign+Bookfair) · `CONTENT_LOG` · `TT_CREATOR_REV` (creator TikTok — manual) · `CREATOR_META_<เดือน>` (creator Meta — auto) · `RAW_DATA` (ads ราย ad)

---

## 🚫 กฎห้ามละเมิด
1. **ชื่อหนังสือ** ใช้ชื่อเต็มจาก `GOALS_DATA` เท่านั้น ห้ามย่อ/เปลี่ยน
2. **Bookfair** ห้ามลบ/overwrite — หักออกจากยอดหน้าร้านกัน double-count (ค่าจาก event report ไม่ใช่ MK13)
3. **MTD Comparison** เทียบ same-period เสมอ (เช่น มิ.ย. 1-N vs พ.ค. 1-N) ไม่ใช่เดือนเต็ม
4. **แก้ index.html** → git commit local ก่อนเสมอ, ทำ local ให้เสร็จแล้ว push ทีเดียว (ระวัง .git/*.lock จาก auto-sync ค้างบ่อย)
5. **%Ads/Margin** ระวังตัวหาร (ADSM44 ≠ MK29) — ดู skill `pct-ads-calc`

---

## 🔄 Pipeline & Automation
- **Auto-sync**: `05_Scripts/opd_runner.py` รันเป็นช่วงๆ → meta_pull → opd_pull → kms_pull → patch_thumbnails → fetch_previews → git push (log: `05_Scripts/meta_pull.log`)
- **Scheduled agents** (รันเมื่อเปิดแอป):
  - `daily-sales-margin-brief` — ทุกวัน 10:00 (สถานะ: ทดลอง, ยังไม่ finalize prompt)
  - `weekly-creator-coaching` — จ/พ/ศ 10:00
- **Creator attribution**: Meta auto (CREATOR_META) · TikTok manual (TT_CREATOR_REV ต้อง map จาก export) · **TikTok ยังไม่มี API auto-pull**

---

## 🧠 Logic สำคัญ (อย่าลืม)
- **Creator 50/50**: ชื่อ Ad มีหลายครีทีมเรา `[มิ้น][ริว]` → แชร์ทุก metric เท่ากัน (เฉพาะ ริว/หมิว/มิ้น/แนน/แก้ม; Central/Influ/Cross Page ไม่แชร์) — `meta_pull.guess_creators()`
- **Run Rate จำนวนเล่ม**: อัตรา = เล่มสะสม ÷ **วันที่มีข้อมูลจริง** (cap วันสุดท้ายที่มียอด ไม่หารด้วยวันที่ยังไม่ sync) · คาดปลายเดือน = อัตรา × วันทั้งเดือน
- **Data lag**: MK13 เข้าของวันก่อนหน้า (ยอดวันนี้มาพรุ่งนี้)
- **Matching gap**: ยอดรวม (OPD_DAILY) ครบ แต่ผลรวมรายเล่ม (OPD_PROD_*) ขาด ~8% เพราะบางชื่อใน MK13 map ไม่เข้า — ยอดภาพรวมเชื่อได้, รายเล่มอาจขาดเล็กน้อย

---

## 📝 Changelog
### 15 มิ.ย. 2026
- เพิ่ม **Run Rate จำนวนเล่ม** (แยกเล่ม+ช่องทาง+ภาพรวม): คอลัมน์เล่มในตารางช่องทาง, การ์ด ขายแล้ว/อัตรา/คาดการณ์, กราฟเล่มรายวัน (Chart.js), คอลัมน์ vs เดือนก่อน same-period, ▲% ใต้ทุกการ์ดสรุป
- **Fix**: run rate per-book หารด้วยวันที่มีข้อมูลจริง (cap lastDate) — เดิมหารเกินวัน
- **50/50 creator split** ใน `meta_pull.py` (guess_creators + แชร์ทุก metric ใน CREATOR_META)
- ตั้ง scheduled agents: Daily Sales & Margin Brief, Creator Coaching
- ตรวจ reconcile: OPD_DAILY ≈ หลังบ้าน (ต่าง ~0.2%), รายเล่มขาด ~8% (matching gap)

<!-- เพิ่ม entry ใหม่ด้านบนสุดของส่วนนี้ทุกครั้งที่แก้ dashboard/pipeline -->
