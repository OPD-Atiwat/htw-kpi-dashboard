# 🧠 HTW Content KPI Dashboard — Project Context
> อัปเดตล่าสุด: 4 เมษายน 2026
> **วิธีใช้:** ทุกครั้งที่เริ่มเซสชันใหม่ ให้บอก Claude ว่า "อ่าน PROJECT_CONTEXT.md ก่อนนะ" เพื่อให้รู้ context ทั้งหมดทันที

## 📂 Memory Files (อ่านเพิ่มเติม)
| ไฟล์ | เนื้อหา |
|---|---|
| `PROJECT_CONTEXT.md` | โครงสร้างระบบ, logic สำคัญ, bug ที่เคยเจอ (ไฟล์นี้) |
| `HTW_DATA_MEMORY.md` | **เป้าหมายยอดขายรายเดือน (GOALS)** + **แผน Content Calendar (KMS)** ทั้งหมด |

> ถ้าต้องการรู้เป้ารายเดือนหรือดูแผน Content ให้อ่าน `HTW_DATA_MEMORY.md` ด้วย

---

## 👤 เกี่ยวกับผู้ใช้
- **ชื่อ:** พี่ (How-to / how-to.cowork@opendurian.com)
- **ตำแหน่ง:** Digital Marketing Supervisor ตลาด How-to
- **บริษัท:** OpenDurian
- **งานหลัก:** ดูแล Content Marketing สายหนังสือ How-to/Education ผ่าน Meta Ads และ TikTok Ads

---

## 📁 โครงสร้าง Folder
```
📁 Excel: Content KPI Dashboard/
│
├── PROJECT_CONTEXT.md              ← ไฟล์นี้ (อ่านก่อนทุกครั้ง)
├── HTW_DATA_MEMORY.md              ← ข้อมูล GOALS + KMS Calendar
├── index.html                      ← Dashboard สำหรับ GitHub Pages (copy จาก 01_Dashboard)
│
├── 01_Dashboard/
│   └── index.html                  ← ✅ Dashboard หลัก (HTW_Content_KPI_Dashboard_Q1FY26)
│
├── 02_Database/                    ← ฐานข้อมูล Export
│   ├── HTW_ContentDB_Master.xlsx
│   └── HTW_ContentDB_Export_Full.xlsx
├── 03_Raw Data/                    ← ข้อมูลดิบรายเดือน (Meta/TikTok CSV)
│   ├── HTW_AdsPerformance_Feb26_IMPORT.csv
│   ├── HTW_Ads_Performance_ALL_Feb26.xlsx
│   ├── HTW_Ads_Performance_Meta_Feb26.xlsx
│   ├── HTW_ContentDB_Feb26_IMPORT.csv
│   └── HTW_Meta_Riw_Feb26_Mapped.xlsx
├── 04_TikTok Mapping/              ← ผลการ map TikTok VDO ID
├── 05_Scripts/                     ← Google Apps Script
└── 06_Assets/                      ← รูปภาพ Creator (ริว, มิ้น, หมิว)
```

---

## 🌐 GitHub Pages
- **Repo:** `OPD-Atiwat/htw-kpi-dashboard`
- **URL:** `https://opd-atiwat.github.io/htw-kpi-dashboard`
- **วิธีอัปเดต:** upload `index.html` (root) ทับใน GitHub repo → หน้าเว็บอัปเดตใน 1-2 นาที

---

## 🎯 Dashboard หลัก: `01_Dashboard/index.html`
Single-file HTML ที่มีทุกอย่างฝังอยู่ใน `<script>` — ไม่ต้องพึ่ง server

### ข้อมูลที่ฝังอยู่
| ตัวแปร | คำอธิบาย | ขนาด |
|---|---|---|
| `RAW_DATA` | ข้อมูลโฆษณารายวัน รายคอนเทนต์ | ~1,649+ rows (Feb-Mar 2026) |
| `GOALS_DATA` | เป้าหมาย KPI รายเดือน | Feb 26, Mar 26, Apr 26 |
| `MONTHLY_FINANCE_GOALS` | เป้าจาก Finance (รวม Consign + Bookfair) | Feb 26, Mar 26, Apr 26 |
| `OPD_DAILY` | ยอด Online รายวัน (ทุก channel รวม) | Feb-Mar 2026 + Apr partial |
| `KMS_CAL_DATA` | แผนงาน Content Calendar | **Feb 26, Mar 26 เท่านั้น** (Apr ยังไม่มีข้อมูล) |
| `WORKLOAD_DATA` | สรุป Workload ต่อ Product ต่อเดือน | Feb 26, Mar 26 |

### Structure ของ RAW_DATA (แต่ละ row)
```json
{
  "month": "Feb26",
  "creator": "มิ้น",
  "ad_date": "2026-02-28",
  "content_name": "ชื่อคอนเทนต์",
  "product": "ชื่อหนังสือ",
  "platform": "Meta",
  "launch_date": "2026-02-04",
  "is_new": false,
  "hit_roas": "ผ่าน | ไม่ผ่าน | ไม่มียอดขาย",
  "reach": 0, "impressions": 0, "clicks": 0,
  "ctr": 0.0, "cpm_thb": 0.0,
  "spend_thb": 0.0, "revenue_thb": 0.0, "roas": 0.0,
  "purchases": 0, "cpa_thb": 0.0,
  "messages": 0, "v2s": 0, "v6s": 0, "ad_variants": 1
}
```

---

## 💰 MONTHLY_FINANCE_GOALS (Finance เป้าจริง)
```javascript
var MONTHLY_FINANCE_GOALS = {
  'Feb 26': { totalGoal:5428571, onlineGoal:3800000, consignGoal:1628571, bookfairGoal:0,      profitGoal:599858  },
  'Mar 26': { totalGoal:6892899, onlineGoal:5000000, consignGoal:1311250, bookfairGoal:581649, profitGoal:897500  },
  'Apr 26': { totalGoal:6900000, onlineGoal:5100000, consignGoal:1800000, bookfairGoal:600000, profitGoal:1308000 }
};
```
> ⚠️ **Apr 26 onlineGoal = 5,100,000** (= 4,500,000 product-sum + 600,000 Bookfair รวมอยู่ด้วย)
> ใน Goal card ใช้ `effGoal = MONTHLY_FINANCE_GOALS[month].onlineGoal` เป็น denominator เสมอ

---

## 👥 ครีเอเตอร์
| Creator | Format | แนวทาง |
|---|---|---|
| **ริว** | VDO | คอนเทนต์วิดีโอ, Meta + TikTok |
| **มิ้น** | VDO | คอนเทนต์วิดีโอ, Meta + TikTok |
| **หมิว** | IMG | คอนเทนต์ภาพนิ่ง (Image), Meta |

> ⚠️ `content_format` ถูกกำหนดจาก creator: `หมิว = IMG`, อื่นๆ = `VDO`

---

## ⚙️ CONFIG สำคัญ (ใน Dashboard)
```javascript
var CONFIG = {
  ROAS_PASS:        2.0,   // VDO per-day threshold (ริว/มิ้น)
  ROAS_PASS_IMG:    1.5,   // IMG per-day threshold (หมิว) — ภาพได้ ROAS ต่ำกว่า
  SUCCESS_REVENUE:  30000, // คอนเทนต์ปัง: ยอดขาย ≥ ฿30,000
  SUCCESS_ROAS:     2.0,   // คอนเทนต์ปัง: ROAS ≥ 2.0x (VDO)
  SUCCESS_ROAS_IMG: 1.5,   // คอนเทนต์ปัง: ROAS ≥ 1.5x (IMG)
};
```

---

## 📊 Logic สำคัญ

### Pass Rate (รายวัน)
- ใช้ `hit_roas` จาก RAW_DATA **โดยตรง** (pre-computed, ห้าม recalculate ใน normalizeRow)
- `passDates` = Set ของ ad_date ที่มีอย่างน้อย 1 row ที่ `hit_roas === "ผ่าน"`
- Pass Rate = `passDates.size / dates.size`

### isSuccess (Content-level ⭐)
- คำนวณ `aggRoas = c.rev / c.spend` (aggregate ทั้งคอนเทนต์)
- `c.isSuccess = (c.rev >= 30000) && (aggRoas >= succRoas)`
- succRoas: IMG ใช้ 1.5x, VDO ใช้ 2.0x

### effGoal (Overview + Goal tab)
```javascript
// ทั้ง renderOverview() และ renderGoal() ใช้ pattern เดียวกัน
var _mfg = MONTHLY_FINANCE_GOALS[month];
var effGoal = (_mfg && _mfg.onlineGoal > 0) ? _mfg.onlineGoal : totalGoal;
// ไม่ใช้ goalTotal (product sum) โดยตรง เพราะไม่รวม Bookfair
```

### Forecast
```javascript
// calcMonthForecast(month) → { actual, days, totalDays, dailyAvg, projected, goalsActual, scaleFactor }
// ใช้ OPD_DAILY เป็น source, ถ้า days < totalDays แสดง forecast bar สีม่วงในกราฟ
```

### Content Format Detection
```javascript
// ใน normalizeRow()
o.content_format = (o.creator === 'หมิว') ? 'IMG' : 'VDO';
// ❌ อย่า recalculate hit_roas ที่นี่ — จะทำให้ Pass Rate พุ่งเป็น 100%
```

---

## 🐛 Bug ที่เคยพบ & วิธีแก้

| Bug | สาเหตุ | วิธีแก้ |
|---|---|---|
| Pass Rate 100% | recalculate hit_roas ใน normalizeRow ด้วย o.roas (aggregate) | ลบ recalculate ออก ใช้ค่าเดิมจาก RAW_DATA |
| Calendar ไม่แสดงข้อมูล | STATE.period ไม่มีใน dashboard | ดึงจาก dropdown `filter-month` + `filter-year` แทน |
| "nan" ใน Calendar chips | Python pandas NaN → str('nan') | clean() function กรอง nan/none ออก |
| TikTok ID ไม่ match | float64 ตัดตัวเลขท้าย 19-digit ID | match ด้วย `str(id)[:15]` prefix |
| Overview เป้าแสดง ฿4.50M แทน ฿5.10M | renderOverview() ใช้ goalTotal แทน Finance onlineGoal | เพิ่ม `_ovMfg` lookup + `ovEffGoal` pattern |
| Calendar cell แสดง raw JSON string | `JSON.stringify(JSON.stringify(item))` มี `"` หลุดออก HTML attribute | เปลี่ยนเป็น `_CAL_ITEMS_STORE[]` array + pass index แทน |
| Calendar grid เพี้ยน (cell ขยายข้ามคอลัมน์) | CSS `repeat(7,1fr)` ไม่ lock minimum width | เปลี่ยนเป็น `repeat(7,minmax(0,1fr))` + `min-width:0;overflow:hidden` |
| Calendar แสดงเมษาแทนมีนา | `KMS_CAL_DATA` มี Apr 26 placeholder 116 items | ลบ Apr 26 ออกจาก KMS_CAL_DATA (ยังไม่มีข้อมูลจริง) |

---

## 📅 ข้อมูลที่มีในระบบ (ณ 4 เม.ย. 2026)
| เดือน | RAW_DATA (Ads) | OPD_DAILY | KMS Calendar | GOALS |
|---|---|---|---|---|
| **Feb 26** | ✅ เต็มเดือน | ✅ เต็มเดือน | ✅ มีข้อมูล | ✅ |
| **Mar 26** | ✅ เต็มเดือน | ✅ เต็มเดือน | ✅ มีข้อมูล | ✅ |
| **Apr 26** | ❌ ยังไม่มี | ⚠️ มีบางส่วน (1-3 เม.ย.) | ❌ ยังไม่ได้ส่ง | ✅ (เป้าเท่านั้น) |

---

## 🗂️ งานที่เคยทำใน Project นี้
1. ✅ สร้าง Dashboard HTML (single-file) ติดตาม KPI Content
2. ✅ เพิ่ม Content Calendar (Visual Plan) จาก KMS files
3. ✅ แก้ Calendar bugs (nan, creator filter, JSON injection, grid layout)
4. ✅ กำหนด Success Content metric (Rev ≥ ฿30K + ROAS ≥ 2.0x/1.5x)
5. ✅ แยก ROAS threshold IMG vs VDO
6. ✅ แก้ 100% Pass Rate bug
7. ✅ TikTok ID Mapping Excel (Mar 26) — 48 matched, 11 no ID, 9 not found
8. ✅ Export RAW_DATA → Excel Database
9. ✅ จัดระเบียบ Folder structure (01–06)
10. ✅ เพิ่ม Forecast bar chart + Forecast card
11. ✅ เชื่อม Finance goals (MONTHLY_FINANCE_GOALS) กับทุก tab
12. ✅ เพิ่ม ยอดออนไลน์ card ใน Overview
13. ✅ Deploy บน GitHub Pages (`https://opd-atiwat.github.io/htw-kpi-dashboard`)
14. ✅ เพิ่มปุ่ม ◀ ▶ สลับเดือนใน Calendar
15. ✅ แก้ Calendar month default (ลบ Apr 26 placeholder ออก)

---

## 💡 แนวทางอนาคต
- **6 เดือน:** ยังใช้ระบบนี้ได้ปกติ (ประมาณ 5,000 rows)
- **12 เดือน:** พิจารณาแยก Dashboard ตามปี หรือ migrate ข้อมูลเก่าออก
- **ระยะยาว:** ย้ายไปใช้ Google Sheets + Apps Script หรือ Database จริง (Airtable/Supabase)
- **Skill ที่น่าสร้าง:** "HTW Dashboard Updater" — รับ CSV ดิบ → อัปเดต OPD_DAILY + GOALS_DATA + KMS_CAL_DATA อัตโนมัติ

---

> 📌 **วิธีเพิ่มข้อมูลเดือนใหม่เข้า Dashboard:**
> 1. Export ข้อมูลจาก Meta / TikTok เป็น CSV
> 2. แปลงเป็น JSON format ตาม structure RAW_DATA ด้านบน
> 3. เพิ่ม rows ต่อท้ายใน `const RAW_DATA = [...]` ใน HTML file
> 4. Update `OPD_DAILY` ด้วยยอดรายวันใหม่
> 5. Update `KMS_CAL_DATA` ถ้ามีแผนเดือนใหม่
> 6. Copy `01_Dashboard/index.html` → root `index.html` แล้ว upload ขึ้น GitHub
