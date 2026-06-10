# 🛠️ Dashboard Development Knowledge — index.html

> ความรู้เชิงเทคนิคสำหรับพัฒนา/แก้ index.html (HTW Content KPI Dashboard)
> อัปเดตล่าสุด: 10 มิ.ย. 2026 · คู่กับ `PROJECT_CONTEXT.md` (โครงสร้างธุรกิจ) และ `HTW_DATA_MEMORY.md` (GOALS + KMS)

---

## 1. สถาปัตยกรรม (Architecture)

- **Single-file HTML** — index.html (~16,200 บรรทัด, ~14 MB) รวม HTML + CSS + JS + data ในไฟล์เดียว
- **Data = JavaScript var blocks** ฝังไว้ตอนต้นไฟล์ (บรรทัด ~1560–3000) ไม่ได้ fetch จากภายนอกตอน runtime → script (Python) เขียนทับ block เหล่านี้
- **Deploy:** copy root `index.html` → GitHub Pages (`https://opd-atiwat.github.io/htw-kpi-dashboard`)
- **Font:** IBM Plex Sans Thai (Google Fonts)

---

## 2. Data Blocks (var declarations ใน index.html)

| Block | บรรทัด~ | เนื้อหา |
|---|---|---|
| `RAW_DATA` | 1563 | Ads performance รายแถว (Meta/TikTok) |
| `AFFILIATE_DATA` | 1566 | ← affiliate_data.json |
| `META_PULL_TS` | 1569 | timestamp sync ล่าสุด (failsafe เขียนทุก run) |
| `GOALS_DATA` | 1576 | เป้ารายเล่ม (ชื่อหนังสือเต็ม — แหล่ง truth ของชื่อ) |
| `MARKETING_STRATEGY` | 1577 | กลยุทธ์รายเดือน |
| `OPD_DAILY` | 1641 | ยอดขายรายวัน (จาก MK13) |
| `OPD_PROD_DATA` / `OPD_PROD_NAME_MAP` | 1642 | ยอดรายเล่ม + map ชื่อ |
| `ADSM_DAILY` / `ADSM_PROD_DAILY` | 1680 | Ad spend (ADSM44) ตัวหารของ %Ads |
| `PROFIT_DATA` | 1762 | กำไร & margin |
| `CONSIGN_DATA` / `CONSIGN_PROD_DATA` | 1856 | Consignment (หลังปก C1) |
| `MONTHLY_FINANCE_GOALS` | 2075 | เป้า Finance (รวม Consign + Bookfair) |
| `CONTENT_LOG` | 2592 | กิจกรรม content |
| `TT_CREATOR_REV` | 2938 | Creator revenue (ริว/มิ้น) |
| `KMS_CAL_DATA` | 2939 | Content calendar |
| `BOOK_COVERS` / `CREATOR_PHOTOS` | 2565/2587 | รูป |

> ⚠️ **ชื่อหนังสือ:** ใช้ชื่อเต็มจาก `GOALS_DATA` เท่านั้น ห้ามย่อ
> ⚠️ **Bookfair:** อยู่แยกใน finance goals — ห้ามลบ/overwrite ต้องหักจากยอดหน้าร้านกัน double-count

---

## 3. Gotchas — จุดพังบ่อย ต้องระวัง

### 3.1 `fmt()` ใส่ `฿` ให้อยู่แล้ว → อย่าเติม `'฿'+` ซ้ำ
```js
function fmt(n){ return '฿'+Math.round(n).toLocaleString(); }
```
- มี `fmt()` ซ้ำหลายตัว (บรรทัด 3043, 14759, 15020, 15103) — ทุกตัวเติม `฿` มาแล้ว
- บั๊ก `฿฿` เกิดจากโค้ดเดิมเขียน `'฿'+fmt(...)` → ลบ `'฿'+` ข้างหน้าออก
- เช็คทุก section ที่แสดงเงิน เพราะ fmt ถูกเรียกหลายที่

### 3.2 Grid columns ต้อง responsive
- Upsell/ตาราง: ใช้ **5 คอลัมน์** จอแคบลดอัตโนมัติ 4/3/2/1
- Calendar grid ต้อง `repeat(7,minmax(0,1fr))` + `min-width:0;overflow:hidden` (ไม่ใช่ `1fr`) ไม่งั้น cell ขยายข้ามคอลัมน์

### 3.3 MTD Comparison ต้อง same-period
- เทียบ **May 1–N vs Apr 1–N** เสมอ (ไม่ใช่เดือนเต็ม)
- จุดที่ต้องตรวจ: `_prevPmTotal`, `_prevChTot`, `_prevConsignMtd`/`_prevBfMtd`, `_prevGdataForProd`

### 3.4 %Ads / Margin — ระวังตัวหาร
- %Ads ใช้ตัวหารถูก (ADSM44 ≠ MK29) — ดู skill `pct-ads-calc`
- spend=0 → chip "ไม่ได้ยิงแอด" (ไม่ใช่ data หาย ถ้ากราฟ ad จบที่วันที่ปิดแอดถือว่าถูก)

---

## 4. Bug Log (ที่เคยแก้)

| Bug | สาเหตุ | วิธีแก้ |
|---|---|---|
| `฿฿` ซ้ำ | โค้ดเดิม `'฿'+fmt(...)` ทั้งที่ fmt เติม ฿ แล้ว | ลบ `'฿'+` ออก |
| Pass Rate 100% | recalc hit_roas ด้วย aggregate roas | ใช้ค่าเดิมจาก RAW_DATA |
| Calendar ไม่แสดง | STATE.period ไม่มี | อ่านจาก dropdown `filter-month`+`filter-year` |
| "nan" chips | pandas NaN → 'nan' | `clean()` กรอง nan/none |
| TikTok ID ไม่ match | float64 ตัดเลขท้าย 19-digit | match ด้วย `str(id)[:15]` |
| เป้า ฿4.50M แทน ฿5.10M | ใช้ goalTotal แทน Finance onlineGoal | `_ovMfg` lookup + `ovEffGoal` |
| Calendar raw JSON ใน cell | double `JSON.stringify` หลุด `"` | `_CAL_ITEMS_STORE[]` + pass index |

---

## 5. Auto-Sync Pipeline (`05_Scripts/opd_runner.py`)

ลำดับ run → เขียนทับ data blocks ใน index.html → git push:

1. `meta_pull.py` — Meta Ads + MK13 sync
2. `opd_pull.py` — Affiliate + Goals
3. `kms_pull.py` — KMS Content Sheet
4. `patch_ao_thumbnails.py` — thumbnails + sheet mapping
5. `fetch_previews.py` — ad preview URLs
6. `fetch_video_sources.py` — MP4 ตรง (เล่นในหน้าได้)

- **Failsafe:** runner เขียน `META_PULL_TS` = เวลาปัจจุบันทุก run กัน timestamp ค้างเมื่อ sub-script fail เงียบ
- **Python path ไม่มี colon:** runner วางที่ `~/opd_runner.py` (path workspace มี `:` รัน via shell ไม่ได้ — Python จัดการ path มี colon ได้ตรง)
- **Git:** ลบ `*.lock` → add → commit `auto: <ts>` → `pull --rebase -X ours` → push (`--force-with-lease` ถ้า rebase abort)
- **Log:** `05_Scripts/meta_pull.log` (ดู skill `dashboard-health-check`)

### Trigger / Schedule
- LaunchAgents: `com.opendurian.dailyrun` / `.hourly` / `.meta-ads-alert` / `.sync-trigger`
- สั่ง sync ทันที: เขียน `.sync_trigger` file (skill `howto-sync-trigger`)

---

## 6. ก่อนแก้ index.html — เช็คลิสต์บังคับ

1. `git status` + commit local ก่อนทุกครั้ง
2. ทำ local ให้เสร็จทั้งหมด → push ทีเดียว (ห้าม push กลางคัน)
3. หลังแก้ → verify syntax (เปิด/grep) ก่อนส่ง
4. บันทึกไฟล์ลง workspace folder เสมอ

---

## 7. วิธีเพิ่มข้อมูลเดือนใหม่

ปกติ auto-sync จัดการให้. ถ้าทำมือ: export CSV → แปลงเป็น JSON ตาม structure → เขียนทับ block (`RAW_DATA`/`OPD_DAILY`/`KMS_CAL_DATA`) → commit → push.

---

## 8. แนวทางอนาคต / scale
- ~5,000 rows (6 เดือน) ยังไหว; 12 เดือนพิจารณาแยก dashboard รายปี
- ระยะยาว: ย้าย data ออกจาก HTML → Google Sheets/Supabase/Airtable
