# 📍 BEP Knowledge — สรุปจากแชต [How-to] BEP

> รวบความรู้ทั้งหมดเรื่อง BEP (และจุดที่แตะ Dashboard) จากแชตที่ปักหมุดไว้

---

## 1. หลักการตั้งราคา (Pricing & Gap Rule)

- **Gap Rule บังคับ:** ราคาต้องห่างกันทุกช่องทาง เรียงจากถูกไปแพง
  `FB < TikTok < Shopee < C1 (หลังปก)` — เน้น Facebook กับ TikTok เป็นช่องหลัก
- **Margin คำนวณจาก List Price เท่านั้น** ⚠️ (เคยพลาดมาแล้ว สำคัญมาก)
  - สูตร: `Margin = LP − Fix − LP×(VarPct + ShipPct + AdsPct)`
  - **Discount ไม่ลด Margin** — ส่วนลด 15% (Shopee/C1/C2) เป็นแค่ Reference เช็ค Gap หลังลดเท่านั้น
- **ทุกช่องทางต้อง Margin เป็นบวก** ห้ามติดลบ
- **ราคาทุกช่องทางต้องอ้างอิง COG** (ต้นทุนพิมพ์/เล่ม)
- ตั้งราคา **หลังปก (C1) ก่อน** แล้วไล่ Gap ลงมา → Shopee → TikTok → FB

---

## 2. โครงสร้าง Cost ต่อช่องทาง

- **ค่าส่งแยกแสดงจากราคาขาย** — ลูกค้าจ่ายจริง = ราคา + ค่าส่ง (ต้องเทียบราคาแบบ "รวมส่ง" เสมอ)
- แต่ละช่องมี: Fix Cost/Unit (COG Variable), %Var, %Shipping, %Ads ต่างกัน
- Shopee/C1/C2 มีส่วนลด 15% เป็น Reference เช็ค Gap

### ตัวอย่าง Channel Margin — Survival Guide (Final)
| Channel | List | รวมส่ง | Margin/Unit | Margin% |
|---|---|---|---|---|
| Facebook | ฿269 | ฿319 | ฿55.90 | 20.78% |
| TikTok | ฿319/329 | ฿345 | ฿8.04 | 2.52% ⚠️ ลงอีกไม่ได้ ติดลบทันที |
| TikTok Live | ฿249 | ฿265 | ฿71.44 | 28.69% |
| Shopee | ฿349/379 | ฿395 | ฿131.85 | 37.78% |
| C1 (หลังปก) | ฿429/419 | ฿419 | ฿216.49 | 50.46% |
| C2 | ฿299 | — | ฿147.02 | 49.17% |

---

## 3. Forecast Structure

- **Input เดียวคือ PEAK units (cell C5)** → ทุกเดือนคำนวณสัดส่วนอัตโนมัติ
  ไม่กรอก units ทีละเดือน — แก้แค่ PEAK cell เดียว แล้ว decay formula วิ่งเอง
- **เดือนแรก Prorated** ตามวันจริง (Launch กลางเดือน ÷ 30) เช่น Launch 14 Sep → factor = 17/30
- **Peak 4 เดือนแรก** (factor 1.0) จากนั้น **Decay** ลงเรื่อยๆ: 80% → 65% → 55% → ... → 5% จนเดือน 24
- **Channel Mix %** กำหนดสัดส่วนการขายแต่ละช่อง
  (ตัวอย่าง SG: FB 58% / TikTok 22% / TKLive 2% / Shopee 12% / C1 4% / C2 2%)

---

## 4. Consignment Rules (ต้องถามทุกเล่ม)

- **C1 = 0 ในเดือนแรก** (ยังไม่ได้วางสต็อกที่ร้านหนังสือ)
- **C2 มียอดเฉพาะบางเดือน** — งานหนังสือ (Book Fair) เช่น เฉพาะ **APR + OCT** (ตัวอย่าง SG ไม่รวม DEC)
- Skill/Model ต้องถามว่าเล่มนี้มี C1/C2 ไหม เริ่มเดือนไหน

---

## 5. Target & BEP

- **แต่ละโปรเจกต์มี Launch Date, GM Target, Fixed Cost ต่างกัน** — ห้าม assume เท่ากัน
- **Column O = Cumul GM vs GM Target เท่านั้น** (ไม่ใช่ Full BEP)
- **Full BEP Hurdle = Fixed Cost + GM Target รวมกัน**
- ตัวอย่าง SG: GM Target Year 1 = ฿875,000 (ถึงภายใน ธ.ค. 2026)
  - Full BEP = ฿1,240,712 (Fixed ฿365,712 + GM ฿875,000)
  - PEAK ขั้นต่ำ = 4,179 → Cumul GM ธ.ค. 2026 = ฿875,049 ✅ (สูงกว่าเดิม 3,932 เพราะ Sep ไม่มี C1)

---

## 6. 🗺️ Journey — ตั้งแต่เปิดจนปิดโครงการ

**Phase 1 — Pre-Launch**
1. กำหนด COG Variable + Fixed Production Cost
2. ตั้งราคาหลังปก (C1) ก่อน แล้วไล่ Gap ลง → Shopee → TikTok → FB
3. ตรวจ Margin ทุกช่อง ต้องบวกทุกตัว
4. กำหนด Launch Date → คำนวณ Prorated Month 1
5. กำหนด GM Target Year 1 + Fixed Cost รวม
6. ถาม Consignment Rule (C1/C2 มีไหม เริ่มเดือนไหน)

**Phase 2 — Build Model**
7. สร้าง Channel Table (List, Disc, Shipping, Fix, %Var, %Ads, Mix, Margin)
8. คำนวณ Blended Margin จาก Mix %
9. หา PEAK ขั้นต่ำที่ทำให้ Cumul GM ถึง Target ภายในเดือนที่กำหนด
10. สร้าง Forecast 24 เดือน (Prorated → Peak → Decay)
11. ตั้ง Column O track vs GM Target

**Phase 3 — Tracking**
12. อัปเดตยอดจริงรายเดือนเทียบ Forecast
13. ดู Column O ว่า Cumul GM ถึงเป้าตามแผนไหม
14. ถ้าต่ำกว่าแผน → เพิ่ม PEAK หรือปรับ Mix

**Phase 4 — Close**
15. Cumul GM ≥ Target → บันทึก BEP Date จริง
16. เทียบ BEP Date จริง vs แผน
17. สรุป Total GM, Total Revenue, Blended Margin% ตลอด 24 เดือน

---

## 7. บทเรียน Case Study (ใช้ตอบทีมผลิต)

**เรื่องจำนวนหน้าน้อยกว่าแต่ตั้งราคาสูง (Survival 144 หน้า vs ออฟฟิศ 176 หน้า):**
- ต้นทุนพิมพ์/หน้า: ออฟฟิศ ฿0.097 vs Survival ฿0.127 (**+31%**) → SG production quality สูงกว่า ราคาสูงกว่าจึงมีเหตุผลรองรับ
- ความเสี่ยงจริงอยู่ "หลังได้ของ" (รีวิวว่าบาง) ไม่ใช่ตอนตัดสินใจซื้อ → แก้ด้วย **creative** (flip โชว์ความหนาแน่นเนื้อหา) มากกว่าลดราคา
- ปรับราคาเน้น **FB** (margin มีห้อง) ไม่แตะ TikTok (margin 1-2% ติดผนัง)
- **ระวัง GM buffer บาง** — v8 เหลือ buffer แค่ ฿380 พลาดเป้าง่าย ใช้ v7 (buffer ฿1,838) ดีกว่า
- **TikTok Shop policy:** ราคาห้ามแพงกว่าช่องอื่น ถ้า FB ถูกกว่าเยอะอาจโดน flag/ลด visibility

---

## 8. จุดที่แตะ Dashboard (ADSM44 / %Ads)

⚠️ **อย่ารวม TikTok Live กับ TikTok Shop** เวลาคิด %Ads — คนละช่อง %Ads จะมั่ว
- ตัวเลข %Ads ต้องดึงจาก M44 จริง ไม่ใช่ค่าที่ import บางส่วน (REF เคยขาดไป 6 เท่า: Feb แสดง ฿85K จริง ฿518K)
- ระวังสับสนระหว่าง "channel mix %" / "sales contribution %" กับ "%Ads" — ต้อง confirm ว่าคอลัมน์ไหนคืออะไรก่อนสรุป
- M44 = แหล่ง Ads Cost by Product ที่ถูกต้อง สำหรับอัปเดต %Ads ราย channel (อ้างอิง skill `pct-ads-calc`)

---

## ไฟล์ที่เกี่ยวข้อง
- BEP model: `10_BEP/Survival Guide.../BEP_SurvivalGuide_v2.xlsx`
- REF: `10_BEP/_REF/REF_BEP_กับดัก-ออฟฟิศ.xlsx`
- Skill: `bep-builder`
