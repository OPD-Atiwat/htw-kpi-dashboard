# 10_BEP — Break Even Point Files

## โครงสร้าง Folder

```
10_BEP/
├── _REF/                                        ← Reference data ที่ใช้ร่วมทุกเล่ม
│   └── REF_BEP_กับดัก-ออฟฟิศ.xlsx             ← Channel mix, %Ads, avg price จากเล่ม Ref 2 เล่ม
│
├── กับดักความรู้สึกผิด/                         ← BEP files เล่มนี้
├── ออฟฟิศนี้เฮงซวย (สัตว์)/                    ← BEP files เล่มนี้
└── Survival Guide เอาตัวรอดได้ในเหตุวิกฤติ/   ← BEP files เล่มถัดไป (WIP)
```

## Convention ชื่อไฟล์

| ประเภท | รูปแบบชื่อ | ตัวอย่าง |
|--------|-----------|---------|
| BEP Template | `BEP_[ชื่อเล่ม]_v[N].xlsx` | `BEP_SurvivalGuide_v1.xlsx` |
| Reference Data | `REF_BEP_[เล่ม1]-[เล่ม2].xlsx` | `REF_BEP_กับดัก-ออฟฟิศ.xlsx` |
| Draft / WIP | `DRAFT_BEP_[ชื่อเล่ม].xlsx` | |

## BEP Logic สรุป

- **Channel Priority:** Facebook > TikTok > Shopee > Consignment
- **Price Gap:** FB (ถูกสุด) < TikTok < Shopee < C1 (แพงสุด = ราคาหลังปก)
- **TikTok Live:** ราคาพิเศษ — ไม่อยู่ใน gap rule
- **เดือน 1:** Prorated ตามวันที่เหลือ | **เดือน 2–4:** PEAK | **เดือน 5+:** Decay
- **เกณฑ์ผ่าน:** Gross Profit ≥ Target เล่มนั้น (ไม่ใช่ GM% หรือจำนวนวัน)
- **%Ads Reference:** FB ~43% | TikTok Shop ~40% | TikTok Affi ~42% | Shopee ~3%

## ไฟล์ Reference หลัก

`_REF/REF_BEP_กับดัก-ออฟฟิศ.xlsx` — อ้างอิงจาก:
- กับดักความรู้สึกผิด (Jan–May 19, 2026) Peak = Feb–Apr
- ออฟฟิศนี้เฮงซวย (สัตว์) (May 7–19, 2026) Launch snapshot
