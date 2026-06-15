#!/usr/bin/env python3
"""
Meta Ads API Puller v2 — OpenDurian How-to
─────────────────────────────────────────────────────────────
v2 fixes:
  ✅ ดึง account-level insights ครั้งเดียว (เร็วกว่า campaign-by-campaign ~10-20x)
  ✅ Field names ถูกต้องตรงกับ Ads Manager export ทุก column
  ✅ Product map ครบ — คืนชื่อภาษาไทยเต็ม ตรงกับ dashboard
  ✅ อัปเดต RAW_DATA (เดือนปัจจุบัน) ใน index.html อัตโนมัติ
  ✅ อัปเดต CREATOR_META_[MONTH] variable สำหรับ Creator Summary tab
  ✅ Action types ถูกต้อง (purchases, revenue, messages)
  ✅ Pagination ครบ ไม่หลุดข้อมูล
  ✅ Token error detection — แจ้งวิธีต่ออายุชัดเจน
"""

import requests
import json
import re
import os
import sys
from datetime import datetime
from collections import defaultdict

# ============================================================
# CONFIG — แก้ตรงนี้เท่านั้น
# ============================================================
ACCESS_TOKEN  = "EAACeOZCbphhwBRdzDkHHUEdj9S32nc2ZB1y65lt1U7suL4nS5UIkwAjUrL82AWprNQ1G0bMx6cv0A7Vf1QWdiEQtOHndZAPzDdBdkDyC7fwJugazXUsHx81HyDFnCkgL5FhYZCNoZB6Gvalzd11CbSJq8TQvHgB6ZAFSVZB8j8dWmzTJaFd9h3Hy4E8bJkTFJIUyQZDZD"
AD_ACCOUNT_ID = "act_303814252252288"
API_VERSION   = "v25.0"

ROAS_IMG = 1.8   # เป้า ROAS สำหรับภาพนิ่ง
ROAS_VDO = 2.0   # เป้า ROAS สำหรับวิดีโอ

DASHBOARD_PATH = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"
# ============================================================

# ─── วันที่อัตโนมัติ: ต้นเดือนนี้ → วันนี้ ──────────────────
today     = datetime.now()
DATE_FROM = today.replace(day=1).strftime("%Y-%m-%d")
DATE_TO   = today.strftime("%Y-%m-%d")
BASE_URL  = f"https://graph.facebook.com/{API_VERSION}"

MONTH_LABELS = {
    (2026,1):"Jan26",(2026,2):"Feb26",(2026,3):"Mar26",
    (2026,4):"Apr26",(2026,5):"May26",(2026,6):"Jun26",
    (2026,7):"Jul26",(2026,8):"Aug26",(2026,9):"Sep26",
    (2026,10):"Oct26",(2026,11):"Nov26",(2026,12):"Dec26",
}

# ─── Product code → Full Thai name mapping ────────────────────
# ดึงจาก parentheses สุดท้ายของ ad name: [Date][Creator] Content (ProductCode)
PRODUCT_MAP = {
    "Trap":            "กับดักความรู้สึกผิด",
    "Ghostly":         "Ghostly Brews ยินดีบริการดวงวิญญาณหลังเที่ยงคืน",
    "ghostly":         "Ghostly Brews ยินดีบริการดวงวิญญาณหลังเที่ยงคืน",
    "Remains":         "Ghostly Remains ปลดพันธนาการดวงวิญญาณหลังความตาย",
    "Witches":         "The Witches' Club ชมรมลับเปลี่ยนชีวิต",
    "Demons":          "ปีศาจตัวนั้น คือฉันเอง A Guide to Fighting the Demons in My Heart",
    "ปีศาจตัวนั้น":      "ปีศาจตัวนั้น คือฉันเอง A Guide to Fighting the Demons in My Heart",
    "Tried":           "เหนื่อยมากไหม พักก่อนก็ได้นะวันนี้",
    "Kiwtum":          "ขอใช้ชีวิตที่เหลือ เพื่อตัวเองนะ",
    "Coworkers":       "ออฟฟิศนี้เฮงซวย (สัตว์)",
    "Embrace":         "โอบกอดความไม่สมบูรณ์แบบของเธอ (Embrace Your Flaws)",
    "Capybara":        "ถ้าโลกมันแย่ ก็แค่คิดแบบคาปิบาร่า Think Like a Happybara",
    "Capybara_Boxset": "ถ้าโลกมันแย่ ก็แค่คิดแบบคาปิบาร่า Think Like a Happybara",
    "DevilLove":       "ปีศาจความรักมักเลือกฉันเป็นเหยื่อ How to Fight the Love Demon in My Heart",
    "Mind":            "ฉันจะมีชีวิตที่ดี",
    "Garden":          "วันไหนที่ใจแข็งแรง ดอกไม้จะผลิบาน The Enchanted Garden",
    "Helmet":          "เมื่อโลกทั้งใบซ่อนอยู่ใต้หมวกกันน็อก (Helmet Girl)",
    "Breath":          "กว่าจะคิดได้ ก็ไม่มีลมหายใจแล้ว (It Is Never Too Late to Love Yourself)",
    "Secret":          "ความลับสู่เงินล้านที่โรงเรียนไม่เคยสอน (The Millionaire's Top Secret)",
    "WelcomeBack":     "ยินดีต้อนรับและขอบคุณที่กลับมา",
    "Memory":          "ทุกความทรงจำคือของขวัญจากวันวาน",
    "Gratitude":       "ชีวิตมีเรื่องให้ขอบคุณมากกว่าเสียใจ",
    "E-Book":          "E-Book",
    "Cross Page":      "Cross Page",
    "Promotion":       "Promotion",
    "Promotion_Payday":"Promotion",
    "Sale":            "Sale",
    "Shopify":         "Shopify",
    "Home":            "หน้าร้าน",
    "Engage":          "Engagement",
    "QuickReply":      "QuickReply",
    "Reels":           "Reels",
    "Kidmakk":         "Kidmakk",
}

# Fields ที่ดึงจาก Meta API (ตรงกับ column ใน Ads Manager export)
# spend              → จำนวนเงินที่ใช้จ่ายไป (THB)
# impressions        → อิมเพรสชัน
# reach              → การเข้าถึง
# inline_link_clicks → การคลิกลิงก์
# ctr                → CTR (ทั้งหมด)
# cpm                → CPM
# actions            → ผลลัพธ์ (purchases, messages)
# action_values      → ค่าคอนเวอร์ชั่นการซื้อ (revenue)
# purchase_roas      → ROAS ของการซื้อ
API_FIELDS = ",".join([
    "ad_id", "ad_name", "campaign_name", "date_start",
    "spend", "impressions", "reach",
    "inline_link_clicks", "ctr", "cpm",
    "actions", "action_values", "purchase_roas",
])


# ─── Helpers ──────────────────────────────────────────────────

def api_get(url, params, retries=3):
    for attempt in range(retries):
        try:
            r = requests.get(url, params=params, timeout=120)
            data = r.json()
            if "error" in data:
                err = data["error"]
                code = err.get("code", 0)
                msg  = err.get("message", "unknown")
                # Token หมดอายุ
                if code in (190, 102, 463, 467):
                    print(f"\n❌ Access Token หมดอายุหรือไม่ถูกต้อง (code {code})")
                    print("   วิธีต่ออายุ:")
                    print("   1. ไปที่ https://developers.facebook.com/tools/explorer/")
                    print("   2. เลือก App → Generate Access Token")
                    print("   3. เพิ่ม permissions: ads_read, read_insights")
                    print("   4. คัดลอก token ใหม่ใส่ ACCESS_TOKEN ใน script")
                    sys.exit(1)
                raise Exception(f"API Error {code}: {msg}")
            return data
        except requests.exceptions.Timeout:
            if attempt < retries - 1:
                print(f"   ⏳ Timeout, retry {attempt+2}/{retries}...")
            else:
                raise


def fetch_all_pages(url, params):
    """ดึงข้อมูลครบทุกหน้า (auto pagination + retry เมื่อโดน rate limit)"""
    import time
    rows = []
    data = api_get(url, params)
    rows.extend(data.get("data", []))
    next_url = data.get("paging", {}).get("next")
    page = 1
    while next_url:
        page += 1
        print(f"   📄 Loading page {page}...", end="\r")
        try:
            r2 = requests.get(next_url, timeout=120)
            d2 = r2.json()
            if "error" in d2:
                err_msg  = d2['error'].get('message', '')
                err_code = d2['error'].get('code', 0)
                # Rate limit → รอแล้ว retry สูงสุด 3 ครั้ง
                if err_code in (4, 17, 32, 613) or 'limit' in err_msg.lower():
                    for attempt in range(1, 4):
                        wait = 60 * attempt
                        print(f"\n   ⏳ Rate limit — รอ {wait} วิ แล้ว retry ({attempt}/3)...")
                        time.sleep(wait)
                        r2 = requests.get(next_url, timeout=120)
                        d2 = r2.json()
                        if "error" not in d2:
                            break
                    else:
                        print(f"\n   ❌ Rate limit ไม่หาย หยุด pagination (ได้ {len(rows):,} records)")
                        break
                else:
                    print(f"\n   ⚠️  Pagination error: {err_msg}")
                    break
            rows.extend(d2.get("data", []))
            next_url = d2.get("paging", {}).get("next")
        except Exception as e:
            print(f"\n   ⚠️  หยุด pagination: {e}")
            break
    if page > 1:
        print(f"   📄 โหลดครบ {page} หน้า                    ")
    return rows


def extract_action(arr, *types):
    """ดึงค่าจาก actions/action_values array"""
    if not arr:
        return 0.0
    for atype in types:
        for a in arr:
            if a.get("action_type") == atype:
                try:
                    return float(a.get("value", 0) or 0)
                except (ValueError, TypeError):
                    return 0.0
    return 0.0


def clean_name(name):
    """ลบ suffix ที่ไม่ต้องการออก"""
    name = re.sub(r'\s*\(\s*สำเนา\s*\)', '', name)
    name = re.sub(r'\s*\(\s*Copy\s*\)', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s*-\s*สำเนา\s*$', '', name)
    return name.strip()


def month_key(date_str):
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return MONTH_LABELS.get((dt.year, dt.month),
                                dt.strftime("%b") + str(dt.year)[2:])
    except Exception:
        return "Unknown"


def current_month_key():
    return MONTH_LABELS.get((today.year, today.month),
                            today.strftime("%b") + str(today.year)[2:])


def guess_creator(ad_name):
    """ดึง creator จาก bracket ที่สอง [Date][CreatorTag] ..."""
    # Central / Influ / Cross Page / Our creators
    TAG_MAP = {
        # Our creators
        "ริว":        "ริว",
        "หมิว":       "หมิว",
        "มิ้น":       "มิ้น",
        "แนน":        "แนน",
        "แก้ม":       "แก้ม",
        # Central (inhouse)
        "Course":     "Central",
        "PILOT":      "Central",
        "Edugadget":  "Central",
        # Influ (external influencers)
        "pearpmk":    "Influ",
        "mena.forest":"Influ",
    }
    tags = re.findall(r'\[([^\]]+)\]', ad_name)
    # First tag = date, second tag = creator
    for b in tags[1:]:
        b = b.strip()
        if b in TAG_MAP:
            return TAG_MAP[b]
        # Fuzzy match our creators
        if "ริว" in b: return "ริว"
        if "หมิว" in b or "mew" in b.lower(): return "หมิว"
        if "มิ้น" in b or "min" in b.lower(): return "มิ้น"
        if "แนน" in b: return "แนน"
        if "แก้ม" in b: return "แก้ม"
        # Any other unknown bracket = Influ (external)
        if b and not re.match(r'^\d', b):  # not a date-like tag
            return "Influ"
    # Fallback: keyword scan in full name
    if "หมิว" in ad_name: return "หมิว"
    if "ริว"  in ad_name: return "ริว"
    if "มิ้น" in ad_name: return "มิ้น"
    return "Unknown"


def guess_product(ad_name):
    """ดึง product code จาก parentheses สุดท้าย → full Thai name"""
    brackets = re.findall(r'\(([^)]+)\)', ad_name)
    if brackets:
        code = brackets[-1].strip()
        if code in PRODUCT_MAP:
            return PRODUCT_MAP[code]
        for k, v in PRODUCT_MAP.items():
            if k.lower() == code.lower():
                return v
        return f"[{code}]"   # ยังไม่ map — แสดง code เพื่อ debug
    return "Unknown"


def resolve_creator_by_product(creator, product):
    """ถ้า product = Cross Page → creator = Cross Page เสมอ"""
    if product == "Cross Page":
        return "Cross Page"
    return creator


# ครีเอเตอร์ทีมเรา (เฉพาะกลุ่มนี้ที่แชร์ยอด 50/50 ได้)
TEAM_CREATORS = ["ริว", "หมิว", "มิ้น", "แนน", "แก้ม"]

def guess_creators(ad_name):
    """คืน list ครีเอเตอร์ทีมเราที่เจอใน bracket [Date][Creator1][Creator2]...
       ถ้าเจอหลายคน → แชร์ยอดเท่ากัน (เช่น [มิ้น][ริว] → ['มิ้น','ริว'] = 50/50)
       ถ้าไม่เจอครีทีมเราเลย → fallback เป็น [guess_creator()] (เดี่ยว เช่น Central/Influ)"""
    tags = re.findall(r'\[([^\]]+)\]', ad_name)
    found = []
    for b in tags[1:]:               # ข้าม tag แรก (วันที่)
        b = b.strip()
        m = None
        if "ริว" in b: m = "ริว"
        elif "หมิว" in b or "mew" in b.lower(): m = "หมิว"
        elif "มิ้น" in b or "min" in b.lower(): m = "มิ้น"
        elif "แนน" in b: m = "แนน"
        elif "แก้ม" in b: m = "แก้ม"
        if m and m not in found:
            found.append(m)
    return found if found else [guess_creator(ad_name)]


def guess_format(ad_name):
    """ดึง content format: VDO หรือ IMG"""
    for b in re.findall(r'\[([^\]]+)\]', ad_name):
        b_up = b.strip().upper()
        if b_up in ('VDO', 'VIDEO', 'VID', 'REELS', 'REEL'):
            return 'VDO'
        if b_up in ('IMG', 'IMAGE', 'ALBUM', 'PHOTO', 'CAROUSEL'):
            return 'IMG'
    return 'IMG'


def guess_launch_date(ad_name):
    """ดึงวันที่จาก [DD.MM.YY] ใน ad name"""
    m = re.search(r'\[(\d{2})\.(\d{2})\.(\d{2})\]', ad_name)
    if m:
        d, mo, y = m.groups()
        return f"20{y}-{mo}-{d}"
    return ""


def roas_target(fmt):
    return ROAS_VDO if fmt == 'VDO' else ROAS_IMG


def hit_roas_label(roas, purchases, fmt='IMG'):
    if purchases == 0:
        return "ไม่มียอดขาย"
    return "ผ่าน" if roas >= roas_target(fmt) else "ไม่ผ่าน"


# ─── Pull insights ─────────────────────────────────────────────

def pull_ad_insights():
    """
    ดึง ad-level insights ระดับ account ครั้งเดียว
    — เร็วกว่า campaign-by-campaign ~10-20x
    — filtering spend>0 เพื่อลดขนาดข้อมูล
    """
    print(f"🔗 ดึงข้อมูลจาก Meta Ads API...")
    print(f"   Account: {AD_ACCOUNT_ID}")
    print(f"   ช่วง:    {DATE_FROM} → {DATE_TO}")
    print(f"   ระดับ:   Ad level (per creative, daily)")
    print(f"   Fields:  spend, revenue, ROAS, purchases, messages, impressions,")
    print(f"            reach, clicks, CTR, CPM")

    params = {
        "level":          "ad",
        "fields":         API_FIELDS,
        "time_range":     json.dumps({"since": DATE_FROM, "until": DATE_TO}),
        "time_increment": 1,
        "filtering":      json.dumps([
            {"field": "spend", "operator": "GREATER_THAN", "value": "0"}
        ]),
        "access_token":   ACCESS_TOKEN,
        "limit":          500,
    }

    url = f"{BASE_URL}/{AD_ACCOUNT_ID}/insights"
    raw_rows = fetch_all_pages(url, params)
    print(f"   ✅ ได้ {len(raw_rows):,} records (spend > 0)")
    return raw_rows


def transform_rows(raw_rows):
    """แปลง API response → RAW_DATA format"""
    rows = []
    for day in raw_rows:
        date_str      = day.get("date_start", "")
        ad_name       = clean_name(day.get("ad_name", ""))
        campaign_name = day.get("campaign_name", "")
        spend       = float(day.get("spend", 0) or 0)
        impressions = int(float(day.get("impressions", 0) or 0))
        reach       = int(float(day.get("reach", 0) or 0))
        clicks      = int(float(day.get("inline_link_clicks", 0) or 0))
        # Meta ส่ง CTR เป็น % แล้ว (e.g., 2.89 = 2.89%)
        ctr         = float(day.get("ctr", 0) or 0)
        cpm         = float(day.get("cpm", 0) or 0)
        actions     = day.get("actions", []) or []
        action_vals = day.get("action_values", []) or []
        roas_arr    = day.get("purchase_roas", []) or []

        # ── Purchases ──
        purchases = extract_action(actions,
            "omni_purchase",
            "offsite_conversion.fb_pixel_purchase",
            "purchase",
        )

        # ── Revenue (purchase conversion value) ──
        revenue = extract_action(action_vals,
            "omni_purchase",
            "offsite_conversion.fb_pixel_purchase",
            "purchase",
        )

        # ── Messaging conversations started ──
        messages = extract_action(actions,
            "onsite_conversion.messaging_conversation_started_7d",
            "onsite_conversation.messaging_conversation_started_7d",
            "messaging_first_reply",
            "onsite_conversion.post_save",
        )

        # ── ROAS ──
        if roas_arr:
            try:
                roas = float(roas_arr[0].get("value", 0) or 0)
            except Exception:
                roas = revenue / spend if spend > 0 and revenue > 0 else 0.0
        else:
            roas = revenue / spend if spend > 0 and revenue > 0 else 0.0

        cpa = spend / purchases if purchases > 0 else 0.0
        cvr = purchases / clicks * 100 if clicks > 0 else 0.0

        fmt     = guess_format(ad_name)
        product = guess_product(ad_name)
        launch  = guess_launch_date(ad_name)
        # creators list (รองรับแชร์ 50/50 เมื่อมีหลายครีในชื่อแอด)
        creators = guess_creators(ad_name)
        if product == "Cross Page":
            creators = ["Cross Page"]          # Cross Page ทับเสมอ ไม่แชร์
        creator = creators[0]                  # primary (tag แรก) — backward compat
        share   = round(1.0 / len(creators), 6)

        rows.append({
            "month":        month_key(date_str),
            "creator":      creator,
            "creators":     creators,
            "share":        share,
            "ad_date":      date_str,
            "content_name": ad_name,
            "product":      product,
            "platform":     "Meta",
            "type":         fmt,
            "reach":        reach,
            "impressions":  impressions,
            "spend_thb":    round(spend, 2),
            "purchases":    int(purchases),
            "revenue_thb":  round(revenue, 2),
            "roas":         round(roas, 4),
            "cpm_thb":      round(cpm, 4),
            "cpa_thb":      round(cpa, 2),
            "clicks":       clicks,
            "ctr":          round(ctr, 4),
            "cvr":          round(cvr, 4),
            "messages":     int(messages),
            "v2s":          0,
            "v6s":          0,
            "ad_variants":  1,
            "hit_roas":      hit_roas_label(roas, int(purchases), fmt),
            "is_new":        False,
            "launch_date":   launch,
            "campaign_name": campaign_name,
        })
    return rows


# ─── Dashboard updater ─────────────────────────────────────────

def update_raw_data(rows, html_path):
    """
    อัปเดต RAW_DATA ใน index.html:
    • ลบ Meta records ของเดือนปัจจุบันออก
    • เพิ่ม records ใหม่จาก API
    • คงข้อมูลเดือนอื่นและ platform อื่นไว้
    """
    if not os.path.exists(html_path):
        print(f"⚠️  ไม่พบไฟล์: {html_path}")
        return False

    cur_month = current_month_key()
    print(f"\n📊 อัปเดต RAW_DATA เดือน {cur_month} ใน dashboard...")

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    m = re.search(r'const RAW_DATA = (\[.*?\]);', content, re.DOTALL)
    if not m:
        print("   ⚠️  ไม่พบ const RAW_DATA ใน HTML")
        return False

    try:
        existing = json.loads(m.group(1))
    except json.JSONDecodeError as e:
        print(f"   ⚠️  Parse RAW_DATA ผิดพลาด: {e}")
        return False

    # คง records เดือนอื่น + platform อื่น (TikTok ฯลฯ)
    kept = [r for r in existing
            if not (r.get("month") == cur_month and r.get("platform") == "Meta")]

    new_recs   = [r for r in rows if r.get("platform") == "Meta"]
    new_dates  = set(r.get("ad_date", "") for r in new_recs)

    # ── Rate-limit safeguard ──────────────────────────────────────
    # ถ้า API ส่งมาแค่บางวัน (rate limit) ให้คงข้อมูลวันที่ยังไม่ได้ดึงไว้
    existing_meta  = [r for r in existing
                      if r.get("month") == cur_month and r.get("platform") == "Meta"]
    existing_dates = set(r.get("ad_date", "") for r in existing_meta)
    missed_dates   = existing_dates - new_dates          # วันที่ API ไม่ส่งมา
    rescued        = [r for r in existing_meta if r.get("ad_date", "") in missed_dates]
    if rescued:
        print(f"   ⚠️  API ได้ {len(new_dates)} วัน / มีอยู่แล้ว {len(existing_dates)} วัน"
              f" — คงข้อมูล {len(missed_dates)} วันที่ขาด ({sorted(missed_dates)[0]} → {sorted(missed_dates)[-1]})")

    merged = kept + new_recs + rescued
    merged.sort(key=lambda r: (
        r.get("month", ""), r.get("ad_date", ""), r.get("creator", "")
    ))

    new_json    = json.dumps(merged, ensure_ascii=False, separators=(",", ":"))
    new_content = re.sub(
        r'const RAW_DATA = \[.*?\];',
        f'const RAW_DATA = {new_json};',
        content,
        flags=re.DOTALL,
    )

    if new_content == content:
        print("   ⚠️  ไม่มีการเปลี่ยนแปลง")
        return False

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    print(f"   ✅ คงไว้ (platform อื่น): {len(kept):,} records")
    print(f"   ✅ Meta ใหม่จาก API: {len(new_recs):,} records ({cur_month})")
    if rescued:
        print(f"   ✅ Meta คงไว้ (ขาดจาก API): {len(rescued):,} records")
    print(f"   ✅ รวม: {len(merged):,} records")
    return True


def update_meta_ad_perf(rows, html_path):
    """
    สร้าง/อัปเดต META_AD_PERF จาก RAW_DATA ทั้งหมด (ทุกเดือน)
    Structure: { content_name: { cr, days: [{d, sp, rv, pu, pass}] } }
    """
    if not os.path.exists(html_path):
        return False

    print(f"\n📊 อัปเดต META_AD_PERF...")

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Build META_AD_PERF from all RAW_DATA (pull existing + merge new rows)
    m = re.search(r'const RAW_DATA = (\[.*?\]);', content, re.DOTALL)
    all_rows = []
    if m:
        try:
            all_rows = json.loads(m.group(1))
        except Exception:
            all_rows = []

    # Group by content_name → per-day aggregation
    from collections import defaultdict
    perf = {}  # content_name → {cr, days_map: {date → {sp,rv,pu,pass}}}

    for r in all_rows:
        if r.get("platform") != "Meta":
            continue
        cn = r.get("content_name", "")
        if not cn:
            continue
        dt = r.get("ad_date", "")
        sp = r.get("spend_thb", 0) or 0
        rv = r.get("revenue_thb", 0) or 0
        pu = r.get("purchases", 0) or 0
        roas = rv / sp if sp > 0 else 0
        tgt  = roas_target(r.get("format", "IMG"))
        passed = 1 if (sp > 0 and roas >= tgt) else 0

        if cn not in perf:
            perf[cn] = {"cr": r.get("creator", ""), "days_map": {}}
        day = perf[cn]["days_map"].setdefault(dt, {"sp": 0, "rv": 0, "pu": 0, "pass": 0})
        day["sp"] = round(day["sp"] + sp, 2)
        day["rv"] = round(day["rv"] + rv, 2)
        day["pu"] += pu
        day["pass"] = max(day["pass"], passed)

    # Convert days_map → sorted days list
    result = {}
    for cn, v in perf.items():
        days = [{"d": d, **v["days_map"][d]} for d in sorted(v["days_map"])]
        result[cn] = {"cr": v["cr"], "days": days}

    new_json = json.dumps(result, ensure_ascii=False, separators=(",", ":"))
    new_content = re.sub(
        r'var META_AD_PERF\s*=\s*\{.*?\};',
        f'var META_AD_PERF = {new_json};',
        content,
        flags=re.DOTALL,
    )

    if new_content == content:
        print("   ⚠️  META_AD_PERF ไม่มีการเปลี่ยนแปลง")
        return False

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    # Count May entries
    may_contents = sum(
        1 for v in result.values()
        if any(d["d"].startswith("2026-05") for d in v.get("days", []))
    )
    print(f"   ✅ META_AD_PERF: {len(result):,} content entries, {may_contents} มี May 26 data")
    return True


def update_creator_summary(rows, html_path):
    """อัปเดต CREATOR_META_[MONTH] สำหรับ Creator Summary tab"""
    if not os.path.exists(html_path):
        return

    cur_month  = current_month_key()
    var_name   = f"CREATOR_META_{cur_month.upper()}"
    creators   = ["หมิว", "ริว", "มิ้น"]
    def _crs(r):
        return r.get("creators") or [r["creator"]]
    def _w(r, cr):
        """น้ำหนักยอดของ row นี้ที่ปันให้ cr (0 ถ้าไม่เกี่ยว, share ถ้าเกี่ยว)"""
        c = _crs(r)
        return r.get("share", 1.0) if cr in c else 0.0
    # รวมทุก row ที่มีครีหลักอย่างน้อย 1 คน (รองรับแอดที่แชร์หลายคน)
    month_rows = [r for r in rows
                  if r["month"] == cur_month and any(c in creators for c in _crs(r))]

    if not month_rows:
        print(f"   ⚠️  ไม่มีข้อมูล Creator หลักใน {cur_month} — ข้าม {var_name}")
        return

    # Summary per creator
    summary = {}
    for cr in creators:
        cr_rows = [(r, _w(r, cr)) for r in month_rows if _w(r, cr) > 0]
        if not cr_rows:
            continue
        sp  = sum(r["spend_thb"]   * w for r, w in cr_rows)
        rv  = sum(r["revenue_thb"] * w for r, w in cr_rows)
        pur = sum(r["purchases"]   * w for r, w in cr_rows)
        imp = sum(r["impressions"] * w for r, w in cr_rows)
        rch = sum(r["reach"]       * w for r, w in cr_rows)
        msg = sum(r["messages"]    * w for r, w in cr_rows)
        clk = sum(r["clicks"]      * w for r, w in cr_rows)
        summary[cr] = {
            "spend":       round(sp, 2),
            "revenue":     round(rv, 2),
            "roas":        round(rv/sp, 4) if sp > 0 else 0,
            "purchases":   round(pur),
            "impressions": round(imp),
            "reach":       round(rch),
            "ctr":         round(clk/imp*100, 4) if imp > 0 else 0,
            "cpm":         round(sp/imp*1000, 4) if imp > 0 else 0,
            "messages":    round(msg),
            "link_clicks": round(clk),
            "cpa":         round(sp/pur, 2) if pur > 0 else 0,
        }

    # Daily per creator (แชร์ตาม weight)
    daily = {}
    for cr in creators:
        cr_rows = [(r, _w(r, cr)) for r in month_rows if _w(r, cr) > 0]
        by_date = defaultdict(lambda: {"s":0,"r":0,"p":0,"i":0,"m":0})
        for r, w in cr_rows:
            d = r["ad_date"]
            by_date[d]["s"] += r["spend_thb"]   * w
            by_date[d]["r"] += r["revenue_thb"] * w
            by_date[d]["p"] += r["purchases"]   * w
            by_date[d]["i"] += r["impressions"] * w
            by_date[d]["m"] += r["messages"]    * w
        daily[cr] = [
            {"d":d,"s":round(v["s"],2),"r":round(v["r"],2),
             "p":round(v["p"]),"i":round(v["i"]),"m":round(v["m"])}
            for d, v in sorted(by_date.items())
        ]

    # Products per creator (แชร์ตาม weight)
    products = {}
    for cr in creators:
        cr_rows = [(r, _w(r, cr)) for r in month_rows if _w(r, cr) > 0]
        by_prod = defaultdict(lambda: {"spend":0,"revenue":0,"purchases":0})
        for r, w in cr_rows:
            prod = r.get("product") or "Unknown"
            by_prod[prod]["spend"]     += r["spend_thb"]   * w
            by_prod[prod]["revenue"]   += r["revenue_thb"] * w
            by_prod[prod]["purchases"] += r["purchases"]   * w
        products[cr] = sorted([
            {"name": p,
             "spend": round(v["spend"], 2),
             "revenue": round(v["revenue"], 2),
             "roas": round(v["revenue"]/v["spend"], 4) if v["spend"]>0 else 0,
             "purchases": round(v["purchases"])}
            for p, v in by_prod.items()
        ], key=lambda x: x["spend"], reverse=True)

    new_obj  = json.dumps(
        {"summary": summary, "daily": daily, "products": products},
        ensure_ascii=False, separators=(",", ":")
    )
    new_line = f"const {var_name} = {new_obj};"

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Replace any existing CREATOR_META_* variable
    new_content = re.sub(
        r'const CREATOR_META_[A-Z0-9]+ = \{.*?\};',
        new_line,
        content,
        flags=re.DOTALL,
    )

    if new_content == content:
        print(f"   ⚠️  ไม่พบ CREATOR_META_* ใน HTML — ข้าม")
        return

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    sp_tot = sum(summary[cr]["spend"]   for cr in summary)
    rv_tot = sum(summary[cr]["revenue"] for cr in summary)
    print(f"   ✅ {var_name} อัปเดตแล้ว")
    if sp_tot > 0:
        print(f"      ทีม {cur_month}: Spend ฿{sp_tot:,.0f} | "
              f"Revenue ฿{rv_tot:,.0f} | ROAS {rv_tot/sp_tot:.2f}x")


# ─── Summary report ───────────────────────────────────────────

def print_summary(rows):
    cur_month  = current_month_key()
    month_rows = [r for r in rows if r["month"] == cur_month]
    if not month_rows:
        print("ไม่มีข้อมูลเดือนนี้")
        return

    all_creators = sorted(set(r["creator"] for r in month_rows))

    print(f"\n{'─'*72}")
    print(f"  สรุป {cur_month}  ({DATE_FROM} → {DATE_TO})")
    print(f"{'─'*72}")
    print(f"  {'Creator':<12} {'Spend':>10} {'Revenue':>10} {'ROAS':>6} "
          f"{'Pur':>5} {'Msg':>5} {'Rows':>5}")
    print(f"  {'─'*12} {'─'*10} {'─'*10} {'─'*6} {'─'*5} {'─'*5} {'─'*5}")

    for cr in all_creators:
        cr_rows = [r for r in month_rows if r["creator"] == cr]
        sp  = sum(r["spend_thb"]   for r in cr_rows)
        rv  = sum(r["revenue_thb"] for r in cr_rows)
        pur = sum(r["purchases"]   for r in cr_rows)
        msg = sum(r["messages"]    for r in cr_rows)
        print(f"  {cr:<12} {sp:>10,.0f} {rv:>10,.0f} "
              f"{rv/sp:>6.2f}x {pur:>5} {msg:>5} {len(cr_rows):>5}"
              if sp > 0 else f"  {cr:<12} {'–':>10}")

    sp_all  = sum(r["spend_thb"]   for r in month_rows)
    rv_all  = sum(r["revenue_thb"] for r in month_rows)
    pur_all = sum(r["purchases"]   for r in month_rows)
    msg_all = sum(r["messages"]    for r in month_rows)
    print(f"  {'─'*12} {'─'*10} {'─'*10} {'─'*6} {'─'*5} {'─'*5} {'─'*5}")
    if sp_all > 0:
        print(f"  {'รวม':<12} {sp_all:>10,.0f} {rv_all:>10,.0f} "
              f"{rv_all/sp_all:>6.2f}x {pur_all:>5} {msg_all:>5} {len(month_rows):>5}")

    # Product breakdown
    print(f"\n  Product Breakdown:")
    by_prod = defaultdict(lambda: {"sp": 0, "rv": 0, "pur": 0})
    for r in month_rows:
        p = r.get("product", "Unknown")
        by_prod[p]["sp"]  += r["spend_thb"]
        by_prod[p]["rv"]  += r["revenue_thb"]
        by_prod[p]["pur"] += r["purchases"]
    for p, v in sorted(by_prod.items(), key=lambda x: x[1]["sp"], reverse=True):
        if v["sp"] > 0:
            roas = v["rv"]/v["sp"]
            print(f"    {p[:48]:<48} ฿{v['sp']:>8,.0f}  {roas:.2f}x  {v['pur']}pur")

    # Unknown product warning
    unknown_rows = [r for r in month_rows if r["product"] in ("Unknown", "") or
                    r["product"].startswith("[")]
    if unknown_rows:
        print(f"\n  ⚠️  Product ยังไม่ map ({len(unknown_rows)} rows) — เพิ่มใน PRODUCT_MAP:")
        seen = set()
        for r in unknown_rows[:8]:
            n = r["content_name"]
            if n not in seen:
                seen.add(n)
                # Extract the product code
                brackets = re.findall(r'\(([^)]+)\)', n)
                code = brackets[-1] if brackets else "?"
                print(f"    [{code}] → {n[:60]}")


# ─── Main ──────────────────────────────────────────────────────

def main():
    print("=" * 72)
    print("  Meta Ads API Puller v2 — OpenDurian How-to")
    print(f"  ช่วง: {DATE_FROM} → {DATE_TO}")
    print("=" * 72)

    # 0. sync MK13 + ADSM44 ก่อน (อัปเดต OPD_DAILY + ADSM44_PCTADS locally)
    try:
        import sys, os
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import mk13_sync
        mk13_sync.sync(DASHBOARD_PATH)
    except Exception as e:
        print(f"\n⚠️  mk13_sync ล้มเหลว: {e}")
        print("   — ข้าม MK13/ADSM44 แต่ Meta pull ยังทำงานต่อ")

    # 1. ดึงข้อมูลจาก API
    raw_rows = pull_ad_insights()
    if not raw_rows:
        print("❌ ไม่มีข้อมูลในช่วงนี้ (spend = 0 ทั้งหมด หรือ API error)")
        return

    # 2. แปลงเป็น RAW_DATA format
    print(f"\n🔄 แปลงข้อมูล...")
    rows = transform_rows(raw_rows)
    print(f"   ✅ {len(rows):,} rows พร้อมแล้ว")

    # 3. สรุปผล
    print_summary(rows)

    # 4. อัปเดต dashboard (RAW_DATA + META_AD_PERF + CREATOR_META)
    print(f"\n💾 อัปเดต Dashboard...")
    ok = update_raw_data(rows, DASHBOARD_PATH)
    if ok:
        update_meta_ad_perf(rows, DASHBOARD_PATH)
        update_creator_summary(rows, DASHBOARD_PATH)
        # 5. Generate + update AO_DATA (Ads Optimize tab) จาก adset-level API
        try:
            update_ao_data(DASHBOARD_PATH)
        except Exception as e:
            print(f"   ⚠️  AO_DATA update ไม่สำเร็จ: {e}")
        print(f"\n✅ เสร็จสิ้น!")
    else:
        print(f"\n⚠️  Dashboard ไม่ได้รับการอัปเดต — ตรวจสอบ DASHBOARD_PATH")


# ─── AO_DATA: Pull adset-level + update dashboard ───────────────────────────

def _parse_adset_meta(adset_name, ad_name=""):
    """แยก audience, age, fmt, book จาก adset name / ad name"""
    import re
    name = adset_name or ""
    # Format
    fmt = "VDO" if re.search(r'\[VDO\]|\[vdo\]|video', name, re.I) else "IMG"
    # Age
    age_m = re.search(r'\[(\d{2}\+)\]', name)
    age = age_m.group(1) if age_m else "-"
    # Auto On
    auto_on = bool(re.search(r'AutoON|Auto ON|🔀', name, re.I))
    # Audience group heuristic
    au_lower = name.lower()
    if "retarget" in au_lower or "retargeting" in au_lower:
        if "page" in au_lower: audience_group = "Retarget — Page"
        elif "video" in au_lower: audience_group = "Retarget — Video"
        elif "conversion" in au_lower: audience_group = "Retarget — Conversion"
        else: audience_group = "Retarget — Other"
    elif "broad" in au_lower or "broad" in au_lower:
        audience_group = "Broad"
    elif "interest" in au_lower or "books" in au_lower:
        if "book" in au_lower: audience_group = "Interest — Books"
        else: audience_group = "Interest — Other"
    else:
        audience_group = "Broad"
    audience = re.sub(r'\[.*?\]|\*|🔀|🌏', '', name).strip()[:40]
    return fmt, age, auto_on, audience_group, audience


def _infer_book_from_names(ad_name, adset_name, campaign_name, product_map):
    """หาชื่อหนังสือจากชื่อ ad/adset/campaign"""
    combined = " ".join([ad_name or "", adset_name or "", campaign_name or ""])
    for code, full in product_map.items():
        if code.lower() in combined.lower() or full[:8] in combined:
            return full
    return "ออฟฟิศนี้เฮงซวย (สัตว์)"  # default หลักสำหรับ ออฟฟิศ


def pull_adset_insights_agg():
    """ดึง adset-level aggregate insights สำหรับ AO_DATA"""
    print(f"\n📊 ดึง Ad Set-level insights (AO_DATA)...")
    params = {
        "level":      "ad",
        "fields":     ",".join([
            "ad_id","ad_name","adset_id","adset_name","campaign_name",
            "spend","impressions","reach","inline_link_clicks","ctr","cpm",
            "actions","action_values","purchase_roas",
        ]),
        "time_range":  json.dumps({"since": DATE_FROM, "until": DATE_TO}),
        "time_increment": "all_days",   # aggregate ทั้งช่วง ไม่แยกรายวัน
        "filtering":   json.dumps([
            {"field": "spend", "operator": "GREATER_THAN", "value": "0"}
        ]),
        "access_token": ACCESS_TOKEN,
        "limit":       500,
    }
    url = f"{BASE_URL}/{AD_ACCOUNT_ID}/insights"
    raw = fetch_all_pages(url, params)
    print(f"   ✅ {len(raw):,} ad rows")
    return raw


def update_ao_data(html_path):
    """Generate AO_DATA จาก Meta API แล้ว inject เข้า dashboard"""
    raw = pull_adset_insights_agg()
    if not raw:
        print("   ⚠️  ไม่มีข้อมูล adset — ข้าม AO_DATA")
        return

    # Aggregate per ad
    ads_agg = {}
    for r in raw:
        aid = r.get("ad_id","")
        if not aid: continue
        if aid not in ads_agg:
            ads_agg[aid] = {
                "ad_id": aid,
                "ad_name": r.get("ad_name",""),
                "adset_id": r.get("adset_id",""),
                "adset_name": r.get("adset_name",""),
                "campaign_name": r.get("campaign_name",""),
                "spend":0,"revenue":0,"purchases":0,"impressions":0,
                "reach":0,"clicks":0,"msg":0,
            }
        d = ads_agg[aid]
        d["spend"]       += float(r.get("spend",0) or 0)
        d["impressions"] += int(float(r.get("impressions",0) or 0))
        d["reach"]       += int(float(r.get("reach",0) or 0))
        d["clicks"]      += int(float(r.get("inline_link_clicks",0) or 0))
        for a in (r.get("action_values") or []):
            if a.get("action_type","") in ("omni_purchase","offsite_conversion.fb_pixel_purchase","purchase"):
                d["revenue"] += float(a.get("value",0) or 0)
        for a in (r.get("actions") or []):
            if a.get("action_type","") in ("omni_purchase","offsite_conversion.fb_pixel_purchase","purchase"):
                d["purchases"] += int(float(a.get("value",0) or 0))
            if "messaging_conversation" in a.get("action_type",""):
                d["msg"] += int(float(a.get("value",0) or 0))

    # Build ads list
    ads_list = []
    for d in sorted(ads_agg.values(), key=lambda x: -x["spend"]):
        sp  = d["spend"]
        rev = d["revenue"]
        roas = round(rev/sp, 2) if sp > 0 else 0.0
        fmt, age, auto_on, aud_group, audience = _parse_adset_meta(d["adset_name"], d["ad_name"])
        ctr = round(d["clicks"]/d["impressions"]*100, 2) if d["impressions"] > 0 else 0.0
        cpm = round(d["spend"]/d["impressions"]*1000, 2) if d["impressions"] > 0 else 0.0
        cpa = round(sp/d["purchases"], 0) if d["purchases"] > 0 else 0.0
        cost_msg = round(sp/d["msg"], 2) if d["msg"] > 0 else 0.0
        # Tier
        if d["purchases"] == 0: tier = "np"
        elif roas >= 2.5: tier = "ex"
        elif roas >= 2.0: tier = "go"
        elif roas >= 1.7: tier = "ok"
        else: tier = "po"

        ads_list.append({
            "ad_id":          d["ad_id"],
            "ad_name":        d["ad_name"],
            "adset_id":       d["adset_id"],
            "adset_name":     d["adset_name"],
            "campaign_name":  d["campaign_name"],
            "audience":       audience,
            "audience_group": aud_group,
            "age":            age,
            "fmt":            fmt,
            "auto_on":        auto_on,
            "spend":          round(sp, 2),
            "revenue":        round(rev, 2),
            "purchases":      d["purchases"],
            "roas":           roas,
            "impressions":    d["impressions"],
            "reach":          d["reach"],
            "clicks":         d["clicks"],
            "ctr":            ctr,
            "cpm":            cpm,
            "msg":            d["msg"],
            "cost_per_msg":   cost_msg,
        })

    ao = {
        "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "period":     f"{DATE_FROM} – {DATE_TO}",
        "daily_ads":  [],   # patch_ao_thumbnails จะเติม preview URLs
        "ads":        ads_list,
        "adsets":     [],
    }

    # อ่านไฟล์ + เก็บ daily_ads เดิม (thumbnail/preview ที่ patch แล้ว)
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    # inject AO_DATA
    prefix = "var AO_DATA = "
    idx = content.find(prefix)
    if idx == -1:
        print("   ⚠️  ไม่พบ var AO_DATA")
        return

    m = re.search(r'var AO_DATA = (\{.*?\});', content[idx:idx+10000000], re.DOTALL)
    if not m:
        print("   ⚠️  parse AO_DATA ไม่ได้")
        return

    old_start = idx
    old_end   = idx + m.end()
    new_json  = prefix + json.dumps(ao, ensure_ascii=False, separators=(",",":")) + ";"
    content   = content[:old_start] + new_json + content[old_end:]

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"   ✅ AO_DATA updated: {len(ads_list)} ads, period={ao['period']}")


if __name__ == "__main__":
    main()
