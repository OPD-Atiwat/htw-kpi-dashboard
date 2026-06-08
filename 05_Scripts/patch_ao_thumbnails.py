"""
patch_ao_thumbnails.py
1. Inject thumbnail_url / video_id จาก meta_creative_map.json
2. Map fb_post_url จาก Google Sheet โดย:
   - Priority 1: Ads FB Name → ad_name (exact)
   - Priority 2: reel/video ID จาก Link FB Post URL → video_id ใน AO_DATA
รัน: python3 05_Scripts/patch_ao_thumbnails.py
"""
import json, re, os, sys, urllib.parse, urllib.request, csv, io, difflib, time
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE      = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard"
HTML_PATH = os.path.join(BASE, "index.html")
MAP_PATH  = os.path.join(BASE, "05_Scripts", "meta_creative_map.json")

SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRAcel-DjxU3ORc9dYBQcFa9ipKuArdbjmj6ygHn2pKC8xgLSB0fQCVn9Gnwu2-hTNKZIipOofFQrq8/pub?output=csv"

def extract_fb_video_id(url):
    """Extract video/reel ID from FB URL"""
    url = url.strip()
    # reel: facebook.com/reel/12345
    m = re.search(r'facebook\.com/reel/(\d+)', url)
    if m: return m.group(1)
    # watch: facebook.com/watch?v=12345
    m = re.search(r'[?&]v=(\d+)', url)
    if m: return m.group(1)
    # video: facebook.com/video/12345
    m = re.search(r'facebook\.com/video/(\d+)', url)
    if m: return m.group(1)
    return None

print("=== Patch AO_DATA: thumbnails + Sheet mapping ===")

# 1. Load creative map
with open(MAP_PATH, "r", encoding="utf-8") as f:
    creative_map = json.load(f)
print(f"Creative map: {len(creative_map)} ads")

# 2. Fetch Sheet
print("Fetching Sheet...")
with urllib.request.urlopen(SHEET_URL) as r:
    raw = r.read().decode("utf-8")
rows = list(csv.reader(io.StringIO(raw)))
headers = rows[0]
idx = {h.strip(): i for i, h in enumerate(headers)}

# Build lookup maps:
# name_map:    ads_fb_name  → entry
# vid_map:     fb_video_id  → entry
# keymsg_list: [(creator, key_msg, entry), ...]  สำหรับ substring match
name_map    = {}
vid_map     = {}
keymsg_list = []

for row in rows[1:]:
    if not row: continue
    fb_url   = row[idx["Link FB Post"]].strip() if len(row) > idx["Link FB Post"] else ""
    ads_name = row[idx["Ads FB Name"]].strip()  if len(row) > idx["Ads FB Name"]  else ""
    creator  = row[idx["Creator"]].strip()      if len(row) > idx["Creator"]      else ""
    key_msg  = row[idx["Key Message"]].strip()  if len(row) > idx["Key Message"]  else ""
    if not fb_url:
        continue
    entry = {
        "fb_post_url": fb_url,
        "creator":  creator,
        "product":  row[idx["Product"]].strip() if len(row) > idx["Product"] else "",
        "key_msg":  key_msg,
    }
    # Priority 1: by Ads FB Name
    if ads_name and ads_name not in (".", ""):
        name_map[ads_name] = entry
    # Priority 2: by extracted video ID
    vid_id = extract_fb_video_id(fb_url)
    if vid_id:
        vid_map[vid_id] = entry
    # Priority 3: key_msg list (สำหรับ substring match)
    if key_msg and len(key_msg) > 5:
        keymsg_list.append((creator.strip(), key_msg, entry))

print(f"Name map entries: {len(name_map)}")
print(f"Video ID map entries: {len(vid_map)}")
print(f"Key Message entries: {len(keymsg_list)}")

# 3. Load HTML
with open(HTML_PATH, "r", encoding="utf-8") as f:
    content = f.read()

# 4. Extract AO_DATA
m = re.search(r'(var AO_DATA = )(\{.*?\});', content, re.DOTALL)
if not m:
    print("❌ ไม่พบ var AO_DATA")
    exit(1)
prefix = m.group(1)
ao_data = json.loads(m.group(2))

# 5. Pre-build fuzzy index: sheet_names list สำหรับ difflib
sheet_name_keys = list(name_map.keys())  # Ads FB Name จาก Sheet

def fuzzy_match_name(ad_name, threshold=0.82):
    """หาชื่อ Sheet ที่ใกล้เคียง ad_name ที่สุด"""
    if not sheet_name_keys or not ad_name: return None
    matches = difflib.get_close_matches(ad_name, sheet_name_keys, n=1, cutoff=threshold)
    return matches[0] if matches else None

# Inject
injected_thumb = 0
matched_name   = 0
matched_fuzzy  = 0
matched_vid    = 0
matched_km     = 0

def apply_sheet_entry(ad, entry):
    if entry.get("fb_post_url"): ad["fb_post_url"]      = entry["fb_post_url"]
    if entry.get("creator"):     ad["_sheet_creator"]   = entry["creator"]
    if entry.get("product"):     ad["_sheet_product"]   = entry["product"]
    if entry.get("key_msg"):     ad["_sheet_key_msg"]   = entry["key_msg"]

for ad in ao_data.get("ads", []):
    aid      = ad.get("ad_id", "")
    ad_name  = ad.get("ad_name", "").strip()
    cr       = creative_map.get(aid, {})

    # thumbnail + video_id
    thumb = cr.get("thumb") or ""
    vid   = cr.get("video_id") or ""
    obj_t = cr.get("type", "")
    if thumb: ad["thumbnail_url"] = thumb; injected_thumb += 1
    if vid:   ad["video_id"] = vid
    if obj_t: ad["creative_type"] = obj_t

    # P1: exact name
    if ad_name in name_map:
        apply_sheet_entry(ad, name_map[ad_name])
        matched_name += 1
    # P2: fuzzy name
    elif best := fuzzy_match_name(ad_name):
        apply_sheet_entry(ad, name_map[best])
        matched_fuzzy += 1
    # P3: video_id
    elif vid and vid in vid_map:
        apply_sheet_entry(ad, vid_map[vid])
        matched_vid += 1
    else:
        # P4: key_msg substring
        ad_creator = ad.get("_creator","") or ""
        for cr_sheet, km, entry in keymsg_list:
            if km and km in ad_name:
                if not cr_sheet or cr_sheet in ad_name or cr_sheet.strip() == ad_creator.strip():
                    apply_sheet_entry(ad, entry)
                    matched_km += 1
                    break

# daily_ads
for ad in ao_data.get("daily_ads", []):
    aid = ad.get("ad_id", "")
    cr  = creative_map.get(aid, {})
    if cr.get("thumb"):     ad["thumbnail_url"] = cr["thumb"]
    if cr.get("video_id"):  ad["video_id"]      = cr["video_id"]
    vid = ad.get("video_id","")
    nm  = ad.get("ad_name","").strip()
    if nm in name_map:           apply_sheet_entry(ad, name_map[nm])
    elif best := fuzzy_match_name(nm): apply_sheet_entry(ad, name_map[best])
    elif vid in vid_map:         apply_sheet_entry(ad, vid_map[vid])

print(f"\nResults:")
print(f"  thumbnail injected  : {injected_thumb}/{len(ao_data.get('ads',[]))}")
print(f"  matched exact name  : {matched_name}")
print(f"  matched fuzzy name  : {matched_fuzzy}")
print(f"  matched video_id    : {matched_vid}")
print(f"  matched key_message : {matched_km}")
print(f"  total fb_post_url   : {matched_name+matched_fuzzy+matched_vid+matched_km}")

# ────────────────────────────────────────────────────────────
# 7. Fetch Ad Preview URLs (Meta Ad Preview API)
#    เฉพาะ active ads (auto_on=true) — ~40 ads, ใช้ parallel fetch
# ────────────────────────────────────────────────────────────
META_TOKEN  = "EAACeOZCbphhwBRdzDkHHUEdj9S32nc2ZB1y65lt1U7suL4nS5UIkwAjUrL82AWprNQ1G0bMx6cv0A7Vf1QWdiEQtOHndZAPzDdBdkDyC7fwJugazXUsHx81HyDFnCkgL5FhYZCNoZB6Gvalzd11CbSJq8TQvHgB6ZAFSVZB8j8dWmzTJaFd9h3Hy4E8bJkTFJIUyQZDZD"
API_VERSION = "v25.0"

def fetch_preview(ad_id):
    url = (f"https://graph.facebook.com/{API_VERSION}/{ad_id}/previews"
           f"?ad_format=MOBILE_FEED_STANDARD&access_token={META_TOKEN}")
    try:
        with urllib.request.urlopen(url, timeout=12) as r:
            data = json.loads(r.read())
        body = (data.get("data") or [{}])[0].get("body", "")
        mx = re.search(r'src="([^"]*business\.facebook\.com[^"]*)"', body)
        if mx:
            return ad_id, mx.group(1).replace("&amp;", "&")
    except Exception:
        pass
    return ad_id, None

# Build map: ad_id → ad object for active ads
active_ads = {a["ad_id"]: a for a in ao_data.get("ads", [])
              if str(a.get("auto_on","")).lower() == "true" and a.get("ad_id")}

print(f"\n[Preview] Fetching {len(active_ads)} active ads...")
p_ok = p_fail = 0
with ThreadPoolExecutor(max_workers=5) as ex:
    futures = {ex.submit(fetch_preview, aid): aid for aid in active_ads}
    for fut in as_completed(futures):
        aid, purl = fut.result()
        if purl:
            active_ads[aid]["preview_url"] = purl
            p_ok += 1
        else:
            p_fail += 1
print(f"  preview_url fetched: {p_ok}, failed: {p_fail}")

# Also clear stale preview_url from inactive ads (optional — keeps data lean)
for ad in ao_data.get("ads", []):
    if str(ad.get("auto_on","")).lower() != "true":
        ad.pop("preview_url", None)

# 8. Write back
new_json    = json.dumps(ao_data, ensure_ascii=False, separators=(",", ":"))
new_content = content[:m.start()] + f"{prefix}{new_json};" + content[m.end():]
with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(new_content)

print(f"\n✅ บันทึก index.html แล้ว — กด Cmd+Shift+R")
