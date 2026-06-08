"""
fetch_previews.py — ดึง Ad Preview URL (Meta Ad Preview API)
ลอง multiple formats จนกว่าจะได้ preview จริง (ไม่ใช่ Story Unavailable)
รัน: python3 05_Scripts/fetch_previews.py
"""
import json, re, urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE       = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard"
HTML_PATH  = f"{BASE}/index.html"
META_TOKEN = "EAACeOZCbphhwBRdzDkHHUEdj9S32nc2ZB1y65lt1U7suL4nS5UIkwAjUrL82AWprNQ1G0bMx6cv0A7Vf1QWdiEQtOHndZAPzDdBdkDyC7fwJugazXUsHx81HyDFnCkgL5FhYZCNoZB6Gvalzd11CbSJq8TQvHgB6ZAFSVZB8j8dWmzTJaFd9h3Hy4E8bJkTFJIUyQZDZD"
API_VER    = "v25.0"

# ลอง format เรียงตามลำดับ — หยุดเมื่อได้ preview ที่ไม่มี error
AD_FORMATS = [
    "MOBILE_FEED_STANDARD",
    "DESKTOP_FEED_STANDARD",
    "MOBILE_INTERSTITIAL",       # Story / vertical
    "INSTAGRAM_STANDARD",
    "FACEBOOK_REELS_MOBILE",
]

def try_format(ad_id, fmt):
    url = (f"https://graph.facebook.com/{API_VER}/{ad_id}/previews"
           f"?ad_format={fmt}&access_token={META_TOKEN}")
    try:
        with urllib.request.urlopen(url, timeout=12) as r:
            data = json.loads(r.read())
        body = (data.get("data") or [{}])[0].get("body", "")
        mx = re.search(r'src="([^"]*business\.facebook\.com[^"]*)"', body)
        if mx:
            return mx.group(1).replace("&amp;", "&")
    except Exception:
        pass
    return None

def fetch_preview(ad_id):
    """ลอง formats ทีละอัน — return (ad_id, url, format_used)"""
    for fmt in AD_FORMATS:
        purl = try_format(ad_id, fmt)
        if purl:
            return ad_id, purl, fmt
    return ad_id, None, None

print("=== Fetch Ad Preview URLs ===")

# 1. Load HTML
with open(HTML_PATH, "r", encoding="utf-8") as f:
    content = f.read()

# 2. Extract AO_DATA
m = re.search(r'(var AO_DATA = )(\{.*?\});', content, re.DOTALL)
if not m:
    print("❌ ไม่พบ var AO_DATA")
    exit(1)

prefix  = m.group(1)
ao_data = json.loads(m.group(2))
ads     = ao_data.get("ads", [])

# 3. Active ads เท่านั้น
active = [(a["ad_id"], a) for a in ads
          if str(a.get("auto_on","")).lower() == "true" and a.get("ad_id")]
print(f"Active ads: {len(active)}")

# 4. Parallel fetch
ok = fail = 0
fmt_counts = {}
ad_map = {aid: ad for aid, ad in active}

with ThreadPoolExecutor(max_workers=5) as ex:
    futures = {ex.submit(fetch_preview, aid): aid for aid, _ in active}
    for fut in as_completed(futures):
        aid, purl, fmt = fut.result()
        if purl:
            ad_map[aid]["preview_url"] = purl
            ad_map[aid]["preview_format"] = fmt
            fmt_counts[fmt] = fmt_counts.get(fmt, 0) + 1
            ok += 1
        else:
            ad_map[aid].pop("preview_url", None)
            fail += 1

# Clear stale preview_url จาก inactive ads
for ad in ads:
    if str(ad.get("auto_on","")).lower() != "true":
        ad.pop("preview_url", None)
        ad.pop("preview_format", None)

print(f"  ✅ fetched: {ok}  ❌ no preview: {fail}")
for fmt, cnt in fmt_counts.items():
    print(f"     {fmt}: {cnt}")

# 5. Write back
new_json    = json.dumps(ao_data, ensure_ascii=False, separators=(",", ":"))
new_content = content[:m.start()] + f"{prefix}{new_json};" + content[m.end():]
with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(new_content)

print("✅ บันทึก index.html แล้ว — กด Cmd+Shift+R")
