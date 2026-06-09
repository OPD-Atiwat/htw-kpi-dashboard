"""
fetch_video_sources.py — ดึง "ไฟล์วิดีโอ MP4 ตรงๆ" ของแต่ละ Ad แล้วฝัง <video> ได้เลย
ทางเลือกแทน Ad Preview iframe (ที่ Reels/Story คืน "Story Unavailable")

หลักการ:
  1. GET /{ad_id}?fields=creative{video_id,thumbnail_url,image_url,object_story_spec}
  2. หา video_id (creative.video_id หรือ object_story_spec.video_data.video_id)
  3. GET /{video_id}?fields=source,picture  → source = MP4 ตรง, picture = thumbnail
  4. เขียน ad.video_url + ad.thumbnail_url กลับเข้า AO_DATA ใน index.html

หมายเหตุ: URL ไฟล์วิดีโอจาก Meta CDN หมดอายุใน ~ชั่วโมง/วัน → ต้องรันในไปป์ไลน์รายวัน
รัน: python3 05_Scripts/fetch_video_sources.py
"""
import json, re, urllib.request, urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE       = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard"
HTML_PATH  = f"{BASE}/index.html"
META_TOKEN = "EAACeOZCbphhwBRdzDkHHUEdj9S32nc2ZB1y65lt1U7suL4nS5UIkwAjUrL82AWprNQ1G0bMx6cv0A7Vf1QWdiEQtOHndZAPzDdBdkDyC7fwJugazXUsHx81HyDFnCkgL5FhYZCNoZB6Gvalzd11CbSJq8TQvHgB6ZAFSVZB8j8dWmzTJaFd9h3Hy4E8bJkTFJIUyQZDZD"
API_VER    = "v25.0"


def _get(path, fields):
    url = (f"https://graph.facebook.com/{API_VER}/{path}"
           f"?fields={urllib.parse.quote(fields)}&access_token={META_TOKEN}")
    try:
        with urllib.request.urlopen(url, timeout=15) as r:
            return json.loads(r.read())
    except Exception:
        return None


def resolve_video(ad_id):
    """คืน (ad_id, video_url, thumb, post_url)
       ลองหลายทาง: creative.video_id → object_story_spec → effective_object_story_id (post)"""
    cr = _get(ad_id, "creative{video_id,thumbnail_url,image_url,effective_object_story_id,"
                     "object_story_id,instagram_permalink_url,"
                     "object_story_spec{video_data{video_id},template_data{video_id}}}")
    if not cr:
        return ad_id, None, None, None
    creative = cr.get("creative", {}) or {}
    thumb = creative.get("thumbnail_url") or creative.get("image_url") or ""
    post_url = creative.get("instagram_permalink_url") or ""

    # 1) หา video_id จากหลายที่
    spec = creative.get("object_story_spec", {}) or {}
    vid = (creative.get("video_id")
           or (spec.get("video_data", {}) or {}).get("video_id")
           or (spec.get("template_data", {}) or {}).get("video_id"))

    src = ""
    if vid:
        v = _get(vid, "source,picture,permalink_url")
        if v:
            src = v.get("source") or ""
            thumb = thumb or v.get("picture") or ""
            post_url = post_url or v.get("permalink_url") or ""

    # 2) ผ่าน effective_object_story_id (โพสต์/Reel ที่บูสต์) — เผื่อ video source + permalink
    story_id = creative.get("effective_object_story_id") or creative.get("object_story_id")
    if story_id:
        p = _get(story_id, "permalink_url,source,picture,full_picture")
        if p:
            post_url = post_url or p.get("permalink_url") or ""
            src      = src or p.get("source") or ""
            thumb    = thumb or p.get("picture") or p.get("full_picture") or ""

    return ad_id, (src or None), (thumb or None), (post_url or None)


print("=== Fetch Video Sources (MP4) ===")

with open(HTML_PATH, "r", encoding="utf-8") as f:
    content = f.read()

m = re.search(r'(var AO_DATA = )(\{.*?\});', content, re.DOTALL)
if not m:
    print("❌ ไม่พบ var AO_DATA"); exit(1)

prefix, ao_data = m.group(1), json.loads(m.group(2))
ads = ao_data.get("ads", [])

# ดึงเฉพาะ ad ที่มี spend (ลดจำนวน API call) + มี ad_id
targets = [a for a in ads if a.get("ad_id") and (a.get("spend") or 0) > 0]
ad_map  = {a["ad_id"]: a for a in targets}
print(f"Target ads: {len(targets)}")

ok = thumb_ok = post_ok = fail = 0
with ThreadPoolExecutor(max_workers=5) as ex:
    futures = {ex.submit(resolve_video, aid): aid for aid in ad_map}
    for fut in as_completed(futures):
        aid, vurl, thumb, post_url = fut.result()
        if vurl:
            ad_map[aid]["video_url"] = vurl
            ok += 1
        if thumb and not ad_map[aid].get("thumbnail_url"):
            ad_map[aid]["thumbnail_url"] = thumb
            thumb_ok += 1
        if post_url and not ad_map[aid].get("fb_post_url"):
            ad_map[aid]["fb_post_url"] = post_url
            post_ok += 1
        if not vurl and not thumb and not post_url:
            fail += 1

print(f"  ✅ video: {ok}  🖼 thumb: {thumb_ok}  🔗 post link: {post_ok}  ❌ none: {fail}")

new_json    = json.dumps(ao_data, ensure_ascii=False, separators=(",", ":"))
new_content = content[:m.start()] + f"{prefix}{new_json};" + content[m.end():]
with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(new_content)

print("✅ บันทึก index.html แล้ว — กด Cmd+Shift+R")
