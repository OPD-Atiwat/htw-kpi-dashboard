#!/usr/bin/env python3
"""
Meta Ads Fetcher — OpenDurian How-to
ดึง 2 ชุดข้อมูล:
  1. เมื่อวาน (daily) — เทียบรายคนว่าใครดีแค่ไหนวันนั้น
  2. MTD (ตั้งแต่ต้นเดือนถึงเมื่อวาน) — leaderboard สะสม
บันทึกเป็น meta_ads_data.json ให้ Cowork Scheduled Task อ่านต่อ
"""

import json
import urllib.request
import urllib.parse
from datetime import datetime, timedelta

# ─── CONFIG ───────────────────────────────────────────────────────────────────
ACCESS_TOKEN  = "EAACeOZCbphhwBRSQPDoqzio8GbcosSyhFmd5y2vZChSWU61DLBGZC3qRwKFoOaO81ggGa13QewrrtTnImZBAqCPMwZC4B43aSc6kt3Nr7NCf5OajVSuIiOo0WTYMCv2fesCqdls5bFvMyZBh64S60onLsemWLZBa4IUtCFHKM4iqWXYAjXN8lt23EpBVhV5VIZCsFwZDZD"
AD_ACCOUNT_ID = "act_303814252252288"
OUTPUT_FILE   = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/meta_ads_data.json"

API_VERSION   = "v19.0"
BASE_URL      = f"https://graph.facebook.com/{API_VERSION}"

# Threshold ต้นทุน/Messenger (บาท)
COST_MSG_GOOD   = 100   # ต่ำกว่านี้ = ดี (เขียว)
COST_MSG_DANGER = 300   # สูงกว่านี้ = ควร pause (แดง)
# ──────────────────────────────────────────────────────────────────────────────


def fetch(url):
    with urllib.request.urlopen(url, timeout=30) as resp:
        return json.loads(resp.read())


INSIGHT_FIELDS_COMMON = [
    "spend", "impressions", "reach", "clicks", "ctr", "cpm",
    "actions", "cost_per_action_type", "purchase_roas", "action_values",
]

def _fetch_insights(date_start, date_end, level, extra_fields, limit=200):
    """generic insight fetcher — level = 'ad' หรือ 'adset'"""
    fields = ",".join(extra_fields + INSIGHT_FIELDS_COMMON)
    params = urllib.parse.urlencode({
        "fields":       fields,
        "level":        level,
        "time_range":   json.dumps({"since": date_start, "until": date_end}),
        "limit":        limit,
        "access_token": ACCESS_TOKEN,
    })
    all_rows = []
    url = f"{BASE_URL}/{AD_ACCOUNT_ID}/insights?{params}"
    while url:
        data = fetch(url)
        all_rows.extend(data.get("data", []))
        url  = data.get("paging", {}).get("next")
    return all_rows


def get_insights(date_start, date_end, limit=200):
    """ดึง insights ระดับ Ad (ad_id + adset_id เพื่อ join)"""
    return _fetch_insights(
        date_start, date_end,
        level="ad",
        extra_fields=["ad_id", "ad_name", "adset_id", "adset_name", "campaign_name"],
        limit=limit,
    )


def get_adset_insights(date_start, date_end, limit=200):
    """ดึง insights ระดับ Ad Set (aggregate per audience group)"""
    return _fetch_insights(
        date_start, date_end,
        level="adset",
        extra_fields=["adset_id", "adset_name", "campaign_name"],
        limit=limit,
    )


def parse_msg(row):
    """จำนวน Messenger conversations started"""
    for item in row.get("actions", []):
        if item.get("action_type") == "onsite_conversion.messaging_conversation_started_7d":
            return int(float(item.get("value", 0)))
    return 0


def parse_cost_per_msg(row):
    """ต้นทุนต่อ Messenger"""
    for item in row.get("cost_per_action_type", []):
        if item.get("action_type") == "onsite_conversion.messaging_conversation_started_7d":
            return float(item.get("value", 0))
    spend = float(row.get("spend", 0))
    msg   = parse_msg(row)
    if msg > 0:
        return round(spend / msg, 2)
    return None


def parse_revenue(row):
    """ดึง Revenue จาก action_values (ตรงกับ meta_pull.py)"""
    for item in row.get("action_values", []) or []:
        if item.get("action_type") in (
            "omni_purchase",
            "offsite_conversion.fb_pixel_purchase",
            "purchase",
        ):
            return float(item.get("value", 0))
    return 0.0


def parse_purchases(row):
    for item in row.get("actions", []) or []:
        if item.get("action_type") in (
            "omni_purchase",
            "offsite_conversion.fb_pixel_purchase",
            "purchase",
        ):
            return int(float(item.get("value", 0)))
    return 0


def parse_roas(row):
    revenue = parse_revenue(row)
    spend   = float(row.get("spend", 0))
    if revenue > 0 and spend > 0:
        return round(revenue / spend, 2)
    return None


def parse_adset_meta(adset_name):
    """แยก audience / age / format / auto_on จากชื่อ Ad Set"""
    import re as _re
    audience = 'No Audience'
    age      = '-'
    fmt      = '-'
    auto_on  = bool(_re.search(r'auto.?on', adset_name, _re.IGNORECASE))

    for b in _re.findall(r'\[([^\]]+)\]', adset_name):
        b = b.strip()
        if b.upper() in ('IMG', 'VDO', 'LIVE'):
            fmt = b.upper(); continue
        if _re.match(r'^Age\s*\d', b, _re.IGNORECASE) or _re.match(r'^\d{2}[\+\-]\d*$', b):
            age = _re.sub(r'^[Aa]ge\s*', '', b).strip(); continue
        if _re.match(r'^\d+\s*[A-Za-z]{3}', b):
            continue
        if _re.match(r'^no.?audience', b, _re.IGNORECASE):
            audience = 'No Audience'; continue
        if _re.match(r'^auto.?on', b, _re.IGNORECASE) or b in ('Cross Page', 'Kiwtum', 'KMS'):
            continue
        audience = b

    def _group(aud):
        if 'Books&' in aud or aud == 'Books': return 'Interest — Books'
        if aud in ('Motivation','Reading','Emotional','Motherhood','Comics'): return 'Interest — Other'
        if 'Engage' in aud or 'LikeFollow' in aud or 'Like&Follow' in aud: return 'Retarget — Page'
        if 'Video-View' in aud or 'VDO View' in aud or 'Ghostly VDO' in aud: return 'Retarget — Video'
        if 'Purchase' in aud or 'Inbox' in aud or 'LLK' in aud: return 'Retarget — Conversion'
        if aud == 'No Audience': return 'Broad'
        return 'Other'

    return {
        'audience':       audience,
        'audience_group': _group(audience),
        'age':            age,
        'fmt':            fmt,
        'auto_on':        auto_on,
    }


def enrich_ads(rows, min_spend=50):
    """enrich ระดับ Ad — parse audience/age/fmt จาก adset_name"""
    result = []
    for r in rows:
        spend = float(r.get("spend", 0))
        if spend < min_spend:
            continue
        adset_meta = parse_adset_meta(r.get("adset_name", ""))
        result.append({
            "ad_id":          r.get("ad_id", ""),
            "ad_name":        r.get("ad_name", ""),
            "adset_id":       r.get("adset_id", ""),
            "adset_name":     r.get("adset_name", ""),
            "campaign_name":  r.get("campaign_name", ""),
            "audience":       adset_meta["audience"],
            "audience_group": adset_meta["audience_group"],
            "age":            adset_meta["age"],
            "fmt":            adset_meta["fmt"],
            "auto_on":        adset_meta["auto_on"],
            "spend":          round(spend, 2),
            "revenue":        round(parse_revenue(r), 2),
            "purchases":      parse_purchases(r),
            "roas":           parse_roas(r),
            "impressions":    int(r.get("impressions", 0)),
            "reach":          int(r.get("reach", 0)),
            "clicks":         int(r.get("clicks", 0)),
            "ctr":            round(float(r.get("ctr", 0)), 2),
            "cpm":            round(float(r.get("cpm", 0)), 2),
            "msg":            parse_msg(r),
            "cost_per_msg":   parse_cost_per_msg(r),
        })
    return result


def enrich_adsets(rows, min_spend=50):
    """enrich ระดับ Ad Set — metric รวมต่อ audience group"""
    result = []
    for r in rows:
        spend = float(r.get("spend", 0))
        if spend < min_spend:
            continue
        revenue   = parse_revenue(r)
        purchases = parse_purchases(r)
        cpp = round(spend / purchases, 2) if purchases > 0 else None
        result.append({
            "adset_id":     r.get("adset_id", ""),
            "adset_name":   r.get("adset_name", ""),
            "campaign_name":r.get("campaign_name", ""),
            "spend":        round(spend, 2),
            "revenue":      round(revenue, 2),
            "purchases":    purchases,
            "roas":         parse_roas(r),
            "cpp":          cpp,
            "impressions":  int(r.get("impressions", 0)),
            "reach":        int(r.get("reach", 0)),
            "clicks":       int(r.get("clicks", 0)),
            "ctr":          round(float(r.get("ctr", 0)), 2),
            "cpm":          round(float(r.get("cpm", 0)), 2),
            "msg":          parse_msg(r),
            "cost_per_msg": parse_cost_per_msg(r),
        })
    return result


# backward-compat alias
def enrich(rows, min_spend=50):
    return enrich_ads(rows, min_spend)


def fetch_ad_creatives(ad_ids):
    """
    ดึง creative assets สำหรับแต่ละ ad_id:
      - thumbnail_url  (รูป preview)
      - image_url      (รูปเต็มสำหรับ image ads)
      - video_id       (สำหรับ video ads → ดึง video_url ต่อ)
      - object_type    (VIDEO / IMAGE / etc.)
    Return: dict {ad_id: {...}}
    """
    result = {}
    ad_ids = list(set(ad_ids))
    BATCH  = 50

    for i in range(0, len(ad_ids), BATCH):
        chunk   = ad_ids[i:i+BATCH]
        ids_str = ",".join(chunk)
        params  = urllib.parse.urlencode({
            "ids":          ids_str,
            "fields":       "id,creative{thumbnail_url,image_url,video_id,object_type}",
            "access_token": ACCESS_TOKEN,
        })
        url = f"{BASE_URL}/?{params}"
        try:
            data = fetch(url)
            for ad_id, ad_data in data.items():
                cr = ad_data.get("creative", {})
                result[ad_id] = {
                    "thumbnail_url": cr.get("thumbnail_url"),
                    "image_url":     cr.get("image_url"),
                    "video_id":      cr.get("video_id"),
                    "object_type":   cr.get("object_type", ""),
                }
        except Exception as e:
            print(f"  ⚠ creative fetch error (batch {i}): {e}")
    return result


def fetch_video_urls(video_ids):
    """
    ดึง MP4 source URL จาก video_id
    Return: dict {video_id: {"video_url": ..., "picture": ...}}
    """
    result = {}
    video_ids = list(set(v for v in video_ids if v))
    if not video_ids:
        return result
    BATCH = 50

    for i in range(0, len(video_ids), BATCH):
        chunk   = video_ids[i:i+BATCH]
        ids_str = ",".join(chunk)
        params  = urllib.parse.urlencode({
            "ids":          ids_str,
            "fields":       "id,source,picture",   # source = MP4 URL
            "access_token": ACCESS_TOKEN,
        })
        url = f"{BASE_URL}/?{params}"
        try:
            data = fetch(url)
            for vid_id, vid_data in data.items():
                result[vid_id] = {
                    "video_url": vid_data.get("source"),
                    "picture":   vid_data.get("picture"),
                }
        except Exception as e:
            print(f"  ⚠ video URL fetch error (batch {i}): {e}")
    return result


# backward-compat alias for old code
def fetch_ad_thumbnails(ad_ids):
    return fetch_ad_creatives(ad_ids)


def merge_thumbnails(ads, creative_map, video_map=None):
    """inject thumbnail_url / image_url / video_url into each ad dict"""
    video_map = video_map or {}
    for ad in ads:
        c = creative_map.get(ad.get("ad_id", ""), {})
        ad["thumbnail_url"]  = c.get("thumbnail_url")
        ad["image_url"]      = c.get("image_url")
        ad["creative_type"]  = c.get("object_type", "")
        vid_id = c.get("video_id")
        if vid_id and vid_id in video_map:
            ad["video_url"]  = video_map[vid_id].get("video_url")
            # use video picture as fallback thumbnail
            if not ad["thumbnail_url"]:
                ad["thumbnail_url"] = video_map[vid_id].get("picture")
        else:
            ad["video_url"]  = None
    return ads


def main():
    today     = datetime.now()
    yesterday = today - timedelta(days=1)
    mtd_start = today.replace(day=1)

    yesterday_str = yesterday.strftime("%Y-%m-%d")
    mtd_start_str = mtd_start.strftime("%Y-%m-%d")
    # MTD ถึงเมื่อวาน (ข้อมูลวันนี้ยังไม่ครบ)
    mtd_end_str   = yesterday_str

    now_str = today.strftime("%Y-%m-%d %H:%M")
    print(f"[{now_str}] ดึงข้อมูล Meta Ads...")
    print(f"  เมื่อวาน: {yesterday_str}")
    print(f"  MTD:      {mtd_start_str} → {mtd_end_str}")

    # ── Ad level (daily + MTD) ──────────────────────────────────
    daily_rows = enrich_ads(get_insights(yesterday_str, yesterday_str), min_spend=50)
    mtd_rows   = enrich_ads(get_insights(mtd_start_str, mtd_end_str),   min_spend=100)

    # ── Ad Set level (MTD เท่านั้น — ใช้ใน Ads Optimize tab) ───
    print(f"  ดึง Ad Set level MTD...")
    mtd_adsets = enrich_adsets(get_adset_insights(mtd_start_str, mtd_end_str), min_spend=100)

    # ── Creatives: thumbnail + video_url ────────────────────────
    all_ad_ids = list(set(
        [a["ad_id"] for a in daily_rows if a.get("ad_id")] +
        [a["ad_id"] for a in mtd_rows   if a.get("ad_id")]
    ))
    print(f"  ดึง creative assets สำหรับ {len(all_ad_ids)} ads...")
    creative_map = fetch_ad_creatives(all_ad_ids)

    # ดึง video source URLs สำหรับ video ads
    video_ids = list(set(
        v["video_id"] for v in creative_map.values()
        if v.get("video_id")
    ))
    print(f"  ดึง video URLs สำหรับ {len(video_ids)} videos...")
    video_map  = fetch_video_urls(video_ids)

    daily_rows = merge_thumbnails(daily_rows, creative_map, video_map)
    mtd_rows   = merge_thumbnails(mtd_rows,   creative_map, video_map)

    vid_count   = sum(1 for a in mtd_rows if a.get("video_url"))
    thumb_count = sum(1 for a in mtd_rows if a.get("thumbnail_url"))
    print(f"  video_url: {vid_count}  thumbnail: {thumb_count}  / {len(mtd_rows)} MTD ads")

    print(f"  daily ads: {len(daily_rows)}, MTD ads: {len(mtd_rows)}, MTD adsets: {len(mtd_adsets)}")

    output = {
        "fetched_at":    now_str,
        "yesterday":     yesterday_str,
        "mtd_start":     mtd_start_str,
        "mtd_end":       mtd_end_str,
        "daily_ads":     daily_rows,
        "mtd_ads":       mtd_rows,
        "mtd_adsets":    mtd_adsets,   # ← ใหม่: Ad Set level สำหรับ Ads Optimize tab
        "thresholds": {
            "cost_msg_good":   COST_MSG_GOOD,
            "cost_msg_danger": COST_MSG_DANGER,
        }
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"  บันทึกแล้ว → {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
