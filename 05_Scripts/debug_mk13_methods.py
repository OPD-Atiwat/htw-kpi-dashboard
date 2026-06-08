#!/usr/bin/env python3
"""Debug: print unique (Sale Channel, Sale Method) combinations จาก MK13"""
import csv, io, ssl, urllib.request
ssl._create_default_https_context = ssl._create_unverified_context

SHEET_ID = "1qYdwXuCHDHeHN6a8vU_RVFBYT5MuYq-8K5AMK3cP4DM"
MK13_GID = "964123706"

url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={MK13_GID}"
req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
with urllib.request.urlopen(req, timeout=60) as r:
    text = r.read().decode("utf-8-sig")

rows = list(csv.reader(io.StringIO(text)))
header = [h.strip() for h in rows[0]]
print(f"Total rows: {len(rows)} | Total cols: {len(header)}")
print(f"Header cols 20-24: {header[20:25]}\n")

ch_col = next(i for i, h in enumerate(header) if "Sale Channel" in h)
mt_col = next((i for i, h in enumerate(header) if "Sale Method" in h), -1)

combos = {}
for row in rows[1:]:
    ch = row[ch_col].strip() if ch_col < len(row) else ""
    mt = row[mt_col].strip() if 0 <= mt_col < len(row) else ""
    key = (ch, mt)
    combos[key] = combos.get(key, 0) + 1

print("Unique (Sale Channel, Sale Method) combos:")
for (ch, mt), cnt in sorted(combos.items()):
    if ch in ("TikTok", "Shopee", "TikTok Live", "Shopee Live"):
        print(f"  [{cnt:5d}x]  ch={ch!r:20s}  mt={mt!r}")
