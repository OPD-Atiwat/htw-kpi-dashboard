#!/usr/bin/env python3
"""
upsell_pull.py — ดึง MK13 (order-level) จาก API mms-admin แล้วสรุป Upsell/ซื้อคู่ ราย "เล่ม"
เขียนผลลง var UPSELL_DATA + UPSELL_TS ใน index.html (แทนการ parse CSV)

- จับคู่จาก "ออเดอร์เดียวกัน" (order field) · anchor = เล่มราคาสูงสุดในชุด
- bundle = หลายเล่มใน order เดียว หรือ โปรขึ้นต้น UP####/BOXSET
- freebie = is_giveaway
- รันใน opd_runner.py (มี freshness guard: ข้ามถ้าเพิ่งอัปเดต < REFRESH_HOURS ชม.)

หมายเหตุ: ถ้า field order/promotion ของ API เปลี่ยน → ดู log บรรทัด "keys:" แล้วปรับ ORDER_KEYS/PROMO_KEYS
"""
import re, json, datetime, urllib.request
from collections import defaultdict

API   = "https://mms-admin.opendurian.com/api/report/marketing_team"
HTML  = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/index.html"
M07   = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard/05_Scripts/m07_pull.py"
TEAM  = 35
LOOKBACK_DAYS = 120        # หน้าต่างข้อมูลที่ใช้วิเคราะห์ pattern
REFRESH_HOURS = 20         # ข้ามถ้าเพิ่งอัปเดตไม่นาน (pattern เปลี่ยนช้า ไม่ต้องดึงทุกชม.)
MIN_UNITS = 50             # เล่มต้องขาย >= นี้ ถึงสรุป

ORDER_KEYS = ["order_no","order_number","sale_order_no","sale_order","transaction_no",
              "transactions_no","so_no","order","ref_no","bill_no","bill_id","document_no"]
PROMO_KEYS = ["promotion","promotion_name","promo","promo_name","promotion_title"]
NAME_KEYS  = ["product_name","product","name"]


def get_token():
    return re.search(r'TOKEN\s*=\s*"([^"]+)"', open(M07, encoding="utf-8").read()).group(1)


def map_channel(r):
    """ตรง logic กับ mk13_pull.map_channel — ให้ key ตรงกับ OPD_PROD_DATA"""
    if r.get("is_tiktok_affiliate"):
        return "TikTok Affi"
    ch = (r.get("sale_channel") or "").strip()
    mt = (r.get("sale_method") or "").strip()
    ml = mt.lower()
    if ch == "TikTok":
        if mt == "Affiliate":
            return "TikTok Affi"
        if "live" in ml:
            return "TikTok Live"
        return "TikTok"
    if ch == "Shopee":
        return "Shopee Live" if "live" in ml else "Shopee"
    if ch == "Facebook":
        return "Shopify" if mt in ("Salepage", "Shopify", "FBSalepage") else "Facebook"
    return {"Instagram": "Instagram", "LINE": "LINE", "Line": "LINE", "YouTube": "YouTube",
            "Lazada": "Lazada", "Web": "Web", "Website": "Web", "หน้าร้าน": "หน้าร้าน"}.get(ch)


def fetch_day(day, token):
    rows, page = [], 1
    while True:
        body = {"reportsID": 13, "marketingTeamID": TEAM, "startDate": day, "endDate": day,
                "panal": 1, "page": page, "row_per_pages": 500,
                "sorting": "transactions_time", "sorting_option": False}
        req = urllib.request.Request(API, data=json.dumps(body).encode(),
              headers={"Content-Type": "application/json", "Authorization": token})
        d = json.load(urllib.request.urlopen(req, timeout=90))
        m = d.get("metric", [])
        rows += m
        if len(m) < 500:
            break
        page += 1
    return rows


def pick_key(sample, candidates):
    for k in candidates:
        if k in sample:
            return k
    # fuzzy: หา key ที่ contain คำใน candidates
    low = {k.lower(): k for k in sample}
    for c in candidates:
        for lk, orig in low.items():
            if c.lower() in lk:
                return orig
    return None


def read_var(html, name):
    m = re.search(r"var %s = (.*?);\n" % name, html, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except Exception:
        return None


def write_var(html, name, value_js):
    pat = r"var %s = .*?;" % name
    new = "var %s = %s;" % (name, value_js)
    if re.search(pat, html):
        return re.subn(pat, new, html, count=1)
    # ถ้ายังไม่มี var → แทรกหลัง var _CAMP_MONUM (จุดอ้างอิงที่มีแน่)
    anchor = "var _CAMP_MONUM="
    i = html.find(anchor)
    if i < 0:
        return html, 0
    j = html.find("\n", i) + 1
    return html[:j] + new + "\n" + html[j:], 1


def main():
    # freshness guard
    with open(HTML, encoding="utf-8") as f:
        html = f.read()
    m = re.search(r'var UPSELL_TS = "([^"]+)";', html)
    if m:
        try:
            last = datetime.datetime.strptime(m.group(1), "%Y-%m-%d %H:%M")
            if (datetime.datetime.now() - last).total_seconds() < REFRESH_HOURS * 3600:
                print(f"  ℹ️  UPSELL อัปเดตล่าสุด {m.group(1)} (< {REFRESH_HOURS}ชม.) — ข้าม")
                return
        except Exception:
            pass

    token = get_token()
    today = datetime.date.today()
    start = today - datetime.timedelta(days=LOOKBACK_DAYS)

    orders = defaultdict(list)
    okey = pkey = nkey = None
    d = start
    total = 0
    while d <= today:
        try:
            rows = fetch_day(d.isoformat(), token)
        except Exception as e:
            print(f"  ⚠️  {d}: {e}")
            d += datetime.timedelta(days=1)
            continue
        for r in rows:
            if okey is None and r:
                okey = pick_key(r, ORDER_KEYS)
                pkey = pick_key(r, PROMO_KEYS)
                nkey = pick_key(r, NAME_KEYS)
                print(f"  keys: order={okey} promo={pkey} name={nkey}")
                print(f"  (all keys: {sorted(r.keys())})")
            if okey is None:
                continue
            name = (r.get(nkey) or "").strip()
            ono  = (r.get(okey) or "")
            if not name or not ono or name.startswith("*"):
                continue
            try:
                qty = int(float(r.get("quantity") or 0))
            except Exception:
                qty = 0
            try:
                price = float(r.get("amount") or 0)
            except Exception:
                price = 0
            promo = (r.get(pkey) or "").strip() if pkey else ""
            free = bool(r.get("is_giveaway"))
            orders[str(ono)].append({"book": name, "qty": qty, "price": price,
                                     "free": free, "promo": promo,
                                     "date": d.isoformat(), "ch": map_channel(r)})
            total += 1
        d += datetime.timedelta(days=1)

    if okey is None or not orders:
        print("  ⚠️  หา order field ไม่เจอ หรือไม่มีข้อมูล — ยกเลิก (ดู keys ด้านบน)")
        return
    print(f"  orders={len(orders)} lines={total}")

    book = defaultdict(lambda: {"u": 0, "bu": 0, "fu": 0, "ao": 0, "bo": 0, "pairs": defaultdict(int)})
    core = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))  # core[book][date][ch] = ยอดไม่รวม upsell
    pday = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))    # pday[book][date][co-book] = #orders พ่วงคู่กัน
    for ono, lines in orders.items():
        distinct = set(l["book"] for l in lines)
        multi = len(distinct) > 1
        mainbook = max(lines, key=lambda l: l["price"])["book"] if lines else None
        order_bundle = multi or any(re.match(r'\s*(UP\d|BOXSET)', l["promo"], re.I) for l in lines)
        odate = lines[0]["date"] if lines else ""
        if not order_bundle:
            for l in lines:
                if l["ch"]:
                    core[l["book"]][l["date"]][l["ch"]] += l["price"]
        elif multi and odate:
            for b in distinct:
                for ob in distinct:
                    if ob != b:
                        pday[b][odate][ob] += 1
        for b in distinct:
            bl = [l for l in lines if l["book"] == b]
            u = sum(l["qty"] for l in bl)
            isb = multi or any(re.match(r'\s*(UP\d|BOXSET)', l["promo"], re.I) for l in bl)
            book[b]["u"] += u
            book[b]["fu"] += sum(l["qty"] for l in bl if l["free"])
            if isb:
                book[b]["bu"] += u
                book[b]["bo"] += 1
                if b == mainbook:
                    book[b]["ao"] += 1
                for ob in distinct:
                    if ob != b:
                        book[b]["pairs"][ob] += 1

    out = {}
    for b, x in book.items():
        if x["u"] < MIN_UNITS or "E-BOOK" in b or b.startswith("SHEET") or b.startswith("คอร์ส"):
            continue
        pairs = sorted(x["pairs"].items(), key=lambda p: -p[1])[:3]
        out[b] = {
            "bun":  round(x["bu"] / x["u"] * 100),
            "free": round(x["fu"] / x["u"] * 100),
            "anc":  round(x["ao"] / x["bo"] * 100) if x["bo"] else 0,
            "pairs": [[p[0][:46], p[1]] for p in pairs if p[1] >= 10],
        }

    # CORE_PROD_DATA: ยอด "ไม่รวม upsell" ราย เล่ม/วัน/ช่อง (เฉพาะเล่มใน out)
    core_out = {}
    for b in out:
        if b in core:
            core_out[b] = {dt: {ch: round(v) for ch, v in chs.items() if v}
                           for dt, chs in core[b].items()}

    # UPSELL_PAIR_DAILY: top 2 เล่มที่พ่วงคู่ ราย เล่ม/วัน
    pday_out = {}
    for b in out:
        if b in pday:
            dd = {}
            for dt, cobooks in pday[b].items():
                top = sorted(cobooks.items(), key=lambda p: -p[1])[:2]
                if top:
                    dd[dt] = [[c[:30], n] for c, n in top]
            if dd:
                pday_out[b] = dd

    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    html, n1 = write_var(html, "UPSELL_DATA", json.dumps(out, ensure_ascii=False, separators=(",", ":")))
    html, n3 = write_var(html, "CORE_PROD_DATA", json.dumps(core_out, ensure_ascii=False, separators=(",", ":")))
    html, n4 = write_var(html, "UPSELL_PAIR_DAILY", json.dumps(pday_out, ensure_ascii=False, separators=(",", ":")))
    html, n2 = write_var(html, "UPSELL_TS", '"%s"' % ts)
    with open(HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✅ UPSELL_DATA {len(out)} เล่ม · CORE_PROD_DATA {len(core_out)} เล่ม (window {LOOKBACK_DAYS}วัน) · TS {ts}")


if __name__ == "__main__":
    main()
