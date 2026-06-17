#!/usr/bin/env python3
"""
OpenDurian Dashboard Auto-Sync
รัน: meta_pull → opd_pull → kms_pull → git push
วาง script นี้ที่ ~/opd_runner.py (path ไม่มี colon)
Python จัดการ path ที่มี colon ได้โดยตรง ไม่ผ่าน shell
"""
import subprocess, os, sys, datetime

BASE    = "/Users/opendurian/Documents/Claude/Projects/Excel: Content KPI Dashboard"
SCRIPTS = os.path.join(BASE, "05_Scripts")
PYTHON  = "/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
LOG     = os.path.join(SCRIPTS, "meta_pull.log")

def w(msg):
    with open(LOG, "a") as f:
        f.write(f"{msg}\n")

def run_py(name):
    r = subprocess.run([PYTHON, os.path.join(SCRIPTS, name)], cwd=SCRIPTS)
    return r.returncode == 0

w("========================================")
w(f"  Run: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
w("========================================")

w("[1/5] Meta Ads + MK13 Sync...")
ok1 = run_py("meta_pull.py")
w("[1.5] MK13 API top-up (เติมวันที่ Sheet ยังไม่มี)...")
ok1b = run_py("mk13_pull.py")
w("[2/5] OPD Pull (Affiliate + Goals)...")
ok2 = run_py("opd_pull.py")
w("[2.5] M07 Margin (Actual + FC)...")
ok2b = run_py("m07_pull.py")
w("[2.6] M44 Ads cost แยกเล่ม (ADSM)...")
ok2c = run_py("m44_pull.py")
w("[3/5] KMS Sheet Pull...")
ok3 = run_py("kms_pull.py")
w("[4/5] Patch AO_DATA: thumbnails + Sheet mapping...")
ok4 = run_py("patch_ao_thumbnails.py")
w("[5/6] Fetch Ad Preview URLs (active ads)...")
ok5 = run_py("fetch_previews.py")
w("[6/6] Fetch Video Sources (MP4 ตรง — เล่นคลิปในหน้าได้)...")
ok6 = run_py("fetch_video_sources.py")

# Failsafe: อัปเดต META_PULL_TS ให้ตรงกับเวลา run ปัจจุบัน
# (ป้องกันกรณี mk13_sync fail แบบเงียบ ทำให้ timestamp ค้าง)
import re as _re
_dash = os.path.join(BASE, "index.html")
try:
    with open(_dash, "r", encoding="utf-8") as _f:
        _html = _f.read()
    _now_ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    _html2 = _re.sub(r'var META_PULL_TS\s*=\s*"[^"]*"',
                     f'var META_PULL_TS = "{_now_ts}"', _html)
    if _html2 != _html:
        with open(_dash, "w", encoding="utf-8") as _f:
            _f.write(_html2)
        w(f"  META_PULL_TS → {_now_ts}")
except Exception as _e:
    w(f"  META_PULL_TS update fail: {_e}")

w("--- สรุป ---")
w(f"  {'OK' if ok1 else 'FAIL'} meta_pull.py")
w(f"  {'OK' if ok1b else 'FAIL'} mk13_pull.py (top-up)")
w(f"  {'OK' if ok2 else 'FAIL'} opd_pull.py")
w(f"  {'OK' if ok2b else 'FAIL'} m07_pull.py")
w(f"  {'OK' if ok2c else 'FAIL'} m44_pull.py (ADSM)")
w(f"  {'OK' if ok3 else 'FAIL'} kms_pull.py")
w(f"  {'OK' if ok4 else 'FAIL'} patch_ao_thumbnails.py")
w(f"  {'OK' if ok5 else 'FAIL'} fetch_previews.py")
w(f"  {'OK' if ok6 else 'FAIL'} fetch_video_sources.py")

# Git push — ใช้ git -C แทน cd (ไม่ต้อง cd เข้า path ที่มี colon)
w("--- Git push ---")
subprocess.run(["find", os.path.join(BASE, ".git"), "-name", "*.lock", "-delete"],
               capture_output=True)
subprocess.run(["git", "-C", BASE, "add", "index.html"], capture_output=True)
diff = subprocess.run(["git", "-C", BASE, "diff", "--cached", "--quiet"])
if diff.returncode == 0:
    w("  skip: ไม่มีการเปลี่ยนแปลง")
else:
    ts_c = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    subprocess.run(["git", "-C", BASE, "commit", "-m", f"auto: {ts_c}"],
                   capture_output=True)
    # pull --rebase -X ours → ถ้า conflict ให้ยึด local version เสมอ (script เพิ่งอัปเดต)
    pull = subprocess.run(["git", "-C", BASE, "pull", "--rebase", "-X", "ours",
                            "origin", "main"], capture_output=True)
    if pull.returncode != 0:
        # rebase ล้มเหลวสนิท → abort แล้ว force push local
        subprocess.run(["git", "-C", BASE, "rebase", "--abort"], capture_output=True)
        push = subprocess.run(["git", "-C", BASE, "push", "--force-with-lease"],
                               capture_output=True)
        w(f"  Git push --force-with-lease: {'ok' if push.returncode==0 else 'FAIL'}")
    else:
        push = subprocess.run(["git", "-C", BASE, "push"], capture_output=True)
        w(f"  Git push: {'ok' if push.returncode==0 else 'FAIL'}")
    w("  Git push done")

w(f"Done: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
w("========================================")
