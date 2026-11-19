"""Keep backfilling vesselreport RPA history until Jan 2026 is reached.

Each round: acquire git lock -> rob_mail_sync (324d, --apply) -> rebuild ->
safe push -> release lock -> check RPA_成功 oldest mail date. Stops when
oldest <= 2026-01-01 or MAX_ROUNDS reached. Safe to re-run.
"""
import os
import subprocess
import sys
import time
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
PY = r"C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
OWNER = "rob_untiljan"
MAX_ROUNDS = 10
TARGET = "2026-01-01"
LOG = os.path.join(BASE, "rob_data", "untiljan.log")


def log(msg):
    line = "[%s] %s" % (datetime.now().strftime("%H:%M:%S"), msg)
    print(line, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(args, env=None):
    e = dict(os.environ)
    e.update(env or {})
    p = subprocess.run(args, cwd=BASE, env=e, stdout=subprocess.PIPE,
                       stderr=subprocess.STDOUT, text=True, errors="replace")
    for ln in p.stdout.splitlines():
        if any(k in ln for k in ("mail cache", "parsed ", "wrote ",
                                 "DONE", "ERROR", "Traceback",
                                 "HIT")):
            log("  " + ln[:150])
    return p.returncode


def oldest_rpa_date():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    for i in range(1, ns.Stores.Count + 1):
        s = ns.Stores(i)
        if "vesselreport" not in (s.DisplayName or "").lower():
            continue
        for c in s.GetRootFolder().Folders:
            if c.Name == "RPA_成功" and c.Items.Count:
                items = c.Items
                items.Sort("[ReceivedTime]", True)  # newest first
                return str(items(items.Count).ReceivedTime)[:10], items.Count
    return None, 0


env = {
    "CLAW_LOCK_OWNER": OWNER,
    "ROB_FOLDER_CAP": "30000",
    "ROB_PER_VESSEL": "2000",
}

log("=== backfill-until-jan start, target <= %s ===" % TARGET)
for rnd in range(1, MAX_ROUNDS + 1):
    oldest, count = oldest_rpa_date()
    log("round %d/%d: RPA_成功 n=%d oldest=%s" % (rnd, MAX_ROUNDS, count, oldest))
    if oldest and oldest <= TARGET:
        log("target reached, stopping")
        break

    rc = run([PY, "git_op_lock.py", "acquire", "--owner", OWNER, "--wait", "180"], env)
    if rc != 0:
        log("lock acquire failed rc=%d, retry in 120s" % rc)
        time.sleep(120)
        continue
    try:
        rc = run([PY, "-u", "rob_mail_sync.py", "--days", "324", "--apply"], env)
        if rc != 0:
            log("mail_sync failed rc=%d" % rc)
            continue
        run([PY, "-u", "rob_fast_rebuild.py"], env)
        run([PY, "git_safe_push.py", "-m",
             "RPA history backfill round %d (until-jan)" % rnd], env)
    finally:
        run([PY, "git_op_lock.py", "release", "--owner", OWNER], env)

    oldest, count = oldest_rpa_date()
    log("round %d done: RPA_成功 n=%d oldest=%s" % (rnd, count, oldest))
    if oldest and oldest <= TARGET:
        log("target reached, stopping")
        break
    time.sleep(60)

log("=== backfill-until-jan finished ===")
