#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Daily Bunker Fuel Report auto-update (standalone, no WorkBuddy needed).

Pipeline (mirrors the former WorkBuddy automation):
  1) get_prices.py          -> fetch real oil/bunker prices (web)
  2) generate_bunker_report.py -> build HTML into Claw-Report/reports/
  3) send_via_win32com.py   -> email via local Outlook (non-fatal if Outlook busy)
  4) git add/commit/push reports/ to LeahLiuL/Claw-Report main

Run by Windows Task Scheduler. Requires VPN not needed (web prices), but
Outlook must be running for the email step.
"""
import os, sys, subprocess
from datetime import datetime

AUTO_DIR   = os.path.dirname(os.path.abspath(__file__))          # Claw-Report
CLAMATION  = r"C:\Users\culadmin\WorkBuddy\Claw\automation"
REPO       = r"C:\Users\culadmin\Claw-Report"
PY         = sys.executable
LOG        = os.path.join(AUTO_DIR, "auto_bunker.log")
TOKEN_FILES = [os.path.join(AUTO_DIR, ".update_token"),
               os.path.expanduser("~/.update_token")]


def log(m):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {m}"
    print(line)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def get_token():
    for p in TOKEN_FILES:
        if os.path.exists(p):
            t = open(p, encoding="utf-8").read().strip()
            if t:
                return t
    return None


def run(cmd, cwd=None, env=None, fatal=True):
    log("+ " + " ".join(cmd))
    r = subprocess.run(cmd, cwd=cwd, env=env,
                       capture_output=True, text=True)
    out = (r.stdout or "") + (r.stderr or "")
    if out.strip():
        for ln in out.strip().splitlines()[-12:]:
            log("   " + ln)
    if r.returncode != 0 and fatal:
        raise RuntimeError(f"command failed ({r.returncode}): {' '.join(cmd)}")
    return r


def ensure_git_auth(repo):
    tok = get_token()
    if tok:
        url = f"https://LeahLiuL:{tok}@github.com/LeahLiuL/Claw-Report.git"
        run(["git", "-C", repo, "remote", "set-url", "origin", url], fatal=False)
    run(["git", "config", "--global", "credential.helper", "store"], fatal=False)


def main():
    log("=== Bunker auto-update START ===")
    ensure_git_auth(REPO)

    # Pull latest (in case reports changed on remote)
    run(["git", "-C", REPO, "pull", "--ff-only", "origin", "main"], fatal=False)

    # Step 1: fetch prices
    log("[1/3] fetch real oil/bunker prices")
    run([PY, os.path.join(CLAMATION, "get_prices.py")], cwd=CLAMATION)

    # Step 2: generate HTML report
    log("[2/3] generate bunker fuel report HTML")
    run([PY, os.path.join(CLAMATION, "generate_bunker_report.py")], cwd=CLAMATION)

    # Step 3: send via Outlook (non-fatal)
    log("[3/3] send report via Outlook")
    try:
        run([PY, os.path.join(CLAMATION, "send_via_win32com.py")], cwd=CLAMATION, fatal=False)
    except Exception as e:
        log(f"send email step failed (non-fatal): {e}")

    # Step 4: git push (only if changed)
    # 2026-09-24 起持 git_op_lock: 本机多自动化并发操作 Claw-Report 曾两次损坏 .git
    log("push to Claw-Report")
    import git_op_lock
    try:
        git_op_lock.acquire("bunker", wait_s=180)
    except git_op_lock.LockHeld as e:
        log(f"!! git op lock busy ({e}), skip push this round (next run retries)")
        log("=== Bunker auto-update DONE (push deferred) ===")
        return 0
    try:
        run(["git", "-C", REPO, "add", "reports/bunker-fuel-report*.html"], fatal=False)
        st = subprocess.run(["git", "-C", REPO, "status", "--porcelain"],
                            capture_output=True, text=True)
        if st.stdout.strip():
            today = datetime.now().strftime("%Y-%m-%d")
            run(["git", "-C", REPO, "commit", "-m", f"chore: bunker report {today}"], fatal=False)
            run(["git", "-C", REPO, "push", "origin", "main"], fatal=False)
            log("changes pushed")
        else:
            log("no report changes to push")
    finally:
        git_op_lock.release("bunker")

    log("=== Bunker auto-update DONE ===")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        log(f"FATAL: {e}")
        sys.exit(1)
