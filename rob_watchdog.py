# -*- coding: utf-8 -*-
"""
ROB watchdog -- detects SILENT failures of the scheduled ROB update.

Why this exists: on 2026-10-22 a UTF-8 Chinese comment inside
rob_fast_update.bat made cmd.exe die on line 1, so the task returned 255
four days in a row WITHOUT writing a single log line. Nobody noticed until
the user saw stale data on the page.

This script checks how long ago the last SUCCESSFUL run was
(a "===== date =====" header followed by "DONE"). If it is older than
MAX_AGE_HOURS it drops a loud warning file on the Desktop and prints it.
If everything is fine it removes the warning file.

Usage:  python rob_watchdog.py            (check + report)
        python rob_watchdog.py --quiet    (no desktop file, stdout only)
"""
import io
import os
import re
import sys
import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(BASE, "rob_data", "fastupdate.log")
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
WARN = os.path.join(DESKTOP, "ROB_UPDATE_FAILED.txt")
MAX_AGE_HOURS = 20          # two runs/day (01:00 + 13:00) -> 20h is a safe gap

HDR = re.compile(r"^=====\s*(\d{4})/(\d{2})/(\d{2})\D+(\d{1,2}):(\d{2}):(\d{2})")


def last_success(path):
    """Return datetime of the last run header that is followed by DONE."""
    if not os.path.exists(path):
        return None, "log file missing"
    with io.open(path, "r", encoding="utf-8", errors="replace") as f:
        lines = f.readlines()
    blocks = []                      # [(start_idx, datetime)]
    for i, ln in enumerate(lines):
        m = HDR.match(ln.strip())
        if m:
            y, mo, d, h, mi, s = (int(x) for x in m.groups())
            try:
                blocks.append((i, datetime.datetime(y, mo, d, h, mi, s)))
            except ValueError:
                pass
    for idx in range(len(blocks) - 1, -1, -1):
        start, dt = blocks[idx]
        end = blocks[idx + 1][0] if idx + 1 < len(blocks) else len(lines)
        if any(l.strip() == "DONE" for l in lines[start:end]):
            return dt, None
    if blocks:
        return None, "latest run (%s) did not reach DONE" % blocks[-1][1]
    return None, "no run header found in log"


def main():
    quiet = "--quiet" in sys.argv
    dt, err = last_success(LOG)
    now = datetime.datetime.now()

    if dt is None:
        msg = []
        msg.append("ROB AUTO UPDATE IS NOT WORKING")
        msg.append("checked: %s" % now.strftime("%Y-%m-%d %H:%M"))
        msg.append("reason : %s" % err)
        msg.append("")
        msg.append("The scheduled tasks ClawReport_ROB_0100 / _1300 run")
        msg.append("C:\\Users\\culadmin\\Claw-Report\\rob_fast_update.bat")
        msg.append("Check:  %s" % LOG)
        msg.append("Fix:    run that .bat manually and read the error output.")
        msg.append("HINT:   .bat files must be PURE ASCII. UTF-8 Chinese makes")
        msg.append("        cmd.exe die before writing anything (happened 10-22).")
        text = "\n".join(msg)
        print(text)
        if not quiet:
            with io.open(WARN, "w", encoding="utf-8") as f:
                f.write(text + "\n")
        return 1

    age = (now - dt).total_seconds() / 3600.0
    print("last successful ROB run: %s (%.1f h ago)" % (dt, age))
    if age > MAX_AGE_HOURS:
        text = ("ROB AUTO UPDATE IS LATE\n"
                "checked            : %s\n"
                "last successful run: %s\n"
                "gap                : %.1f hours (limit %d)\n\n"
                "Tasks: ClawReport_ROB_0100 / ClawReport_ROB_1300\n"
                "Log  : %s\n" % (now.strftime("%Y-%m-%d %H:%M"), dt, age,
                                 MAX_AGE_HOURS, LOG))
        print(text)
        if not quiet:
            with io.open(WARN, "w", encoding="utf-8") as f:
                f.write(text)
        return 1

    if os.path.exists(WARN):
        os.remove(WARN)
        print("removed stale warning file")
    print("OK - update pipeline healthy")
    return 0


if __name__ == "__main__":
    sys.exit(main())
