# -*- coding: utf-8 -*-
"""回退 5 处 FWE 填报异常: 按原始附件证据, 把该油列改回 POB 值。
只改这 5 个精确 (vessel, report_time, oil) 单元格, 不碰其它行/列。
默认 dry-run; --apply 才写, 写前强制备份 CSV。

判定依据(逐封打开原始 xlsx 确认, 见 diag_fwe_final.py 输出):
  - FWE 值与同报告其它阶段(POB/ANCHOR/DROP)众数偏离巨大 -> FWE 填报错误, 退回 POB
  - FWE 与其它阶段一致(旧值是 RESUME SAILING 抓错) -> 正确修复, 不动
"""
import sys, os, csv, shutil, argparse
from datetime import datetime
sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import rob_refresh as R

HISTORY = os.path.join(BASE, "rob_data", "rob_history.csv")
# (vessel 关键词, report_time, oil列, 当前错误值, 回退到的 POB 值, 依据)
FIXES = [
    ("MEDKON LIA",     "2026-07-09 15:32:40", "mgo",   318.0,  33.2,  "POB MGO=33.2, FWE MGO=318 是错填(量级错)"),
    ("MEDKON LIA",     "2026-06-24 21:33:53", "ulsfo", 251.2, 396.0,  "POB ULSFO=396, FWE ULSFO=251.2 偏低37%"),
    ("STAR BLESSING",  "2026-07-22 22:15:54", "lsfo",  294.7, 894.8,  "POB/ANCHOR/DROP 都894-899, FWE=294.7 错填"),
    ("CUL HOCHIMINH",  "2026-08-16 16:25:36", "lsfo",  247.8, 191.55, "POB/ANCHOR=191.55, FWE=247.8 偏高"),
    ("MYD TIANJIN",    "2026-10-10 20:20:44", "mgo",   414.5, 167.4,  "POB MGO=167.4, FWE MGO=414.5 填成=FWE LSFO"),
]

def num(x):
    try: return float(x)
    except: return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    print("=== FWE 填报异常回退 (%s) ===" % ("APPLY" if args.apply else "DRY-RUN"))

    with open(HISTORY, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    hdr = rows[0]
    ci = {n: hdr.index(n) for n in hdr}

    changes = []
    for vkey, rt, oil, cur, newv, why in FIXES:
        matched = []
        for i, r in enumerate(rows[1:], start=1):
            if len(r) < len(hdr): continue
            if (r[ci["report_time"]] or "")[:19] != rt: continue
            if R.norm(vkey) not in R.norm(r[ci["vessel"]]) and R.norm(r[ci["vessel"]]) not in R.norm(vkey):
                continue
            oldv = num(r[ci[oil]])
            if oldv is None or abs(oldv - cur) > 0.05:
                continue
            matched.append((i, r, oldv))
        if len(matched) != 1:
            print("  [WARN] %s %s %s 匹配 %d 行(期望1), 跳过" % (vkey, rt, oil, len(matched)))
            for i, r, ov in matched[:3]:
                print("      row%d vessel=%s %s=%s" % (i, r[ci["vessel"]], oil, ov))
            continue
        i, r, ov = matched[0]
        print("  [OK] row%d %-16s %s %s: %.2f -> %.2f | %s" % (
            i, r[ci["vessel"]][:16], rt, oil, ov, newv, why))
        changes.append((i, ci[oil], newv, r[ci["vessel"]], rt, oil, ov))

    print("\n可回退 %d / %d 处" % (len(changes), len(FIXES)))
    if not args.apply:
        print("DRY-RUN: 未写盘。确认后加 --apply。")
        return 0

    if len(changes) != len(FIXES):
        print("[FATAL] 有 FIX 未匹配, 为安全起见中止写盘")
        return 2

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak_dir = os.path.join(BASE, "rob_data", "tpl_backup", "fwe_misfill_revert_" + stamp)
    os.makedirs(bak_dir, exist_ok=True)
    shutil.copy2(HISTORY, os.path.join(bak_dir, "rob_history.csv"))
    print("已备份 -> %s" % bak_dir)

    for i, col, newv, *_ in changes:
        rows[i][col] = ("%.3f" % newv).rstrip("0").rstrip(".")
    with open(HISTORY, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    print("已写回 %d 处 -> %s" % (len(changes), HISTORY))
    return 0

if __name__ == "__main__":
    sys.exit(main())
