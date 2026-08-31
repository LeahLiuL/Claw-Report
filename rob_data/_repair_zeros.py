# -*- coding: utf-8 -*-
"""修复历史上被 BERTH 报告解析 bug 清零的 ROB 行(一次性, 确认后可删)。

背景: extract_rob 向右找数值时不会在单位单元格('MT')停下, 一行多组的 BERTH 报告里
空白处会误抓到下一组的 0(如 'ANCHOR AWEIGH ROB LSFO | (空) | MT | ... HSFO | 0'),
导致 LSFO/MGO 被写成 0, Trend 趋势图出现假的断崖。解析逻辑已修, 本脚本回填历史。

做法:
  1. 找出可疑行 —— 某船绝大多数行都有油量, 但个别行 lsfo/mgo 为 0(该船零值占比 <50%)
  2. 按 ReceivedTime 精确定位原邮件(全文件夹 Restrict 一次, 建 {收件时间: 邮件} 索引)
  3. 用修好的 extract_rob 重新解析, 就地覆盖 CSV 的 lsfo/hsfo/mgo/ulsfo/bw/fw
  4. 原文件先备份为 rob_history.csv.bak_repair
"""
import csv, os, sys
sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
import rob_refresh as R

CSV = os.path.join(BASE, "rob_data", "rob_history.csv")
FIELDS = ["date", "vessel", "code", "lane", "pic",
          "lsfo", "hsfo", "mgo", "ulsfo", "bw", "fw", "refeer",
          "found", "report_time"]
OIL_COLS = {"LSFO": 5, "HSFO": 6, "MGO": 7, "ULSFO": 8, "BW": 9, "FW": 10}


def zero(v):
    return (v or "").strip() in ("", "0", "0.0", "None")


def main(apply=False):
    rows = list(csv.reader(open(CSV, encoding="utf-8", newline="")))
    header, data = rows[0], [r for r in rows[1:] if len(r) >= 14]
    print("rows:", len(data))

    # 1) 每船零值占比
    stat = {}
    for r in data:
        if r[12] != "1":
            continue
        s = stat.setdefault(r[1], {"n": 0, "lsfo0": 0, "mgo0": 0})
        s["n"] += 1
        if zero(r[5]):
            s["lsfo0"] += 1
        if zero(r[7]):
            s["mgo0"] += 1
    targets = {}   # report_time -> [row idx]
    for i, r in enumerate(data):
        if r[12] != "1" or not r[13]:
            continue
        s = stat.get(r[1])
        if not s or s["n"] < 3:
            continue
        bad = []
        if zero(r[5]) and s["lsfo0"] / s["n"] < 0.5:
            bad.append("LSFO")
        if zero(r[7]) and s["mgo0"] / s["n"] < 0.5:
            bad.append("MGO")
        if bad:
            targets.setdefault(r[13][:19], []).append((i, bad, r[1]))
    print("suspicious rows: %d (vessels: %d)"
          % (sum(len(v) for v in targets.values()),
             len({x[2] for v in targets.values() for x in v})))
    if not targets:
        return

    # 2) 全文件夹一次 Restrict, 建 {收件时间: 邮件}
    store = R.connect_outlook()
    inbox = store.GetDefaultFolder(6)
    folders, seen = [inbox], {id(inbox)}

    def walk(f):
        try:
            for c in f.Folders:
                if id(c) in seen:
                    continue
                seen.add(id(c))
                folders.append(c)
                walk(c)
        except Exception:
            pass
    walk(inbox)
    times = sorted(targets)
    from datetime import datetime, timedelta
    lo = (datetime.strptime(times[0][:10], "%Y-%m-%d") - timedelta(days=1)) \
        .strftime("%m/%d/%Y %H:%M %p")
    hi = (datetime.strptime(times[-1][:10], "%Y-%m-%d") + timedelta(days=1)) \
        .strftime("%m/%d/%Y %H:%M %p")
    filt = "[ReceivedTime] >= '%s' AND [ReceivedTime] <= '%s'" % (lo, hi)
    print("folders: %d, window %s ~ %s" % (len(folders), lo, hi))
    found = {}
    for f in folders:
        try:
            items = f.Items.Restrict(filt)
        except Exception:
            continue
        for it in items:
            try:
                t = it.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                continue
            if t in targets and t not in found:
                found[t] = (it, f.Name)
    print("located mails: %d / %d" % (len(found), len(targets)))

    # 3) 重新解析并覆盖
    fixed = missing = 0
    for t, lst in targets.items():
        if t not in found:
            missing += len(lst)
            continue
        it, fname = found[t]
        att = R.pick_report_attachment(it)
        if att is None:
            missing += len(lst)
            continue
        rob = R.extract_rob(att)
        if not any(k in rob for k in R.OIL_KEYS):
            missing += len(lst)
            continue
        for i, bad, vessel in lst:
            old = [data[i][5], data[i][7]]
            for k, col in OIL_COLS.items():
                if k in rob and rob[k] is not None:
                    data[i][col] = str(rob[k])
            print("  [FIX] %-24s %s  LSFO %s -> %s | MGO %s -> %s   (%s)"
                  % (vessel[:24], t, old[0], data[i][5], old[1], data[i][7], fname))
            fixed += 1
    print("\nfixed rows: %d, still missing: %d" % (fixed, missing))

    if apply and fixed:
        import shutil
        bak = CSV + ".bak_repair"
        shutil.copy2(CSV, bak)
        tmp = CSV + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(FIELDS)
            w.writerows(data)
        os.replace(tmp, CSV)
        print("APPLIED -> backup:", bak)
    elif not apply:
        print("(dry run, add --apply to write)")


if __name__ == "__main__":
    main(apply="--apply" in sys.argv)
