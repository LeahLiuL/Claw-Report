# -*- coding: utf-8 -*-
"""一次性清理 rob_history.csv 脏数据(确认无误后可删除本脚本)。

脏数据来源: 两台机器用不同去重键并发写同一 CSV, 同一份报告被写两次;
第二次写入时油种被清零(lsfo=0/空, bw/fw/refeer 空), 污染 Trend 趋势图。

清洗规则(只删坏行, 保留好行, 不丢历史):
 1. 按 (vessel, report_time) 分组去重: 同一份报告保留"字段最全"(非空列最多)的那行,
    同分保留先出现的那行。
 2. 删除 found=1 但四个油种(lsfo/hsfo/mgo/ulsfo)全为空/0 的行 —— 抓到了却没数,
    属于写入中断的残行。
 3. 表头与列数保持 14 列不变。
 4. MISS 行(report_time 为空): 按 (date, vessel) 去重, 每天每船保留一条。
"""
import csv, os, sys, shutil
sys.stdout.reconfigure(encoding="utf-8")

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV_IN = os.path.join(BASE, "rob_data", "rob_history.csv")
FIELDS = ["date", "vessel", "code", "lane", "pic",
          "lsfo", "hsfo", "mgo", "ulsfo", "bw", "fw", "refeer",
          "found", "report_time"]
OIL_IDX = [5, 6, 7, 8]


def has_oil(row):
    for i in OIL_IDX:
        v = (row[i] or "").strip()
        if v and v not in ("0", "0.0", "None"):
            return True
    return False


def score(row):
    return sum(1 for c in row if (c or "").strip() not in ("", "None"))


def main(apply=False):
    with open(CSV_IN, encoding="utf-8", newline="") as f:
        rd = csv.reader(f)
        header = next(rd)
        rows = [r[:14] for r in rd if len(r) >= 14]
    print("header:", header)
    print("total data rows:", len(rows))

    dropped_empty, dropped_dup, dropped_miss = 0, 0, 0
    groups, miss_groups = {}, {}
    for r in rows:
        rt = (r[13] or "").strip()[:19]
        if rt:
            # 规则 2: 标称抓到但四个油种全空 -> 残行, 直接丢
            if (r[12] or "").strip() == "1" and not has_oil(r):
                dropped_empty += 1
                print("  [DROP empty] %-24s %s  %s" % (r[1][:24], rt, r[5:12]))
                continue
            groups.setdefault((r[1], rt), []).append(r)
        else:
            k = (r[0], r[1])
            if k in miss_groups:
                dropped_miss += 1
                continue
            miss_groups[k] = r

    keep = []
    for key, items in groups.items():
        if len(items) > 1:
            # 按 (有油种, 字段完整度) 排序, 取最好的一条
            items_sorted = sorted(items, key=lambda x: (has_oil(x), score(x)), reverse=True)
            dropped_dup += len(items) - 1
            print("  [DEDUP] %-24s %s  x%d -> keep %s"
                  % (key[0][:24], key[1], len(items), items_sorted[0][5:9]))
            keep.append(items_sorted[0])
        else:
            keep.append(items[0])
    keep.extend(miss_groups.values())

    # 保持原有时间顺序
    idx = {id(r): i for i, r in enumerate(rows)}
    keep.sort(key=lambda r: (idx.get(id(r), 0)))

    print("\n-- summary --")
    print("dropped: empty-oil=%d, dup=%d, dup-miss=%d" % (dropped_empty, dropped_dup, dropped_miss))
    print("keep: %d rows (from %d)" % (len(keep), len(rows)))

    if apply:
        bak = CSV_IN + ".bak_preclean"
        shutil.copy2(CSV_IN, bak)
        tmp = CSV_IN + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(FIELDS)
            w.writerows(keep)
        os.replace(tmp, CSV_IN)
        print("\nAPPLIED. backup ->", bak)
        print("final rows:", len(keep))
    else:
        print("\n(dry run, use --apply to write)")


if __name__ == "__main__":
    main(apply="--apply" in sys.argv)
