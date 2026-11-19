# -*- coding: utf-8 -*-
"""
报告覆盖率三级核对(只读, 不写盘, 不解析 xlsx 附件)。

对rob_history.csv里每条 BERTH/NOON 记录, 到 vesselreport@culines.com 邮箱
全部文件夹找对应邮件, 分三级:
  1) 精确命中: 秒级 report_time + 船名 都对上 -> 邮件在, 时间戳没变
  2) 模糊命中: 秒级对不上, 但同船 + 同日有邮件 -> 邮件在, 只是被 RPA 改写/转发改了时间
  3) 真缺失:   同船同日都找不到邮件 -> 疑似真没发/没进这个邮箱

只判断邮件定位, 不打开附件, 几分钟可完成。

用法:
  python report_coverage.py --types BERTH,NOON
  python report_coverage.py --types BERTH,NOON --since 2026-08-01
  COVERAGE_SCAN_CAP=200 python report_coverage.py   # 冒烟测试
"""
import sys, os, csv, time, argparse
from collections import defaultdict, Counter
from datetime import datetime, timedelta
import win32com.client as win32
import rob_refresh as R

HISTORY = os.path.join("rob_data", "rob_history.csv")
MAX_DEPTH = 4
SCAN_CAP = int(os.environ.get("COVERAGE_SCAN_CAP", "0"))


def log(m):
    print(m, flush=True)


def norm(s):
    return R.norm(s or "")


def load_targets(types, since=""):
    with open(HISTORY, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    idx, bytime, byday = {}, defaultdict(list), defaultdict(list)
    for i, r in enumerate(rows):
        typ = (r.get("report_type") or "").strip().upper()
        if typ not in types:
            continue
        rt = (r.get("report_time") or "")[:19]
        if not rt:
            continue
        if since and rt[:10] < since:
            continue
        v = r.get("vessel") or ""
        vn = norm(v)
        key = (vn, rt, typ, i)
        idx[key] = {"vessel": v, "rt": rt, "type": typ, "row": i, "vn": vn}
        bytime[rt].append(key)
        byday[(vn, rt[:10])].append(key)   # 船名+日期 -> keys(模糊命中用)
    return idx, bytime, byday


def pick_key(rt, subject, attnames, bytime, seen):
    """秒级时间命中后, 单候选直接用, 多候选靠船名 token 判船。"""
    keys = bytime.get(rt, [])
    if not keys:
        return None
    fresh = [k for k in keys if k not in seen]
    if not fresh:
        return None
    if len(fresh) == 1:
        return fresh[0]
    blob = norm(subject) + " " + " ".join(norm(a) for a in attnames)
    for k in fresh:
        if k[0] and k[0] in blob:
            return k
    return None


def walk_folders(f, depth=0, out=None):
    if out is None:
        out = []
    if depth > MAX_DEPTH:
        return out
    try:
        for c in f.Folders:
            out.append(c)
            walk_folders(c, depth + 1, out)
    except Exception:
        pass
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--types", default="BERTH,NOON")
    ap.add_argument("--since", default="")
    args = ap.parse_args()
    types = set(t.strip().upper() for t in args.types.split(",") if t.strip())

    idx, bytime, byday = load_targets(types, args.since)
    total = Counter(v["type"] for v in idx.values())
    log("=== 报告覆盖率三级核对(只读, 不解析附件) ===")
    for t in sorted(types):
        log("  %-8s 目标 %d 行" % (t, total[t]))
    if not bytime:
        log("无目标, 退出")
        return 0

    # 邮件里可能出现的船名 token -> 该船在 idx 里的所有 vn (同一船多种写法归一)
    all_vn = set(k[0] for k in idx if k[0])
    times = sorted(bytime)
    lo = datetime.strptime(times[0], "%Y-%m-%d %H:%M:%S") - timedelta(days=1)
    hi = datetime.strptime(times[-1], "%Y-%m-%d %H:%M:%S") + timedelta(days=1)
    filt = "[ReceivedTime] >= '%s' AND [ReceivedTime] <= '%s'" % (
        lo.strftime("%m/%d/%Y %H:%M %p"), hi.strftime("%m/%d/%Y %H:%M %p"))
    log("时间过滤: %s ~ %s" % (times[0][:10], times[-1][:10]))

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    target = None
    for si in range(1, ns.Stores.Count + 1):
        st = ns.Stores(si)
        if "vesselreport" in (st.DisplayName or "").lower():
            target = st
            break
    if target is None:
        log("[FATAL] 找不到 vesselreport@culines.com store")
        return 2
    log("[store] %s" % target.DisplayName)
    root = target.GetRootFolder()
    folders = [root] + walk_folders(root)
    log("待扫文件夹: %d 个" % len(folders))

    seen_exact, seen_fuzzy = set(), set()
    t0 = time.time()
    n_scanned = 0
    for fi, fld in enumerate(folders):
        try:
            items = fld.Items
            items.Sort("[ReceivedTime]", True)
            sub = items.Restrict(filt) if filt else items
        except Exception:
            continue
        ne = nf = n = 0
        for it in sub:
            n += 1
            n_scanned += 1
            if SCAN_CAP and n > SCAN_CAP:
                break
            try:
                t = it.ReceivedTime
                rt = t.strftime("%Y-%m-%d %H:%M:%S") if hasattr(t, "strftime") else ""
            except Exception:
                continue
            if not rt:
                continue
            try:
                subj = str(it.Subject or "")
            except Exception:
                subj = ""
            attnames = []
            try:
                for a in it.Attachments:
                    try:
                        attnames.append(a.FileName)
                    except Exception:
                        pass
            except Exception:
                pass
            blob = norm(subj) + " " + " ".join(norm(a) for a in attnames)
            # 1) 精确
            if rt in bytime:
                k = pick_key(rt, subj, attnames, bytime, seen_exact)
                if k:
                    seen_exact.add(k)
                    seen_fuzzy.add(k)
                    ne += 1
            # 2) 模糊: 同船 + 同日 (秒级对不上时)
            day = rt[:10]
            for vn in all_vn:
                if len(vn) >= 4 and vn in blob:
                    for k in byday.get((vn, day), []):
                        if k not in seen_fuzzy:
                            seen_fuzzy.add(k)
                            nf += 1
        if ne or nf:
            log("  [%d/%d] %s 精确+%d 模糊+%d" % (fi + 1, len(folders), fld.Name, ne, nf))

    exact = Counter(idx[k]["type"] for k in seen_exact if k in idx)
    fuzzy = Counter(idx[k]["type"] for k in seen_fuzzy if k in idx)
    log("")
    log("=== 结果(用时 %.0fs) ===" % (time.time() - t0))
    for t in sorted(types):
        tot = total[t]
        e, f = exact[t], fuzzy[t]
        miss = tot - f
        log("  %-8s 总数 %d" % (t, tot))
        log("    精确命中 %5d (%.1f%%)  秒级时间+船名都对上" % (e, 100.0 * e / tot if tot else 0))
        log("    模糊命中 %5d (%.1f%%)  同船同日有(时间被改/转发)" % (f - e, 100.0 * (f - e) / tot if tot else 0))
        log("    真缺失   %5d (%.1f%%)  同船同日都找不到" % (miss, 100.0 * miss / tot if tot else 0))
    for t in sorted(types):
        missed = [v for k, v in idx.items() if v["type"] == t and k not in seen_fuzzy]
        missed.sort(key=lambda v: v["rt"], reverse=True)
        log("  真缺失样本 %s (最近12条):" % t)
        for v in missed[:12]:
            log("    %s %s" % (v["vessel"], v["rt"]))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
