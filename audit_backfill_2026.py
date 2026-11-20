# -*- coding: utf-8 -*-
"""
以邮箱为准的补抓审计(只读, 不解析附件内容, 不写 history)。

对 vesselreport@culines.com 2026 年全部邮件, 只看 主题/时间/附件名:
  - 判报告类型: 附件名含 NOON REPORT -> NOON, BERTH REPORT -> BERTH
  - 判船: 船名/code token 匹配 主题+附件名, 兜底 sender
  - key = (归一化船名, 秒级 report_time)
对比 rob_history.csv 里已有的 (船, 时间, 类型), 算出:
  - 邮箱里 NOON/BERTH 报告总数
  - history 已有 / history 缺失(可补)
  - 按船、按月分布, 导出可补明细

用法:
  python audit_backfill_2026.py
  AUDIT_CAP=200 python audit_backfill_2026.py    # 冒烟
"""
import sys, os, csv, time, json, argparse
from collections import defaultdict, Counter
from datetime import datetime
import win32com.client as win32
import rob_refresh as R

HISTORY = os.path.join("rob_data", "rob_history.csv")
RESULTS = os.path.join("rob_data", "rob_results.json")
SENDERS = os.path.join("rob_data", "vessel_senders.json")
OUT = os.path.join("rob_data", "backfill_2026_fillable.csv")
MAX_DEPTH = 4
CAP = int(os.environ.get("AUDIT_CAP", "0"))


def log(m):
    print(m, flush=True)


def norm(s):
    return R.norm(s or "")


def kind_of_att(attnames, subject):
    """返回 NOON / BERTH / None。以附件名为准(最可靠), 主题兜底。"""
    blob = " ".join(attnames).upper()
    sub = (subject or "").upper()
    if "NOON" in blob or "NOON REPORT" in sub:
        return "NOON"
    if "BERTH" in blob or "BERTH" in sub or "BERTHING" in blob:
        return "BERTH"
    return None


def walk(f, out, d=0):
    if d > MAX_DEPTH:
        return
    try:
        for c in f.Folders:
            out.append(c)
            walk(c, out, d + 1)
    except Exception:
        pass
    return out


def load_history():
    have = defaultdict(set)   # type -> set((vn, rt))
    for r in csv.DictReader(open(HISTORY, encoding="utf-8-sig")):
        t = (r.get("report_type") or "").strip().upper()
        if t not in ("NOON", "BERTH"):
            continue
        rt = (r.get("report_time") or "")[:19]
        if not rt.startswith("2026"):
            continue
        have[t].add((norm(r.get("vessel")), rt))
    return have


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--types", default="NOON,BERTH")
    args = ap.parse_args()
    types = set(t.strip().upper() for t in args.types.split(",") if t.strip())

    have = load_history()
    results = json.load(open(RESULTS, encoding="utf-8"))
    toks = []          # token -> canonical vessel name
    for r in results:
        v = r.get("vessel") or ""
        if not v:
            continue
        nv = norm(v)
        if len(nv) >= 4:
            toks.append((nv, v))
        c = norm(r.get("code") or "")
        if len(c) >= 5:
            toks.append((c, v))
    toks.sort(key=lambda x: -len(x[0]))

    # sender map: history sender + vessel_senders.json
    smap = {}
    for r in results:
        s = (r.get("sender") or "").strip().lower()
        if s and "@" in s:
            smap[s] = r.get("vessel") or ""
    if os.path.exists(SENDERS):
        for k, v in json.load(open(SENDERS, encoding="utf-8")).items():
            if isinstance(v, str) and "@" in v:
                smap.setdefault(v.strip().lower(), k)
    log("fleet tokens=%d senders=%d | history have NOON=%d BERTH=%d"
        % (len(toks), len(smap), len(have["NOON"]), len(have["BERTH"])))

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    st = None
    for i in range(1, ns.Stores.Count + 1):
        s = ns.Stores(i)
        if "vesselreport" in (s.DisplayName or "").lower():
            st = s
            break
    if st is None:
        log("[FATAL] vesselreport store not found")
        return 2
    root = st.GetRootFolder()
    folders = [root] + walk(root, [])
    log("store=%s folders=%d" % (st.DisplayName, len(folders)))

    filt = "[ReceivedTime] >= '01/01/2026 12:00 AM'"
    mail = defaultdict(set)     # type -> set((vn, rt)) 邮箱里存在的
    no_vessel = Counter()
    n = 0
    t0 = time.time()
    for fi, f in enumerate(folders):
        try:
            items = f.Items
            items.Sort("[ReceivedTime]", True)
            sub = items.Restrict(filt)
        except Exception:
            continue
        for it in sub:
            n += 1
            if CAP and n > CAP:
                break
            try:
                t = it.ReceivedTime
                rt = t.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                continue
            try:
                subj = str(it.Subject or "")
            except Exception:
                subj = ""
            att = []
            try:
                for a in it.Attachments:
                    try:
                        att.append(a.FileName)
                    except Exception:
                        pass
            except Exception:
                pass
            kind = kind_of_att(att, subj)
            if kind not in types:
                continue
            blob = norm(subj) + " " + " ".join(norm(a) for a in att)
            vname = None
            for tok, vn in toks:
                if tok in blob:
                    vname = vn
                    break
            if not vname:
                try:
                    sa = (R.get_sender(it) or "").strip().lower()
                    vname = smap.get(sa)
                except Exception:
                    vname = None
            if not vname:
                no_vessel[kind] += 1
                continue
            mail[kind].add((norm(vname), rt))
        if fi % 20 == 0 and fi:
            log("  [%d/%d] scanned=%d mailNOON=%d mailBERTH=%d (%.0fs)"
                % (fi, len(folders), n, len(mail["NOON"]), len(mail["BERTH"]), time.time() - t0))
        try:
            del sub
        except Exception:
            pass

    log("")
    log("=== 结果 (用时 %.0fs, 扫 %d 封) ===" % (time.time() - t0, n))
    fillable = {}
    for t in sorted(types):
        m, h = mail[t], have[t]
        miss = m - h
        fillable[t] = miss
        log("  %-6s 邮箱 %5d | history已有 %5d | 可补(邮箱有history无) %5d"
            % (t, len(m), len(h), len(miss)))
        log("         判不出船(跳过): %d" % no_vessel[t])

    # 按船、按月分布(可补)
    allmiss = []
    for t, s in fillable.items():
        for vn, rt in s:
            allmiss.append((vn, rt[:10], t))
    bymonth = Counter((rt[:7], t) for vn, rt, t in allmiss)
    byvessel = Counter((vn, t) for vn, rt, t in allmiss)
    log("")
    log("=== 可补按月 ===")
    for m in sorted(set(x[0] for x in bymonth)):
        log("  %s  NOON %4d  BERTH %4d" % (m, bymonth[(m, "NOON")], bymonth[(m, "BERTH")]))
    log("")
    log("=== 可补按船 TOP30 ===")
    for (vn, t), c in byvessel.most_common(30):
        log("  %-24s %-6s %d" % (vn[:24], t, c))

    with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["vessel_key", "report_time", "report_type"])
        for vn, rt, t in sorted(allmiss):
            w.writerow([vn, rt, t])
    log("")
    log("可补明细已导出: %s (%d 条)" % (OUT, len(allmiss)))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
