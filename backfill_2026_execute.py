# -*- coding: utf-8 -*-
"""
执行 2026 全量补抓: 读 audit 导出的真缺口清单 backfill_2026_real.csv
(船+日+类型), 到 vesselreport 邮箱定位对应邮件, 解析 xlsx 附件 ROB, 幂等追加 rob_history.csv。

定位维度用 (船, 日期, 报告类型) 而非秒级时间 —— 因为 RPA 搬运会改时间戳, 秒级对不上;
同一天同船同类型有多封(更正件)时取最新一封。写入 report_time 用邮件真实秒级时间。

安全护栏:
  - 默认 dry-run, 不显式 --apply 绝不写盘
  - 写前自动备份 rob_history.csv 到 rob_data/tpl_backup/backfill_2026_<时间戳>/
  - append 只新增 (vessel, report_time) 不存在的行, 不改/删已有行
  - 只锁 vesselreport store, Restrict 时间过滤, for it in sub 迭代(非索引)
  - 写盘时再按 (船, 日期, 类型) 二次校验, 防止 history 里已有同船同日

用法:
  python backfill_2026_execute.py                      # dry-run
  python backfill_2026_execute.py --apply --limit 20   # 小批验证
  python backfill_2026_execute.py --apply              # 全量
"""
import sys, os, csv, time, json, shutil, argparse
from collections import defaultdict, Counter
from datetime import datetime, timedelta
import win32com.client as win32
import pythoncom
import rob_refresh as R

ROB_DIR = "rob_data"
HISTORY = os.path.join(ROB_DIR, "rob_history.csv")
FILLABLE = os.path.join(ROB_DIR, "backfill_2026_real.csv")
RESULTS = os.path.join(ROB_DIR, "rob_results.json")
MAX_DEPTH = 4
SNAP_FIELDS = ["date", "vessel", "code", "lane", "pic", "lsfo", "hsfo", "mgo",
               "ulsfo", "bw", "fw", "refeer", "found", "report_time", "speed", "report_type"]


def log(m):
    print(m, flush=True)


def norm(s):
    return R.norm(s or "")


def kind_of_att(attnames, subject):
    blob = " ".join(attnames).upper()
    sub = (subject or "").upper()
    if "NOON" in blob or "NOON REPORT" in sub:
        return "NOON"
    if "BERTH" in blob or "BERTHING" in blob or "BERTH" in sub:
        return "BERTH"
    return None


def walk(f, out, d=0):
    if d > MAX_DEPTH:
        return out
    try:
        for c in f.Folders:
            out.append(c)
            walk(c, out, d + 1)
    except Exception:
        pass
    return out


def load_need():
    """(vn, date, type) -> 出现次数。清单已是日期级真缺口。"""
    need = Counter()
    with open(FILLABLE, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            key = (norm(r["vessel_key"]), (r["report_time"] or "")[:10],
                   (r["report_type"] or "").strip().upper())
            need[key] += 1
    return need


def build_maps():
    results = json.load(open(RESULTS, encoding="utf-8"))
    toks, meta = [], {}
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
        meta[v] = {"code": r.get("code"), "lane": r.get("lane"), "pic": r.get("pic")}
    toks.sort(key=lambda x: -len(x[0]))
    smap = {}
    for r in results:
        s = (r.get("sender") or "").strip().lower()
        if s and "@" in s:
            smap[s] = r.get("vessel") or ""
    sp = os.path.join(ROB_DIR, "vessel_senders.json")
    if os.path.exists(sp):
        for k, v in json.load(open(sp, encoding="utf-8")).items():
            if isinstance(v, str) and "@" in v:
                smap.setdefault(v.strip().lower(), k)
    return toks, smap, meta


def history_day_keys():
    """history 里已有的 (船, 日, 类型) 集合, 用于写盘二次校验。"""
    s = set()
    for r in csv.DictReader(open(HISTORY, encoding="utf-8-sig")):
        t = (r.get("report_type") or "").strip().upper()
        if t not in ("NOON", "BERTH"):
            continue
        rt = (r.get("report_time") or "")[:10]
        if rt:
            s.add((norm(r.get("vessel")), rt, t))
    return s


def existing_time_keys():
    ex = set()
    for r in csv.DictReader(open(HISTORY, encoding="utf-8-sig")):
        rt = (r.get("report_time") or "")[:19]
        if rt:
            ex.add((r.get("vessel"), rt))
    return ex


def backup_history():
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(ROB_DIR, "tpl_backup", "backfill_2026_" + ts)
    os.makedirs(dst, exist_ok=True)
    shutil.copy2(HISTORY, os.path.join(dst, "rob_history.csv"))
    return dst


def new_row(vessel, meta, rob, recv, report_type):
    return {
        "date": recv[:10], "vessel": vessel,
        "code": meta.get(vessel, {}).get("code") or "",
        "lane": meta.get(vessel, {}).get("lane") or "",
        "pic": meta.get(vessel, {}).get("pic") or "",
        "lsfo": rob.get("LSFO"), "hsfo": rob.get("HSFO"),
        "mgo": rob.get("MGO"), "ulsfo": rob.get("ULSFO"),
        "bw": rob.get("BW"), "fw": rob.get("FW"),
        "refeer": rob.get("REFEER"), "found": "1",
        "report_time": recv, "speed": rob.get("SPEED"),
        "report_type": report_type,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    need = load_need()
    toks, smap, meta = build_maps()
    hist_days = history_day_keys()
    ex_time = existing_time_keys()
    log("need(船+日+类型)=%d | history已有天数key=%d" % (len(need), len(hist_days)))

    dates = sorted(set(k[1] for k in need))
    lo = datetime.strptime(dates[0], "%Y-%m-%d")
    hi = datetime.strptime(dates[-1], "%Y-%m-%d")
    filt = "[ReceivedTime] >= '%s' AND [ReceivedTime] <= '%s'" % (
        (lo - timedelta(days=1)).strftime("%m/%d/%Y %H:%M %p"),
        (hi + timedelta(days=1)).strftime("%m/%d/%Y 11:59 PM"))
    log("time filter: %s ~ %s" % (dates[0], dates[-1]))

    pythoncom.CoInitialize()
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
    log("folders=%d" % len(folders))

    located = {}          # (vn,date,type) -> (canonical, recv秒, rob, kind)
    nmail = 0
    t0 = time.time()
    for fi, f in enumerate(folders):
        try:
            items = f.Items
            items.Sort("[ReceivedTime]", True)
            sub = items.Restrict(filt)
        except Exception:
            continue
        for it in sub:
            nmail += 1
            try:
                t = it.ReceivedTime
                rt = t.strftime("%Y-%m-%d %H:%M:%S")
                day = rt[:10]
            except Exception:
                continue
            try:
                subj = str(it.Subject or "")
            except Exception:
                subj = ""
            att_names = []
            try:
                for a in it.Attachments:
                    try:
                        att_names.append(a.FileName)
                    except Exception:
                        pass
            except Exception:
                pass
            kind = kind_of_att(att_names, subj)
            if kind is None:
                continue
            # 该邮件的 (船,日,类型) 是否在需求里
            blob = norm(subj) + " " + " ".join(norm(a) for a in att_names)
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
                continue
            key = (norm(vname), day, kind)
            if key not in need:
                continue
            if key in located:
                continue      # 已定位(邮件按时间倒序, 第一封=最新)
            # 解析附件
            try:
                att = R.pick_report_attachment(it)
                if att is None:
                    continue
                rob = R.extract_rob(att)
            except Exception:
                continue
            if not any(k in rob for k in R.OIL_KEYS):
                continue
            located[key] = (vname, rt, rob, kind)
            if len(located) % 50 == 0:
                log("  progress: located=%d/%d (%.0fs)" % (len(located), len(need), time.time() - t0))
            if args.limit and len(located) >= args.limit:
                break
        if args.limit and len(located) >= args.limit:
            break
        try:
            del sub
        except Exception:
            pass

    # 组装待写行 + 二次校验
    rows = []
    skipped_day = 0
    for key, (vname, rt, rob, kind) in located.items():
        if key in hist_days:
            skipped_day += 1
            continue
        if (vname, rt) in ex_time:
            skipped_day += 1
            continue
        rk = (rob.get("REPORT_KIND") or "").strip().upper() or kind
        if rk not in ("NOON", "BERTH"):
            continue      # 用户只要 NOON/BERTH, 排除 DRIFT/ANCHOR/SAILING
        rows.append(new_row(vname, meta, rob, rt, rk))
        ex_time.add((vname, rt))
        hist_days.add(key)

    log("")
    log("=== 结果 (用时 %.0fs, 扫 %d 封) ===" % (time.time() - t0, nmail))
    log("  需求=%d | 已定位=%d | 未定位=%d"
        % (len(need), len(located), len(need) - len(located)))
    log("  二次校验跳过=%d | 可写新增行=%d" % (skipped_day, len(rows)))
    log("  新增按类型: %s" % dict(Counter(r["report_type"] for r in rows)))
    if rows:
        log("  样本(前8):")
        for r in sorted(rows, key=lambda x: x["report_time"])[:8]:
            log("    %s %s %s lsfo=%s mgo=%s"
                % (r["vessel"], r["report_time"], r["report_type"], r["lsfo"], r["mgo"]))

    if not args.apply:
        log("(dry-run) add --apply to write. 数据未改动。")
        return 0
    if not rows:
        log("无新增行, 不写盘。")
        return 0

    bak = backup_history()
    log("备份: %s" % bak)
    write_header = not os.path.exists(HISTORY) or os.path.getsize(HISTORY) == 0
    with open(HISTORY, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=SNAP_FIELDS, extrasaction="ignore")
        if write_header:
            w.writeheader()
        for r in rows:
            w.writerow(r)
    log("已追加 %d 行到 %s" % (len(rows), HISTORY))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
