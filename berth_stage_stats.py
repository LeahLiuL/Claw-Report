# -*- coding: utf-8 -*-
r"""只读统计: 历史 BERTH 报告 ROB 最终采纳值来自哪个阶段(POB/FWE/ANCHOR/DROP/普通)。

背景: 用户口径 = "BERTH 优先取 FWE ROB, 没有 FWE 才取 POB";
      当前代码 = POB(50) > FWE(40) > ANCHOR(30) > DROP(20), 相反。
      extract_rob 内部算 rob_prio 但未写入 CSV, 故需重解析邮件统计。

输出(按主油种 LSFO>ULSFO>HSFO>MGO):
  1. 当前规则(POB优先)下采纳值的阶段分布 -> 回答"多少取的是 POB"
  2. 每封是否同时出现 POB 与 FWE 值 -> 评估改成 FWE 优先的影响面

复用 backfill_berth_rob.py 的扫描框架(已验证定位正确、不卡死)。默认只统计不写盘。
"""
import sys, os, csv, time, argparse, tempfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import win32com.client as win32
import rob_refresh as R

HISTORY_CSV = os.path.join(BASE, "rob_data", "rob_history.csv")
SCAN_CAP = int(os.environ.get("BERTH_STAGE_SCAN_CAP", "0"))
MAX_DEPTH = 4
PRIO_NAME = {50: "POB", 40: "FWE", 30: "ANCHOR AWEIGH", 20: "DROP ANCHOR", 0: "(普通ROB)"}
OIL_CODES = ("LSFO", "HSFO", "ULSFO", "MGO", "BW", "FW", "REFEER", "REEFER")
STAGE_PRIO = {"POB": 50, "FWE": 40, "ANCHOR AWEIGH": 30, "DROP ANCHOR": 20}

import re as _re


def log(msg):
    print(msg, flush=True)


def _as_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if not s or not _re.match(r"^-?\d+(\.\d+)?$", s):
        return None
    return float(s)


def extract_rob_stage(att):
    """返回 {kind, chosen:{oil:prio}, seen:{oil:set(prio)}}。逻辑与 R.extract_rob 一致。"""
    import openpyxl
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        att.SaveAsFile(p)
        wb = openpyxl.load_workbook(p, data_only=True)
    except Exception:
        try:
            os.remove(p)
        except Exception:
            pass
        return None

    def _stage_of(cu, pos):
        prefix = cu[:pos].upper()
        for name in sorted(STAGE_PRIO, key=lambda x: -len(x)):
            if name in prefix:
                return STAGE_PRIO[name]
        return 0

    def _take(cu, i, row):
        if "ROB" not in cu:
            return None
        pos = cu.rfind("ROB")
        tail = cu[pos + 3:].strip()
        if not tail:
            return None
        tok = tail.split()[0].strip(" :|-")
        if tok not in OIL_CODES:
            return None
        prio = _stage_of(cu, pos)
        for j in range(i + 1, len(row)):
            v = row[j]
            if v is None:
                continue
            if isinstance(v, str) and not v.strip():
                continue
            n = _as_num(v)
            if n is None:
                break
            return (tok, round(n, 3), prio)
        return None

    chosen = {}                      # oil -> 采纳阶段的 prio
    seen = defaultdict(set)          # oil -> 出现过的所有阶段 prio
    full_text = []
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for i, c in enumerate(row):
                    if c and isinstance(c, str):
                        cu = c.upper().strip()
                        full_text.append(cu)
                        got = _take(cu, i, row)
                        if got:
                            k, v, prio = got
                            key = "REFEER" if k == "REEFER" else k
                            seen[key].add(prio)
                            if prio >= chosen.get(key, -1):
                                chosen[key] = prio
    finally:
        try:
            os.remove(p)
        except Exception:
            pass
    kind = R.detect_report_kind(" ".join(full_text))
    if "ULSFO" in chosen and "LSFO" not in chosen:
        chosen["LSFO"] = chosen["ULSFO"]
        seen["LSFO"] = seen["ULSFO"]
    return {"kind": (kind or "").strip().upper(), "chosen": chosen, "seen": dict(seen)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--vessel", default="", help="只统计某船, 如 CUL XIAMEN")
    ap.add_argument("--limit", type=int, default=0, help="最多处理 N 封(小样本验证)")
    args = ap.parse_args()

    log("=== BERTH ROB 阶段统计 (只读) ===")
    # 目标 BERTH 行(秒级时间索引)
    with open(HISTORY_CSV, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    hdr = rows[0]
    ci = {n: hdr.index(n) for n in hdr}
    bytime = defaultdict(list)
    n_target = 0
    for r in rows[1:]:
        if len(r) < len(hdr):
            continue
        rt = (r[ci["report_time"]] or "")[:19]
        if (r[ci["report_type"]] or "").strip().upper() != "BERTH" or not rt:
            continue
        vessel = r[ci["vessel"]]
        if args.vessel:
            k = args.vessel.upper()
            if k not in vessel.upper() and k != (r[ci["code"]] or "").upper():
                continue
        bytime[rt].append((R.norm(vessel), vessel, rt))
        n_target += 1
    log("目标 BERTH 行: %d" % n_target)

    times = sorted(bytime)
    lo = datetime.strptime(times[0], "%Y-%m-%d %H:%M:%S") - timedelta(days=1)
    hi = datetime.strptime(times[-1], "%Y-%m-%d %H:%M:%S") + timedelta(days=1)
    filt = "[ReceivedTime] >= '%s' AND [ReceivedTime] <= '%s'" % (
        lo.strftime("%m/%d/%Y %H:%M %p"), hi.strftime("%m/%d/%Y %H:%M %p"))

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    target = None
    for si in range(1, ns.Stores.Count + 1):
        st = ns.Stores(si)
        if "vesselreport" in (st.DisplayName or "").lower():
            target = st
            break
    if target is None:
        log("[FATAL] 找不到 vesselreport store")
        return 2
    log("[store] %s" % target.DisplayName)

    root = target.GetRootFolder()
    folders = [root]
    stack = [(root, 0)]
    while stack:
        f, d = stack.pop()
        if d >= MAX_DEPTH:
            continue
        try:
            for c in f.Folders:
                folders.append(c)
                stack.append((c, d + 1))
        except Exception:
            pass
    log("待扫文件夹: %d" % len(folders))

    seen_keys = set()   # 已统计的目标 key(norm vessel, rt)
    stage_cnt = Counter()          # 主油种采纳阶段
    both_pob_fwe = 0               # 主油种同时出现 POB 和 FWE
    only_pob = 0                   # 主油种只有 POB 没 FWE
    only_fwe = 0                   # 主油种只有 FWE 没 POB
    fwe_would_change = 0           # 改成 FWE 优先会变(采纳POB但同报告有FWE)
    no_main_oil = 0
    kind_cnt = Counter()
    t0 = time.time()
    stop = False

    for fi, fld in enumerate(folders):
        if stop:
            break
        try:
            cnt = fld.Items.Count
        except Exception:
            continue
        if not cnt:
            continue
        log("[%d/%d] %s (%s)" % (fi + 1, len(folders), fld.Name, cnt))
        try:
            items = fld.Items
            items.Sort("[ReceivedTime]", True)
            sub = items.Restrict(filt) if filt else items
        except Exception:
            continue
        n = 0
        for it in sub:
            n += 1
            if SCAN_CAP and n > SCAN_CAP:
                break
            try:
                t = it.ReceivedTime
                rt = t.strftime("%Y-%m-%d %H:%M:%S") if hasattr(t, "strftime") else ""
            except Exception:
                continue
            cands = [c for c in bytime.get(rt, []) if c not in seen_keys]
            if not cands:
                continue
            try:
                subj = str(it.Subject or "")
            except Exception:
                subj = ""
            blob = R.norm(subj)
            chosen_cand = None
            if len(cands) == 1:
                chosen_cand = cands[0]
            else:
                for c in cands:
                    if c[0] and c[0] in blob:
                        chosen_cand = c
                        break
            if chosen_cand is None:
                continue
            try:
                att = R.pick_report_attachment(it)
            except Exception:
                att = None
            if att is None:
                continue
            st = extract_rob_stage(att)
            if st is None:
                continue
            seen_keys.add(chosen_cand)
            kind_cnt[st["kind"] or "(空)"] += 1
            if st["kind"] and st["kind"] != "BERTH":
                continue
            # 选主油种
            main = None
            for oil in ("LSFO", "ULSFO", "HSFO", "MGO"):
                if oil in st["chosen"]:
                    main = oil
                    break
            if main is None:
                no_main_oil += 1
                continue
            p = st["chosen"][main]
            stage_cnt[p] += 1
            seen_p = st["seen"].get(main, set())
            hp, hf = 50 in seen_p, 40 in seen_p
            if hp and hf:
                both_pob_fwe += 1
            if hp and not hf:
                only_pob += 1
            if hf and not hp:
                only_fwe += 1
            if p == 50 and hf:
                fwe_would_change += 1
            if args.limit and len(seen_keys) >= args.limit:
                stop = True
                break

    total = sum(stage_cnt.values())
    log("")
    log("=== 结果 (用时 %.0fs, 统计 %d 封 BERTH) ===" % (time.time() - t0, total))
    log("目标行数 %d, 命中 %d, 未命中 %d" % (n_target, len(seen_keys), n_target - len(seen_keys)))
    log("REPORT_KIND 分布: %s" % dict(kind_cnt))
    log("无机油值的 BERTH: %d" % no_main_oil)
    log("")
    log("【当前规则 POB 优先】主油种采纳阶段分布:")
    for p in (50, 40, 30, 20, 0):
        c = stage_cnt.get(p, 0)
        log("  %-14s %5d  (%.1f%%)" % (PRIO_NAME[p], c, 100.0 * c / total if total else 0))
    log("")
    log("【FWE 阶段出现情况】(主油种):")
    log("  同时有 POB 和 FWE 值 : %d" % both_pob_fwe)
    log("  只有 POB 没有 FWE    : %d" % only_pob)
    log("  只有 FWE 没有 POB    : %d" % only_fwe)
    log("")
    log("【改成 FWE 优先的影响】当前采纳 POB、但同报告存在 FWE 值 -> 会变: %d 条" % fwe_would_change)
    return 0


if __name__ == "__main__":
    sys.exit(main())
