# -*- coding: utf-8 -*-
r"""全量 BERTH ROB 强制回刷: 用新 extract_rob(POB 优先) 重解析历史 BERTH 报告覆盖 CSV。

== 背景 ==
2026-11-19 修复 BERTH 报告 ROB 取值: 旧规则"同一油种后值覆盖前值"会取到附件表格
最下方的 ANCHOR AWEIGH(起锚)/DROP ANCHOR(抛锚) ROB, 而航次期初应取靠泊 POB ROB。
实例 CUL XIAMEN 2625: POB 635.57 被 ANCHOR AWEIGH 的 651.18 覆盖。
新规则(prio POB50>FWE40>ANCHOR AWEIGH30>DROP20)对【今后抓取】已生效, 但历史 CSV
里 3600 条 BERTH 行仍是旧值 —— 三重去重(每日/回补/报告归档)都跳过已存在记录,
不会自动回刷。本脚本做这一次性的强制重解析。

== 做法 ==
1. 读 rob_history.csv 全部 report_type=BERTH 行建索引:
   key=(norm(vessel), report_time[:19]) -> 行号 + 当前 oil 值 + 船名。
   (CSV 的 report_time 就是邮件 ReceivedTime, 见 rob_refresh.apply_hit, 秒级精确匹配。)
2. 扫 Outlook 所有 CULINES.COM / VESSELREPORT Store 的收件箱【一级船名夹及 Master 子夹】:
   对每封邮件用 ReceivedTime[:19] + 船名 token(主题或文件夹名) 在索引里查, 命中才处理。
3. 用新的 R.extract_rob(att) 重解析附件, 与 CSV 当前 oil 值对比, 有差异才记录。
4. 默认 dry-run(只报告不写); --apply 才写, 且写前强制备份 CSV。

== 安全护栏 ==
- 默认 dry-run; 只改索引命中的 BERTH 行的 oil 列(lsfo/hsfo/mgo/ulsfo/bw/fw);
  绝不新增/删除行, 不碰 NOON/SAILING 等其他行, 不动 date/vessel/report_time/type。
- 新值为空(解析不出该油种)时不覆盖该油种, 留旧值 —— 宁可留旧不写空。
- 解析出明确非 BERTH 类型(REPORT_KIND)时跳过, 防误伤。
- 全程单线程串行访问 Outlook(禁止并发, 否则 RPC 崩溃)。本脚本不 git, 由正常 bat 推送。

== 用法(culadmin 那台, OUTLOOK.EXE 已打开并加载主界面) ==
  python backfill_berth_rob.py --dry-run                       # 全量预演, 看命中率+差异清单
  python backfill_berth_rob.py --dry-run --vessel CUL XIAMEN   # 单船预演(先验证定位是否正确)
  python backfill_berth_rob.py --dry-run --limit 5             # 前5条差异预演
  python backfill_berth_rob.py --apply --vessel CUL XIAMEN     # 单船正式回刷
  python backfill_berth_rob.py --apply                         # 全量正式回刷
"""
import sys, os, csv, argparse, time, shutil
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import win32com.client as win32
import rob_refresh as R

HISTORY_CSV = os.path.join(BASE, "rob_data", "rob_history.csv")
OIL_COLS = ("lsfo", "hsfo", "mgo", "ulsfo", "bw", "fw")
OIL_KEYS = ("LSFO", "HSFO", "MGO", "ULSFO", "BW", "FW")
MAX_ITEMS = int(os.environ.get("BERTH_REPARSE_MAX_ITEMS", "8000"))
TOP_INBOX_SKIP = 20000   # 顶层平铺收件箱(6万封)不扫, BERTH 报告都在船名夹


def log(msg):
    print(msg, flush=True)


def load_targets(vessel_filter="", since=""):
    with open(HISTORY_CSV, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    hdr = rows[0]
    ci = {name: hdr.index(name) for name in hdr}
    idx = {}
    for i, r in enumerate(rows[1:], start=1):
        if len(r) < len(hdr):
            continue
        rt = (r[ci["report_time"]] or "")[:19]
        if (r[ci["report_type"]] or "").strip().upper() != "BERTH" or not rt:
            continue
        if since and rt[:10] < since:
            continue
        vessel = r[ci["vessel"]]
        if vessel_filter:
            k = vessel_filter.upper()
            if k not in vessel.upper() and k != (r[ci["code"]] or "").upper():
                continue
        key = (R.norm(vessel), rt)
        idx[key] = {"row": i, "old": {o: r[ci[o]] for o in OIL_COLS},
                    "vessel": vessel, "rt": rt}
    return rows, hdr, ci, idx


def same_num(a, b, tol=0.005):
    if a in ("", None):
        return b in ("", None)
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip() == str(b).strip()


def parse_att_new(att):
    try:
        rob = R.extract_rob(att)
    except Exception as e:
        log("    [warn] extract_rob 异常: %s" % str(e)[:60])
        return None, ""
    if not rob:
        return None, ""
    kind = (rob.get("REPORT_KIND") or "").strip().upper()
    d = {}
    for k in OIL_KEYS:
        v = rob.get(k)
        if isinstance(v, (int, float)):
            d[k] = round(float(v), 2)
    return d, kind


def build_vessel_index():
    fleet = R.load_fleet()
    idx = {}
    for r in fleet:
        for tok in (R.norm(r["vessel"]), R.norm(r.get("code") or "")):
            if tok and len(tok) >= 4:
                idx[tok] = r["vessel"]
    return idx


def scan_folder(fld, targets, idx, seen, changes, args):
    """扫描单个文件夹的邮件, 命中目标则重解析对比。targets 为船名 token 集合(空=靠主题判船)。"""
    try:
        cnt = fld.Items.Count
    except Exception:
        return 0, 0
    if cnt > MAX_ITEMS:
        log("  [skip big] %s (%d items)" % (fld.Name, cnt))
        return 0, 0
    matched_folder_vessels = set()
    if targets:
        fn = R.norm(fld.Name or "")
        for tok, vname in idx.items():
            if tok in fn:
                matched_folder_vessels.add(vname)
    hit = 0
    for i in range(1, cnt + 1):
        try:
            it = fld.Items.Item(i)
        except Exception:
            continue
        if it is None:
            continue
        try:
            recv = it.ReceivedTime
            rt = recv.strftime("%Y-%m-%d %H:%M:%S") if hasattr(recv, "strftime") else ""
        except Exception:
            continue
        if not rt:
            continue
        try:
            subj = str(it.Subject or "")
        except Exception:
            subj = ""
        normsubj = R.norm(subj)
        # 判船: 文件夹名命中的船 + 主题 token 命中的船
        cands = set(matched_folder_vessels)
        for tok, vname in idx.items():
            if tok in normsubj:
                cands.add(vname)
        hit_key = None
        for vname in cands:
            k = (R.norm(vname), rt)
            if k in idx and k not in seen:
                hit_key = k
                break
        if hit_key is None:
            continue
        try:
            att = R.pick_report_attachment(it)
        except Exception:
            att = None
        if att is None:
            continue
        d, kind = parse_att_new(att)
        if d is None:
            continue
        if kind and kind not in ("BERTH",):
            seen.add(hit_key)
            continue
        tgt = idx[hit_key]
        diffs = {}
        for o, k in zip(OIL_COLS, OIL_KEYS):
            new = d.get(k)
            if new is None:
                continue
            if not same_num(tgt["old"][o], new):
                diffs[o] = (tgt["old"][o], new)
        seen.add(hit_key)
        hit += 1
        if diffs:
            changes.append((tgt["row"], tgt["vessel"], tgt["rt"], diffs, kind or "BERTH?"))
            log("  [diff] %s %s | %s -> %s" % (
                tgt["vessel"], tgt["rt"][:16],
                {o: p[0] for o, p in diffs.items()},
                {o: p[1] for o, p in diffs.items()}))
            if args.limit and len(changes) >= args.limit:
                return hit, 1
    return hit, 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true", help="真正写 CSV(默认 dry-run)")
    ap.add_argument("--dry-run", action="store_true", help="显式预演(默认)")
    ap.add_argument("--vessel", default="", help="只处理某船/代码, 如 CUL XIAMEN / H708")
    ap.add_argument("--limit", type=int, default=0, help="最多记录 N 条差异(小批验证, 0=不限)")
    ap.add_argument("--since", default="", help="只处理 report_time>=YYYY-MM-DD")
    args = ap.parse_args()

    apply = args.apply and not args.dry_run
    log("=== BERTH ROB 强制回刷 (%s) ===" % ("APPLY" if apply else "DRY-RUN"))
    rows, hdr, ci, idx = load_targets(args.vessel, args.since)
    log("目标 BERTH 行: %d" % len(idx))
    if not idx:
        return 0
    vidx = build_vessel_index()

    seen, changes = set(), []
    t0 = time.time()
    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for st in ns.Stores:
        name = (st.DisplayName or "").upper()
        if "CULINES.COM" not in name and "VESSELREPORT" not in name:
            continue
        log("[store] %s" % st.DisplayName)
        try:
            root = st.GetRootFolder()
            inbox, bn = None, -1
            for f in root.Folders:
                if f.Name.strip() in ("收件箱", "Inbox", "INBOX"):
                    try:
                        c = f.Items.Count
                    except Exception:
                        c = 0
                    if c > bn:
                        inbox, bn = f, c
            if inbox is None:
                continue
            # 只扫一级船名夹 + 其 Master 子夹; 顶层大收件箱跳过
            for sf in inbox.Folders:
                fn = R.norm(sf.Name or "")
                is_vessel_folder = any(tok in fn for tok in vidx)
                if not is_vessel_folder:
                    continue
                flds = [sf]
                try:
                    for ssf in sf.Folders:
                        if "MASTER" in (ssf.Name or "").upper():
                            flds.append(ssf)
                except Exception:
                    pass
                for fld in flds:
                    _, stop = scan_folder(fld, True, idx, seen, changes, args)
                    if stop:
                        break
                if args.limit and len(changes) >= args.limit:
                    break
        except Exception as e:
            log("  [warn] store 遍历异常: %s" % str(e)[:80])
        if args.limit and len(changes) >= args.limit:
            break

    log("")
    log("命中 %d / %d 封, 有差异 %d 条, 未命中 %d 条, 用时 %.0fs"
        % (len(seen), len(idx), len(changes), len(idx) - len(seen), time.time() - t0))
    if not changes:
        log("无改动。")
        return 0
    if not apply:
        log("DRY-RUN: 未写盘。确认命中率和差异后加 --apply 执行。")
        return 0

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak_dir = os.path.join(BASE, "rob_data", "tpl_backup", "berth_reparse_" + stamp)
    os.makedirs(bak_dir, exist_ok=True)
    shutil.copy2(HISTORY_CSV, os.path.join(bak_dir, "rob_history.csv"))
    log("已备份 CSV -> %s" % bak_dir)

    byrow = {}
    for row_i, vessel, rt, diffs, kind in changes:
        byrow.setdefault(row_i, {}).update({o: p[1] for o, p in diffs.items()})
    for row_i, newvals in byrow.items():
        for o, v in newvals.items():
            rows[row_i][ci[o]] = str(v)
    with open(HISTORY_CSV, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    log("已写回 %d 行 -> %s" % (len(byrow), HISTORY_CSV))
    log("检查 git diff 确认只改 BERTH 行 oil 列, 再走正常 bat 提交推送。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
