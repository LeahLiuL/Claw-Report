# -*- coding: utf-8 -*-
r"""全量 BERTH ROB 强制回刷: 用现行 extract_rob(FWE 优先) 重解析历史 BERTH 报告覆盖 CSV。

== 背景 ==
BERTH 报告 ROB 取值经两次修正:
  1) 2026-11-19: 修掉旧规则"同一油种后值覆盖前值"(会取到表格最下方的 ANCHOR AWEIGH
     起锚 / DROP ANCHOR 抛锚 ROB)。实例 CUL XIAMEN 2625 POB 635.57 被起锚 651.18 覆盖。
  2) 2026-11-23(用户定稿, commit ade8239f): 改为 FWE 优先 —— prio
     FWE60 > POB50 > ANCHOR AWEIGH30 > DROP20, 即优先取 FWE ROB, 没有 FWE 才退回 POB。
现行规则对【今后抓取】已生效, 但历史 CSV 里 BERTH 行仍是旧值 —— 三重去重(每日/回补/
报告归档)都跳过已存在记录, 不会自动回刷。本脚本做这一次性的强制重解析。
统计(berth_stage_stats.py, 3991封): 若按 FWE 优先, 1413 条原取 POB 的报告存在 FWE 值会改。

== 2026-11-19 本机修正(重要) ==
原版有两处致命问题, 在本机实测卡死/零命中, 已修:
  1. scan_folder 把同一个 idx 参数当两种结构用(船名token索引 vs (船名,时间)行索引),
     判船处 `tok in fn` 对元组 key 抛 TypeError -> 外层吞掉 -> 命中0/假成功。
  2. 只扫 Inbox.Folders 下的船名夹。但 vesselreport 邮箱里 RPA_成功(16748)/RPA_失败(1951)
     是 Inbox 的【兄弟节点】, 位于 root 层, 且平铺无船名夹 -> 原逻辑一封都定位不到;
     又因逐封 Items.Item(i) 索引访问 + 全 store 扫描 -> COM 挂死(CPU 近 0)。
  现改为: 只锁 vesselreport store, 从 GetRootFolder() 遍历所有文件夹, 用
  Items.Restrict([ReceivedTime]范围) 过滤, `for it in sub` 迭代(非索引), 先按秒级
  时间把候选收窄到 CSV 目标行, 再判船+重解析。与 rob_mail_sync.py 口径一致。

== 做法 ==
1. 读 rob_history.csv 全部 report_type=BERTH 行建索引:
   key=(norm(vessel), report_time[:19]) -> 行号 + 当前 oil 值 + 船名。
   并建 bytime: report_time[:19] -> [key,...] 供秒级时间预筛。
2. 扫 vesselreport@culines.com 的全部文件夹(root + 递归), 每夹 Restrict 时间范围。
3. 命中(时间秒级相同 + 船名 token 匹配主题/附件名)才用 R.extract_rob(att) 重解析, 与
   CSV 当前值对比, 有差异才记录。
4. 默认 dry-run(只报告不写); --apply 才写, 且写前强制备份 CSV。

== 安全护栏 ==
- 默认 dry-run; 只改索引命中的 BERTH 行的 oil 列(lsfo/hsfo/mgo/ulsfo/bw/fw);
  绝不新增/删除行, 不碰 NOON/SAILING 等其他行, 不动 date/vessel/report_time/type。
- 新值为空(解析不出该油种)时不覆盖该油种, 留旧值 —— 宁可留旧不写空。
- 解析出明确非 BERTH 类型(REPORT_KIND)时跳过, 防误伤。
- 全程单线程串行访问 Outlook(禁止并发, 否则 RPC 崩溃)。本脚本不 git, 由正常 bat 推送。

== 用法(culadmin 那台, OUTLOOK.EXE 已打开并加载主界面) ==
  python backfill_berth_rob.py --dry-run                       # 全量预演, 看命中率+差异清单
  python backfill_berth_rob.py --dry-run --vessel "CUL XIAMEN"  # 单船预演(先验证定位是否正确)
  python backfill_berth_rob.py --dry-run --limit 5             # 前5条差异预演
  python backfill_berth_rob.py --apply --vessel "CUL XIAMEN"    # 单船正式回刷
  python backfill_berth_rob.py --apply                          # 全量正式回刷
  调试: env BERTH_REPARSE_SCAN_CAP=N 限制每夹最多迭代 N 封(0=不限)
"""
import sys, os, csv, argparse, time, shutil
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import win32com.client as win32
import rob_refresh as R

HISTORY_CSV = os.path.join(BASE, "rob_data", "rob_history.csv")
OIL_COLS = ("lsfo", "hsfo", "mgo", "ulsfo", "bw", "fw")
OIL_KEYS = ("LSFO", "HSFO", "MGO", "ULSFO", "BW", "FW")
SCAN_CAP = int(os.environ.get("BERTH_REPARSE_SCAN_CAP", "0"))   # 0=不限
MAX_DEPTH = 4


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


def build_bytime(idx):
    """report_time(秒) -> [key,...], 供秒级时间预筛。"""
    d = {}
    for key in idx:
        d.setdefault(key[1], []).append(key)
    return d


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


def pick_key(rt, subject, attnames, bytime, vidx, seen):
    """秒级时间命中后选 key; 单候选直接用, 多候选靠主题/附件名 token 判船。"""
    keys = bytime.get(rt, [])
    if not keys:
        return None
    fresh = [k for k in keys if k not in seen]
    if not fresh:
        return None
    if len(fresh) == 1:
        return fresh[0]
    blob = R.norm(subject) + " " + " ".join(R.norm(a) for a in attnames)
    for k in fresh:
        if k[0] and k[0] in blob:
            return k
    return None


def scan_folder(fld, vidx, idx, bytime, seen, changes, args, filt):
    """扫描单个文件夹, 时间预筛命中目标行才重解析对比。"""
    try:
        items = fld.Items
        items.Sort("[ReceivedTime]", True)
        sub = items.Restrict(filt) if filt else items
    except Exception as e:
        log("  [warn] %s 取子集失败: %s" % (fld.Name, str(e)[:40]))
        return 0, False   # 系统文件夹(日历/联系人等)无 ReceivedTime, 跳过
    hit = 0
    n = 0
    for it in sub:                      # 迭代而非 Items.Item(i) 索引(避免 COM 挂死)
        n += 1
        if SCAN_CAP and n > SCAN_CAP:
            break
        try:
            t = it.ReceivedTime
            rt = t.strftime("%Y-%m-%d %H:%M:%S") if hasattr(t, "strftime") else ""
        except Exception:
            continue
        if not rt or rt not in bytime:
            continue                    # 秒级时间预筛: 绝大多数邮件在此被排除
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
        hit_key = pick_key(rt, subj, attnames, bytime, vidx, seen)
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
                return hit, True
    return hit, False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true", help="真正写 CSV(默认 dry-run)")
    ap.add_argument("--dry-run", action="store_true", help="显式预演(默认)")
    ap.add_argument("--vessel", default="", help="只处理某船/代码, 如 \"CUL XIAMEN\" / CUXM")
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
    bytime = build_bytime(idx)

    # 时间范围 -> Restrict 过滤串(与 CSV report_time 同为本地时间)
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

    seen, changes = set(), []
    t0 = time.time()
    for fi, fld in enumerate(folders):
        try:
            cnt = fld.Items.Count
        except Exception:
            cnt = -1
        if not cnt:
            continue
        log("[%d/%d] %s (%s items)" % (fi + 1, len(folders), fld.Name, cnt))
        hit, stop = scan_folder(fld, vidx, idx, bytime, seen, changes, args, filt)
        if hit:
            log("   命中 %d 封" % hit)
        if stop:
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
