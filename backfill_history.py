# -*- coding: utf-8 -*-
"""历史 ROB 报告回填: 扫描 Outlook 历史邮件, 补全 rob_history.csv 的 speed/report_type/BERTH-ROB。

目标(用户 2026-09-28 需求):
  1) TDR vs Actual 需要历史每期 NOON 报告的 speed -> 补 speed + report_type。
  2) Voyage 航次油耗按 BERTH 报告 ROB 计算 -> 需要历史上所有 BERTH 报告的 ROB 行。

做法(性能优化版, 单船秒级):
  对每艘船 v:
    (a) match_folder 定位其 Vessel 文件夹树(含 '<X> Master' 等子文件夹, 递归),
        对该树每个文件夹 Restrict([ReceivedTime]>=since) 后遍历, 主题须含船名 norm token
        (防共享/前缀文件夹串船) -> extract_rob。
    (b) 补充: 收件箱顶层 Restrict 主题含船名 token -> extract_rob (覆盖落在收件箱顶层的报告)。
  与现有 rob_history.csv 按 (vessel, report_time) 去重合并:
    - 已有行: 仅补 speed / report_type 两列(若缺失)
    - 新行: 新增(保留完整历史 ROB, 不重复每日最新快照)
  --dry-run 只统计不写盘。
"""
import sys, os, csv, argparse, json
sys.stdout.reconfigure(encoding="utf-8")
from datetime import datetime, timedelta
from collections import Counter, defaultdict

import rob_refresh as R

BASE = os.path.dirname(os.path.abspath(__file__))
HISTORY_CSV = R.HISTORY_CSV
SNAP_FIELDS = R.SNAP_FIELDS
FOUND_JSON = os.path.join(BASE, "rob_data", "backfill_found.json")


def _blank(v):
    """判断单元格是否为空(兼容 None / 空串 / 数字 0/0.0)。"""
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip() == ""
    # 数字: 0/0.0 视为空(可被回填覆盖); 非零数字视为已有值
    try:
        return float(v) == 0.0
    except Exception:
        return False


def collect_vessels(history_csv):
    vs = {}
    with open(history_csv, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            v = (row.get("vessel") or "").strip()
            if not v or R.is_offline(v):
                continue
            vs.setdefault(v, (row.get("code") or "").strip())
    return vs


def scan_vessel(vessel, code, inbox, folder, since_dt):
    """扫描单船: 递归文件夹树 + 收件箱主题。返回报告列表。"""
    out = []
    cutoff = since_dt.strftime("%m/%d/%Y %H:%M %p")
    nv = R.norm(vessel); nc = R.norm(code or "")
    tokens = [t for t in (nv, nc) if len(t) >= 4]
    if not tokens:
        return out

    def scan_folder(fo):
        try:
            items = fo.Items.Restrict("[ReceivedTime] >= '%s'" % cutoff)
            items.Sort("[ReceivedTime]", True)
        except Exception:
            return
        for it in items:
            try:
                rt = it.ReceivedTime
            except Exception:
                continue
            if rt < since_dt:
                continue
            try:
                subj = it.Subject or ""
            except Exception:
                subj = ""
            sj = R.norm(subj)
            if not any(t in sj for t in tokens):
                continue
            att = R.pick_report_attachment(it)
            if att is None:
                continue
            rob = R.extract_rob(att)
            if not any(k in rob for k in R.OIL_KEYS):
                continue
            su = subj.upper()
            # 报告类型优先用附件正文判定(REPORT_KIND), 主题关键词仅作兜底
            rtype = (rob.get("REPORT_KIND") or "").strip().upper()
            if not rtype:
                for k in R.REPORT_KEYS:
                    if k in su:
                        rtype = k
                        break
            out.append({
                "report_time": rt.strftime("%Y-%m-%d %H:%M:%S"),
                "report_type": rtype,
                "speed": rob.get("SPEED"),
                "lsfo": rob.get("LSFO"), "hsfo": rob.get("HSFO"),
                "mgo": rob.get("MGO"), "ulsfo": rob.get("ULSFO"),
                "bw": rob.get("BW"), "fw": rob.get("FW"),
            })

    # (a) 文件夹树(含子文件夹递归)
    if folder is not None:
        scan_folders = [folder]
        stack = []
        try:
            stack = list(folder.Folders)
        except Exception:
            pass
        while stack:
            sub = stack.pop(0)
            scan_folders.append(sub)
            try:
                stack.extend(sub.Folders)
            except Exception:
                pass
        for fo in scan_folders:
            scan_folder(fo)
    # (b) 收件箱顶层(主题含船名 token)
    try:
        sql = "@SQL=\"urn:schemas:httpmail:subject\" like '%%%s%%'" % nv
        items = inbox.Items.Restrict(sql)
        items.Sort("[ReceivedTime]", True)
        for it in items:
            try:
                rt = it.ReceivedTime
            except Exception:
                continue
            if rt < since_dt:
                continue
            att = R.pick_report_attachment(it)
            if att is None:
                continue
            rob = R.extract_rob(att)
            if not any(k in rob for k in R.OIL_KEYS):
                continue
            su = (it.Subject or "").upper()
            # 报告类型优先用附件正文判定(REPORT_KIND), 主题关键词仅作兜底
            rtype = (rob.get("REPORT_KIND") or "").strip().upper()
            if not rtype:
                for k in R.REPORT_KEYS:
                    if k in su:
                        rtype = k
                        break
            out.append({
                "report_time": rt.strftime("%Y-%m-%d %H:%M:%S"),
                "report_type": rtype,
                "speed": rob.get("SPEED"),
                "lsfo": rob.get("LSFO"), "hsfo": rob.get("HSFO"),
                "mgo": rob.get("MGO"), "ulsfo": rob.get("ULSFO"),
                "bw": rob.get("BW"), "fw": rob.get("FW"),
            })
    except Exception:
        pass
    return out


def merge(history_csv, found, found_meta, dry_run):
    rows = []
    with open(history_csv, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    existing = {}
    for i, r in enumerate(rows):
        key = ((r.get("vessel") or "").strip(), (r.get("report_time") or "").strip())
        existing.setdefault(key, i)
    added = 0; updated_speed = 0; updated_type = 0
    for vessel, reps in found.items():
        code = found_meta.get(vessel, "")
        for rep in reps:
            key = (vessel, rep["report_time"])
            if key in existing:
                i = existing[key]
                if _blank(rows[i].get("speed")):
                    if rep["speed"] not in (None, ""):
                        rows[i]["speed"] = rep["speed"]; updated_speed += 1
                if _blank(rows[i].get("report_type")):
                    if rep["report_type"]:
                        rows[i]["report_type"] = rep["report_type"]; updated_type += 1
            else:
                newrow = {f: "" for f in SNAP_FIELDS}
                newrow.update({
                    "date": rep["report_time"][:10],
                    "vessel": vessel, "code": code,
                    "lsfo": rep["lsfo"], "hsfo": rep["hsfo"], "mgo": rep["mgo"],
                    "ulsfo": rep["ulsfo"], "bw": rep["bw"], "fw": rep["fw"],
                    "found": "1", "report_time": rep["report_time"],
                    "speed": rep["speed"] or "", "report_type": rep["report_type"],
                })
                rows.append(newrow); added += 1
    if not dry_run:
        with open(history_csv, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=SNAP_FIELDS)
            w.writeheader()
            for r in rows:
                w.writerow({k: r.get(k, "") for k in SNAP_FIELDS})
    return added, updated_speed, updated_type


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", default="2026-01-01", help="YYYY-MM-DD 起始(回溯窗口)")
    ap.add_argument("--vessel", default=None, help="只回填某船(调试)")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--resume", action="store_true",
                    help="跳过 Outlook 扫描, 直接读取上次 dump 的 backfill_found.json 做合并")
    args = ap.parse_args()

    since_dt = datetime.strptime(args.since, "%Y-%m-%d")
    try:
        since_dt = since_dt.replace(tzinfo=datetime.now().astimezone().tzinfo)
    except Exception:
        pass
    stores = R.connect_outlook()
    if not stores:
        print("[ERR] Outlook 未连接"); return 1
    inbox = stores[0].GetDefaultFolder(6)
    idx = {}
    def walk(f, d=0):
        if d > 5:
            return
        try:
            for c in f.Folders:
                idx.setdefault(R.norm(c.Name), c)
                walk(c, d + 1)
        except Exception:
            pass
    walk(inbox)
    vessels = collect_vessels(HISTORY_CSV)
    if args.vessel:
        vessels = {args.vessel: vessels.get(args.vessel, "")}

    found = {}
    found_meta = {}
    if args.resume and os.path.exists(FOUND_JSON):
        print("=== resume from %s ===" % FOUND_JSON)
        with open(FOUND_JSON, encoding="utf-8") as f:
            tmp = json.load(f)
        found = tmp.get("found", {})
        found_meta = tmp.get("meta", {})
    else:
        for vessel, code in vessels.items():
            folder, exact = R.match_folder(idx, vessel, code)
            if folder is None:
                # 退而求其次: 收件箱顶层主题 Restrict(scan_vessel 内已含)
                pass
            reps = scan_vessel(vessel, code, inbox, folder, since_dt)
            if reps:
                found[vessel] = reps
                found_meta[vessel] = code
        try:
            with open(FOUND_JSON, "w", encoding="utf-8") as f:
                json.dump({"found": found, "meta": found_meta}, f, ensure_ascii=False)
            print("scan dump -> %s (%d vessels)" % (FOUND_JSON, len(found)))
        except Exception as e:
            print("[WARN] dump found json failed:", e)

    rc = Counter()
    spc = 0
    for v, reps in found.items():
        for r in reps:
            rc[r["report_type"] or "(none)"] += 1
            if r["speed"] not in (None, ""):
                spc += 1
    print("=== scan stats ===")
    print("vessels matched:", len(found))
    print("total reports parsed:", sum(len(v) for v in found.values()))
    print("by type:", dict(rc))
    print("with speed:", spc)
    if args.vessel:
        for v, reps in found.items():
            print("  %s: %d reports" % (v, len(reps)))
            for r in sorted(reps, key=lambda x: x["report_time"])[:20]:
                print("    ", r["report_time"][:10], r["report_type"], "sp=", r["speed"],
                      "ls=", r["lsfo"], "hs=", r["hsfo"])
    added, us, ut = merge(HISTORY_CSV, found, found_meta, dry_run=args.dry_run)
    print("merge: added=%d updated_speed=%d updated_type=%d (dry_run=%s)" % (added, us, ut, args.dry_run))
    return 0


if __name__ == "__main__":
    sys.exit(main())
