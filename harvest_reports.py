# -*- coding: utf-8 -*-
"""harvest_reports.py — 全量报告归档(报告级, 一封报告一行)

背景
----
rob_history.csv 是「船 × 日期 = 1 行」的日快照, 每日刷新只保留时间最新的一封报告。
实测 CUL HUANGPU 近 30 天邮箱实收 NOON 27 / BERTH 25 / SAILING 22 / DEPARTURE 14
共 88 封, 入库仅 30 条, 留存率 34%。BERTH 几乎全丢, 而航次油耗的期初/期末边界
恰恰只能取 BERTH 或 NOON。

本脚本不改动任何现有表/网页/算法, 只**旁路**产出报告级归档:
    rob_data/rob_reports.jsonl     一封报告一行

设计要点
--------
* 先一次性扫描全部文件夹建内存池, 再按船匹配 —— 避免"每船重扫一遍邮箱"的 O(N×M)。
* 只按 主题 / 所在文件夹 认船, 逐封解析发件人(X.500)太慢且收益低。
* 可断点续跑: 按 vessel+received+subject 去重, 重复执行只补新增。

用法
----
  python harvest_reports.py --days 30
  python harvest_reports.py --days 90 --vessel CUHP
  python harvest_reports.py --days 30 --rebuild
"""
import os
import sys
import json
import hashlib
import argparse
import time
from datetime import datetime, timedelta

import rob_refresh as R

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = os.path.join(HERE, "rob_data", "rob_reports.jsonl")
OIL_KEYS = ("LSFO", "HSFO", "MGO", "ULSFO", "BW", "FW", "REEFER")


KIND_PHRASES = (
    ("SAILING", ("SAILING REPORT", "SAILING REP", "DEPARTURE REPORT", "DEPARTURE REP")),
    ("BERTH", ("BERTH REPORT", "BERTHING REPORT", "BERTH REP")),
    ("NOON", ("NOON REPORT", "NOON REP")),
    ("ANCHOR", ("ANCHOR REPORT", "ANCHOR REP")),
    ("DRIFT", ("DRIFT REPORT",)),
)


def refine_kind(body_kind, attach_name, subject):
    """修正报道类型。优先级: 附件名完整短语 > 主题完整短语 > 正文特征判定。

    为什么需要这一层: rob_refresh.detect_report_kind 的 BERTH 正文标记里含 "ALL FAST",
    而 SAILING REPORT 也会记录离泊的 ALL FAST 时间, 导致 CUHP 近 30 天 42 条 BERTH 中
    21 条实为 SAILING(2026-11-10 实测)。附件名(如 'CUL SAILING REPORT.xlsx')和主题
    ('... sailing report at THLCH')都写得很明确, 作为更高优先级的信号更可靠。
    这里只影响归档, 不改动线上判定, 避免动到已在用的航次数字。
    """
    for label, src in (("attach", attach_name), ("subject", subject)):
        u = (src or "").upper()
        for kind, phrases in KIND_PHRASES:
            if any(p in u for p in phrases):
                return kind, label
    return (body_kind or "").strip().upper(), "body"


def load_done():
    done = set()
    if not os.path.exists(OUT_FILE):
        return done
    try:
        with open(OUT_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    done.add(json.loads(line).get("key", ""))
                except Exception:
                    pass
    except Exception:
        pass
    return done


def make_key(vessel, recv, subject):
    raw = "%s|%s|%s" % (vessel, recv, subject or "")
    return hashlib.md5(raw.encode("utf-8", "replace")).hexdigest()[:20]


def build_pool(folders, cutoff, label=""):
    """扫一遍给定文件夹, 返回带附件的邮件池。每封只读一次必要属性。"""
    filt = "[ReceivedTime] >= '%s'" % cutoff.strftime("%m/%d/%Y %I:%M %p")
    pool = []
    t_all = time.time()
    for idx, fo in enumerate(folders, 1):
        try:
            items = fo.Items.Restrict(filt)
        except Exception:
            continue
        try:
            fname = fo.Name or ""
        except Exception:
            fname = "?"
        fnorm = R.norm(fname)
        n = 0
        try:
            for it in items:
                try:
                    if it.Attachments.Count == 0:
                        continue
                except Exception:
                    continue
                try:
                    recv = it.ReceivedTime
                    subj = it.Subject or ""
                    eid = it.EntryID
                except Exception:
                    continue
                pool.append({"eid": eid, "it": it, "recv": recv, "subj": subj,
                             "sn": R.norm(subj), "fn": fnorm, "folder": fname})
                n += 1
        except Exception:
            pass
        if n:
            print("   %s[%3d/%3d] %-28s +%-4d (pool=%d)"
                  % (label, idx, len(folders), fname[:28], n, len(pool)), flush=True)
    print("   pool built: %d mails with attachments in %.0fs"
          % (len(pool), time.time() - t_all), flush=True)
    return pool


def pick(rec, pool):
    """按 精确文件夹名 / 文件夹名前缀 / 主题含船名 三条规则挑该船的候选。"""
    nv = R.norm(rec["vessel"])
    nc = R.norm(rec.get("code") or "")
    out = []
    for p in pool:
        fn, sn = p["fn"], p["sn"]
        if fn and ((nv and fn == nv) or (nc and fn == nc)):
            out.append(p)
            continue
        prefix = (nv and fn and (nv in fn or fn in nv)) or (nc and fn and nc in fn)
        if prefix:
            # 共用/变体文件夹必须主题里也带船名, 否则会串船(MEDKON DON vs LIA)
            if any(t and t in sn for t in (nv, nc)):
                out.append(p)
            continue
        if any(t and t in sn for t in (nv, nc)):
            out.append(p)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=30)
    ap.add_argument("--vessel", default="")
    ap.add_argument("--rebuild", action="store_true")
    ap.add_argument("--max-per-vessel", type=int, default=600)
    ap.add_argument("--no-extra", action="store_true")
    ap.add_argument("--force", action="store_true",
                    help="邮件池小于阈值时仍继续(默认中止, 防止静默产出空归档)")
    args = ap.parse_args()

    if args.rebuild and os.path.exists(OUT_FILE):
        os.remove(OUT_FILE)
    done = load_done()

    fleet = R.load_fleet()
    if args.vessel:
        k = args.vessel.upper()
        fleet = [r for r in fleet
                 if k in r["vessel"].upper() or k == (r.get("code") or "").upper()]
    print("fleet=%d days=%d archived=%d" % (len(fleet), args.days, len(done)), flush=True)

    stores = R.connect_outlook()
    if not stores:
        raise SystemExit("no CULINES store")

    folders = []
    for s in stores:
        try:
            _inb = s.GetDefaultFolder(6)
        except Exception:
            continue
        folders.append(_inb)
        folders.extend(R.build_folder_list(_inb))
        if not args.no_extra:
            folders.extend(R.build_extra_folders(s, _inb))
    print("folders=%d" % len(folders), flush=True)

    cutoff = datetime.now() - timedelta(days=args.days)
    pool = build_pool(folders, cutoff)

    # 池大小闸门 —— 与 rob_refresh 的同步闸门同理, 但这里后果更隐蔽:
    # Outlook 句柄泄漏(实测跑满 8 天达 2.2 万句柄)或联机存档掉线时, Restrict 会返回
    # 空集或只返回极小一部分, 脚本照常跑完并"归档 0 条", 看起来像当天没报告。
    # 正常量级参考: 30 天约 8,000 封带附件邮件; 按天折算后再打 6 折作为下限。
    min_pool = min(4000, max(200, int(100 * args.days)))
    if len(pool) < min_pool and not args.force:
        print("!! 中止: 邮件池只有 %d 封 < 阈值 %d" % (len(pool), min_pool))
        print("!! Outlook 可能资源耗尽(句柄过多)或联机存档未同步完整。")
        print("!! 请重启 Outlook 后重跑; 确需继续加 --force。")
        return 1
    if not pool:
        print("!! 邮件池为空, 中止")
        return 1

    n_new = 0
    fout = open(OUT_FILE, "a", encoding="utf-8")
    for i, rec in enumerate(fleet, 1):
        t0 = time.time()
        cands = pick(rec, pool)
        rows = []
        tried = 0
        for p in cands:
            if tried >= args.max_per_vessel:
                break
            key = make_key(rec["vessel"], str(p["recv"]), p["subj"])
            if key in done:
                continue
            try:
                att = R.pick_report_attachment(p["it"])
            except Exception:
                att = None
            if att is None:
                continue
            tried += 1
            try:
                rob = R.extract_rob(att)
            except Exception:
                continue
            if not any(k in rob for k in OIL_KEYS):
                continue
            try:
                aname = att.FileName or ""
            except Exception:
                aname = ""
            kind, ksrc = refine_kind(rob.get("REPORT_KIND"), aname, p["subj"])
            rows.append({
                "key": key,
                "vessel": rec["vessel"],
                "code": rec.get("code") or "",
                "received": str(p["recv"])[:19],
                "subject": p["subj"],
                "folder": p["folder"],
                "attach": aname,
                "report_type": kind,
                "kind_src": ksrc,
                "lsfo": rob.get("LSFO"), "hsfo": rob.get("HSFO"),
                "mgo": rob.get("MGO"), "ulsfo": rob.get("ULSFO"),
                "bw": rob.get("BW"), "fw": rob.get("FW"),
                "refeer": rob.get("REEFER"), "speed": rob.get("SPEED"),
            })
            done.add(key)
        for r in rows:
            fout.write(json.dumps(r, ensure_ascii=False) + "\n")
        fout.flush()
        n_new += len(rows)
        print("[%2d/%2d] %-24s cand=%-4d new=%-3d cum=%-5d (%.0fs)"
              % (i, len(fleet), rec["vessel"], len(cands), len(rows), n_new,
                 time.time() - t0), flush=True)
    fout.close()
    print("\nDONE archived_new=%d total=%d -> %s" % (n_new, len(done), OUT_FILE))


if __name__ == "__main__":
    main()
