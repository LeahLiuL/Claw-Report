# -*- coding: utf-8 -*-
"""
Abstract Log 每日采集入库 (直接从 Outlook vesselreport 抓, 不依赖 reports.db).

背景
----
abstract_log_db.py 是从 180M 的 reports.db 读 PDF/xlsx 原件解析的, 而
reports.db 不是每天更新 -> 做不到每日同步. 本脚本仿 rob_mail_sync.py 的模式,
每天直接从 Outlook vesselreport@culines.com 抓 Abstract Log 邮件,
解析后增量写 rob_data/abstract_log.db (只有 ~0.5M, 可每天上传).

要点
----
* mail_id = report_db._mail_id(it) (EntryID md5[:20]), 与历史库一致:
  同一封邮件 INSERT OR REPLACE 覆盖, 不会重复.
* 附件选择: 主题含 ABSTRACT 的邮件, 附件未必是 Abstract Log (常见 CII/NOON).
  先按文件名含 ABSTRACT 选; 找不到则对每个 xlsx 逐个 parse_report 校验,
  取第一个真正是 Abstract Log 模板的原件.
* Outlook 门控: 复用 rob_mail_sync.ensure_outlook (mode>=700), 未同步完 exit 2
  让 bat 跳过本轮, 下一轮回补.
* 回补窗口: --days 5, 即使漏跑几天也能补回来.

用法
----
  python abstract_daily_sync.py --days 5            # dry-run, 不写库
  python abstract_daily_sync.py --days 5 --apply    # 真正入库
  python abstract_daily_sync.py --days 5 --apply --limit 3   # 限量验证
"""
import os
import sys
import gc
import time
import shutil
import sqlite3
import argparse
import datetime
import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

OUT_DB = os.path.join(HERE, "rob_data", "abstract_log.db")
ARCHIVE = os.path.join(HERE, "rob_data", "report_archive")
LOG = os.path.join(HERE, "rob_data", "abstract_daily_sync.log")
MAX_DEPTH = 4
PER_FOLDER_CAP = 400


def _load(mod, name):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, mod))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def log(msg):
    line = "[abs-sync] %s" % msg
    print(line, flush=True)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write("%s %s\n" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), line))
    except Exception:
        pass


def is_abstract_mail(subject, att_names):
    """候选: 主题/附件含 ABSTRACT 且至少有 xlsx 附件 (排除纯文字回复邮件)."""
    has_xlsx = any(a.lower().endswith((".xlsx", ".xlsm", ".xls")) for a in (att_names or []))
    if not has_xlsx:
        return False
    blob = ((subject or "") + " " + " ".join(att_names or "")).upper()
    return "ABSTRACT" in blob


def pick_abstract_attachment(it, A):
    """返回 (att, filename): 优先文件名含 ABSTRACT 的 xlsx; 否则逐个 parse 校验."""
    atts = []
    try:
        for a in it.Attachments:
            try:
                fn = str(a.FileName or "")
            except Exception:
                continue
            if fn.lower().endswith((".xlsx", ".xlsm", ".xls")):
                atts.append((a, fn))
    except Exception:
        return None
    if not atts:
        return None
    # 1) 文件名含 ABSTRACT
    for a, fn in atts:
        if "ABSTRACT" in fn.upper():
            return a, fn
    # 2) 逐个存临时并校验是不是 Abstract Log 模板
    import tempfile
    for a, fn in atts:
        tmp = os.path.join(tempfile.gettempdir(), "abs_probe_%d.xlsx" % os.getpid())
        try:
            a.SaveAsFile(tmp)
            res, _ = A.parse_report(tmp)
            if res is not None:
                return a, fn
        except Exception:
            pass
        finally:
            try:
                os.remove(tmp)
            except Exception:
                pass
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=5, help="回补窗口(天), 默认5")
    ap.add_argument("--apply", action="store_true", help="写库(默认 dry-run)")
    ap.add_argument("--limit", type=int, default=0, help="最多处理N封(验证用)")
    ap.add_argument("--out", default=OUT_DB)
    args = ap.parse_args()

    R = _load("rob_mail_sync.py", "rms")
    RD = _load("report_db.py", "rdb")
    A = _load("abstract_log_db.py", "aldb")

    log("START days=%d apply=%s limit=%d" % (args.days, args.apply, args.limit))

    app, ns = R.ensure_outlook()  # mode>=700 门控, 未同步完会 sys.exit(2)

    # 定位 vesselreport store
    target = None
    for si in range(1, ns.Stores.Count + 1):
        st = ns.Stores(si)
        if "vesselreport" in (st.DisplayName or "").lower():
            target = st
            break
    if target is None:
        log("ERROR vesselreport store not found")
        return 1
    root = target.GetRootFolder()
    folders = [root] + R.walk_folders(root)
    log("store=%s folders=%d" % (target.DisplayName, len(folders)))

    since = datetime.datetime.now() - datetime.timedelta(days=args.days)
    filt = "[ReceivedTime] >= '%s'" % since.strftime("%m/%d/%Y %H:%M %p")
    log("filter=%s" % filt)

    # 已有 mail_id (增量跳过)
    db, leg_cols = A.init_db(args.out)
    have = {r[0] for r in db.execute("SELECT mail_id FROM voyages")}
    log("existing voyages=%d" % len(have))

    inv_mail, toks = RD._build_vessel_index()

    found = added = updated = skipped = failed = 0
    t0 = time.time()
    for fi, f in enumerate(folders):
        try:
            items = f.Items
            items.Sort("[ReceivedTime]", True)
            try:
                sub = items.Restrict(filt)
            except Exception:
                sub = items
        except Exception:
            continue
        n = 0
        try:
            for it in sub:
                n += 1
                if n > PER_FOLDER_CAP:
                    break
                try:
                    subj = str(it.Subject or "")
                except Exception:
                    subj = ""
                try:
                    att_names = [str(a.FileName or "") for a in it.Attachments]
                except Exception:
                    att_names = []
                if not is_abstract_mail(subj, att_names):
                    continue
                found += 1
                try:
                    mid = RD._mail_id(it)
                except Exception:
                    continue
                if mid in have:
                    skipped += 1
                    continue
                if args.limit and (added + updated) >= args.limit:
                    continue
                picked = pick_abstract_attachment(it, A)
                if picked is None:
                    failed += 1
                    log("  [no-att] %s %s" % (mid[:10], subj[:45]))
                    continue
                att, fn = picked
                try:
                    rt = it.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    rt = ""
                ym = rt[:7].replace("-", "/") if len(rt) >= 7 else "unknown/00"
                dest_dir = os.path.join(ARCHIVE, ym)
                os.makedirs(dest_dir, exist_ok=True)
                dest = os.path.join(dest_dir, mid + "_abstract.xlsx")
                try:
                    att.SaveAsFile(dest)
                except Exception as e:
                    failed += 1
                    log("  [save-err] %s %s" % (mid[:10], str(e)[:50]))
                    continue
                res, legs = A.parse_report(dest)
                if res is None:
                    failed += 1
                    log("  [parse-fail] %s %s" % (mid[:10], str(legs)[:45]))
                    try:
                        os.remove(dest)
                    except Exception:
                        pass
                    continue
                voy = res
                sender = ""
                try:
                    sender = R.get_sender(it) or ""
                except Exception:
                    pass
                vessel, code = RD._guess_vessel(subj, sender, "", inv_mail, toks)
                if not vessel:
                    vessel = voy.get("ship_name", "") or ""
                voy.update(mail_id=mid, received=rt, vessel=vessel or "",
                           subject=subj[:300], attach_name=fn[:200], file_path=dest,
                           leg_count=len(legs))
                if args.apply:
                    cols = list(voy.keys())
                    db.execute("INSERT OR REPLACE INTO voyages(%s) VALUES(%s)" % (
                        ",".join(cols), ",".join("?" * len(cols))),
                        [voy.get(k) for k in cols])
                    for seq, leg in enumerate(legs, 1):
                        leg.update(mail_id=mid, vessel=vessel or "", received=rt,
                                   voyage_no=voy.get("voyage_no", ""), seq=seq)
                        vals = [leg.get(k) for k in leg_cols]
                        db.execute("INSERT OR REPLACE INTO legs(%s) VALUES(%s)" % (
                            ",".join(leg_cols), ",".join("?" * len(leg_cols))), vals)
                    db.commit()
                    db.execute("INSERT OR REPLACE INTO meta(k,v) VALUES(?,?)",
                               ("last_parse", datetime.datetime.now().isoformat(timespec="seconds")))
                    db.commit()
                    have.add(mid)
                added += 1
                log("  [ok] %s %-14s voy=%-8s legs=%d %s" % (
                    mid[:10], (vessel or "")[:14], voy.get("voyage_no"), len(legs),
                    "SAVED" if args.apply else "dry"))
        except Exception as e:
            log("  [warn] folder %d: %s" % (fi, str(e)[:60]))
        try:
            del sub
            del items
        except Exception:
            pass
        if fi % 20 == 0:
            gc.collect()

    db.close()
    log("DONE found=%d new=%d skipped=%d failed=%d (%.0fs) apply=%s" % (
        found, added, skipped, failed, time.time() - t0, args.apply))
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
