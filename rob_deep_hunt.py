# -*- coding: utf-8 -*-
"""
Deep hunt for vessels whose ROB is stale.

Unlike rob_mail_sync.py (which only walks the inbox subtree, depth<=4), this walks
EVERY store in the profile -- including the online archive / PST files -- at greater
depth. Used to close the gap when vessels stay stale after the normal sync.

Usage:
  python rob_deep_hunt.py --days 10             # dry run, report what is found
  python rob_deep_hunt.py --days 10 --apply     # write results + history
"""
import argparse
import csv
import datetime
import gc
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R
import rob_mail_sync as M

BASE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(BASE, "rob_data")
RESULTS = os.path.join(ROB_DIR, "rob_results.json")
HISTORY_CSV = os.path.join(ROB_DIR, "rob_history.csv")

MAX_DEPTH = 6
PER_FOLDER_CAP = 1200
PER_VESSEL = 40
CACHE = os.path.join(ROB_DIR, "_deep_hunt_cache.json")


def all_stores(ns):
    """Every top-level store root in the profile."""
    out = []
    try:
        for s in ns.Folders:
            try:
                out.append((s.Name, s))
            except Exception:
                continue
    except Exception:
        pass
    return out


def walk(f, depth=0, out=None, limit=600):
    if out is None:
        out = []
    if depth > MAX_DEPTH or len(out) > limit:
        return out
    try:
        for c in f.Folders:
            out.append(c)
            walk(c, depth + 1, out, limit)
    except Exception:
        pass
    return out


def scan(ns, days, toks, sender_map, wantset):
    """Walk EVERY store (incl. online archive) and collect candidate report mails."""
    since = datetime.datetime.now() - datetime.timedelta(days=days)
    filt = "[ReceivedTime] >= '%s'" % since.strftime("%m/%d/%Y %H:%M %p")
    print("filter: %s" % filt)

    best = {}          # vessel -> [(ts, store_id, entry_id, subject, via)]
    total_folders = 0
    for sname, store in all_stores(ns):
        folders = [store] + walk(store)
        total_folders += len(folders)
        print("  store '%s': %d folders" % (sname, len(folders)))
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
                        t = it.ReceivedTime
                    except Exception:
                        continue
                    ts = t.strftime("%Y-%m-%d %H:%M:%S")
                    try:
                        subj = it.Subject or ""
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
                    xl = [a for a in att
                          if a.lower().endswith((".xlsx", ".xls", ".xlsm"))]
                    if not xl or not M.is_report_att(xl):
                        continue
                    blob = M.norm(subj) + " " + " ".join(M.norm(a) for a in xl)
                    vname = M.match_vessel(blob, toks)
                    via = "name"
                    if not vname and M.looks_like_report(xl):
                        try:
                            sa = (R.get_sender(it) or "").strip().lower()
                        except Exception:
                            sa = ""
                        if sa:
                            vname = sender_map.get(sa)
                            via = "sender"
                            if vname:
                                # shared mailbox guard
                                if [t for t, vn in toks if t in blob and vn != vname]:
                                    vname = None
                    if not vname or vname not in wantset:
                        continue
                    lst = best.setdefault(vname, [])
                    if len(lst) >= PER_VESSEL:
                        continue
                    try:
                        eid = it.EntryID
                        sid = store.StoreID
                    except Exception:
                        continue
                    lst.append((ts, sid, eid, subj, via))
            except Exception:
                pass
            try:
                del sub
                del items
            except Exception:
                pass
            if fi % 30 == 0:
                gc.collect()
    print("\ntotal folders scanned: %d" % total_folders)
    return best


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=10)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--vessels", default="", help="comma separated vessel names; default = all stale")
    ap.add_argument("--use-cache", action="store_true",
                    help="reuse the candidate list from the last scan (skip the 15-min walk)")
    args = ap.parse_args()

    import pythoncom
    pythoncom.CoInitialize()

    results = json.load(open(RESULTS, encoding="utf-8"))
    byname = {r["vessel"]: r for r in results}
    toks = M.build_fleet_tokens(results)
    sender_map = M.build_sender_map(results)

    if args.vessels:
        want = [v.strip() for v in args.vessels.split(",") if v.strip()]
    else:
        today = datetime.date.today().isoformat()
        want = [r["vessel"] for r in results
                if (r.get("report_time") or "")[:10] < today]
    print("target vessels: %d" % len(want))
    wantset = set(want)

    app, ns = M.ensure_outlook()

    cur = {r["vessel"]: (r.get("report_time") or "")[:19] for r in results}
    best = {}          # vessel -> [(ts, store_id, entry_id, subject, via)]

    if args.use_cache and os.path.exists(CACHE):
        try:
            best = json.load(open(CACHE, encoding="utf-8"))
            print("cache hit: %d vessels from %s (skipping the walk)"
                  % (len(best), os.path.basename(CACHE)))
        except Exception as e:
            print("[WARN] cache load failed: %s" % e)
            best = {}

    if not best:
        best = scan(ns, args.days, toks, sender_map, wantset)
        try:
            json.dump(best, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False)
            print("candidate list cached -> %s" % CACHE)
        except Exception as e:
            print("[WARN] cache write failed: %s" % e)

    print("vessels with candidate mail: %d / %d" % (len(best), len(wantset)))
    for v in sorted(wantset):
        if v not in best:
            print("   [none] %s (still %s)" % (v, cur.get(v) or "MISS"))
    updates = []
    newhist = []
    for vname, lst in best.items():
        lst.sort(key=lambda x: x[0], reverse=True)
        rec = byname.get(vname)
        if rec is None:
            continue
        cur_rt = cur.get(vname, "")
        for ts, sid, eid, subj, via in lst:
            got = None
            for _ in range(2):
                try:
                    it = ns.GetItemFromID(eid, sid)
                except Exception:
                    it = None
                if it is None:
                    try:
                        it = ns.GetItemFromID(eid)
                    except Exception:
                        continue
                try:
                    att = R.pick_report_attachment(it)
                    if att is None:
                        break
                    rob = R.extract_rob(att)
                    if not any(k in rob for k in R.OIL_KEYS):
                        break
                    got = (rob, it.ReceivedTime, it.Subject or "", R.get_sender(it))
                    break
                except Exception:
                    continue
            if got is None:
                continue
            rob, recv, subj2, se = got
            new = recv.strftime("%Y-%m-%d %H:%M:%S")
            rt = (rob.get("REPORT_KIND") or "").strip().upper()
            if not rt:
                subjU = (subj2 or "").upper()
                for k in R.REPORT_KEYS:
                    if k in subjU:
                        rt = k
                        break
            newhist.append({
                "date": new[:10], "vessel": vname, "code": rec.get("code"),
                "lane": rec.get("lane"), "pic": rec.get("pic"),
                "lsfo": rob.get("LSFO"), "hsfo": rob.get("HSFO"),
                "mgo": rob.get("MGO"), "ulsfo": rob.get("ULSFO"),
                "bw": rob.get("BW"), "fw": rob.get("FW"), "refeer": rob.get("REFEER"),
                "found": "1", "report_time": new, "speed": rob.get("SPEED"),
                "report_type": rt,
            })
            if new > cur_rt and not any(u[0] == vname for u in updates):
                print("  [HIT] %-22s %s -> %s (%s) | %s"
                      % (vname, cur_rt[5:16] or "-", new[5:16], via, (subj2 or "")[:40]))
                rec2 = dict(rec)
                R.apply_hit(rec2, (rob, recv, subj2, se))
                for k in ("rob_lsfo", "rob_hsfo", "rob_mgo", "rob_ulsfo",
                          "rob_bw", "rob_fw", "rob_refeer", "rob_speed",
                          "rob_draft_fwd", "rob_draft_mid", "rob_draft_aft"):
                    if rec2.get(k) is None and rec.get(k) is not None:
                        rec2[k] = rec[k]
                updates.append((vname, rec, rec2))

    print("\nwould update %d vessels, add up to %d history rows"
          % (len(updates), len(newhist)))
    if not args.apply:
        print("(dry-run) add --apply to write")
        return

    for vname, old, new in updates:
        old.update(new)
    json.dump(results, open(RESULTS, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    n = M.append_history(results, newhist)
    print("wrote %d vessels, %d new history rows" % (len(updates), n))


if __name__ == "__main__":
    main()
