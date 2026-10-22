# -*- coding: utf-8 -*-
"""
ROB sync from ORIGINAL crew report emails (not the 1data.info summary).

Why this exists: the summary emails (servicesc-3@1data.info) miss a large share of
crew reports -- on 2026-10-22 the mailbox held 44 original report emails but the
summary path only refreshed 14 vessels. This script scans the mailbox directly,
matches report emails to fleet vessels, and parses ROB from the Excel attachments.

Design notes (learned the hard way):
  - Use Items.Restrict("[ReceivedTime] >= ...") instead of walking every item.
  - Only KEEP (folder, EntryID) for mails that could actually improve a vessel;
    holding thousands of COM objects triggers "out of memory or system resources".
  - Release COM objects + gc.collect() periodically.
  - Match by longest fleet token to avoid short-name false positives.

Usage:
  python rob_mail_sync.py --days 3            # dry run
  python rob_mail_sync.py --days 3 --apply    # write rob_results.json + rob_history.csv
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
import rob_refresh as R  # reuse norm / pick_report_attachment / extract_rob

BASE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(BASE, "rob_data")
RESULTS = os.path.join(ROB_DIR, "rob_results.json")
HISTORY_CSV = os.path.join(ROB_DIR, "rob_history.csv")
SNAP_FIELDS = R.SNAP_FIELDS

MAX_DEPTH = 4
PER_FOLDER_CAP = 1200
PER_VESSEL = 60      # max report mails kept per vessel in one run


def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def ensure_outlook():
    """Make sure Outlook is running and the MAPI cache is warm."""
    import subprocess
    import time
    import win32com.client

    try:
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
        try:
            ns.Logon("", "", False, False)
        except Exception:
            pass
        return app, ns
    except Exception:
        pass

    # Not available -> try to launch Outlook
    paths = [
        r"C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Root\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
    ]
    p = next((x for x in paths if os.path.exists(x)), None)
    if p:
        print("[outlook] not running, launching: %s" % p)
        subprocess.Popen([p])
        for _ in range(60):  # up to ~5 min
            time.sleep(5)
            try:
                app = win32com.client.Dispatch("Outlook.Application")
                ns = app.GetNamespace("MAPI")
                ns.Logon("", "", False, False)
                try:
                    if ns.ExchangeConnectionMode == 700:
                        print("[outlook] ready (mode=700)")
                        return app, ns
                except Exception:
                    pass
                print("[outlook] waiting for sync ...")
            except Exception:
                continue
        raise RuntimeError("Outlook did not become ready in time")
    raise RuntimeError("Outlook unavailable and OUTLOOK.EXE not found")


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


def build_fleet_tokens(results):
    """vessel name / code -> (token, vessel). Longest token wins on match."""
    toks = []
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
    return toks


def build_sender_map(results):
    """sender email -> vessel (from runtime history; richer than vessel_senders.json)."""
    m = {}
    for r in results:
        s = (r.get("sender") or "").strip().lower()
        if s and "@" in s:
            m[s] = r["vessel"]
    return m


def is_report_att(att_names):
    """Only parse Excel files that actually look like a crew report.

    Filters out terminal sailing reports (TSR.xls), load lists, DG declarations,
    booking reports etc. -- parsing those wastes minutes and never yields ROB.
    """
    up = [a.upper() for a in att_names]
    if any(k in a for k in R.REPORT_KEYS for a in up):
        return True
    return any("CUL" in a and "REPORT" in a for a in up)


def looks_like_report(att_names):
    return any("CUL" in a.upper() and any(k in a.upper() for k in R.REPORT_KEYS)
               for a in att_names)


def match_vessel(blob, toks):
    for tok, vname in toks:
        if tok and tok in blob:
            return vname
    return None


def collect_candidates(ns, days, results, toks, sender_map):
    """Scan recent mail; return {vessel: (recv_str, folder, entryID, subject)} newest only.

    Two match strategies:
      1. fleet name/code token inside subject or attachment name
      2. for reports with NO vessel name in the subject (e.g. "CUL NOON REPORT20261022"),
         resolve the sender's SMTP address and look it up in the runtime sender map
    """
    inbox = ns.GetDefaultFolder(6)
    folders = [inbox] + walk_folders(inbox)
    print("folders: %d" % len(folders))

    since = datetime.datetime.now() - datetime.timedelta(days=days)
    filt = "[ReceivedTime] >= '%s'" % since.strftime("%m/%d/%Y %H:%M %p")
    print("filter: %s" % filt)

    cur = {}
    for r in results:
        rt = (r.get("report_time") or "")[:19]
        if rt:
            cur[r["vessel"]] = rt

    best = {}     # vessel -> (recv_str, folder, entryID, subject)
    scanned = 0
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
                scanned += 1
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
                xl = [a for a in att if a.lower().endswith((".xlsx", ".xls", ".xlsm"))]
                if not xl:
                    continue
                if not is_report_att(xl):
                    continue
                blob = norm(subj) + " " + " ".join(norm(a) for a in xl)
                vname = match_vessel(blob, toks)
                via = "name"
                if not vname and looks_like_report(xl):
                    # report with no vessel name in subject -> resolve by sender
                    try:
                        sa = (R.get_sender(it) or "").strip().lower()
                    except Exception:
                        sa = ""
                    if sa:
                        vname = sender_map.get(sa)
                        via = "sender"
                if not vname:
                    continue
                # guard against shared mailboxes: if a *different* fleet vessel is
                # named in the subject, the sender match is not trustworthy
                if via == "sender":
                    other = [t for t, vn in toks if t in blob and vn != vname]
                    if other:
                        continue
                lst = best.setdefault(vname, [])
                if len(lst) >= PER_VESSEL:
                    continue
                try:
                    eid = it.EntryID
                except Exception:
                    continue
                lst.append((ts, f, eid, subj, via))
        except Exception as e:
            print("  [warn] folder %d: %s" % (fi, str(e)[:60]))
        try:
            del sub
            del items
        except Exception:
            pass
        if fi % 20 == 0:
            gc.collect()
        print("  scanned folders %d/%d (rows=%d, candidates=%d)"
              % (fi + 1, len(folders), scanned, len(best)), end="\r")
    print()
    return best


def parse_one(ns, folder, eid, retries=2):
    for attempt in range(retries):
        try:
            it = ns.GetItemFromID(eid)
        except Exception:
            it = None
        if it is None:
            try:
                it = ns.GetItemFromID(eid, folder.StoreID)
            except Exception:
                it = None
        if it is None:
            continue
        try:
            att = R.pick_report_attachment(it)
            if att is None:
                return None
            rob = R.extract_rob(att)
            if not any(k in rob for k in R.OIL_KEYS):
                return None
            try:
                subj = it.Subject or ""
            except Exception:
                subj = ""
            try:
                recv = it.ReceivedTime
            except Exception:
                return None
            try:
                se = R.get_sender(it)
            except Exception:
                se = ""
            return rob, recv, subj, se
        except Exception as e:
            print("    [parse retry %d] %s" % (attempt + 1, str(e)[:60]))
    return None


def append_history(results, newhist):
    """Append rows to rob_history.csv, idempotent on (vessel, report_time)."""
    existing = set()
    if os.path.exists(HISTORY_CSV):
        try:
            with open(HISTORY_CSV, encoding="utf-8-sig", newline="") as f:
                for row in csv.DictReader(f):
                    rt = (row.get("report_time") or "")[:19]
                    if rt:
                        existing.add((row.get("vessel"), rt))
        except Exception:
            pass
    rows = []
    for r in newhist:
        key = (r["vessel"], r["report_time"][:19])
        if key in existing:
            continue
        existing.add(key)
        rows.append(r)
    if not rows:
        return 0
    write_header = not os.path.exists(HISTORY_CSV) or os.path.getsize(HISTORY_CSV) == 0
    with open(HISTORY_CSV, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=SNAP_FIELDS, extrasaction="ignore")
        if write_header:
            w.writeheader()
        for r in rows:
            w.writerow(r)
    return len(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=3)
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    import pythoncom
    pythoncom.CoInitialize()

    results = json.load(open(RESULTS, encoding="utf-8"))
    byname = {r["vessel"]: r for r in results}
    toks = build_fleet_tokens(results)
    sender_map = build_sender_map(results)
    # merge the curated master->vessel map (vessel_senders.json), keyed by norm code
    try:
        vs = json.load(open(os.path.join(ROB_DIR, "vessel_senders.json"),
                            encoding="utf-8"))
        for k, v in vs.items():
            if isinstance(v, str) and "@" in v:
                sender_map.setdefault(v.strip().lower(), k)
    except Exception:
        pass
    print("fleet: %d vessels, %d match tokens, %d senders"
          % (len(results), len(toks), len(sender_map)))

    app, ns = ensure_outlook()
    cands = collect_candidates(ns, args.days, results, toks, sender_map)
    print("\ncandidate vessels with newer mail: %d" % len(cands))

    updates = []      # (vessel, old_rec, new_rec)
    newhist = []      # history rows (dedup happens at write time)
    nparsed = nhit = nnorob = 0

    for vname in sorted(cands):
        lst = sorted(cands[vname], key=lambda x: x[0], reverse=True)  # newest first
        rec = byname.get(vname)
        if rec is None:
            continue
        cur_rt = (rec.get("report_time") or "")[:19]
        for ts, folder, eid, subj, via in lst:
            got = parse_one(ns, folder, eid)
            nparsed += 1
            if got is None:
                nnorob += 1
                continue
            rob, recv, subj2, se = got
            new = recv.strftime("%Y-%m-%d %H:%M:%S")
            nhit += 1
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
                "bw": rob.get("BW"), "fw": rob.get("FW"),
                "refeer": rob.get("REFEER"),
                "found": "1", "report_time": new,
                "speed": rob.get("SPEED"), "report_type": rt,
            })
            if new <= cur_rt:
                continue   # older than what we already show -> history only
            # this one is the newest for the vessel (lst is newest-first)
            if any(u[0] == vname for u in updates):
                continue
            old = cur_rt[:16]
            print("  [HIT] %-22s %s -> %s (%s) | %s"
                  % (vname, old or "-", new[5:16], via, (subj2 or "")[:38]))
            rec2 = dict(rec)
            R.apply_hit(rec2, (rob, recv, subj2, se))
            # never blank out a value we already have
            for k in ("rob_lsfo", "rob_hsfo", "rob_mgo", "rob_ulsfo",
                      "rob_bw", "rob_fw", "rob_refeer", "rob_speed",
                      "rob_draft_fwd", "rob_draft_mid", "rob_draft_aft"):
                if rec2.get(k) is None and rec.get(k) is not None:
                    rec2[k] = rec[k]
            updates.append((vname, rec, rec2))

    print("\nparsed %d mails | with ROB: %d | no-rob: %d" % (nparsed, nhit, nnorob))
    print("would update %d vessels, add up to %d history rows"
          % (len(updates), len(newhist)))
    if not args.apply:
        print("(dry-run) add --apply to write")
        return

    for vname, old, new in updates:
        old.update(new)
    json.dump(results, open(RESULTS, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    n = append_history(results, newhist)
    print("wrote %d vessels, %d new history rows" % (len(updates), n))


if __name__ == "__main__":
    main()
