# -*- coding: utf-8 -*-
"""
导出 audit 里"判不出船(跳过)"的 NOON/BERTH 邮件明细(只读, 不解析 xlsx)。
并尝试用发件人邮箱前缀 / 主题原文 匹配船队 token, 看能救回多少。

输出 rob_data/unmatched_reports.csv 列:
  report_time, type, sender, guess_by_sender, guess_by_subject, subject, attachments

用法:
  python audit_unmatched.py
  AUDIT_CAP=300 python audit_unmatched.py
"""
import sys, os, csv, time, json
from collections import Counter, defaultdict
import win32com.client as win32
import rob_refresh as R

ROB_DIR = "rob_data"
RESULTS = os.path.join(ROB_DIR, "rob_results.json")
SENDERS = os.path.join(ROB_DIR, "vessel_senders.json")
OUT = os.path.join(ROB_DIR, "unmatched_reports.csv")
MAX_DEPTH = 4
CAP = int(os.environ.get("AUDIT_CAP", "0"))


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


def main():
    results = json.load(open(RESULTS, encoding="utf-8"))
    toks, names = [], []
    for r in results:
        v = r.get("vessel") or ""
        if not v:
            continue
        nv = norm(v)
        if len(nv) >= 4:
            toks.append((nv, v))
            names.append((nv, v))
        c = norm(r.get("code") or "")
        if len(c) >= 5:
            toks.append((c, v))
    toks.sort(key=lambda x: -len(x[0]))
    smap = {}
    for r in results:
        s = (r.get("sender") or "").strip().lower()
        if s and "@" in s:
            smap[s] = r["vessel"]
    if os.path.exists(SENDERS):
        for k, v in json.load(open(SENDERS, encoding="utf-8")).items():
            if isinstance(v, str) and "@" in v:
                smap.setdefault(v.strip().lower(), k)
    log("tokens=%d names=%d senders=%d" % (len(toks), len(names), len(smap)))

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    st = None
    for i in range(1, ns.Stores.Count + 1):
        s = ns.Stores(i)
        if "vesselreport" in (s.DisplayName or "").lower():
            st = s
            break
    root = st.GetRootFolder()
    folders = [root] + walk(root, [])
    log("folders=%d" % len(folders))
    filt = "[ReceivedTime] >= '01/01/2026 12:00 AM'"

    rows = []
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
                rt = it.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
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
            if kind is None:
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
                    sa = ""
            if vname:
                continue      # 已能判船, 不是目标
            # --- 判不出的: 记录 + 尝试从 sender / subject 猜 ---
            try:
                sender = (R.get_sender(it) or "").strip()
            except Exception:
                sender = ""
            local = sender.split("@")[0] if "@" in sender else ""
            nlocal = norm(local)
            guess_s = ""
            # sender 前缀(去 mv/m/v 等) 含船名 token
            for tok, vn in names:
                if tok and (tok in nlocal or nlocal in tok and len(nlocal) >= 5):
                    guess_s = vn
                    break
            if not guess_s and smap.get(sender.strip().lower()):
                guess_s = smap[sender.strip().lower()]
            # 主题原文: 去掉常见前缀后整体匹配船名(放宽 min len)
            guess_j = ""
            for tok, vn in names:
                if len(tok) >= 4 and tok in blob:
                    guess_j = vn
                    break
            rows.append([rt, kind, sender, guess_s, guess_j, subj, " | ".join(att)])
        if CAP and n > CAP:
            break
        try:
            del sub
        except Exception:
            pass

    log("")
    log("扫 %d 封, 判不出船 %d 封 (%.0fs)" % (n, len(rows), time.time() - t0))
    with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["report_time", "type", "sender", "guess_by_sender",
                    "guess_by_subject", "subject", "attachments"])
        for r in rows:
            w.writerow(r)
    log("导出: %s" % OUT)

    gs = sum(1 for r in rows if r[3])
    gj = sum(1 for r in rows if r[4])
    either = sum(1 for r in rows if r[3] or r[4])
    log("")
    log("sender 前缀可猜船: %d | subject 可猜船: %d | 任一可猜: %d | 仍无法: %d"
        % (gs, gj, either, len(rows) - either))
    log("按类型: %s" % dict(Counter(r[1] for r in rows)))
    # sender 分布 TOP
    log("")
    log("=== 发件人 TOP30 ===")
    for s, c in Counter((r[2] or "(空)").lower() for r in rows).most_common(30):
        log("  %-48s %d" % (s[:48], c))
    log("")
    log("=== 猜出船的样本(前15) ===")
    for r in [x for x in rows if x[3] or x[4]][:15]:
        log("  %s %s sender=%s -> by_sender=%s by_subj=%s | %s"
            % (r[0][:10], r[1], r[2][:32], r[3], r[4], r[5][:45]))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
