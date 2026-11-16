# -*- coding: utf-8 -*-
"""2026 早期(1-7 月)ROB 报告高效回补。

背景 / 为什么要单独写:
  - 1-7 月的 NOON/BERTH/SAILING ROB 报告**散在联机存档的中文收件箱顶层**(66023 封),
    没有被归入各船名文件夹; 主邮箱和存档的船名文件夹里 1-7 月只有 TDR/DG/STOWAGE,
    没有 ROB 报告。
  - 老的 backfill_history.py 逐封遍历文件夹 + Python 判断主题, 66023 封全表扫描极慢
    (单船 20 分钟跑不完)。harvest_reports.py 的 connect_outlook() 又故意排除存档。
  - 本脚本用 DASL @SQL "subject LIKE '%NOON%'" + 日期做 **server-side 过滤**(实测 0.9s),
    只把命中的 ~2000 封报告拿来解析附件, 命中率 ~70%, 每封 ~2s。

输出: rob_data/rob_reports_backfill.jsonl (行结构同 rob_reports.jsonl, voyage_consumption 直接消费)
      用 make_key(vessel, received, subject) 与主归档及自身去重, 可反复运行(--resume)。

用法:
  python backfill_early_reports.py --since 2026-01-01 --until 2026-07-27
  python backfill_early_reports.py --since 2026-01-01 --until 2026-07-27 --vessel CUHP
  python backfill_early_reports.py --dry-run
"""
import sys, os, json, time, argparse, hashlib
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")
import win32com.client as win32
import rob_refresh as R

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "rob_data", "rob_reports_backfill.jsonl")
MAIN_OUT = os.path.join(BASE, "rob_data", "rob_reports.jsonl")

REPORT_KWS = ("NOON", "BERTH", "SAILING")   # DASL 主题关键词
OIL_KEYS = R.OIL_KEYS


def make_key(vessel, recv, subject):
    raw = "%s|%s|%s" % (vessel, recv, subject or "")
    return hashlib.md5(raw.encode("utf-8", "replace")).hexdigest()[:20]


def load_done():
    """主归档 + 本次回补归档里已有的 key, 全部跳过。"""
    done = set()
    for p in (MAIN_OUT, OUT):
        if not os.path.exists(p):
            continue
        for line in open(p, encoding="utf-8"):
            line = line.strip()
            if not line:
                continue
            try:
                done.add(json.loads(line).get("key", ""))
            except Exception:
                pass
    return done


def largest_inboxes():
    """每个 CULINES store 选 Items.Count 最大的收件箱(存档有英文 Inbox 空壳 + 中文收件箱)。"""
    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    out = []
    for st in ns.Stores:
        if "CULINES.COM" not in (st.DisplayName or "").upper():
            continue
        try:
            root = st.GetRootFolder()
            best, best_n = None, -1
            for f in root.Folders:
                fn = (f.Name or "").strip()
                if fn in ("收件箱", "Inbox", "INBOX") or fn.upper() == "INBOX":
                    try:
                        cnt = f.Items.Count
                    except Exception:
                        cnt = 0
                    if cnt > best_n:
                        best, best_n = f, cnt
            if best is not None:
                out.append((st.DisplayName, best))
                print("[store] %s -> %s (%d)" % (st.DisplayName, best.Name, best_n), flush=True)
        except Exception as e:
            print("[store] skip %s: %s" % (st.DisplayName, e), flush=True)
    return out


def match_vessel(subject, fleet_idx, sender=""):
    """按【主题】或【发件人邮箱】匹配船队。返回 (vessel, code) 或 None。

    早期(1-7 月)大量 NOON 报告主题只有 'CUL NOON REPORT', 船名不在主题里,
    必须靠发件人识别: m.odyssey@canopus-mail.com -> M. ODYSSEY;
    m.mariner@... -> M. MARINER; hdxhongyuan@msatmail.com -> HDX HONGYUAN(已下线, 不收录)。
    2026-11-16 加 sender 参数。"""
    hits = []
    # 1) 主题 token
    sn = R.norm(subject or "")
    if sn:
        for token, (vessel, code) in fleet_idx.items():
            if len(token) >= 4 and token in sn:
                hits.append((vessel, code, len(token)))
    # 2) 发件人邮箱 local part token (m.odyssey -> modyssey; captain.cuhp -> captaincuhp 含 cuhp)
    if sender:
        local = str(sender).split("@")[0] if "@" in str(sender) else str(sender)
        sl = R.norm(local)
        if sl:
            for token, (vessel, code) in fleet_idx.items():
                if len(token) >= 4 and token in sl:
                    hits.append((vessel, code, len(token)))
    if not hits:
        return None
    # 最长 token 命中优先(避免 CUHP 前缀误配), 同长度主题命中优先
    hits.sort(key=lambda x: -x[2])
    return hits[0][0], hits[0][1]


def build_fleet_idx(vessel_filter=""):
    fleet = R.load_fleet()
    if vessel_filter:
        k = vessel_filter.upper()
        fleet = [r for r in fleet if k in r["vessel"].upper() or k == (r.get("code") or "").upper()]
    idx = {}
    for r in fleet:
        v, c = r["vessel"], (r.get("code") or "")
        for token in (R.norm(v), R.norm(c)):
            if token and len(token) >= 4:
                idx[token] = (v, c)
    print("fleet=%d tokens=%d filter=%r" % (len(fleet), len(idx), vessel_filter), flush=True)
    return fleet, idx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", default="2026-01-01")
    ap.add_argument("--until", default="2026-07-27", help="上界(不含)")
    ap.add_argument("--vessel", default="", help="只抓某船/代码(调试)")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--max-per-kw", type=int, default=4000, help="每个关键词最多解析多少封")
    args = ap.parse_args()

    done = load_done()
    print("done keys: %d" % len(done), flush=True)
    fleet, fleet_idx = build_fleet_idx(args.vessel)

    since_s = args.since.replace("-", "/")
    until_s = args.until.replace("-", "/")
    boxes = largest_inboxes()

    fout = None if args.dry_run else open(OUT, "a", encoding="utf-8")
    n_new = 0
    t_start = time.time()
    try:
        for sname, inbox in boxes:
            # 存档船名文件夹 1-7 月没有 ROB 报告, 只扫收件箱顶层; 主邮箱顶层 1-7 月为 0, 跳过也行
            for kw in REPORT_KWS:
                sql = ("@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%%%s%%') AND "
                       "(\"urn:schemas:httpmail:datereceived\" >= '%s') AND "
                       "(\"urn:schemas:httpmail:datereceived\" < '%s')"
                       % (kw, since_s, until_s))
                try:
                    items = inbox.Items.Restrict(sql)
                    total = items.Count
                except Exception as e:
                    print("  [%s/%s] Restrict ERR %s" % (sname, kw, e), flush=True)
                    continue
                print("[%s] %s: %d candidates" % (sname[:12], kw, total), flush=True)
                if total == 0:
                    continue
                it = items.GetFirst()
                n = 0
                nrob = 0
                skipped = 0
                kw_t0 = time.time()
                while it is not None and n < args.max_per_kw:
                    n += 1
                    try:
                        recv = str(it.ReceivedTime)[:19]
                        subj = it.Subject or ""
                    except Exception:
                        it = items.GetNext()
                        continue
                    try:
                        sender = R.get_sender(it)
                    except Exception:
                        sender = ""
                    vm = match_vessel(subj, fleet_idx, sender)
                    if vm is None:
                        it = items.GetNext()
                        continue
                    vessel, code = vm
                    key = make_key(vessel, recv, subj)
                    if key in done:
                        skipped += 1
                        it = items.GetNext()
                        continue
                    try:
                        att = R.pick_report_attachment(it)
                    except Exception:
                        att = None
                    if att is None:
                        it = items.GetNext()
                        continue
                    try:
                        aname = att.FileName or ""
                        rob = R.extract_rob(att)
                    except Exception:
                        it = items.GetNext()
                        continue
                    if not any(k in rob for k in OIL_KEYS):
                        it = items.GetNext()
                        continue
                    kind = (rob.get("REPORT_KIND") or "").strip().upper()
                    if not kind:
                        su = subj.upper()
                        for k in R.REPORT_KEYS:
                            if k in su:
                                kind = k
                                break
                    row = {
                        "key": key, "vessel": vessel, "code": code,
                        "received": recv, "subject": subj,
                        "folder": "TOPLEVEL", "attach": aname,
                        "report_type": kind, "kind_src": "body/file",
                        "lsfo": rob.get("LSFO"), "hsfo": rob.get("HSFO"),
                        "mgo": rob.get("MGO"), "ulsfo": rob.get("ULSFO"),
                        "bw": rob.get("BW"), "fw": rob.get("FW"),
                        "refeer": rob.get("REEFER"), "speed": rob.get("SPEED"),
                    }
                    if fout is not None:
                        fout.write(json.dumps(row, ensure_ascii=False) + "\n")
                        fout.flush()
                    done.add(key)
                    n_new += 1
                    nrob += 1
                    try:
                        it = items.GetNext()
                    except Exception as e:
                        print("      [WARN] GetNext 中断(RPC?): %s" % str(e)[:60], flush=True)
                        break   # 本关键词剩余跳过, 不拖垮整个脚本
                print("      解析 %d 封, 命中 %d, 去重跳过 %d, 新增 %d, 耗时 %.0fs"
                      % (n, nrob, skipped, n_new, time.time() - kw_t0), flush=True)
    finally:
        if fout is not None:
            fout.close()
    print("DONE new=%d total=%.0fs -> %s" % (n_new, time.time() - t_start, OUT), flush=True)


if __name__ == "__main__":
    main()
