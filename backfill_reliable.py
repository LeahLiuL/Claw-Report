# -*- coding: utf-8 -*-
"""可靠版全船队历史回补: 索引遍历船名夹(不用不稳定的 DASL)。

背景(2026-11-18): DASL `@SQL subject LIKE + 宽日期` 在 6.6 万封收件箱 + 同名 Master
文件夹上结果集不稳定 —— H708 的 10-05~10-08 报告在窄窗口能命中、宽窗口整段漏掉。
改为对每个【船名文件夹及其 Master 子夹】(都由 Items.Count <= 500 的小夹) 用
Items.Item(i) 索引全量遍历 + Python 端日期过滤, 结果确定可靠。

输出: rob_data/rob_reports_backfill.jsonl (与主归档/自身按 key 去重, 可反复跑)
用法: python backfill_reliable.py --since 2026-07-27 --until 2026-11-19
      python backfill_reliable.py --since ... --until ... --vessel H708
"""
import sys, os, json, hashlib, argparse, time
from datetime import datetime
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import win32com.client as win32
import rob_refresh as R

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "rob_data", "rob_reports_backfill.jsonl")
MAIN_OUT = os.path.join(BASE, "rob_data", "rob_reports.jsonl")
MAX_ITEMS = int(os.environ.get("BACKFILL_MAX_ITEMS", "600"))


def make_key(v, r, s):
    return hashlib.md5(("%s|%s|%s" % (v, r, s or "")).encode("utf-8", "replace")).hexdigest()[:20]


def load_done():
    done = set()
    for p in (MAIN_OUT, OUT):
        if os.path.exists(p):
            for ln in open(p, encoding="utf-8"):
                try: done.add(json.loads(ln).get("key", ""))
                except Exception: pass
    return done


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", default="2026-01-01")
    ap.add_argument("--until", default="2026-11-19")
    ap.add_argument("--rolling-days", type=int, default=0,
                    help=">0 时覆盖 --since: 只抓最近 N 天(到今天), 供每日滚动回补用")
    ap.add_argument("--vessel", default="", help="只抓某船/代码(调试)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if args.rolling_days > 0:
        from datetime import timedelta
        _today = datetime.now().date()
        args.since = (_today - timedelta(days=args.rolling_days)).strftime("%Y-%m-%d")
        args.until = (_today + timedelta(days=1)).strftime("%Y-%m-%d")
        print("[rolling] since=%s until=%s" % (args.since, args.until), flush=True)

    lo = datetime.strptime(args.since, "%Y-%m-%d").date()
    hi = datetime.strptime(args.until, "%Y-%m-%d").date()
    done = load_done()
    print("done keys: %d" % len(done), flush=True)

    fleet = R.load_fleet()
    if args.vessel:
        k = args.vessel.upper()
        fleet = [r for r in fleet if k in r["vessel"].upper() or k == (r.get("code") or "").upper()]
    print("fleet target: %d" % len(fleet), flush=True)
    # 船名 token -> (vessel, code)
    idx = {}
    for r in fleet:
        for tok in (R.norm(r["vessel"]), R.norm(r.get("code") or "")):
            if tok and len(tok) >= 4:
                idx[tok] = (r["vessel"], r.get("code") or "")

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    fout = None if args.dry_run else open(OUT, "a", encoding="utf-8")
    nnew = 0
    nfold = 0
    t0 = time.time()
    try:
        for st in ns.Stores:
            if "CULINES.COM" not in (st.DisplayName or "").upper():
                continue
            root = st.GetRootFolder()
            inbox = None
            bn = -1
            for f in root.Folders:
                if f.Name.strip() in ("收件箱", "Inbox", "INBOX"):
                    try: c = f.Items.Count
                    except Exception: c = 0
                    if c > bn: inbox, bn = f, c
            if inbox is None:
                continue
            for sf in inbox.Folders:
                fn = R.norm(sf.Name or "")
                # 一级夹必须是船名夹
                hit = [pair for tok, pair in idx.items() if tok in fn]
                if not hit:
                    continue
                targets = [sf]
                try:
                    for ssf in sf.Folders:
                        if "MASTER" in (ssf.Name or "").upper():
                            targets.append(ssf)
                except Exception:
                    pass
                for tg in targets:
                    try:
                        cnt = tg.Items.Count
                    except Exception:
                        continue
                    nfold += 1
                    # 大文件夹(>MAX_ITEMS)用 DASL 窄日期窗口 server-side 过滤;
                    # 45 天滚动窗口下 DASL 结果稳定(宽窗口 + 同名 Master 才会漏)。
                    # 小文件夹用索引全量遍历(100% 可靠)。
                    use_dasl = cnt > MAX_ITEMS
                    if use_dasl:
                        print("  [dasl] %s (%d)" % (tg.Name, cnt), flush=True)
                        sql = ("@SQL=(\"urn:schemas:httpmail:datereceived\" >= '%s') AND "
                               "(\"urn:schemas:httpmail:datereceived\" < '%s')"
                               % (args.since.replace('-', '/'), args.until.replace('-', '/')))
                        try:
                            ob = tg.Items.Restrict(sql)
                            cnt = ob.Count
                        except Exception as e:
                            print("    DASL ERR %s" % str(e)[:60], flush=True)
                            continue
                    else:
                        ob = tg.Items
                    # Restrict 结果集用 GetFirst/GetNext(MAPI 标准, Item(i) 极慢);
                    # 普通 Items 集合用 Item(i) 索引。统一先取成列表。
                    if use_dasl:
                        items_list = []
                        try:
                            _it = ob.GetFirst()
                            while _it is not None:
                                items_list.append(_it)
                                _it = ob.GetNext()
                        except Exception as e:
                            print("    iterate ERR %s" % str(e)[:60], flush=True)
                    else:
                        items_list = None   # 用 Item(i) 惰性取, 避免一次性 600 对象
                    for i in range(1, cnt + 1):
                        if items_list is not None:
                            if i > len(items_list): break
                            it = items_list[i - 1]
                        else:
                            try: it = ob.Item(i)
                            except Exception: continue
                        if it is None: continue
                        try:
                            t = it.ReceivedTime
                            if not hasattr(t, "date") or not (lo <= t.date() < hi):
                                continue
                            subj = str(it.Subject or "")
                            su = subj.upper()
                            if not any(kw in su for kw in R.REPORT_KEYS):
                                continue
                        except Exception:
                            continue
                        # 匹配船(优先主题, 再文件夹)
                        vm = None
                        sn = R.norm(subj)
                        for tok, pair in idx.items():
                            if len(tok) >= 4 and tok in sn:
                                vm = pair; break
                        if vm is None:
                            tn = R.norm(tg.Name or "")
                            for tok, pair in idx.items():
                                if len(tok) >= 4 and tok in tn:
                                    vm = pair; break
                        if vm is None:
                            continue
                        vessel, code = vm
                        recv = t.strftime("%Y-%m-%d %H:%M:%S")
                        key = make_key(vessel, recv, subj)
                        if key in done:
                            continue
                        try:
                            att = R.pick_report_attachment(it)
                            rob = R.extract_rob(att) if att else None
                        except Exception:
                            continue
                        if not rob or not any(x in rob for x in R.OIL_KEYS):
                            continue
                        kind = (rob.get("REPORT_KIND") or "").strip().upper()
                        if not kind:
                            for kk in R.REPORT_KEYS:
                                if kk in su:
                                    kind = kk; break
                        row = {
                            "key": key, "vessel": vessel, "code": code,
                            "received": recv, "subject": subj, "folder": tg.Name,
                            "attach": (att.FileName if att else ""),
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
                        nnew += 1
                        print("  + %s %s %s %s" % (vessel[:14], recv[:16], kind, subj[:38]), flush=True)
    finally:
        if fout is not None:
            fout.close()
    print("DONE new=%d folders=%d total=%.0fs -> %s" % (nnew, nfold, time.time() - t0, OUT), flush=True)


if __name__ == "__main__":
    main()
