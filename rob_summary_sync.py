# -*- coding: utf-8 -*-
"""ROB 快速同步: 解析『[中联航运-船舶管理报告]-报表汇总』邮件(每天 09:00 / 15:00 各一封),
直接取每艘船的 ROB, 不再全邮箱扫描 51 个船长的来信。

背景
----
旧链路 rob_refresh.py 用 Outlook COM 遍历 14 天 x 上百个文件夹 x 逐封解析附件,
单次 8~13 分钟, 且 Outlook 必须开着、缓存必须同步完。
新链路: 该汇总邮件是系统侧已经解析好的成品(按 Noon/Berth/Sailing/Anchor/
Bunker Change-over 分表, 每行带 ROB LSFO/MGO/HSFO/ULSFO + Email Received Time),
抓 N 封汇总邮件(每天 2 封)即可覆盖 N/2 天的全船队, 耗时秒级。

用法
----
  python rob_summary_sync.py                  # 干跑: 只打印会更新什么, 不写盘
  python rob_summary_sync.py --days 14        # 回看 14 天(默认 10)
  python rob_summary_sync.py --apply          # 真正写入 rob_results.json + rob_history.csv

安全: 默认不写任何文件; --apply 才写, 且只在新报告时间 > 现有时间时覆盖。
"""
import os
import re
import csv
import sys
import json
import argparse
import datetime
import tempfile
import time

BASE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(BASE, "rob_data")
RESULTS = os.path.join(ROB_DIR, "rob_results.json")
HISTORY_CSV = os.path.join(ROB_DIR, "rob_history.csv")
SENDER_MAP_FILE = os.path.join(BASE, "vessel_senders.json")

SNAP_FIELDS = ["date", "vessel", "code", "lane", "pic",
               "lsfo", "hsfo", "mgo", "ulsfo", "bw", "fw", "refeer",
               "found", "report_time", "speed", "report_type"]

SUBJECT_KEY = "报表汇总"          # 汇总邮件主题特征(中联航运-船舶管理报告)

# 含 ROB 的分表 -> report_type
ROB_SHEETS = {
    "Sailing Report": "Sailing",
    "Berth Report": "Berth",
    "Anchor Drifting Report": "Anchor",
    "Resume Sailing Report": "Resume Sailing",
    "Bunker Fuel Change-over Report": "Bunker Change-over",
}
OILS = ("LSFO", "HSFO", "MGO", "ULSFO")


def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


# ---------------------------------------------------------------- 取汇总邮件
OUTLOOK_EXE = r"C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE"


def _outlook_alive():
    """Outlook 进程是否已在跑(用 GetActiveObject, 不会把没开的 Outlook 拉起来)。"""
    import win32com.client
    try:
        win32com.client.GetActiveObject("Outlook.Application")
        return True
    except Exception:
        return False


def ensure_outlook(max_wait=300, minimize=True):
    """保证 Outlook 可用: 没开就拉起, 并等 Exchange 缓存同步到可用为止。

    为什么要等: Outlook 刚启动时"进程在、能连 COM", 但缓存还没同步完,
    此时 Restrict 只返回已同步的那一小部分 —— 刷新会照常跑完却抓不到新报告,
    且看不出任何异常(2026-09-22 / 10-09 都栽在这里)。
    """
    import time
    import subprocess
    import win32com.client

    started = False
    if not _outlook_alive():
        print("[outlook] not running -> launching")
        try:
            os.startfile("outlook")           # 走 App Paths 注册, 最稳
        except Exception:
            subprocess.Popen([OUTLOOK_EXE])
        started = True
    else:
        print("[outlook] already running")

    app = None
    t0 = time.time()
    mode = None
    n = 0
    while time.time() - t0 < max_wait:
        try:
            app = win32com.client.Dispatch("Outlook.Application")
            ns = app.GetNamespace("MAPI")
            mode = None
            try:
                mode = ns.ExchangeConnectionMode
            except Exception:
                pass
            if mode:                       # 0 = 还没建立 Exchange 连接
                try:
                    n = ns.GetDefaultFolder(6).Items.Count
                except Exception:
                    n = 0
                if n > 200:                # 收件箱有货 = 缓存可用
                    break
        except Exception:
            pass
        time.sleep(5)

    print("[outlook] ready: mode=%s inbox=%d (%.0fs)" % (mode, n, time.time() - t0))
    if mode and mode < 700:
        print("[outlook][WARN] mode=%s < 700, 缓存可能还没全同步, 结果可能不全" % mode)

    if started and minimize and app is not None:
        try:
            app.ActiveExplorer().WindowState = 1      # 1 = olMinimized
        except Exception:
            pass
    return app


def find_summary_emails(days):
    """返回 [(received_datetime, mail_item, subject)] 按时间升序。"""
    import win32com.client
    ensure_outlook()
    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    since = datetime.datetime.now() - datetime.timedelta(days=days)
    cutoff = since.strftime("%m/%d/%Y %H:%M %p")

    out = []
    # 主邮箱 + 联机存档都找(更早的汇总邮件会归档到存档 store)
    stores = [s for s in ns.Stores]
    for st in stores:
        try:
            inbox = st.GetDefaultFolder(6)
        except Exception:
            continue
        items = None
        # 先试 DASL 主题过滤(快); 失败/不支持则退回按时间 Restrict 后在 Python 侧过滤
        try:
            dasl = ('@SQL=' + '"http://schemas.microsoft.com/mapi/proptag/0x0037001f"'
                    + " LIKE '%" + SUBJECT_KEY + "%'"
                    + " AND \"urn:schemas:httpmail:datereceived\" > '"
                    + since.strftime("%Y-%m-%dT%H:%M:%S") + "'")
            items = inbox.Items.Restrict(dasl)
            _ = items.Count
        except Exception:
            items = None
        if items is None:
            try:
                items = inbox.Items.Restrict("[ReceivedTime] >= '%s'" % cutoff)
            except Exception:
                continue
        try:
            for m in items:
                try:
                    subj = m.Subject or ""
                except Exception:
                    continue
                if SUBJECT_KEY not in subj:
                    continue
                try:
                    rt = m.ReceivedTime
                    rt = datetime.datetime(rt.year, rt.month, rt.day,
                                           rt.hour, rt.minute, rt.second)
                except Exception:
                    continue
                out.append((rt, m, subj))
        except Exception as e:
            print("  [WARN] iterate %s failed: %s" % (st.DisplayName, e))
    # 硬过滤: DASL 的日期条件在部分 Outlook 版本上不生效, 这里再在 Python 侧兜一次,
    # 否则会把几十天前的汇总邮件也捞进来(结果一样但白下载附件、变慢)
    cutoff_dt = datetime.datetime.now() - datetime.timedelta(days=days)
    out = [x for x in out if x[0] >= cutoff_dt]
    out.sort(key=lambda x: x[0])
    return out


# ---------------------------------------------------------------- 解析
def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None   # 容错: 源表偶发把表头当值写进单元格(如 'ROB LSFO')


def _parse_recv(s):
    """'2026/10/22 12:26:45,215' -> datetime"""
    if s is None:
        return None
    if isinstance(s, datetime.datetime):
        return s.replace(microsecond=0)
    t = str(s).strip()
    m = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?", t)
    if not m:
        return None
    g = [int(x) if x else 0 for x in m.groups()]
    try:
        return datetime.datetime(g[0], g[1], g[2], g[3], g[4], g[5])
    except Exception:
        return None


def _pick_rob_cols(hdr):
    """ROB 列名不统一(ROB LSFO / FWE ROB LSFO / COMMENCE ROB LSFO / COMPLETION ROB LSFO),
    按优先级挑: COMPLETION > FWE > 裸 ROB > COMMENCE。"""
    up = [str(h or "").upper() for h in hdr]
    cols = {}
    for oil in OILS:
        cands = [i for i, h in enumerate(up) if "ROB" in h and oil in h]
        if not cands:
            continue
        def rank(i):
            h = up[i]
            if "COMPLETION" in h or "COMPLETE" in h:
                return 0
            if "FWE" in h:
                return 1
            if "COMMENCE" in h or "START" in h:
                return 3
            return 2
        cols[oil] = sorted(cands, key=rank)[0]
    return cols


def parse_summary_xlsx(path):
    """-> [(code, lane, voyage, report_type, rob_dict, recv_dt, date_str)]"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    rows_out = []
    zero_skipped = []
    for sn, rtype in ROB_SHEETS.items():
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        allrows = list(ws.iter_rows(values_only=True))
        if len(allrows) < 2:
            continue
        hdr = [str(h or "").strip() for h in allrows[0]]
        up = [h.upper() for h in hdr]
        try:
            iv = up.index("VESSEL NAME")
        except ValueError:
            continue
        ilane = up.index("LANE NAME") if "LANE NAME" in up else None
        ivoy = up.index("VOYAGE NO.") if "VOYAGE NO." in up else (
            up.index("VOYAGE") if "VOYAGE" in up else None)
        idate = up.index("DATE") if "DATE" in up else None
        irecv = up.index("EMAIL RECEIVED TIME") if "EMAIL RECEIVED TIME" in up else None
        istat = up.index("STATUS") if "STATUS" in up else None
        robcols = _pick_rob_cols(hdr)
        for r in allrows[1:]:
            code = (r[iv] or "")
            code = str(code).strip().upper()
            if not code:
                continue
            if istat is not None:
                st = str(r[istat] or "").strip()
                if st and st not in ("成功", "OK", "Success"):
                    continue          # 解析失败的行不采信
            recv = _parse_recv(r[irecv]) if irecv is not None else None
            d = r[idate] if idate is not None else None
            if isinstance(d, datetime.datetime):
                dstr = d.strftime("%Y-%m-%d")
            elif d:
                dstr = str(d)[:10]
            else:
                dstr = (recv.strftime("%Y-%m-%d") if recv else "")
            rob = {}
            for oil in OILS:
                i = robcols.get(oil)
                rob[oil] = _num(r[i]) if i is not None else None
            if all(v is None for v in rob.values()):
                continue
            # 保护1: 四条油全为 0 = 船长没填(源表照填 0), 采信会污染页面
            if all(v == 0 for v in rob.values()):
                zero_skipped.append((code, sn))
                continue
            rows_out.append({
                "code": code,
                "lane": (str(r[ilane] or "").strip() if ilane is not None else ""),
                "voyage": (str(r[ivoy] or "").strip() if ivoy is not None else ""),
                "report_type": rtype,
                "rob": rob,
                "recv": recv,
                "date": dstr,
            })
    return rows_out, zero_skipped


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=10, help="回看天数(每天 2 封汇总邮件)")
    ap.add_argument("--apply", action="store_true", help="真正写盘(默认干跑)")
    args = ap.parse_args()

    results = json.load(open(RESULTS, encoding="utf-8")) if os.path.exists(RESULTS) else []
    by_code = {}
    by_name = {}
    for r in results:
        if r.get("code"):
            by_code[str(r["code"]).strip().upper()] = r
        by_name[norm(r["vessel"])] = r

    print("existing results: %d vessels (code map %d)" % (len(results), len(by_code)))

    mails = find_summary_emails(args.days)
    # 同步完整性校验: 每天应有 2 封(09:00 / 15:00)。明显偏少 = 缓存还没同步完,
    # 等一会儿重扫(状态能自愈, 别一次定生死 —— 2026-10-09 就因为只扫一次漏了 3 艘)
    expect = args.days * 2
    for _try in range(2):
        if len(mails) >= max(2, int(expect * 0.6)):
            break
        print("[WARN] only %d summary emails (expect ~%d) - cache may still be "
              "syncing; wait 60s and rescan (%d/2)" % (len(mails), expect, _try + 1))
        time.sleep(60)
        mails = find_summary_emails(args.days)
    print("summary emails found (last %dd): %d" % (args.days, len(mails)))
    if not mails:
        print("nothing to do")
        return

    tmpd = tempfile.mkdtemp(prefix="rob_sum_")
    latest = {}       # code -> record
    unknown = set()
    zero_all = []
    nrows = 0
    for rt, m, subj in mails:
        got = None
        for a in m.Attachments:
            if a.FileName.lower().endswith((".xlsx", ".xls", ".xlsm")):
                fp = os.path.join(tmpd, a.FileName)
                try:
                    a.SaveAsFile(fp)
                    got = fp
                except Exception as e:
                    print("  [WARN] save %s failed: %s" % (a.FileName, e))
                break
        if not got:
            continue
        try:
            rows, zeros = parse_summary_xlsx(got)
            zero_all += zeros
        except Exception as e:
            print("  [WARN] parse %s failed: %s" % (os.path.basename(got), e))
            continue
        nrows += len(rows)
        for r in rows:
            key = r["code"]
            prev = latest.get(key)
            if prev is None or (r["recv"] and prev["recv"] and r["recv"] > prev["recv"]):
                latest[key] = r
            if key not in by_code:
                unknown.add(key)
        try:
            os.remove(got)
        except Exception:
            pass

    print("parsed ROB rows: %d | vessels in summaries: %d" % (nrows, len(latest)))
    if zero_all:
        print("[skip] all-zero ROB rows (unfilled by crew): %d" % len(zero_all))
    if unknown:
        print("[note] codes not in current fleet: %s" % ", ".join(sorted(unknown)))

    # 对比现有值, 列出会更新的船
    updates, newhist = [], []
    for code, r in sorted(latest.items()):
        cur = by_code.get(code)
        if cur is None:
            continue
        curt = (cur.get("report_time") or "")[:19]
        newt = r["recv"].strftime("%Y-%m-%d %H:%M:%S") if r["recv"] else ""
        if not newt:
            continue
        if curt and newt <= curt:
            continue
        updates.append((cur["vessel"], code, curt or "(none)", newt,
                        r["report_type"], r["rob"]))
        newhist.append((cur, r, newt))

    print("")
    print("=== would update %d vessels ===" % len(updates))
    for v, c, ot, nt, rt_, rob in updates:
        print("  %-24s %-6s %s -> %s | %-16s LSFO=%-8s MGO=%-8s HSFO=%-6s ULSFO=%-6s"
              % (v[:24], c, ot, nt, rt_,
                 rob["LSFO"], rob["MGO"], rob["HSFO"], rob["ULSFO"]))

    if not args.apply:
        print("")
        print("(dry-run) add --apply to write")
        return

    # ---- 写盘 ----
    # 保护2: 新记录某油品为 None(源表解析瑕疵)时保留旧值, 绝不把已有数据抹成空
    kept_old = 0
    for cur, r, newt in newhist:
        for oil, key in (("LSFO", "rob_lsfo"), ("HSFO", "rob_hsfo"),
                         ("MGO", "rob_mgo"), ("ULSFO", "rob_ulsfo")):
            v = r["rob"][oil]
            if v is None and cur.get(key) is not None:
                v = cur.get(key)
                kept_old += 1
            cur[key] = v
        cur["report_type"] = r["report_type"]
        cur["report_time"] = newt
        cur["source"] = "summary:" + r["report_type"]
        cur["found"] = True
        if r["lane"]:
            cur["lane"] = r["lane"]
        cur["voyage"] = r["voyage"]

    json.dump(results, open(RESULTS, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # 历史 CSV(按 vessel+report_time 去重追加)
    exist = set()
    if os.path.exists(HISTORY_CSV):
        for row in csv.DictReader(open(HISTORY_CSV, encoding="utf-8-sig")):
            exist.add((row.get("vessel"), (row.get("report_time") or "")[:19]))
    addrows = []
    for cur, r, newt in newhist:
        k = (cur["vessel"], newt)
        if k in exist:
            continue
        addrows.append({
            "date": r["date"], "vessel": cur["vessel"], "code": cur.get("code", ""),
            "lane": r["lane"] or cur.get("lane", ""), "pic": cur.get("pic", ""),
            "lsfo": r["rob"]["LSFO"], "hsfo": r["rob"]["HSFO"],
            "mgo": r["rob"]["MGO"], "ulsfo": r["rob"]["ULSFO"],
            "bw": cur.get("rob_bw"), "fw": cur.get("rob_fw"),
            "refeer": cur.get("rob_refeer"),
            "found": "1", "report_time": newt,
            "speed": cur.get("rob_speed"),
            "report_type": r["report_type"].upper(),
        })
    if addrows:
        fresh = not os.path.exists(HISTORY_CSV) or os.path.getsize(HISTORY_CSV) == 0
        with open(HISTORY_CSV, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=SNAP_FIELDS)
            if fresh:
                w.writeheader()
            w.writerows(addrows)
    print("")
    print("APPLIED: updated %d vessels, appended %d history rows" % (len(newhist), len(addrows)))


if __name__ == "__main__":
    main()
