# -*- coding: utf-8 -*-
"""只读诊断(可靠版): 对 11 行按秒级时间定位邮件, 完整 dump 每个 ROB 单元格及右侧值。
不做船名过滤(秒级时间唯一), 复用已验证的宽松 dump 逻辑。不写数据。"""
import sys, os, csv, tempfile
from datetime import datetime, timedelta
sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import win32com.client as win32
import rob_refresh as R
import openpyxl

CSV = os.path.join(BASE, "rob_data", "berth_fwe_suspicious_values.csv")

def dump_full(att):
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        att.SaveAsFile(p)
        wb = openpyxl.load_workbook(p, data_only=True)
    except Exception as e:
        try: os.remove(p)
        except Exception: pass
        print("    [ERR] %s" % str(e)[:60]); return
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for i, c in enumerate(row):
                if c and isinstance(c, str) and "ROB" in c.upper():
                    cu = c.upper().strip()
                    right = [str(row[j]).strip()[:16] for j in range(i+1, min(i+7, len(row)))
                             if row[j] is not None and str(row[j]).strip()]
                    if right:
                        print("      [%s] %s" % (cu[:48], " | ".join(right)))
                        n += 1
                        if n > 40: break
            if n > 40: break
        if n > 40: break
    try: os.remove(p)
    except Exception: pass

rows = list(csv.DictReader(open(CSV, encoding="utf-8-sig")))
secs_by_day = {}
for r in rows:
    rt = r["report_time"][:19]
    secs_by_day.setdefault(rt[:10], set()).add(rt)

ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
st = next((ns.Stores(i) for i in range(1, ns.Stores.Count+1)
           if "vesselreport" in (ns.Stores(i).DisplayName or "").lower()), None)
root = st.GetRootFolder()
folders = [root]; stack = [(root, 0)]
while stack:
    f, d = stack.pop()
    if d >= 4: continue
    try:
        for c in f.Folders:
            folders.append(c); stack.append((c, d+1))
    except Exception: pass

for day, secs in sorted(secs_by_day.items()):
    y, m, dd = day.split("-")
    lo = "%s/%s/%s 12:00 AM" % (m, dd, y)
    nextd = (datetime(int(y), int(m), int(dd)) + timedelta(days=1)).strftime("%m/%d/%Y 12:00 AM")
    filt = "[ReceivedTime] >= '%s' AND [ReceivedTime] <= '%s'" % (lo, nextd)
    print("\n========== %s ==========" % day)
    for fld in folders:
        try:
            items = fld.Items; items.Sort("[ReceivedTime]", True); sub = items.Restrict(filt)
        except Exception:
            continue
        for it in sub:
            try: rt = it.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
            except Exception: continue
            if rt[:19] not in secs: continue
            try: subj = str(it.Subject or "")
            except Exception: subj = ""
            try: att = R.pick_report_attachment(it)
            except Exception: continue
            if att is None: continue
            print("\n  [mail] %s | %s | %s" % (fld.Name[:14], rt, subj[:52]))
            dump_full(att)
print("\nDONE")
