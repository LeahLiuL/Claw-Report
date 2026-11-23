# -*- coding: utf-8 -*-
"""只读诊断: 打开可疑大跳变 BERTH 报告的原始 xlsx, dump 所有 ROB 行的阶段前缀+原始值。
按精确秒级时间定位邮件(Restrict 按天收窄), 不写任何数据。"""
import sys, os, tempfile
from datetime import datetime, timedelta
sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import win32com.client as win32
import rob_refresh as R
import openpyxl

# (船名关键词, 目标秒级时间, 天数窗口前后)
TARGETS = [
    ("CUL YANGPU",   "2025-12-31 13:13:23", 1),
    ("STAR BLESSING","2026-07-22 22:15:54", 1),
    ("SHENGTANG",    "2026-09-14 16:28:47", 1),
    ("SHENGTANG",    "2026-09-14 16:28:49", 1),
    ("MYD TIANJIN",  "2026-10-10 20:20:30", 1),
    ("MYD TIANJIN",  "2026-10-10 20:20:44", 1),
    ("medkonlia",    "2026-07-09 15:32:40", 1),
]

def dump_att(att, tag):
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        att.SaveAsFile(p)
        wb = openpyxl.load_workbook(p, data_only=True)
    except Exception as e:
        print("    [ERR save/open] %s" % str(e)[:80])
        try: os.remove(p)
        except Exception: pass
        return
    print("    === %s 附件 ROB 行 ===" % tag)
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for i, c in enumerate(row):
                if c and isinstance(c, str) and "ROB" in c.upper():
                    cu = c.upper().strip()
                    # 取右侧最多 6 个单元格
                    right = []
                    for j in range(i+1, min(i+7, len(row))):
                        v = row[j]
                        if v is not None and str(v).strip():
                            right.append(str(v).strip()[:18])
                    if right:
                        print("      [%s] %s" % (cu[:52], " | ".join(right)))
                        n += 1
                        if n > 45: break
            if n > 45: break
        if n > 45: break
    # 当前 extract_rob 结果
    try:
        rob = R.extract_rob(att)
        print("    >>> extract_rob(FWE优先): LSFO=%s HSFO=%s MGO=%s ULSFO=%s kind=%s" % (
            rob.get("LSFO"), rob.get("HSFO"), rob.get("MGO"), rob.get("ULSFO"), rob.get("REPORT_KIND")))
    except Exception as e:
        print("    [ERR extract] %s" % str(e)[:80])
    try: os.remove(p)
    except Exception: pass

ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
st = next((ns.Stores(i) for i in range(1, ns.Stores.Count+1)
           if "vesselreport" in (ns.Stores(i).DisplayName or "").lower()), None)
root = st.GetRootFolder()
folders = [root]
stack = [(root, 0)]
while stack:
    f, d = stack.pop()
    if d >= 4: continue
    try:
        for c in f.Folders:
            folders.append(c); stack.append((c, d+1))
    except Exception: pass
print("文件夹 %d 个" % len(folders))

for vessel, rt, win in TARGETS:
    dt = datetime.strptime(rt, "%Y-%m-%d %H:%M:%S")
    lo = (dt - timedelta(days=win)).strftime("%m/%d/%Y 12:00 AM")
    hi = (dt + timedelta(days=win)).strftime("%m/%d/%Y 12:00 AM")
    filt = "[ReceivedTime] >= '%s' AND [ReceivedTime] <= '%s'" % (lo, hi)
    print("\n######## %s  %s ########" % (vessel, rt))
    found = 0
    for fld in folders:
        try:
            items = fld.Items; items.Sort("[ReceivedTime]", True); sub = items.Restrict(filt)
        except Exception:
            continue
        for it in sub:
            try:
                t = it.ReceivedTime
                it_rt = t.strftime("%Y-%m-%d %H:%M:%S") if hasattr(t, "strftime") else ""
            except Exception:
                continue
            # 目标秒级 +/- 60s 内, 且船名 token 在主题/附件名
            try: dsec = abs((datetime.strptime(it_rt, "%Y-%m-%d %H:%M:%S") - dt).total_seconds())
            except Exception: continue
            if dsec > 90: continue
            try: subj = str(it.Subject or "")
            except Exception: subj = ""
            atts = []
            try:
                for a in it.Attachments:
                    try: atts.append(a.FileName)
                    except Exception: pass
            except Exception: pass
            blob = R.norm(subj) + " " + " ".join(R.norm(a) for a in atts)
            if R.norm(vessel) not in blob: continue
            try: att = R.pick_report_attachment(it)
            except Exception: att = None
            if att is None: continue
            print("  [mail] %s | %s | %s | %s" % (fld.Name[:14], it_rt, subj[:42], atts[:1]))
            dump_att(att, "%s %s (diff=%ds)" % (vessel, it_rt, dsec))
            found += 1
            if found >= 3: break
        if found >= 3: break
    if not found:
        print("  (未定位到邮件)")
print("\nDONE")
