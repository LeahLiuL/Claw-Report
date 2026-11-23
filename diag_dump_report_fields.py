# -*- coding: utf-8 -*-
"""只读诊断: 从 vesselreport 邮箱各抓 NOON/BERTH/SAILING 报告若干封,
把原始 xlsx 的所有非空单元格(坐标 + 值)完整 dump 到 txt, 用于设计'全字段报告库'结构。
只读取, 不写任何业务数据、不改邮箱。"""
import sys, os, tempfile
from datetime import datetime, timedelta
sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import win32com.client as win32
import rob_refresh as R
import openpyxl

OUT = os.path.join(BASE, "rob_data", "diag_report_fields_dump.txt")
NEED = 3  # 每种类型抓几封
WANT = ["NOON", "BERTH", "SAILING"]
found = {k: 0 for k in WANT}
lines = []


def dump_att(att, tag):
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        att.SaveAsFile(p)
        wb = openpyxl.load_workbook(p, data_only=True)
    except Exception as e:
        lines.append("  [ERR] %s" % str(e)[:80])
        try: os.remove(p)
        except Exception: pass
        return
    lines.append("  --- sheets: %s ---" % ", ".join(wb.sheetnames))
    for ws in wb.worksheets:
        lines.append("  === sheet [%s] dims=%s ===" % (ws.title, ws.dimensions))
        for row in ws.iter_rows():
            cells = []
            for c in row:
                if c.value is None:
                    continue
                s = str(c.value).strip()
                if not s:
                    continue
                cells.append("%s=%s" % (c.coordinate, s[:40]))
            if cells:
                lines.append("    " + " || ".join(cells))
    try: os.remove(p)
    except Exception: pass


ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
st = next((ns.Stores(i) for i in range(1, ns.Stores.Count + 1)
           if "vesselreport" in (ns.Stores(i).DisplayName or "").lower()), None)
if st is None:
    print("NO vesselreport STORE")
    sys.exit(1)
root = st.GetRootFolder()
folders = [root]
stack = [(root, 0)]
while stack:
    f, d = stack.pop()
    if d >= 4:
        continue
    try:
        for c in f.Folders:
            folders.append(c)
            stack.append((c, d + 1))
    except Exception:
        pass

# 最近 45 天
lo = (datetime.now() - timedelta(days=45)).strftime("%m/%d/%Y 12:00 AM")
filt = "[ReceivedTime] >= '%s'" % lo
for fld in folders:
    if all(found[k] >= NEED for k in WANT):
        break
    try:
        items = fld.Items
        items.Sort("[ReceivedTime]", True)
        sub = items.Restrict(filt)
    except Exception:
        continue
    for it in sub:
        try:
            subj = str(it.Subject or "")
        except Exception:
            continue
        try:
            att = R.pick_report_attachment(it)
        except Exception:
            continue
        if att is None:
            continue
        # 用附件正文判定类型
        try:
            rob = R.extract_rob(att)
            kind = rob.get("REPORT_KIND", "")
        except Exception:
            kind = ""
        if kind not in WANT or found[kind] >= NEED:
            continue
        found[kind] += 1
        try:
            rt = it.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            rt = ""
        lines.append("\n========== [%s] #%d | %s | %s ==========" %
                     (kind, found[kind], fld.Name[:18], rt))
        lines.append("  subject: %s" % subj[:90])
        try:
            lines.append("  attach : %s" % str(att.FileName or "")[:90])
        except Exception:
            pass
        lines.append("  parsed ROB: %s" % {k: v for k, v in rob.items() if k != "REPORT_KIND"})
        dump_att(att, kind)
        if all(found[k] >= NEED for k in WANT):
            break

with open(OUT, "w", encoding="utf-8") as fh:
    fh.write("\n".join(lines))
print("抓取统计: %s" % found)
print("已 dump -> %s (%d 行)" % (OUT, len(lines)))
print("DONE")
