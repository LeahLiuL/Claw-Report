# -*- coding: utf-8 -*-
"""诊断5: 枚举 Outlook 里【全部 store/邮箱】, 找最近 3 天的 NOON 报告。
之前只扫了主邮箱一个 store —— 如果船长发到共享邮箱/另一个账户, 就会全部漏掉。
"""
import sys, os, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R
import win32com.client

OL = win32com.client.Dispatch("Outlook.Application")
NS = OL.GetNamespace("MAPI")
print("mode =", NS.ExchangeConnectionMode, " Offline =", NS.Offline)
print("=== 全部 store ===")
stores = list(NS.Stores)
for i, s in enumerate(stores):
    print("  [%d] %-45s  root=%s" % (i, s.DisplayName, s.GetRootFolder().Name))

cut = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%m/%d/%Y %H:%M")

for i, s in enumerate(stores):
    root = s.GetRootFolder()
    flds = []
    def walk(f, d=0):
        if d > 4:
            return
        try:
            for c in f.Folders:
                flds.append(c)
                walk(c, d + 1)
        except Exception:
            pass
    walk(root)
    rows = []
    for f in flds:
        try:
            for m in f.Items.Restrict("[ReceivedTime] >= '%s'" % cut):
                rt = m.ReceivedTime
                rows.append((rt.strftime("%Y-%m-%d %H:%M"), f.Name,
                             (m.Subject or "")[:70], R.get_sender(m)))
        except Exception:
            pass
    rows.sort(reverse=True)
    noon = [r for r in rows if "NOON" in r[2].upper()]
    print("\n--- [%d] %s : 文件夹 %d, 近3天邮件 %d, 含NOON %d ---"
          % (i, s.DisplayName, len(flds), len(rows), len(noon)))
    if noon:
        print("    按天:", dict(sorted(collections.Counter(r[0][:10] for r in noon).items(), reverse=True)))
        for r in noon[:8]:
            print("      ", " | ".join(r))
    elif rows:
        print("    最新3封:", rows[:3])
