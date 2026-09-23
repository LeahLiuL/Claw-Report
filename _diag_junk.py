# -*- coding: utf-8 -*-
"""诊断3: 全 store(含垃圾邮件/已删除/收件箱外的顶层文件夹)找最近 NOON 报告。只读。"""
import sys, os, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R

store = R.connect_outlook()
root = store.GetRootFolder()

allf = []
def walk(f, depth=0):
    if depth > 5:
        return
    try:
        for c in f.Folders:
            allf.append((depth + 1, c))
            walk(c, depth + 1)
    except Exception:
        pass
walk(root)
print("全 store 文件夹数 =", len(allf))

cut = (datetime.datetime.now() - datetime.timedelta(days=4)).strftime("%m/%d/%Y %H:%M")
hits = []
for d, f in allf:
    try:
        it = f.Items.Restrict("[ReceivedTime] >= '%s'" % cut)
        if it.Count == 0:
            continue
        for m in it:
            subj = (m.Subject or "")
            if "NOON" not in subj.upper():
                continue
            rt = m.ReceivedTime
            t = rt.strftime("%Y-%m-%d %H:%M")
            hits.append((t, f.Name, R.get_sender(m), subj[:60]))
    except Exception:
        pass
hits.sort(reverse=True)
print("\n全 store 近4天含 NOON 的邮件 =", len(hits))
print("按天:", dict(sorted(collections.Counter(h[:10] for h in hits).items(), reverse=True)))
print("按文件夹:", dict(sorted(collections.Counter(h[1] for h in hits).items(),
                              key=lambda x: -x[1])[:15]))
for h in hits[:15]:
    print("  ", " | ".join(h))
