# -*- coding: utf-8 -*-
"""诊断: Outlook 里到底有没有最近 3 天的 NOON / BERTH 报告邮件。
只读, 不改任何数据。
"""
import sys, os, datetime, collections

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R

store = R.connect_outlook()
print("OL_STATE:", R.OL_STATE)
inbox = store.GetDefaultFolder(6)
print("Inbox:", inbox.Name, "items=", inbox.Items.Count)

folder_list = R.build_folder_list(inbox)
print("folders:", len(folder_list))

cut_dt = datetime.datetime.now() - datetime.timedelta(days=3)
cut = cut_dt.strftime("%m/%d/%Y %H:%M")
cut5 = (datetime.datetime.now() - datetime.timedelta(days=6)).strftime("%m/%d/%Y %H:%M")

def count(items_filter):
    n = 0
    for f in [inbox] + list(folder_list):
        try:
            n += f.Items.Restrict(items_filter).Count
        except Exception:
            pass
    return n

print("recent 3d items =", count("[ReceivedTime] >= '%s'" % cut))
print("recent 6d items =", count("[ReceivedTime] >= '%s'" % cut5))

hits = []
for f in [inbox] + list(folder_list):
    try:
        it = f.Items.Restrict("[ReceivedTime] >= '%s'" % cut)
        it.Sort("[ReceivedTime]", True)
        for m in it:
            subj = (m.Subject or "")
            u = subj.upper()
            if not any(k in u for k in ("NOON", "BERTH", "DEPARTURE", "SAILING", "ARRIVAL")):
                continue
            rt = m.ReceivedTime
            t = rt.strftime("%Y-%m-%d %H:%M") if hasattr(rt, "strftime") else str(rt)
            snd = R.get_sender(m)
            hits.append((t, f.Name, snd, subj[:70]))
    except Exception as e:
        print("  skip", f.Name, e)

hits.sort(reverse=True)
print("\n=== 最近3天 NOON/BERTH/DEPARTURE 邮件: %d 封 ===" % len(hits))
byday = collections.Counter(h[:10] for h in hits)
print("按收到日(ReceivedTime):", dict(sorted(byday.items(), reverse=True)))
print("\n最近 40 封:")
for t, fn, snd, subj in hits[:40]:
    print("  %s | %-20s | %-34s | %s" % (t, fn[:20], snd[:34], subj))
