# -*- coding: utf-8 -*-
"""诊断6: 用【船长邮箱 sender】反查。只建一次邮件池(近6天), 按 sender 归组。
回答两个问题:
  1) 收件箱现在到底有多少封、最新一封是什么时候(判断邮箱本身是不是陈旧缓存)
  2) 每个船长邮箱最新一封邮件是哪天、是不是 NOON
"""
import sys, os, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R

store = R.connect_outlook()
NS = store.Session
print("store =", store.DisplayName, " mode =", NS.ExchangeConnectionMode, " offline =", NS.Offline)

inbox = store.GetDefaultFolder(6)
print("收件箱 Items.Count =", inbox.Items.Count)
try:
    it = inbox.Items
    it.Sort("[ReceivedTime]", True)
    print("收件箱最新 5 封:")
    for i, m in enumerate(it):
        if i >= 5:
            break
        rt = m.ReceivedTime
        print("   ", rt.strftime("%Y-%m-%d %H:%M"), "|", R.get_sender(m), "|", (m.Subject or "")[:60])
except Exception as e:
    print("   sort err", e)

folder_list = R.build_folder_list(inbox)
pool, scanned, nf = R.build_deep_pool(inbox, folder_list, 6)
print("\n池: kept=%d scanned=%d folders=%d" % (len(pool), scanned, nf))

byday = collections.Counter()
for row in pool:
    byday[row[0].strftime("%Y-%m-%d")] += 1
print("池按收到日:", dict(sorted(byday.items(), reverse=True)))

smap = R.load_sender_map()
print("\n船长邮箱数 =", len(smap))
by_sender = collections.defaultdict(list)
for row in pool:
    s = row[3]
    if s is None:
        try:
            s = (R.get_sender(row[1]) or "").lower()
        except Exception:
            s = ""
        row[3] = s
    if s:
        by_sender[s].append(row)

print("\n=== 每个船长邮箱: 最新一封 ===")
recent3 = 0
for v, mail in sorted(smap.items()):
    ml = by_sender.get(mail.lower(), [])
    if not ml:
        print("  %-24s %-32s  (近6天无邮件)" % (v[:24], mail[:32]))
        continue
    top = sorted(ml, key=lambda x: x[0], reverse=True)[:2]
    t0 = top[0][0]
    if (datetime.datetime.now() - t0.replace(tzinfo=None)).days <= 3:
        recent3 += 1
    desc = " | ".join("%s %s" % (r[0].strftime("%m-%d %H:%M"), (r[1].Subject or "")[:38]) for r in top)
    print("  %-24s %-32s  %s" % (v[:24], mail[:32], desc))
print("\n近3天有邮件的船长邮箱数 =", recent3, "/", len(smap))
