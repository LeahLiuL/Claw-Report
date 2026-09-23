# -*- coding: utf-8 -*-
"""诊断2: 邮件到底是"没到"还是"没同步下来"。只读。"""
import sys, os, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R

store = R.connect_outlook()
inbox = store.GetDefaultFolder(6)
folder_list = R.build_folder_list(inbox)
folders = [inbox] + list(folder_list)

import win32com.client
NS = store.Session
print("ExchangeConnectionMode =", NS.ExchangeConnectionMode, " Offline =", NS.Offline)
print("  0=NoExchange 100=Offline 200=CachedOffline 300=CachedHeaders 400=CachedDrizzle 500=CachedFull 600=Online")
try:
    for so in NS.SyncObjects:
        print("  SyncObject:", so.Name)
except Exception as e:
    print("  syncobj err", e)

cut = (datetime.datetime.now() - datetime.timedelta(days=5)).strftime("%m/%d/%Y %H:%M")
rows = []
for f in folders:
    try:
        it = f.Items.Restrict("[ReceivedTime] >= '%s'" % cut)
        for m in it:
            rt = m.ReceivedTime
            t = rt.strftime("%Y-%m-%d %H:%M") if hasattr(rt, "strftime") else str(rt)
            rows.append((t, f.Name, (m.Subject or "")[:60], R.get_sender(m)))
    except Exception:
        pass

rows.sort(reverse=True)
print("\n总邮件(近5天) =", len(rows))
byday = collections.Counter(r[0][:10] for r in rows)
print("按天:", dict(sorted(byday.items(), reverse=True)))
noon = collections.Counter(r[0][:10] for r in rows if "NOON" in r[2].upper())
print("含NOON按天:", dict(sorted(noon.items(), reverse=True)))
print("最新一封:", rows[0] if rows else None)

print("\n=== 抽样船长邮箱: 各自最新 3 封 ===")
for snd in ["culnansha@CULINES.COM", "culhochiminh@CULINES.COM", "maogangshanghai@126.com",
            "tenvenus@gtmailplus.com", "shengtang178@163.com", "bamsibeyrek@negmar.com",
            "culxiamen@CULINES.COM", "krtasman@krsline.com"]:
    got = [r for r in rows if r[3].lower() == snd][:3]
    print("--", snd)
    for t, fn, subj, s in got:
        print("     ", t, "|", fn, "|", subj)
    if not got:
        print("      (近5天无)")
