# -*- coding: utf-8 -*-
"""轻量复核: Outlook 当前同步状态 + 最近 3 天 NOON 报告是否真的没有。只读。"""
import sys, os, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R

store = R.connect_outlook()
NS = store.Session
print("mode =", NS.ExchangeConnectionMode, " offline =", NS.Offline)
inbox = store.GetDefaultFolder(6)
folders = [inbox] + list(R.build_folder_list(inbox))

cut = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%m/%d/%Y %H:%M")
rows = []
for f in folders:
    try:
        for m in f.Items.Restrict("[ReceivedTime] >= '%s'" % cut):
            rt = m.ReceivedTime
            rows.append((rt.strftime("%Y-%m-%d %H:%M"), (m.Subject or "").upper()))
    except Exception:
        pass
rows.sort(reverse=True)
print("近3天邮件总数 =", len(rows), " 最新一封 =", rows[0][0] if rows else None)
print("按天:", dict(sorted(collections.Counter(r[0][:10] for r in rows).items(), reverse=True)))
print("含NOON按天:", dict(sorted(collections.Counter(r[0][:10] for r in rows if "NOON" in r[1]).items(), reverse=True)))
