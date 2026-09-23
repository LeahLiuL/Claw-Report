# -*- coding: utf-8 -*-
"""用 DASL SMTP 字段 (urn:schemas:httpmail:senderemail) 检索 culnansha@culines.com 发来的所有邮件,
不分文件夹, 列出最新 15 封。同时统计 09-20/21/22 各多少封。"""
import win32com.client
import collections, datetime

SENDER = "culnansha@culines.com"
DASL = "urn:schemas:httpmail:senderemail"

def main():
    o = win32com.client.Dispatch("Outlook.Application")
    ns = o.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)
    folders = [inbox]
    seen = {id(inbox)}
    def walk(f):
        try:
            for c in f.Folders:
                if id(c) in seen: continue
                seen.add(id(c)); folders.append(c); walk(c)
        except Exception: pass
    walk(inbox)

    filt = "@SQL=\"%s\" = '%s'" % (DASL, SENDER)
    hits = []
    per_day = collections.Counter()
    for f in folders:
        try:
            items = f.Items.Restrict(filt)
        except Exception:
            continue
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        for it in items:
            try:
                rt = it.ReceivedTime
                subj = it.Subject or ""
                hits.append((rt, f.Name, subj))
                per_day[rt.strftime("%Y-%m-%d")] += 1
            except Exception:
                continue
    hits.sort(key=lambda x: x[0], reverse=True)
    print("发件人(SMTP): %s" % SENDER)
    print("命中总数: %d" % len(hits))
    print("按天分布: " + ", ".join("%s=%d" % (k, per_day[k]) for k in sorted(per_day, reverse=True)))
    print("\n最新 15 封:")
    for i, (rt, fname, subj) in enumerate(hits[:15], 1):
        print(" %2d | %s | [%s] | %s" % (i, rt.strftime("%Y-%m-%d %H:%M"), fname, subj))

if __name__ == "__main__":
    main()
