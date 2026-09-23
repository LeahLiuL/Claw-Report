# -*- coding: utf-8 -*-
"""Outlook 不能直接用 SMTP 过滤(它存的是 X.500 内部地址), 改用 SenderName 过滤,
再解析 PrimarySmtp 确认是 culnansha@culines.com, 列出该发件人最新 20 封。"""
import win32com.client
import collections

SENDER_NAME = "CUL Nansha"
WANT_SMTP = "culnansha@culines.com"

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

    filt = "[SenderName]='%s'" % SENDER_NAME
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
                smtp = ""
                try:
                    smtp = it.Sender.GetExchangeUser().PrimarySmtpAddress or ""
                except Exception:
                    smtp = ""
                if smtp.lower() == WANT_SMTP.lower():
                    hits.append((rt, f.Name, subj))
                    per_day[rt.strftime("%Y-%m-%d")] += 1
            except Exception:
                continue
    hits.sort(key=lambda x: x[0], reverse=True)
    print("发件人显示名: %s  (解析 SMTP=%s)" % (SENDER_NAME, WANT_SMTP))
    print("命中总数: %d" % len(hits))
    if per_day:
        print("按天分布: " + ", ".join("%s=%d" % (k, per_day[k]) for k in sorted(per_day, reverse=True)))
    print("\n最新 20 封:")
    for i, (rt, fname, subj) in enumerate(hits[:20], 1):
        print(" %2d | %s | [%s] | %s" % (i, rt.strftime("%Y-%m-%d %H:%M"), fname, subj))

if __name__ == "__main__":
    main()
