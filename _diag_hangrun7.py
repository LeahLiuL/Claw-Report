# -*- coding: utf-8 -*-
"""按主题搜 HANG RUN 7 / HGR7 的全部邮件, 看船长是否发过 NOON/ROB 报告。"""
import datetime, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32com.client

ol = win32com.client.Dispatch("Outlook.Application")
ns = ol.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)

def walk(f, out):
    try:
        out.append(f)
        for c in f.Folders:
            walk(c, out)
    except Exception:
        pass

folders = []
walk(inbox, folders)
print("folders:", len(folders))

cutoff = (datetime.datetime.now() - datetime.timedelta(days=8)).strftime("%m/%d/%Y %H:%M %p")
hits = []
for f in folders:
    try:
        items = f.Items.Restrict("[ReceivedTime] >= '" + cutoff + "'")
    except Exception:
        continue
    for it in items:
        try:
            subj = (it.Subject or "")
        except Exception:
            continue
        u = subj.upper()
        if ("HANG RUN" in u) or ("HGR7" in u) or ("HANGRUN" in u):
            try:
                atts = [a.FileName for a in it.Attachments]
            except Exception:
                atts = []
            # 解析发件人 SMTP
            try:
                sa = it.SenderEmailAddress or ""
                if "@" not in sa:
                    try:
                        sa = it.Sender.GetExchangeUser().PrimarySmtpAddress or sa
                    except Exception:
                        pass
            except Exception:
                sa = ""
            hits.append((it.ReceivedTime, subj, sa, atts, f.Name))

hits.sort(key=lambda x: x[0], reverse=True)
print("hits:", len(hits))
for h in hits[:25]:
    t, subj, sa, atts, folder = h
    print("%s | [%s] | %s | sender=%s | att=%s" % (t, folder, subj[:60], sa, atts[:3]))
