# -*- coding: utf-8 -*-
"""找出 CUL NANSHA 的 NOON 报告邮件, 打印真实 SenderName / SenderEmailAddress / AddressEntryUserType。"""
import win32com.client

def main():
    o = win32com.client.Dispatch("Outlook.Application")
    ns = o.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)

    folders = [inbox]
    seen = {id(inbox)}
    def walk(f):
        try:
            for c in f.Folders:
                if id(c) in seen:
                    continue
                seen.add(id(c))
                folders.append(c)
                walk(c)
        except Exception:
            pass
    walk(inbox)

    print("文件夹数: %d" % len(folders))
    want = ("NOON", "CUL NANSHA", "CULNANSHA", "2635W")
    found = []
    for f in folders:
        try:
            items = f.Items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%NOON%'")
        except Exception:
            continue
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        for it in items:
            try:
                subj = it.Subject or ""
                if "CUL NANSHA" in subj.upper() or "CULNANSHA" in subj.upper():
                    se = it.SenderEmailAddress
                    sn = it.SenderName
                    rt = it.ReceivedTime
                    ae = None
                    try:
                        ae = it.Sender.GetExchangeUser().PrimarySmtpAddress
                    except Exception:
                        ae = "(n/a)"
                    found.append((rt, f.Name, sn, se, ae, subj[:60]))
            except Exception:
                continue
    found.sort(key=lambda x: x[0], reverse=True)
    print("命中 CUL NANSHA NOON 相关: %d 封(显示最新 20)" % len(found))
    for rt, fname, sn, se, ae, subj in found[:20]:
        print("-" * 80)
        print("时间:%s | 文件夹:%s" % (rt, fname))
        print("  SenderName      : %s" % sn)
        print("  SenderEmailAddress: %s" % se)
        print("  PrimarySmtp     : %s" % ae)
        print("  主题            : %s" % subj)

if __name__ == "__main__":
    main()
