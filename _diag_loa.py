# -*- coding: utf-8 -*-
"""找出 LOA HARMONY 的真实船长发件人邮箱: 按主题含 LOA HARMONY 搜全部文件夹,
列出最新若干封的 发件人显示名 + 真实 SMTP(经 get_sender 解析 X.500)。"""
import win32com.client

def get_sender(it):
    try:
        v = it.SenderEmailAddress
        if v and "@" in v:
            return v
    except Exception:
        pass
    try:
        ex = it.Sender
        if ex:
            smtp = ex.GetExchangeUser().PrimarySmtpAddress
            if smtp and "@" in smtp:
                return smtp
    except Exception:
        pass
    return ""

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
    print("文件夹数: %d\n" % len(folders))

    q = "@SQL=\"urn:schemas:httpmail:subject\" like '%LOA HARMONY%'"
    hits = []
    for f in folders:
        try:
            items = f.Items.Restrict(q)
        except Exception:
            continue
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        for it in items:
            try:
                hits.append((it.ReceivedTime, f.Name, it.SenderName or "", get_sender(it), it.Subject or "", it.Attachments.Count))
            except Exception:
                continue
    hits.sort(key=lambda x: x[0], reverse=True)
    print("LOA HARMONY 主题命中: %d 封, 最新 15 封:\n" % len(hits))
    for rt, fn, sn, sa, subj, att in hits[:15]:
        print(" %s | [%s] | %s | %s | att=%d | %s" % (rt.strftime("%Y-%m-%d %H:%M"), fn, sn, sa, att, subj[:55]))

if __name__ == "__main__":
    main()
