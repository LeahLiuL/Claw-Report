# -*- coding: utf-8 -*-
"""完全按发件人邮箱检索(不看主题关键词), 列出该邮箱发来的最新 15 封(主题+附件名+日期)。
覆盖: HONG DA XIN 708 / NEW THINKER / 以及之前漏抓的 3 艘对照。"""
import win32com.client

TARGETS = {
    "hongdaxin708@msatmail.com": "HONG DA XIN 708",
    "newthinker@lianjieship.com.cn": "NEW THINKER",
}

def get_sender(it):
    try:
        v = it.SenderEmailAddress
        if v and "@" in v: return v
    except Exception: pass
    try:
        ex = it.Sender
        if ex:
            s = ex.GetExchangeUser().PrimarySmtpAddress
            if s and "@" in s: return s
    except Exception: pass
    return ""

def main():
    o = win32com.client.Dispatch("Outlook.Application")
    ns = o.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)
    flds=[inbox]
    seen={id(inbox)}
    def walk(f):
        try:
            for c in f.Folders:
                if id(c) in seen: continue
                seen.add(id(c)); flds.append(c); walk(c)
        except Exception: pass
    walk(inbox)
    print("文件夹数: %d\n" % len(flds))
    # 收集: email -> list[(rt, folder, subject, [atts])]
    found={k:[] for k in TARGETS}
    for f in flds:
        try: its=f.Items
        except Exception: continue
        try: its.Sort("[ReceivedTime]",True)
        except Exception: pass
        n=0
        for it in its:
            try:
                if n>=400: break
                n+=1
                sa=get_sender(it).lower()
                if sa in found:
                    atts=[a.FileName for a in it.Attachments]
                    found[sa].append((it.ReceivedTime, f.Name, it.Subject or "", atts))
            except Exception: continue
    for email, name in TARGETS.items():
        hits=sorted(found[email], key=lambda x:x[0], reverse=True)
        print("="*70)
        print("%s  <%s>  命中 %d 封" % (name, email, len(hits)))
        for rt,fn,subj,atts in hits[:15]:
            print("  %s | [%s] | %s" % (rt.strftime("%Y-%m-%d %H:%M"), fn, subj[:55]))
            if atts:
                print("       附件: %s" % atts)

if __name__=="__main__":
    main()
