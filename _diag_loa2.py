# -*- coding: utf-8 -*-
"""确认 loaharmony@doshiphk.com 是否发过 ROB/NOON 报告:
1) 主题含 LOA HARMONY 且含 NOON/RPT/ROB 的邮件
2) 该发件人全部邮件中带可解析 xlsx 附件的
打印最新若干供判断。"""
import win32com.client

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
    flds = [inbox]
    seen={id(inbox)}
    def walk(f):
        try:
            for c in f.Folders:
                if id(c) in seen: continue
                seen.add(id(c)); flds.append(c); walk(c)
        except Exception: pass
    walk(inbox)
    q = "@SQL=(\"urn:schemas:httpmail:subject\" like '%LOA HARMONY%') AND (\"urn:schemas:httpmail:subject\" like '%NOON%' OR \"urn:schemas:httpmail:subject\" like '%RPT%' OR \"urn:schemas:httpmail:subject\" like '%ROB%')"
    a=[]; b=[]
    for f in flds:
        try: its=f.Items.Restrict(q)
        except Exception: continue
        try: its.Sort("[ReceivedTime]",True)
        except Exception: pass
        for it in its:
            try: a.append((it.ReceivedTime,f.Name,it.Subject or "",it.Attachments.Count))
            except Exception: pass
    a.sort(key=lambda x:x[0],reverse=True)
    print("LOA HARMONY + NOON/RPT/ROB 主题命中: %d 封" % len(a))
    for rt,fn,subj,att in a[:10]:
        print("  %s | [%s] | att=%d | %s" % (rt.strftime("%Y-%m-%d %H:%M"),fn,att,subj[:60]))

    # 该发件人全部带附件邮件
    q2="@SQL=\"urn:schemas:httpmail:senderemail\" = 'loaharmony@doshiphk.com'"
    # 上面 DASL 不一定生效, 改为遍历解析
    for f in flds:
        try: its=f.Items
        except Exception: continue
        try: its.Sort("[ReceivedTime]",True)
        except Exception: pass
        n=0
        for it in its:
            try:
                if n>=200: break
                n+=1
                if get_sender(it).lower()=="loaharmony@doshiphk.com" and it.Attachments.Count>0:
                    b.append((it.ReceivedTime,f.Name,it.Subject or "",it.Attachments.Count))
            except Exception: continue
    b.sort(key=lambda x:x[0],reverse=True)
    print("\nloaharmony@doshiphk.com 带附件邮件: %d 封 (显示最新 10)" % len(b))
    for rt,fn,subj,att in b[:10]:
        print("  %s | [%s] | att=%d | %s" % (rt.strftime("%Y-%m-%d %H:%M"),fn,att,subj[:60]))

if __name__=="__main__":
    main()
