# -*- coding: utf-8 -*-
"""对停在 09-20/09-18 的 10 艘船, 在 leahliu 主邮箱全部文件夹按船名+NOON/RPT 搜最新报告,
判断: 有更新报告却被漏抓(代码问题) vs 邮箱里根本没有(同步/未发)。"""
import win32com.client

VESSELS = ["GLOBAL 01","HONG DA XIN 708","HUA DONG 811","M. MARINER","MAO GANG SHANG HAI",
           "NEW THINKER","TEN VENUS","WAN FU DA","XIN MING ZHOU 106","XING DONG 88"]

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
    for vname in VESSELS:
        q = "@SQL=(\"urn:schemas:httpmail:subject\" like '%%%s%%') AND (\"urn:schemas:httpmail:subject\" like '%%NOON%%' OR \"urn:schemas:httpmail:subject\" like '%%RPT%%' OR \"urn:schemas:httpmail:subject\" like '%%ROB%%')" % vname
        hits=[]
        for f in flds:
            try: its=f.Items.Restrict(q)
            except Exception: continue
            try: its.Sort("[ReceivedTime]",True)
            except Exception: pass
            for it in its:
                try: hits.append((it.ReceivedTime,f.Name,get_sender(it),it.Subject or "",it.Attachments.Count))
                except Exception: pass
        hits.sort(key=lambda x:x[0],reverse=True)
        if hits:
            rt,fn,sa,subj,att=hits[0]
            print("%-20s | 最新NOON: %s | [%s] | att=%d | %s" % (vname, rt.strftime("%Y-%m-%d %H:%M"), fn, att, subj[:48]))
            if len(hits)>1:
                rt2=hits[1][0]
                print("%-20s |   次新:    %s" % ("", rt2.strftime("%Y-%m-%d %H:%M")))
        else:
            print("%-20s | 邮箱里没有任何 NOON/RPT/ROB 报告" % vname)

if __name__=="__main__":
    main()
