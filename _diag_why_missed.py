# -*- coding: utf-8 -*-
"""找出 HUA DONG 811 / MAO GANG SHANG HAI / TEN VENUS 最新 NOON 邮件,
列出附件文件名, 并用 rob_refresh.extract_rob 实际解析, 定位漏抓原因。"""
import win32com.client, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rob_refresh as R

VESSELS = ["HUA DONG 811","MAO GANG SHANG HAI","TEN VENUS"]

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
    for vname in VESSELS:
        print("="*70)
        print("船:", vname)
        q = "@SQL=(\"urn:schemas:httpmail:subject\" like '%%%s%%') AND (\"urn:schemas:httpmail:subject\" like '%%NOON%%')" % vname
        hits=[]
        for f in flds:
            try: its=f.Items.Restrict(q)
            except Exception: continue
            try: its.Sort("[ReceivedTime]",True)
            except Exception: pass
            for it in its:
                try: hits.append((it.ReceivedTime,it))
                except Exception: pass
        hits.sort(key=lambda x:x[0],reverse=True)
        if not hits:
            print("  没找到 NOON 邮件"); continue
        rt, it = hits[0]
        print("  最新NOON: %s" % rt.strftime("%Y-%m-%d %H:%M"))
        atts=[a.FileName for a in it.Attachments]
        print("  附件(%d): %s" % (len(atts), atts))
        # 用真正的 pick_report_attachment + extract_rob
        att = R.pick_report_attachment(it)
        if att is None:
            print("  -> pick_report_attachment 返回 None (没有符合的 xlsx) => 这就是漏抓原因")
        else:
            print("  -> 选中附件: %s" % att.FileName)
            try:
                res = R.extract_rob(att)
                print("  -> extract_rob 结果: %s" % res)
            except Exception as e:
                print("  -> extract_rob 抛异常: %s" % e)

if __name__=="__main__":
    main()
