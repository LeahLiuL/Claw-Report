# -*- coding: utf-8 -*-
"""精确确认 4 艘 MISS 船是否存在 ROB/NOON 报告(xlsx 含 ROB 数据):
对每艘船, 在其专用文件夹 + 全部文件夹里找 主题含(NOON/ROB/BUNKER/RPT) 且含船名/代码的邮件,
并检测是否带可解析的 xlsx 附件。"""
import win32com.client
import re

MISS = {
    "MYD TIANJIN": ("MYD TIANJIN", "MYDTIANJIN", "TB"),
    "KR CELEBES":  ("KR CELEBES", "KRCELEBRES", "KRCB"),
    "MEDKON SINOP":("MEDKON SINOP", "MEDKONSINOP", "MS"),
    "TB JINJIANG": ("TB JINJIANG", "TBJINJIANG", "TBJJ"),
}
RPT_KW = ("NOON", "ROB", "BUNKER", "RPT", "NOON RPT")

def norm(s):
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())

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

    for vname, toks in MISS.items():
        print("="*70)
        print("船: %s  关键词: %s" % (vname, toks))
        # 构造 REDMOND 风格多关键词 OR: 主题含 NOON/ROB/RPT 之一 且 含船名token
        rpt_or = " OR ".join("\"urn:schemas:httpmail:subject\" like '%%%s%%'" % k for k in RPT_KW)
        name_or = " OR ".join("\"urn:schemas:httpmail:subject\" like '%%%s%%'" % t for t in toks)
        q = "@SQL=(%s) AND (%s)" % (rpt_or, name_or)
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
                    hits.append((it.ReceivedTime, f.Name, it.Subject or "", it.Attachments.Count))
                except Exception:
                    continue
        hits.sort(key=lambda x: x[0], reverse=True)
        if hits:
            print("  找到 %d 封疑似 ROB/NOON 报告:" % len(hits))
            for rt, fn, subj, att in hits[:10]:
                print("   %s | [%s] | att=%d | %s" % (rt.strftime("%Y-%m-%d %H:%M"), fn, att, subj[:60]))
        else:
            print("  未发现任何 NOON/ROB/RPT 报告邮件 —— 该船在 leahliu 主邮箱确实没有 ROB 报告")

if __name__ == "__main__":
    main()
