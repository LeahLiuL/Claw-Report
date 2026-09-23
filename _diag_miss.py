# -*- coding: utf-8 -*-
"""排查 4 艘 MISS 船: 在 leahliu 主邮箱全部文件夹里按船名(主题)和已知船长邮箱检索,
确认 NOON/报告是否存在。若存在却没被抓到, 说明是代码问题; 若不存在, 则是真没发/在别处。"""
import win32com.client

MISS = {
    "MYD TIANJIN":  ("mydtianjin", "mydtianjin@gtmailplus.com"),
    "KR CELEBES":   ("krcelebres", None),
    "MEDKON SINOP": ("medkonsinop", "medkonsinop@medkonlines.net"),
    "TB JINJIANG":  ("tbjinjiang", "tbjinjiang@msatmail.com"),
}

def norm(s):
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())

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
    print("扫描文件夹数: %d\n" % len(folders))

    for vname, (nv, smtp) in MISS.items():
        print("="*70)
        print("船: %s   norm=%s   sender=%s" % (vname, nv, smtp))
        # 按主题含船名
        hits = []
        for f in folders:
            try:
                items = f.Items.Restrict("@SQL=\"urn:schemas:httpmail:subject\" like '%%%s%%'" % vname)
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
                    sa = get_sender(it)
                    att = it.Attachments.Count
                    hits.append((rt, f.Name, sa, subj[:55], att))
                except Exception:
                    continue
        hits.sort(key=lambda x: x[0], reverse=True)
        if hits:
            print("  主题命中 %d 封, 最新 8 封:" % len(hits))
            for rt, fn, sa, subj, att in hits[:8]:
                print("   %s | [%s] | att=%d | %s | %s" % (rt.strftime("%Y-%m-%d %H:%M"), fn, att, sa, subj))
        else:
            print("  主题命中: 0 封 —— leahliu 主邮箱里没有任何含该船名的邮件")

if __name__ == "__main__":
    main()
