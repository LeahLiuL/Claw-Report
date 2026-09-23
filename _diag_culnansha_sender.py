# -*- coding: utf-8 -*-
"""按发件人 culnansha@culines.com 检索, 列出最新 30 封邮件(主题+时间), 不依赖任何过滤。"""
import win32com.client
import datetime

SENDER = "culnansha@culines.com"

def dt(t):
    try:
        return t.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(t)

def main():
    o = win32com.client.Dispatch("Outlook.Application")
    ns = o.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)

    # 全文件夹: 收件箱 + 所有子文件夹(递归)
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

    print("检索发件人: %s" % SENDER)
    print("扫描文件夹数: %d" % len(folders))

    filt = "[SenderEmailAddress]='%s'" % SENDER
    hits = []
    total = 0
    for f in folders:
        try:
            items = f.Items.Restrict(filt)
        except Exception:
            continue
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        n = 0
        for it in items:
            try:
                subj = it.Subject or ""
                rt = it.ReceivedTime
                hits.append((rt, subj, f.Name))
                n += 1
                total += 1
                if n >= 50:
                    break
            except Exception:
                continue
    print("命中总数(每文件夹前50): %d" % total)
    hits.sort(key=lambda x: x[0], reverse=True)
    print("\n最新 30 封:")
    for i, (rt, subj, fname) in enumerate(hits[:30], 1):
        print(" %2d | %s | [%s] | %s" % (i, dt(rt), fname, subj))

if __name__ == "__main__":
    main()
