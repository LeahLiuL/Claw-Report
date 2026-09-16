"""采样验证: 在真实 Outlook 邮件上测试 DRAFT 前/中/后吃水抽取是否生效。

按【发件人邮箱】定位船长报告(与 rob_refresh 主流程一致), 而非按主题
(收件箱里含船名的邮件大多是码头安排/危险品, 不是 ROB 报告)。
只读不写。
用法: python sample_draft_check.py [船名1 船名2 ...]
"""
import sys
import datetime
import rob_refresh as R


def collect(sender, limit_days=45, per_folder=400):
    """按 sender 收集候选邮件(收件箱 + 一层子文件夹), 倒序。"""
    store = R.connect_outlook()
    root = store.GetDefaultFolder(6)
    folders = [root]
    try:
        for i in range(1, root.Folders.Count + 1):
            folders.append(root.Folders.Item(i))
    except Exception:
        pass
    cutoff = datetime.datetime.now() - datetime.timedelta(days=limit_days)
    pool = []
    for fd in folders:
        try:
            items = fd.Items
            items.Sort("[ReceivedTime]", True)
        except Exception:
            continue
        n = 0
        for it in items:
            n += 1
            if n > per_folder:
                break
            try:
                rt = it.ReceivedTime.replace(tzinfo=None)
            except Exception:
                continue
            if rt < cutoff:
                break
            se = (R.get_sender(it) or "").lower()
            if se == sender.lower():
                pool.append((rt, it, fd.Name))
    pool.sort(key=lambda x: x[0], reverse=True)
    return pool


def main(names):
    sm = R.load_sender_map()
    for name in names:
        nv = R.norm(name)
        sender = sm.get(nv)
        print("\n=== %s  sender=%s ===" % (name, sender))
        if not sender:
            print("  无已知 sender, 跳过")
            continue
        pool = collect(sender)
        print("  匹配邮件(近45天):", len(pool))
        if not pool:
            continue
        hit = R.scan_for_rob([it for _, it, _ in pool], max_walk=80)
        if not hit:
            print("  >>> scan_for_rob 未取到 ROB <<<")
            rt, it, fn = pool[0]
            print("  最新一封: [%s] %s | %s" % (fn, rt, (it.Subject or "")[:60]))
            continue
        rob, recv, subj, se = hit
        print("  subject:", subj[:70])
        print("  recv   :", recv)
        print("  ROB    :", {k: v for k, v in rob.items() if not k.startswith("DRAFT")})
        print("  DRAFT  :", {k: v for k, v in rob.items() if k.startswith("DRAFT")})
        if not any(k.startswith("DRAFT") for k in rob):
            print("  >>> 未抓到任何 DRAFT 字段 <<<")


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        args = ["CUL HAIPHONG", "CUL NANSHA", "CUL XIAMEN", "CHANG SHENG JI 7"]
    main(args)
