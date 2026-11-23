# -*- coding: utf-8 -*-
"""
补抓 Abstract Log 原件 (不碰 reports.db / ROB 库).

背景: report_db.py --collect 每封邮件只存一个附件, pick_report_attachment()
按 REPORT_KEYS 选, CII 在关键词里而 Abstract Log 不在 -> "ABSTRACT LOG & CII"
邮件归档的是 CII 附件, 真正的 Abstract Log 原件漏存.

本脚本:
  1. 从 reports.db 找主题/附件含 ABSTRACT 但解析失败的邮件(=归档错了附件)
  2. 连 Outlook vesselreport, 对这些邮件遍历所有附件, 找文件名含 ABSTRACT 的 xlsx
  3. 另存为 <mail_id>_abstract.xlsx, 写 override 映射 json
  4. abstract_log_db.py 解析时优先用 override 里的路径

用法:
  python abstract_collect_missing.py            # dry-run
  python abstract_collect_missing.py --apply    # 真正另存附件
"""
import os
import sys
import json
import time
import hashlib
import sqlite3
import argparse
import datetime
import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_DB = os.path.join(HERE, "rob_data", "reports.db")
ARCHIVE = os.path.join(HERE, "rob_data", "report_archive")
OVERRIDE = os.path.join(HERE, "rob_data", "abstract_file_override.json")


def _load(mod, name):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, mod))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    A = _load("abstract_log_db.py", "aldb")
    R = _load("report_db.py", "rdb")

    src = sqlite3.connect(SRC_DB)
    rows = src.execute("""SELECT mail_id,received,subject,attach_name,file_path FROM reports
        WHERE UPPER(COALESCE(subject,'')) LIKE '%ABSTRACT%'
           OR UPPER(COALESCE(attach_name,'')) LIKE '%ABSTRACT%'""").fetchall()
    need = {}
    for mid, rt, subj, att, fp in rows:
        res, _ = A.parse_report(fp) if fp and os.path.exists(fp) else (None, "")
        if res is None:
            need[mid] = {"received": rt, "subject": subj, "attach": att}
    print("需补抓 Abstract Log 原件: %d 封" % len(need))

    override = {}
    if os.path.exists(OVERRIDE):
        override = json.load(open(OVERRIDE, encoding="utf-8"))
        print("已有 override: %d 条" % len(override))

    import win32com.client as win32
    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    st = next((ns.Stores(i) for i in range(1, ns.Stores.Count + 1)
               if "vesselreport" in (ns.Stores(i).DisplayName or "").lower()), None)
    if st is None:
        print("找不到 vesselreport 邮箱"); return
    root = st.GetRootFolder()
    folders, stack = [root], [(root, 0)]
    while stack:
        f, d = stack.pop()
        if d >= 4:
            continue
        try:
            for ch in f.Folders:
                folders.append(ch)
                stack.append((ch, d + 1))
        except Exception:
            pass
    print("扫描 %d 个文件夹" % len(folders))

    hit = miss = 0
    miss_list = []
    t0 = time.time()
    for fld in folders:
        try:
            items = fld.Items
            sub = items.Restrict("[ReceivedTime] >= '01/01/2026 12:00 AM'")
        except Exception:
            continue
        for it in sub:
            try:
                mid = R._mail_id(it)
            except Exception:
                continue
            if mid not in need:
                continue
            # 找 Abstract Log 附件
            found = None
            for a in it.Attachments:
                fn = str(a.FileName or "")
                if fn.lower().endswith((".xlsx", ".xlsm")) and "ABSTRACT" in fn.upper():
                    found = (a, fn)
                    break
            if found is None:
                miss += 1
                miss_list.append((mid, need[mid]["subject"][:55], need[mid]["attach"][:45]))
                continue
            hit += 1
            a, fn = found
            if args.apply:
                rt = need[mid]["received"]
                ym = rt[:7].replace("-", "/") if len(rt) >= 7 else "unknown/00"
                dest_dir = os.path.join(ARCHIVE, ym)
                os.makedirs(dest_dir, exist_ok=True)
                dest = os.path.join(dest_dir, mid + "_abstract.xlsx")
                try:
                    a.SaveAsFile(dest)
                    override[mid] = dest
                except Exception as e:
                    print("  [SAVE-ERR] %s %s" % (mid[:10], str(e)[:50]))
            if hit <= 8:
                print("  [hit] %s %s -> %s" % (mid[:10], need[mid]["subject"][:38], fn[:40]))

    print("\n找到 Abstract Log 附件: %d, 未找到: %d (%.0fs)" % (hit, miss, time.time() - t0))
    if miss_list:
        print("\n--- 未找到的(可能邮件里确实没有 Abstract Log 附件) ---")
        for x in miss_list[:20]:
            print("   ", x[0][:10], "|", x[1], "|", x[2])
    if args.apply:
        json.dump(override, open(OVERRIDE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("\n已写 override -> %s (共 %d 条)" % (OVERRIDE, len(override)))
    else:
        print("\nDRY-RUN (未另存). 确认后加 --apply")
    src.close()


if __name__ == "__main__":
    main()
