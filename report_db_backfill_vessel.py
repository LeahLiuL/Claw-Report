# -*- coding: utf-8 -*-
"""
把报告库里空船名的报告, 用附件 VESSEL_NAME(code) 对照表回填船名。
只填空 vessel='' 的行, 不改任何已判定船名、不改任何字段/数值。
默认 dry-run, 加 --apply 才写盘。写前备份 reports.db。

对照依据: cul_daily_movement.html 下线船 code→vessel + 主题实测
  UGSZ/UGGZ = UGL SHENZHEN (主题 MV UGL SHENZHEN, UGGZ 3 封同主题)
  CUJCK/CUJK = CUL JAKARTA (已下线)
  HDHY = HDX HONGYUAN / GFMQ = GUO FU MIN QIANG / HOGE = HOOGE
  ZPYY = ZHONG PENG YOU YI / TBFZ = TB FENGZE / CSJ7 = CHANG SHENG JI 7
"""
import os
import sys
import sqlite3
import shutil
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "rob_data", "reports.db")

CODE2VESSEL = {
    "UGSZ": ("UGL SHENZHEN", "UGSZ"),
    "UGGZ": ("UGL SHENZHEN", "UGGZ"),
    "CUJCK": ("CUL JAKARTA", "CUJCK"),
    "CUJK": ("CUL JAKARTA", "CUJK"),
    "HDHY": ("HDX HONGYUAN", "HDHY"),
    "GFMQ": ("GUO FU MIN QIANG", "GFMQ"),
    "HOGE": ("HOOGE", "HOGE"),
    "ZPYY": ("ZHONG PENG YOU YI", "ZPYY"),
    "TBFZ": ("TB FENGZE", "TBFZ"),
    "CSJ7": ("CHANG SHENG JI 7", "CSJ7"),
}

APPLY = "--apply" in sys.argv
conn = sqlite3.connect(DB)

# 收集空船名报告 -> 其附件的 VESSEL_NAME code
rows = conn.execute(
    "SELECT r.mail_id, f.value_text FROM reports r "
    "JOIN fields f ON f.mail_id=r.mail_id "
    "WHERE r.vessel='' AND f.field='VESSEL_NAME'"
).fetchall()

fixes = []
unknown = {}
for mid, code in rows:
    code = str(code or "").strip().upper()
    if code in CODE2VESSEL:
        v, c = CODE2VESSEL[code]
        fixes.append((v, c, mid))
    else:
        unknown[code] = unknown.get(code, 0) + 1

print("空船名且附件有 code 的报告: %d" % len(rows))
print("可回填: %d" % len(fixes))
if unknown:
    print("未覆盖的 code: %s" % unknown)

from collections import Counter
stat = Counter((v, c) for v, c, _ in fixes)
print("\n回填分布:")
for (v, c), n in sorted(stat.items()):
    print("  %-22s %-8s %s 封" % (v, c, n))

if not APPLY:
    print("\nDRY-RUN: 未写盘。确认后加 --apply。")
    sys.exit(0)

bak = DB + ".bak_vessel_" + datetime.now().strftime("%Y%m%d-%H%M%S")
shutil.copy2(DB, bak)
print("\n已备份 -> %s" % bak)
conn.executemany("UPDATE reports SET vessel=?, code=? WHERE mail_id=? AND vessel=''", fixes)
conn.commit()
left = conn.execute("SELECT COUNT(*) FROM reports WHERE vessel=''").fetchone()[0]
print("已回填 %d 封, 剩余空船名 %d 封" % (len(fixes), left))
conn.close()
