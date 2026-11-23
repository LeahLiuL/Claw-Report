# -*- coding: utf-8 -*-
"""
报告库查询工具 - 全部走本地 SQLite, 不连 Outlook。
后续任何统计/回刷/诊断都从这里取数, 不再逐封打开附件。

用法示例:
  python report_query.py --vessel "STAR BLESSING" --field ROB_LSFO
  python report_query.py --vessel "STAR BLESSING" --date 2026-07-22
  python report_query.py --pivot ROB_LSFO            # 各阶段ROB宽表
  python report_query.py --stage-summary             # BERTH 各阶段覆盖统计
  python report_query.py --sql "SELECT ..."          # 自定义SQL(只读)
  python report_query.py --fields                    # 列出全部字段名
"""
import os
import sys
import sqlite3
import argparse

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "rob_data", "reports.db")


def _c():
    return sqlite3.connect(DB)


def list_fields(conn):
    print("=== 库内字段 (field : 出现次数) ===")
    for r in conn.execute("SELECT field,COUNT(*) n FROM fields GROUP BY field ORDER BY n DESC"):
        print("  %-36s %s" % (r[0], r[1]))


def by_vessel(conn, vessel, field=None, date=None):
    sql = ("SELECT r.received, r.report_type, f.stage, f.field, "
           "COALESCE(f.value_num,f.value_text) val, f.unit "
           "FROM fields f JOIN reports r ON r.mail_id=f.mail_id "
           "WHERE r.vessel LIKE ?")
    args = ["%" + vessel + "%"]
    if field:
        sql += " AND f.field=?"; args.append(field)
    if date:
        sql += " AND r.received LIKE ?"; args.append(date + "%")
    sql += " ORDER BY r.received DESC"
    rows = conn.execute(sql, args).fetchall()
    print("=== %s %s %s (%d 行) ===" % (vessel, field or "", date or "", len(rows)))
    for r in rows:
        print("  %s | %-7s | %-16s | %-30s | %s%s" %
              (r[0], r[1] or "", r[2] or "-", r[3], r[4], r[5]))


def pivot(conn, field):
    """宽表: 每个 BERTH 报告的 POB/FWE/ANCHOR/DROP 各阶段该字段值, 一眼看填报异常。"""
    sql = ("SELECT r.received, r.vessel, r.report_type, "
           "MAX(CASE WHEN f.stage='POB' THEN COALESCE(f.value_num,f.value_text) END) pob, "
           "MAX(CASE WHEN f.stage='FWE' THEN COALESCE(f.value_num,f.value_text) END) fwe, "
           "MAX(CASE WHEN f.stage='DROP ANCHOR' THEN COALESCE(f.value_num,f.value_text) END) drop_a, "
           "MAX(CASE WHEN f.stage='ANCHOR AWEIGH' THEN COALESCE(f.value_num,f.value_text) END) anchor_a "
           "FROM fields f JOIN reports r ON r.mail_id=f.mail_id "
           "WHERE f.field=? GROUP BY r.mail_id ORDER BY r.received DESC LIMIT 60")
    rows = conn.execute(sql, (field,)).fetchall()
    print("=== %s 各阶段宽表 (最近 60) ===" % field)
    print("  %-20s %-18s %-7s %10s %10s %10s %10s" %
          ("received", "vessel", "type", "POB", "FWE", "DROP", "ANCHOR"))
    for r in rows:
        print("  %-20s %-18s %-7s %10s %10s %10s %10s" %
              (r[0], (r[1] or "")[:18], r[2] or "", r[3], r[4], r[5], r[6]))


def stage_summary(conn):
    print("=== 报告类型分布 ===")
    for r in conn.execute("SELECT report_type,COUNT(*) FROM reports GROUP BY report_type ORDER BY 2 DESC"):
        print("  %-10s %s" % (r[0] or "(空)", r[1]))
    print("=== BERTH 各阶段覆盖 ===")
    for r in conn.execute(
            "SELECT f.stage,COUNT(DISTINCT f.mail_id) FROM fields f "
            "JOIN reports r ON r.mail_id=f.mail_id "
            "WHERE r.report_type='BERTH' AND f.stage!='' GROUP BY f.stage ORDER BY 2 DESC"):
        print("  %-20s %s 封" % (r[0], r[1]))


def sql_run(conn, q):
    try:
        cur = conn.execute(q)
    except Exception as e:
        print("SQL 错误:", e); return
    cols = [d[0] for d in cur.description] if cur.description else []
    print(" | ".join(cols))
    for r in cur.fetchall()[:200]:
        print(" | ".join("" if v is None else str(v) for v in r))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--vessel")
    ap.add_argument("--field")
    ap.add_argument("--date")
    ap.add_argument("--pivot")
    ap.add_argument("--stage-summary", action="store_true")
    ap.add_argument("--fields", action="store_true")
    ap.add_argument("--sql")
    a = ap.parse_args()
    conn = _c()
    if a.fields:
        list_fields(conn)
    elif a.pivot:
        pivot(conn, a.pivot)
    elif a.stage_summary:
        stage_summary(conn)
    elif a.sql:
        sql_run(conn, a.sql)
    elif a.vessel:
        by_vessel(conn, a.vessel, a.field, a.date)
    else:
        ap.print_help()
    conn.close()
