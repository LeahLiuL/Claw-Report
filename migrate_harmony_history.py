# -*- coding: utf-8 -*-
"""
改名迁移(历史数据): EVERLASTING HARVEST/EVHA -> THE HARMONY/THMY
让 ROB 网页趋势图把改名前后当成同一艘船, 连续显示。
- rob_history.csv   : 网页 history 数据源(v/c 字段), 309 行改名
- draft_history.csv : 吃水历史(vessel 字段), 44 行改名
history/*.json 每日快照是归档(不被 build_html 读取), 保留原样不动。
有备份 _backup_rename_harmony_20261125/, 可回退。幂等。
"""
import os, csv, sys

BASE = os.path.dirname(os.path.abspath(__file__))
ROB = os.path.join(BASE, "rob_data")
OLD_V, NEW_V = "EVERLASTING HARVEST", "THE HARMONY"
OLD_C, NEW_C = "EVHA", "THMY"


def migrate(path, vcol, ccol=None):
    if not os.path.exists(path):
        print("  [skip] %s not found" % path)
        return 0
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return 0
    header = rows[0]
    vi = header.index(vcol) if vcol in header else None
    ci = header.index(ccol) if ccol and ccol in header else None
    n = 0
    for row in rows[1:]:
        if vi is not None and len(row) > vi and row[vi] == OLD_V:
            row[vi] = NEW_V
            n += 1
        if ci is not None and len(row) > ci and row[ci] == OLD_C:
            row[ci] = NEW_C
    with open(path, "w", encoding="utf-8", newline="") as f:
        csv.writer(f).writerows(rows)
    return n


def main():
    n1 = migrate(os.path.join(ROB, "rob_history.csv"), "vessel", "code")
    n2 = migrate(os.path.join(ROB, "draft_history.csv"), "vessel", None)
    print("rob_history.csv: %d rows renamed" % n1)
    print("draft_history.csv: %d rows renamed" % n2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
