#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_pic.py —— 汇总/维护当前船队的 PIC 对照表(P盘大Excel同级)
- 船代码 / 显示名: 从 GitHub vessel.csv 取(权威)。
- PIC:    若 P盘 PIC汇总.xlsx(人工编辑主文件)已存在, 保留其中手工维护的PIC(不覆盖);
          缺失的船留空待补。
仅输出 PIC汇总.xlsx(人工编辑主文件); 不再生成 .csv 镜像。
【完全不依赖旧大Excel】。
"""
import os, glob, argparse
import openpyxl
from build_fleet_movement import (norm, ROUTE_OVERRIDE, ROUTE_ALIAS,
                                  canon_route, load_vessel_csv, load_pic,
                                  DEFAULT_SRC, VESSEL_CSV, DEFAULT_PIC, UPD_DIR, GEN_DIR)

# 输出到 生成结果 子文件夹(随 UPD_DIR 自动切换 P:/Z:), 两机通用。仅生成 xlsx(人工编辑主文件)。
OUT_XLSX = os.path.join(GEN_DIR, "PIC汇总.xlsx")

def latest_xlsx(folder):
    fs = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    if not fs: return None
    fs.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return fs[0]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=DEFAULT_SRC)
    ap.add_argument("--vessel", default=VESSEL_CSV)
    ap.add_argument("--pic", default=DEFAULT_PIC, help="已有PIC表(保留手工PIC)")
    args = ap.parse_args()
    vessel = load_vessel_csv(args.vessel)
    existing_pic = load_pic(args.pic)   # 从 xlsx 主文件保留手工维护的PIC
    src = args.src
    folders = sorted([d for d in os.listdir(src)
                      if os.path.isdir(os.path.join(src, d)) and d != "已下线船舶"])
    rows = []
    for fol in folders:
        p = latest_xlsx(os.path.join(src, fol))
        if not p: continue
        wb = openpyxl.load_workbook(p, data_only=True)
        ws = wb[wb.sheetnames[0]]
        src_route = ws.cell(1, 1).value
        route = canon_route(fol, src_route)
        key = norm(fol)
        vent = vessel.get(key)
        # 船代码: vessel.csv(权威) -> 源R1C9
        code = (vent or {}).get("code") or ws.cell(1, 9).value
        # 显示名: vessel.csv(权威) -> 文件夹名
        disp = ((vent or {}).get("display") or fol) if vent else fol
        # PIC: 已有PIC表(手工)优先; 缺失则留空待补
        pic = existing_pic.get(norm(fol)) or ""
        status = "已有" if pic else "⚠ 缺失待补"
        rows.append((route, disp, fol, code or "", pic, status))

    # 按航线、船名排序
    rows.sort(key=lambda x: (norm(x[0]), norm(x[1])))

    # 写 xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PIC汇总"
    hdr = ["航线", "船名(显示)", "文件夹名", "船代码", "PIC", "状态"]
    fill_h = openpyxl.styles.PatternFill("solid", fgColor="1F4E78")
    bold = openpyxl.styles.Font(bold=True, color="FFFFFF")
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(1, c, h); cell.font = bold; cell.fill = fill_h
        cell.border = openpyxl.styles.Border(*[openpyxl.styles.Side(style="thin", color="BFBFBF")]*4)
    miss_fill = openpyxl.styles.PatternFill("solid", fgColor="FFF2CC")
    for i, (route, disp, fol, code, pic, status) in enumerate(rows, 2):
        vals = [route, disp, fol, code, pic, status]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(i, c, v)
            cell.border = openpyxl.styles.Border(*[openpyxl.styles.Side(style="thin", color="BFBFBF")]*4)
            if status.startswith("⚠"):
                cell.fill = miss_fill
    widths = [10, 24, 22, 10, 18, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    os.makedirs(os.path.dirname(OUT_XLSX), exist_ok=True)
    try:
        wb.save(OUT_XLSX)
    except PermissionError:
        # 文件可能被 Excel 打开而锁定 -> 落到 .new.xlsx 避免覆盖失败
        alt = OUT_XLSX[:-5] + ".new.xlsx"
        wb.save(alt)
        print(f"  [WARN] {OUT_XLSX} 被占用(可能Excel打开), 已写入 {alt}")

    total = len(rows)
    miss = sum(1 for r in rows if r[5].startswith("⚠"))
    print(f"当前船队: {total} 艘 | 有PIC: {total-miss} | 缺失待补: {miss}")
    print(f"  -> {OUT_XLSX}")
    if miss:
        print("缺失PIC的船:")
        for r in rows:
            if r[5].startswith("⚠"):
                print(f"   {r[0]:<6} {r[1]}")

if __name__ == "__main__":
    main()
