#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_pic.py —— 汇总当前船队的 PIC 对照表
按当前 2026/ 各船文件夹(=当前船队)生成 航线|船名|船代码|PIC|状态 表,
PIC 取自旧大Excel(按文件夹名=船名匹配); 旧表里没有的新船标 "⚠ 缺失待补"。
输出 PIC汇总.xlsx 与 PIC汇总.csv。
"""
import os, glob, csv, argparse
import openpyxl
from build_fleet_movement import (norm, FOLDER_ALIAS, ROUTE_OVERRIDE, ROUTE_ALIAS,
                                  canon_route, load_old_ref, DEFAULT_SRC, DEFAULT_OLD)

OUT_XLSX = r"P:/04 上海操作中心/01 船期管理科/船期管理/VSL Daily Movement/更新\PIC汇总.xlsx"
OUT_CSV  = r"P:/04 上海操作中心/01 船期管理科/船期管理/VSL Daily Movement\更新\PIC汇总.csv"

def latest_xlsx(folder):
    fs = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    if not fs: return None
    fs.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return fs[0]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=DEFAULT_SRC)
    ap.add_argument("--old", default=DEFAULT_OLD)
    args = ap.parse_args()
    old_ref, _ = load_old_ref(args.old)
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
        code = ws.cell(1, 9).value
        route = canon_route(fol, src_route)
        key = norm(FOLDER_ALIAS.get(fol, fol))
        ref = old_ref.get(key, {})
        pic_raw = ref.get("pic") or ""
        pic = str(pic_raw).replace("PIC:", "").replace("PIC :", "").strip()
        disp = ref.get("display") or fol
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
    wb.save(OUT_XLSX)

    # 写 csv
    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(hdr)
        for r in rows:
            w.writerow(r)

    total = len(rows)
    miss = sum(1 for r in rows if r[5].startswith("⚠"))
    print(f"当前船队: {total} 艘 | 有PIC: {total-miss} | 缺失待补: {miss}")
    print(f"  -> {OUT_XLSX}")
    print(f"  -> {OUT_CSV}")
    if miss:
        print("缺失PIC的船:")
        for r in rows:
            if r[5].startswith("⚠"):
                print(f"   {r[0]:<6} {r[1]}")

if __name__ == "__main__":
    main()
