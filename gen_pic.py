#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_pic.py —— 汇总/维护当前船队的 PIC 对照表
=======================================================
2026-09-30 改造要点(用户要求: 船代码 / 显示名 / PIC 都以 PIC汇总.xlsx 为唯一权威):
  1. 本脚本是【增量维护工具】, 不再整体重建覆盖 —— 现有 PIC汇总.xlsx 里人工维护的
     船代码(D列) / 显示名(B列) / PIC(E列) 一律保留, 只对扫描到的新船补行。
     (曾有 risk: 旧版每轮重跑会静默抹掉人工登记, 包括已下线船的补登记录)
  2. 扫描范围包含 "2026\\已下线船舶\\" —— 已下线船的数据要参与其它模块统计,
     必须从这张表里取到代码。
  3. 只在表里查不到时才用 vessel.csv / 源 R1C9 兜底, 并打 WARN 提示人工复核。

输出: PIC汇总.xlsx(人工编辑主文件)。列: 航线|船名(显示)|文件夹名|船代码|PIC|状态
"""
import os, glob, argparse
import openpyxl
from build_fleet_movement import (norm, canon_route, collect_vessel_folders, strip_retired,
                                  RETIRED_SUFFIX, load_vessel_csv, _load_pic_col,
                                  DEFAULT_SRC, VESSEL_CSV, DEFAULT_PIC, UPD_DIR, GEN_DIR)

# 输出到 生成结果 子文件夹(随 UPD_DIR 自动切换 P:/Z:), 两机通用。
OUT_XLSX = os.path.join(GEN_DIR, "PIC汇总.xlsx")


def latest_xlsx(folder):
    fs = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    if not fs: return None
    fs.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return fs[0]


def load_existing_table(path):
    """读已有人力维护的 PIC汇总.xlsx 全行: key = norm(文件夹名) -> 各列原始值。
    保留人工登记是这次改造的核心 —— 不在表里的新船才需要兜底补值。"""
    out = {}
    if not os.path.exists(path):
        return out
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        print(f"  [WARN] 读取现有 PIC汇总 失败({e}), 视为空表。")
        return out
    ws = wb[wb.sheetnames[0]]
    for r in range(2, ws.max_row + 1):
        route = ws.cell(r, 1).value
        disp  = ws.cell(r, 2).value
        fol   = ws.cell(r, 3).value
        code  = ws.cell(r, 4).value
        pic   = ws.cell(r, 5).value
        stat  = ws.cell(r, 6).value
        if not fol or not str(fol).strip():
            continue
        k = norm(str(fol).strip())
        if not k:
            continue
        out[k] = {"route": str(route or "").strip(),
                  "disp":  str(disp or "").strip(),
                  "folder": str(fol).strip(),
                  "code":  str(code or "").strip(),
                  "pic":   str(pic or "").strip(),
                  "status": str(stat or "").strip()}
    return out


def read_src_meta(path, base_name):
    """取源文件的航线与船代码候选(仅新船兜底用)。
    用 read_source 的完整解析(含段标题检测 + lanes_in_file 继承), 避免只取 R1C1
    得到港口码/日期之类的脏值(如 TB FENGZE 曾取到 INNSA)。"""
    from build_fleet_movement import read_source
    try:
        d = read_source(path, vessel_code=None, folder_name=base_name)
        # 优先取任意一行所属的真实航线(row_route 已做段标题识别与继承)
        routes = [rr.get("row_route") for rr in d.get("rows", []) if rr.get("row_route")]
        return (routes[0] if routes else d.get("route", "")), d.get("code")
    except Exception as e:
        print(f"  [WARN] 读源失败跳过 lane 解析: {base_name} ({e})")
        return "", None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=DEFAULT_SRC)
    ap.add_argument("--vessel", default=VESSEL_CSV)
    ap.add_argument("--pic", default=DEFAULT_PIC, help="已有PIC表(保留人工维护的代码/显示名/PIC)")
    args = ap.parse_args()

    vessel = load_vessel_csv(args.vessel)
    existing = load_existing_table(args.pic)
    print(f"  现有 PIC汇总 条目: {len(existing)} 条 -> {args.pic}")

    fleet = collect_vessel_folders(args.src)   # 含 "已下线船舶/" 子目录
    rows, scanned, new_cnt = [], set(), 0
    for folder_path, fol, base, retired in fleet:
        p = latest_xlsx(folder_path)
        if not p: continue
        key = norm(base)
        scanned.add(key)
        ex = existing.get(key)
        # ── 表里有 -> 完全保留人工登记 ──
        if ex:
            rows.append((ex["route"], ex["disp"], ex["folder"], ex["code"], ex["pic"],
                         ex["status"] or ("已下线" if retired else "已有")))
            continue
        # ── 表里没有 -> 兜底补值并提醒人工复核 ──
        new_cnt += 1
        c1, c9 = read_src_meta(p, base)
        vent = vessel.get(key)
        code = (vent or {}).get("code") or (str(c9).strip() if c9 else "")
        disp = strip_retired((vent or {}).get("display") or strip_retired(fol))
        route = canon_route(base, str(c1).strip() if c1 else "")
        print(f"  [NEW] 新船待登记: {disp} (lane={route}, code={code or '?'}) —— 请人工复核代码后再使用")
        rows.append((route, disp, base, code, "", "已下线" if retired else "⚠ 缺失待补"))

    # 表里登记了、但磁盘目录里已找不到的船(含人工补登的下线船) —— 原样保留, 不静默删除
    kept_only = []
    for k, ex in existing.items():
        if k not in scanned:
            rows.append((ex["route"], ex["disp"], ex["folder"], ex["code"], ex["pic"],
                         ex["status"] or "登记无目录"))
            kept_only.append(ex["folder"])

    rows.sort(key=lambda x: (norm(x[0]), strip_retired(str(x[1]))))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PIC汇总"
    hdr = ["航线", "船名(显示)", "文件夹名", "船代码", "PIC", "状态"]
    fill_h = openpyxl.styles.PatternFill("solid", fgColor="1F4E78")
    bold = openpyxl.styles.Font(bold=True, color="FFFFFF")
    thin = openpyxl.styles.Side(style="thin", color="BFBFBF")
    bd = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(1, c, h); cell.font = bold; cell.fill = fill_h; cell.border = bd
    miss_fill   = openpyxl.styles.PatternFill("solid", fgColor="FFF2CC")
    retired_fill = openpyxl.styles.PatternFill("solid", fgColor="EDEDED")
    for i, (route, disp, fol, code, pic, status) in enumerate(rows, 2):
        for c, v in enumerate([route, disp, fol, code, pic, status], 1):
            cell = ws.cell(i, c, v); cell.border = bd
            if status.startswith("⚠"):
                cell.fill = miss_fill
            elif status == "已下线":
                cell.fill = retired_fill
    for i, w in enumerate([10, 24, 22, 10, 18, 14], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    os.makedirs(os.path.dirname(OUT_XLSX), exist_ok=True)
    try:
        wb.save(OUT_XLSX)
    except PermissionError:
        alt = OUT_XLSX[:-5] + ".new.xlsx"
        wb.save(alt)
        print(f"  [WARN] {OUT_XLSX} 被占用(可能Excel打开), 已写入 {alt}")

    miss = sum(1 for r in rows if str(r[5]).startswith("⚠"))
    print(f"  表条目: {len(rows)} (本次新增 {new_cnt}) | 缺PIC待补: {miss} | 仅登记无目录: {len(kept_only)}")
    if kept_only:
        print("   仅登记在表、目录里没有的船(已保留): " + ", ".join(sorted(kept_only)[:10]))
    print(f"  -> {OUT_XLSX}")

if __name__ == "__main__":
    main()
