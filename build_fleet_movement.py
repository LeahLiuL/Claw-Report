#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_fleet_movement.py  —— 重建版 CUL DAILY MOVEMENT 大Excel
=====================================================================
按当前 2026/ 各船文件夹(=当前船队, 增删自动反映)重建大Excel, 不再把旧大Excel当模板。

规则(用户2026-07-30确认):
  1. 港口行: 仅显示 ETB 在 今日±WINDOW天 窗口内的行 (默认±30)。
  2. PIC:    从旧大Excel按"文件夹名(=船名)"取; 新船无PIC则留空(手动补)。
  3. 船块顺序: 按 ROUTE_ORDER(由旧大Excel航线顺序推导, 可编辑) 分组;
             组内按船名(文件夹名)排序; 一船多航线(多文件夹)则显示多块。
  4. Remark: 取源第一个sheet中 ETB 最接近今日 那行的 C18,
             重建成 "Remark:航次 港口 原文"。

重要事实(已核查):
  - 源 R1 C4(船名) 经常为空/错 -> 船名以【文件夹名】为准。
  - 源 R1 C1(航线码) 与旧块头航线码大量改名 -> 用源R1C1作"当前航线",
    顺序用下方 ROUTE_ORDER 常量(旧位置推导, 可改)。
  - 源 R1 C1 可能为空(ZBM) -> 回退用旧航线码分组。

用法:
  python build_fleet_movement.py --src "P:\\...\\2026" --old "P:\\...\\CUL DAILY MOVEMENT.xlsx" --output "CUL DAILY MOVEMENT.rebuilt.xlsx"
  (culadmin 那台默认 Z: 盘, 直接 python build_fleet_movement.py)
"""
import argparse, os, glob, shutil
from datetime import datetime, date, timedelta
import openpyxl

DEFAULT_SRC = r"Z:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新\2026"
DEFAULT_OLD = r"Z:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新\CUL DAILY MOVEMENT.xlsx"
DEFAULT_OUT = r"CUL DAILY MOVEMENT.rebuilt.xlsx"

# 当前航线码 默认顺序(仅当旧大Excel读不到块顺序时作回退)。
# 实际分组顺序由 build_route_order() 按旧大Excel块顺序推导(更贴合你习惯的排法)。
ROUTE_ORDER = ["ST3","CHT","HDT","CST","NSX","NP2","REX","RTS","SGX","SL1",
               "CGX","HLX","CGS","AEM","AM1","ZGCD","IMR","NAX","CCT"]
ROUTE_FALLBACK = "RTS"   # 源航线码为空时回退
WINDOW_DAYS = 30

# 文件夹名 -> 旧块头船名 别名(让文件夹能匹配旧PIC/显示名)
FOLDER_ALIAS = {"HDX 728": "HONG DA XIN 728", "DONG FANG MING HAI": "DONG FANG MIN HAI"}
# 航线修正(用户2026-07-30确认): 文件夹 -> 规范航线码(覆盖源R1C1的改名/错误)
ROUTE_OVERRIDE = {"CUL NANSHA": "CCT"}     # 源R1C1误为HDT, 实为CCT
# 航线合并: 源航线码 -> 规范航线码(同一条航线在源里有不同叫法)
ROUTE_ALIAS = {"AM1": "AEM"}                # AEM 与 AM1 是同一航线

# 大Excel列头(沿用旧文件标签, 与源C1..C16位置一一对应)
COL_HEADERS = ["PORT","man in","wait","Proforma","ltm eta","ltm etd","VOY. NO",
               "date","ETA","ETB","ETD","run","Port Stay(hr)","fsp distance",
               "speed","ETA Delay/Ahead"]

SEGMENT_TITLE = "CUL VESSEL DAILY MOVEMENT  "

def norm(s): return (s or "").strip().upper().replace(" ", "")
def norm_voy(s):
    if s is None: return ""
    return str(s).strip().upper().replace(" ", "")
REV_ALIAS = {norm(v): k for k, v in FOLDER_ALIAS.items()}   # 旧显示名(规范) -> 文件夹名

def nearest_voy_upward(rows, idx):
    """ETB最近行航次号为空时, 向上(行号更小=表中更靠上)找最近的、有航次号的行。"""
    for j in range(idx - 1, -1, -1):
        if rows[j]["voy"]:
            return rows[j]["voy"]
    return ""

def latest_xlsx(folder):
    files = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    if not files: return None
    files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return files[0]

def read_source(path):
    """读源第一个sheet。返回 dict: route, code, rows[(c1..c16), etb, voy, port, remark]。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    route = ws.cell(1, 1).value
    code = ws.cell(1, 9).value
    # 找列头行
    hr = None
    for r in range(1, min(ws.max_row, 80) + 1):
        if norm(ws.cell(r, 1).value) == "PORT":
            hr = r; break
    rows = []
    if hr is None:
        return {"route": route, "code": code, "rows": rows}
    for r in range(hr + 1, ws.max_row + 1):
        c1 = ws.cell(r, 1).value
        if c1 is None: continue
        s1 = norm(c1)
        if s1 == "PORT": continue          # 多段航次表, 跳过重复列头继续读
        if s1 == "" or s1.startswith("REMARK"): continue
        vals = [ws.cell(r, c).value for c in range(1, 17)]   # C1..C16
        etb = ws.cell(r, 10).value
        rows.append({
            "vals": vals,
            "etb": etb if isinstance(etb, datetime) else None,
            "voy": norm_voy(ws.cell(r, 7).value),
            "port": s1,
            "remark": ws.cell(r, 18).value,
        })
    return {"route": route, "code": code, "rows": rows}

def load_old_ref(old_path):
    """从旧大Excel取: 船名(文件夹别名化)->(显示名, PIC), 以及块出现顺序 old_order(规范后用于分组排序)。"""
    wb = openpyxl.load_workbook(old_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    ref = {}
    order = []
    r = 1
    while r <= ws.max_row:
        if r >= ws.max_row: break
        if norm(ws.cell(r + 1, 1).value) != "PORT":
            r += 1; continue
        c4 = ws.cell(r, 4).value
        if not c4 or "VESSEL" in norm(c4):
            r += 1; continue
        ship = norm(c4)
        ref[ship] = {"display": str(c4).strip(), "pic": ws.cell(r, 16).value}
        if ship not in order:
            order.append(ship)
        r += 2
    return ref, order

def canon_route(folder, src_route):
    """算规范航线码: 先应用文件夹级覆盖, 再做同航线合并, 空则回退。"""
    if folder in ROUTE_OVERRIDE:
        r = ROUTE_OVERRIDE[folder]
    else:
        r = src_route
    r = ROUTE_ALIAS.get(norm(r), r)
    if norm(r) in ("", None):
        r = ROUTE_FALLBACK
    return r

def build_route_order(old_order, folder_to_canon):
    """分组顺序: 跟随旧大Excel块顺序(旧显示名->当前文件夹->规范航线码),
    旧模板没有的新船文件夹追加到末尾。"""
    seen = []
    for ship_norm in old_order:
        fol = REV_ALIAS.get(ship_norm, ship_norm)   # 旧显示名 -> 文件夹名
        r = folder_to_canon.get(norm(fol))
        if r and r not in seen:
            seen.append(r)
    for fol, r in folder_to_canon.items():           # 新船追加
        if r not in seen:
            seen.append(r)
    return seen

def main():
    ap = argparse.ArgumentParser(description="按当前船队重建 CUL DAILY MOVEMENT 大Excel")
    ap.add_argument("--src", default=DEFAULT_SRC)
    ap.add_argument("--old", default=DEFAULT_OLD, help="旧大Excel(取PIC/显示名)")
    ap.add_argument("--output", default=DEFAULT_OUT)
    ap.add_argument("--today", default=None, help="基准日 YYYY-MM-DD, 默认今天")
    ap.add_argument("--window", type=int, default=WINDOW_DAYS)
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()

    today = datetime.strptime(args.today, "%Y-%m-%d").date() if args.today else date.today()
    lo = today - timedelta(days=args.window)
    hi = today + timedelta(days=args.window)

    print(f"[基准日] {today}  窗口 ±{args.window}天 -> [{lo} ~ {hi}]")
    print("=== 1/4 读旧大Excel(PIC/显示名/块顺序) ===")
    old_ref, old_order = load_old_ref(args.old)
    print(f"  旧船块数(参考): {len(old_ref)}")

    print("=== 2/4 扫描当前船队(2026/文件夹) ===")
    folders = sorted([d for d in os.listdir(args.src)
                      if os.path.isdir(os.path.join(args.src, d)) and d != "已下线船舶"])
    ships = []   # {folder, route, code, rows}
    for fol in folders:
        p = latest_xlsx(os.path.join(args.src, fol))
        if not p:
            print(f"  [WARN] 无xlsx跳过: {fol}"); continue
        d = read_source(p)
        route = canon_route(fol, d["route"])   # 应用覆盖+合并
        ships.append({"folder": fol, "route": route, "code": d["code"], "rows": d["rows"]})
    print(f"  当前船文件夹数: {len(ships)}")

    print("=== 3/4 分组排序 + 写表 ===")
    # 分组
    groups = {}
    folder_to_canon = {}
    for s in ships:
        groups.setdefault(norm(s["route"]), []).append(s)
        folder_to_canon[norm(s["folder"])] = s["route"]
    # 顺序: 跟随旧大Excel块顺序(映射当前航线码), 新船追加末尾
    ordered_routes = build_route_order(old_order, folder_to_canon)
    if not ordered_routes:   # 旧模板读不到时回退到 ROUTE_ORDER
        ordered_routes = [r for r in ROUTE_ORDER if norm(r) in groups]
        ordered_routes += sorted([r for r in groups if norm(r) not in [norm(x) for x in ROUTE_ORDER]])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CUL DAILY MOVEMENT"
    # 列宽
    widths = [10,8,7,10,11,11,10,10,16,16,16,7,11,11,8,14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    row = 1
    seg_fill = openpyxl.styles.PatternFill("solid", fgColor="1F4E78")
    hdr_fill = openpyxl.styles.PatternFill("solid", fgColor="D9E1F2")
    bold = openpyxl.styles.Font(bold=True)
    thin = openpyxl.styles.Side(style="thin", color="BFBFBF")
    border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)

    def setc(r, c, v, font=None, fill=None, bd=False, merge_to=None):
        cell = ws.cell(r, c, v)
        if font: cell.font = font
        if fill: cell.fill = fill
        if bd: cell.border = border
        if merge_to:
            ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=merge_to)
        return cell

    blocks_written = 0
    for route in ordered_routes:
        grp = sorted(groups[norm(route)], key=lambda s: norm(s["folder"]))
        # 段标题(C1='#VALUE!' 与 C4=标题 为两个独立单元格, 不合并)
        setc(row, 1, "#VALUE!", font=bold, fill=seg_fill)
        setc(row, 4, SEGMENT_TITLE, font=bold, fill=seg_fill)
        row += 1
        for s in grp:
            fol = s["folder"]
            key = norm(FOLDER_ALIAS.get(fol, fol))
            disp = old_ref.get(key, {}).get("display") or fol
            pic = old_ref.get(key, {}).get("pic") or ""
            # 块头
            setc(row, 1, route, font=bold)
            setc(row, 4, disp, font=bold)
            setc(row, 8, "DATE")
            setc(row, 9, s["code"])
            setc(row, 16, pic)
            row += 1
            # 列头
            for c, h in enumerate(COL_HEADERS, 1):
                setc(row, c, h, font=bold, fill=hdr_fill, bd=True)
            row += 1
            # 港口行(窗口内)
            shown = 0
            for rr in s["rows"]:
                if rr["etb"] is None:
                    continue
                d = rr["etb"].date()
                if not (lo <= d <= hi):
                    continue
                for c, v in enumerate(rr["vals"], 1):
                    setc(row, c, v, bd=True)
                row += 1; shown += 1
            # Remark: ETB 最接近今日 那行; 航次号为空则向上找最近的有航次号的行
            best_idx = None; best_diff = None
            for i, rr in enumerate(s["rows"]):
                if rr["etb"] is None: continue
                diff = abs((rr["etb"].date() - today).days)
                if best_diff is None or diff < best_diff:
                    best_diff = diff; best_idx = i
            if best_idx is not None:
                best = s["rows"][best_idx]
                voy = best["voy"] or nearest_voy_upward(s["rows"], best_idx)
                rem = best["remark"]
                txt = f"Remark:{voy} {best['port']}" + (f" {rem}" if rem else "")
                setc(row, 1, txt)
            row += 1
            row += 1   # 块间空行
            blocks_written += 1
        row += 1   # 段间空行

    print("=== 4/4 保存 ===")
    if os.path.exists(args.output) and not args.no_backup:
        bak = args.output + ".bak_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(args.output, bak)
        print(f"  已备份 -> {bak}")
    wb.save(args.output)
    print(f"  航线组数: {len(ordered_routes)}  写出船块数: {blocks_written}")
    print(f"  输出: {args.output}")

if __name__ == "__main__":
    raise SystemExit(main())
