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

输出: PIC汇总.xlsx(人工编辑主文件)。列: 航线|船名(显示)|重命名|文件夹名|船代码|PIC|状态

2026-11-06 改造(用户新增「重命名」列 + 船名变更需求):
  1. 读写一律【按表头名定位】, 不再硬编码 1..6 列 —— 用户在 B 后插入「重命名」列后,
     旧代码回写时会把整表压回 6 列, 每天静默清掉人工登记的曾用名。
  2. 写回时【沿用现有表头顺序】, 未知列(用户自己加的备注列等)原样保留每行原值。
  3. 改船名后文件夹名变了、但表里「文件夹名」列还没改时, 用「重命名」列自动认领该行
     (把文件夹名更新为新名), 避免同一艘船在表里裂成"旧名无目录 + 新名缺代码"两行。
"""
import os, glob, argparse
import openpyxl
from build_fleet_movement import (norm, canon_route, collect_vessel_folders, strip_retired,
                                  RETIRED_SUFFIX, load_vessel_csv, _load_pic_col, _pic_cols,
                                  DEFAULT_SRC, VESSEL_CSV, DEFAULT_PIC, UPD_DIR, GEN_DIR)

# 输出到 生成结果 子文件夹(随 UPD_DIR 自动切换 P:/Z:), 两机通用。
OUT_XLSX = os.path.join(GEN_DIR, "PIC汇总.xlsx")

# 表头缺失时的默认列(旧布局); 实际写回以现有表头为准
DEFAULT_HEADER = ["航线", "船名(显示)", "文件夹名", "船代码", "PIC", "状态"]


def latest_xlsx(folder):
    fs = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    if not fs: return None
    fs.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return fs[0]


def _header_key(h):
    """表头单元格 -> 内部字段名; 认不出返回 None(该列按"未知列"原样保留)。"""
    hs = str(h or "").upper().replace(" ", "")
    if not hs:
        return None
    if any(k in hs for k in ("重命名", "RENAME", "曾用名", "旧船名", "原名", "FORMERNAME", "OLDNAME")):
        return "rename"
    if "文件夹" in hs or hs.startswith("FOLDER"):
        return "folder"
    if "代码" in hs or "CODE" in hs:
        return "code"
    if "船名" in hs or "VESSEL" in hs or "SHIP" in hs or "DISPLAY" in hs:
        return "disp"
    if "PIC" in hs:
        return "pic"
    if "状态" in hs or "STATUS" in hs:
        return "status"
    if hs in ("航线", "SERVICE", "ROUTE", "LANE", "服务"):
        return "route"
    return None


def load_existing_table(path):
    """读已有人力维护的 PIC汇总.xlsx: (rows, header)
      rows   : key=norm(文件夹名) -> {字段..., "_raw": {列号: 原值}}
      header : [(列号, 表头原文, 字段名或None)] —— 写回时按此顺序, 不丢用户新增的列
    保留人工登记是这次改造的核心 —— 不在表里的新船才需要兜底补值。"""
    out, header = {}, []
    if not os.path.exists(path):
        return out, header
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        print(f"  [WARN] 读取现有 PIC汇总 失败({e}), 视为空表。")
        return out, header
    ws = wb[wb.sheetnames[0]]
    for c in range(1, ws.max_column + 1):
        h = ws.cell(1, c).value
        header.append((c, h, _header_key(h)))
    cols = {k: c for (c, _h, k) in header if k}
    fol_c = cols.get("folder")
    if not fol_c:
        print("  [WARN] 现有 PIC汇总 表头找不到「文件夹名」列, 视为空表(请检查第 1 行表头)")
        return {}, []
    def _g(r, key):
        c = cols.get(key)
        return ws.cell(r, c).value if c else None
    for r in range(2, ws.max_row + 1):
        fol = _g(r, "folder")
        if not fol or not str(fol).strip():
            continue
        k = norm(str(fol).strip())
        if not k:
            continue
        raw = {c: ws.cell(r, c).value for (c, _h, key) in header if key is None}
        out[k] = {"route":  str(_g(r, "route") or "").strip(),
                  "disp":   str(_g(r, "disp") or "").strip(),
                  "rename": str(_g(r, "rename") or "").strip(),
                  "folder": str(fol).strip(),
                  "code":   str(_g(r, "code") or "").strip(),
                  "pic":    str(_g(r, "pic") or "").strip(),
                  "status": str(_g(r, "status") or "").strip(),
                  "_raw": raw}
    return out, header


def claim_by_rename(existing, key, base):
    """改船名场景: 文件夹名改成新名后, 表里「文件夹名」列可能仍是旧名, 直接按 key 查不到。
    此时用「重命名」列认领 —— 登记值(规范化后)等于新文件夹名, 说明这一行就是它。
    认领成功则把该行 folder 更新为新名(下次按新名即可命中), 并返回该行。"""
    hit = None
    for k, ex in existing.items():
        names = [x.strip() for x in str(ex.get("rename") or "").replace("；", ";").replace("，", ",")
                 .replace("、", ";").replace("/", ";").split(";") if x.strip()]
        if any(norm(n) == key for n in names):
            hit = ex
            break
    if not hit:
        return None
    print(f"  [改名跟随] 表内旧行「{hit['folder']}」的重命名列含「{base}」→ 文件夹名更新为新名")
    hit["folder"] = base
    for k in [k for k, v in existing.items() if v is hit]:
        existing.pop(k, None)
    existing[key] = hit
    return hit



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
    existing, header = load_existing_table(args.pic)
    print(f"  现有 PIC汇总 条目: {len(existing)} 条 -> {args.pic}")

    fleet = collect_vessel_folders(args.src)   # 含 "已下线船舶/" 子目录
    rows, scanned, new_cnt = [], set(), 0
    for folder_path, fol, base, retired in fleet:
        p = latest_xlsx(folder_path)
        if not p: continue
        key = norm(base)
        scanned.add(key)
        ex = existing.get(key)
        # 改船名后文件夹名变了、表里还写旧文件夹名的情况: 用「重命名」列认领
        if not ex:
            ex = claim_by_rename(existing, key, base)
        # ── 表里有 -> 完全保留人工登记 ──
        if ex:
            rows.append({"route": ex["route"], "disp": ex["disp"], "rename": ex.get("rename", ""),
                         "folder": base, "code": ex["code"], "pic": ex["pic"],
                         "status": ex["status"] or ("已下线" if retired else "已有"),
                         "_raw": ex.get("_raw", {})})
            continue
        # ── 表里没有 -> 兜底补值并提醒人工复核 ──
        new_cnt += 1
        c1, c9 = read_src_meta(p, base)
        vent = vessel.get(key)
        code = (vent or {}).get("code") or (str(c9).strip() if c9 else "")
        disp = strip_retired((vent or {}).get("display") or strip_retired(fol))
        route = canon_route(base, str(c1).strip() if c1 else "")
        print(f"  [NEW] 新船待登记: {disp} (lane={route}, code={code or '?'}) —— 请人工复核代码后再使用")
        rows.append({"route": route, "disp": disp, "rename": "", "folder": base,
                     "code": code, "pic": "", "status": "已下线" if retired else "⚠ 缺失待补",
                     "_raw": {}})

    # 表里登记了、但磁盘目录里已找不到的船(含人工补登的下线船) —— 原样保留, 不静默删除
    kept_only = []
    for k, ex in existing.items():
        if k not in scanned:
            rows.append({"route": ex["route"], "disp": ex["disp"], "rename": ex.get("rename", ""),
                         "folder": ex["folder"], "code": ex["code"], "pic": ex["pic"],
                         "status": ex["status"] or "登记无目录", "_raw": ex.get("_raw", {})})
            kept_only.append(ex["folder"])

    rows.sort(key=lambda x: (norm(x["route"]), strip_retired(str(x["disp"]))))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PIC汇总"
    # 表头: 沿用现有表头(含用户新增的「重命名」列), 缺失的默认列补到末尾
    if header:
        hdr_cols = [(c, h, k) for (c, h, k) in header]
        have = {k for (_c, _h, k) in hdr_cols if k}
        for d in DEFAULT_HEADER:
            k = _header_key(d)
            if k and k not in have:
                hdr_cols.append((len(hdr_cols) + 1, d, k))
                have.add(k)
    else:
        hdr_cols = [(i, h, _header_key(h)) for i, h in enumerate(DEFAULT_HEADER, 1)]
    fill_h = openpyxl.styles.PatternFill("solid", fgColor="1F4E78")
    bold = openpyxl.styles.Font(bold=True, color="FFFFFF")
    thin = openpyxl.styles.Side(style="thin", color="BFBFBF")
    bd = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, (_src, h, _k) in enumerate(hdr_cols, 1):
        cell = ws.cell(1, c, h); cell.font = bold; cell.fill = fill_h; cell.border = bd
    miss_fill   = openpyxl.styles.PatternFill("solid", fgColor="FFF2CC")
    retired_fill = openpyxl.styles.PatternFill("solid", fgColor="EDEDED")
    for i, row in enumerate(rows, 2):
        for c, (src_c, _h, k) in enumerate(hdr_cols, 1):
            # 已知列按字段取值; 未知列(用户自建列)沿用该行原值
            v = row.get(k, "") if k else row.get("_raw", {}).get(src_c, "")
            cell = ws.cell(i, c, v); cell.border = bd
            if str(row["status"]).startswith("⚠"):
                cell.fill = miss_fill
            elif row["status"] == "已下线":
                cell.fill = retired_fill
    widths = {1: 10, 2: 24, 3: 14, 4: 22, 5: 10, 6: 18, 7: 14}
    for i in range(1, len(hdr_cols) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = widths.get(i, 16)

    os.makedirs(os.path.dirname(OUT_XLSX), exist_ok=True)
    try:
        wb.save(OUT_XLSX)
    except PermissionError:
        alt = OUT_XLSX[:-5] + ".new.xlsx"
        wb.save(alt)
        print(f"  [WARN] {OUT_XLSX} 被占用(可能Excel打开), 已写入 {alt}")

    miss = sum(1 for r in rows if str(r["status"]).startswith("⚠"))
    print(f"  表条目: {len(rows)} (本次新增 {new_cnt}) | 缺PIC待补: {miss} | 仅登记无目录: {len(kept_only)}")
    if kept_only:
        print("   仅登记在表、目录里没有的船(已保留): " + ", ".join(sorted(kept_only)[:10]))
    print(f"  -> {OUT_XLSX}")

if __name__ == "__main__":
    main()
