# -*- coding: utf-8 -*-
"""船名变更(rename)端到端回归 —— 覆盖用户真实会做的 3 种操作, 结果必须一致。

背景: 用户改船名的方式是【只把磁盘上的文件夹名改成新名】, 源 Excel 里旧航段
仍写旧船名、之后的新航段改新名。PIC汇总.xlsx 的「重命名」列用来登记旧名。

  S1 只改 folder 名 (PIC表/源文件都不动)                  -> gen_pic 规则B 认领
  S2 改 folder + PIC 的「文件夹名」「显示名」, 重命名列填旧名 -> 按文件夹名直接命中
  S3 改 folder, 重命名列填新名 (表里文件夹名仍是旧名)     -> gen_pic 规则A 认领

每个场景检查:
  ① PIC 表最终只有一行 (不裂成"旧名无目录 + 新名缺代码")
  ② 船代码 / PIC 人 / 显示名 = 新名, 旧名进了「重命名」列
  ③ build 出两个船块: 旧块 C4=旧名、新块 C4=新名, 两块 C16 的 VID 相同
  ④ 页面 SUMMARY 只有这一艘船, 但全船期里旧/新两个名字都在

用法: python test_rename_regression.py
"""
import os
import re
import sys
import json
import shutil
import subprocess
from datetime import datetime

import openpyxl

REPO = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, REPO)
import gen_pic                                        # noqa: E402
from build_fleet_movement import norm, strip_retired  # noqa: E402

PY = sys.executable
BASE = os.path.join(REPO, ".cache", "renametest")
SRC = os.path.join(BASE, "src")                       # 旧段旧名 + 新段新名
SRC_NEW_ONLY = os.path.join(BASE, "src_newonly")      # 源文件已完全切成新名
OLD, NEW, CODE, PICGUY = "TEST OLD NAME", "TEST NEW NAME", "TNMN", "Leah Liu"
HDR = ["PORT", "man in", "wait", "Proforma", "ltm eta", "ltm etd", "VOY. NO",
       "date", "ETA", "ETB", "ETD", "run", "Port Stay(hr)", "fsp distance",
       "speed", "ETA Delay/Ahead"]
FAIL = []


def check(cond, msg):
    print(("  PASS  " if cond else "  FAIL  ") + msg)
    if not cond:
        FAIL.append(msg)


def make_src(root, new_only=False):
    """旧段(AEM, 2026-10)写旧名, 新段(REX, 2026-11/12)写新名 —— 与用户描述一致。
    new_only=True 时连旧段也写新名(模拟"源文件已全部改成新名, 只剩船代码能对身份")。"""
    shutil.rmtree(root, ignore_errors=True)
    os.makedirs(os.path.join(root, NEW))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Official"

    def seg(row, lane, c4, c9):
        ws.cell(row, 1, lane)
        ws.cell(row, 4, c4)
        ws.cell(row, 9, c9)

    def data(row, port, voy, eta, etb, etd):
        ws.cell(row, 1, port)
        ws.cell(row, 7, voy)
        ws.cell(row, 9, eta)
        ws.cell(row, 10, etb)
        ws.cell(row, 11, etd)

    seg(1, "AEM", NEW if new_only else OLD, "TOLD")
    for i, h in enumerate(HDR, 1):
        ws.cell(2, i, h)
    data(3, "CNSHA", "2601E", datetime(2026, 10, 1, 8), datetime(2026, 10, 1, 12), datetime(2026, 10, 2, 6))
    data(4, "CNNGB", "2601E", datetime(2026, 10, 6, 8), datetime(2026, 10, 6, 12), datetime(2026, 10, 7, 6))
    seg(6, "REX", NEW, "TNMN")
    for i, h in enumerate(HDR, 1):
        ws.cell(7, i, h)
    data(8, "CNSHA", "2602W", datetime(2026, 11, 20, 8), datetime(2026, 11, 20, 12), datetime(2026, 11, 21, 6))
    data(9, "CNXMN", "2602W", datetime(2026, 11, 25, 8), datetime(2026, 11, 25, 12), datetime(2026, 11, 26, 6))
    data(10, "MYPKG", "2602W", datetime(2026, 12, 1, 8), datetime(2026, 12, 1, 14), datetime(2026, 12, 2, 6))
    wb.save(os.path.join(root, NEW, "TEST-2026Y.xlsx"))


def make_pic(path, folder, disp=OLD, rename=""):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PIC汇总"
    ws.append(["航线", "船名(显示)", "重命名", "文件夹名", "船代码", "PIC", "状态"])
    ws.append(["AEM", disp, rename, folder, CODE, PICGUY, "已有"])
    wb.save(path)


def run_gen_pic(pic_in, pic_out, root):
    gen_pic.OUT_XLSX = pic_out
    argv = sys.argv
    sys.argv = ["gen_pic.py", "--src", root, "--pic", pic_in]
    try:
        gen_pic.main()
    finally:
        sys.argv = argv
    ws = openpyxl.load_workbook(pic_out, data_only=True).active
    hdr = [str(ws.cell(1, c).value or "") for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        rows.append({hdr[c - 1]: ws.cell(r, c).value for c in range(1, len(hdr) + 1)})
    return rows


def run_build(pic, out, root):
    r = subprocess.run([PY, os.path.join(REPO, "build_fleet_movement.py"),
                        "--src", root, "--pic", pic, "--output", out, "--no-backup"],
                       capture_output=True, text=True, cwd=REPO)
    if r.returncode != 0:
        print(r.stdout[-2000:], r.stderr[-2000:])
    return r


def blocks(path):
    """[(route, c4_name, code, vid, cur, n_rows)]"""
    ws = openpyxl.load_workbook(path, data_only=True).active
    out, i = [], 1
    while i <= ws.max_row:
        c16 = ws.cell(i, 16).value
        if c16 and isinstance(c16, str) and "PIC" in c16:
            meta = {}
            for seg in str(c16).split("|")[1:]:
                if "=" in seg:
                    k, v = seg.split("=", 1)
                    meta[k.strip().upper()] = v.strip()
            n = 0
            j = i + 2
            while j <= ws.max_row:
                c16j = ws.cell(j, 16).value
                if c16j and isinstance(c16j, str) and "PIC" in c16j:
                    break
                if isinstance(ws.cell(j, 10).value, datetime):
                    n += 1
                j += 1
            out.append((str(ws.cell(i, 1).value), str(ws.cell(i, 4).value), str(ws.cell(i, 9).value),
                        meta.get("VID", ""), meta.get("CUR", ""), n, str(c16)))
            i = j
        else:
            i += 1
    return out


def run_gen_html(xlsx, html):
    r = subprocess.run([PY, os.path.join(REPO, "gen_html.py"), "--excel", xlsx, "--out", html],
                       capture_output=True, text=True, cwd=REPO)
    if r.returncode != 0:
        print(r.stdout[-2000:], r.stderr[-2000:])
    h = open(os.path.join(REPO, html), encoding="utf-8").read()
    m = re.search(r"const TODAY_DATA = (\{.*?\});\s*\n", h, re.S)
    return json.loads(m.group(1)) if m else {}


def scenario(tag, pic_kwargs, root=None, blocks_expect=2, expect_names=None):
    expect_names = expect_names or {OLD, NEW}
    root = root or SRC
    print(f"\n=== {tag} ===")
    pic_in = os.path.join(BASE, f"pic_{tag}.xlsx")
    pic_out = os.path.join(BASE, f"pic_{tag}_out.xlsx")
    make_pic(pic_in, **pic_kwargs)
    rows = run_gen_pic(pic_in, pic_out, root)
    check(len(rows) == 1, f"PIC 表只有 1 行 (实际 {len(rows)}: {[r.get('文件夹名') for r in rows]})")
    if not rows:
        return
    r0 = rows[0]
    check(str(r0.get("文件夹名") or "").strip() == NEW, f"文件夹名 = {NEW} (实际 {r0.get('文件夹名')})")
    check(str(r0.get("船名(显示)") or "").strip() == NEW, f"显示名 = {NEW} (实际 {r0.get('船名(显示)')})")
    check(str(r0.get("船代码") or "").strip() == CODE, f"船代码未丢 = {CODE} (实际 {r0.get('船代码')})")
    check(str(r0.get("PIC") or "").strip() == PICGUY, f"PIC 未丢 = {PICGUY} (实际 {r0.get('PIC')})")
    ren = str(r0.get("重命名") or "")
    check(norm(OLD) in {norm(x) for x in re.split(r"[;；,，/、|]+", ren) if x.strip()},
          f"旧名已补进「重命名」列 (实际「{ren}」)")

    xls = os.path.join(BASE, f"built_{tag}.xlsx")
    run_build(pic_out, xls, root)
    bl = blocks(xls)
    names = [(b[1], b[4]) for b in bl]
    check(len(bl) == blocks_expect, f"拆成 {blocks_expect} 个船块 (实际 {names})")
    if len(bl) == blocks_expect and bl:
        got = {strip_retired(b[1]) for b in bl}
        check(got == expect_names, f"块头船名 = {sorted(expect_names)} (实际 {sorted(got)})")
        check(len({b[3] for b in bl}) == 1 and bl[0][3] == NEW,
              f"所有块 VID 相同且=船身份 (实际 {[b[3] for b in bl]})")
        check({b[4] for b in bl} == {NEW}, f"所有块 CUR 都是当前名 (实际 {[b[4] for b in bl]})")
    d = run_gen_html(xls, os.path.join(".cache", f"renametest/page_{tag}.html"))
    vs = d.get("vessels") or []
    fs = d.get("fullSchedule") or []
    check(len(vs) == 1, f"Summary 只有 1 艘船 (实际 {[(v.get('vessel'), v.get('vesselCur')) for v in vs]})")
    if len(expect_names) >= 2:
        check(len({(f.get("vessel"), f.get("vesselCur"), f.get("vid")) for f in fs}) >= 2,
              "全船期里旧名/新名都出现")
    if vs:
        check(vs[0].get("vesselCur") == NEW, f"Summary 当前名 = {NEW} (实际 {vs[0].get('vesselCur')})")


def main():
    print("构造测试源:", BASE)
    make_src(SRC)
    make_src(SRC_NEW_ONLY, new_only=True)
    scenario("S1", dict(folder=OLD))                             # 只改 folder
    scenario("S2", dict(folder=NEW, disp=NEW, rename=OLD))       # 表也改好
    scenario("S3", dict(folder=OLD, disp=OLD, rename=NEW))       # 重命名列填新名
    # 源文件已完全切成新名: 段标题里没有旧名, 只能靠船代码认领
    scenario("S4", dict(folder=OLD), root=SRC_NEW_ONLY, blocks_expect=2, expect_names={NEW})
    print("\n==== 结果:", "全部通过" if not FAIL else f"{len(FAIL)} 项失败")
    for f in FAIL:
        print("  -", f)
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
