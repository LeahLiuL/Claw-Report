r"""从 燃油添加日志.xlsx 抓取【加油量】(仅体积 MT, 不含价格) 生成 rob_data/bunkering.json。

数据源: 燃油添加日志.xlsx (多机共用同一份网络盘, 盘符不同内容相同):
  - 本机(leahliu): P:\04 上海操作中心\01 船期管理科\Bunker\加油Claw文件\  (原 '加油' 目录于 2026-09 改名为 'Bunker', 脚本双路径兼容)
  - 自动机(culadmin): Z:/... 同一目录 (自动探测 P: -> Z: -> 相对路径; 兼容 '加油'/'Bunker')
  - 列: 文件名, 船名, 航次, 加油地点, 加油时间, 申请人, 加油公司, 帐期,
        然后每种油: 价格/价格单位/数量/数量单位/合计 (VLSFO,MGO,HSFO,ULSFO,LNG,驳船费,CREDIT),
        付款总额, 已制单, 已提交, ...
  - 我们【只取数量(体积 MT)】, 不碰任何价格/金额/合计列 —— 满足『价格保密, 不上 GitHub』。

输出: rob_data/bunkering.json  (结构兼容 rob_refresh.build_html / 网页趋势/航次油耗)
       { "<船名>": { "<YYYY-MM-DD>": 体积MT(该日各油种数量之和), ... }, ... }

注: 同一船同一天可能有多行(不同油种分别记录), 按 (船名, 日期) 汇总。
"""
import os
import json
import csv
import openpyxl
from collections import defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(BASE, "rob_data")
HISTORY_CSV = os.path.join(ROB_DIR, "rob_history.csv")
_REL_OLD = r"04 上海操作中心\01 船期管理科\加油\加油Claw文件\燃油添加日志.xlsx"   # 旧目录名(部分机器仍用)
_REL_NEW = r"04 上海操作中心\01 船期管理科\Bunker\加油Claw文件\燃油添加日志.xlsx"  # 2026-09 起 '加油' 改名为 'Bunker'


def find_bunker_xlsx():
    """多机盘符自适应: 本机 P: -> culadmin Z: -> 环境变量覆盖; 兼容 '加油'/'Bunker' 目录改名。"""
    rels = [os.environ.get("BUNKER_REL"), _REL_OLD, _REL_NEW]
    rels = [r for r in rels if r]
    drives = [os.environ.get("BUNKER_DRIVE"), "P:\\", "Z:\\", ""]
    drives = [d for d in drives if d is not None]
    cand = [os.environ.get("BUNKER_XLSX")]     # 完整路径最高优先级
    for rel in rels:
        for drv in drives:
            cand.append(os.path.join(drv, rel) if drv else rel)
    for c in cand:
        if c and os.path.exists(c):
            return c
    return None


BUNKER_JSON = os.path.join(ROB_DIR, "bunkering.json")
BUNKER_TYPES_JSON = os.path.join(ROB_DIR, "bunkering_types.json")

# 我们要抓取体积的油种列(对应 ROB 的 LSFO/HSFO/MGO/ULSFO)
OIL_LABELS = ["VLSFO 0.5%", "MGO", "HSFO 3.5%", "ULSFO 0.1%"]
# 映射到 ROB 的 4 个油种键(供航次油耗按油种拆分展示)
OIL_KEY = {"VLSFO 0.5%": "ls", "HSFO 3.5%": "hs", "MGO": "mg", "ULSFO 0.1%": "us"}
OFFLINE_MARK = "已下线"


def is_offline(name):
    return OFFLINE_MARK in (name or "")


def find_qty_cols(header_row):
    """header_row 里每种油名所在列 ci, 数量列 = ci+2 (价格/价格单位/数量)。"""
    cols = {}
    for ci, val in enumerate(header_row):
        if val in OIL_LABELS:
            cols[val] = ci
    return cols


def build_bunkering(xlsx_path=None):
    xlsx_path = xlsx_path or find_bunker_xlsx()
    if not xlsx_path or not os.path.exists(xlsx_path):
        raise FileNotFoundError(
            "找不到 燃油添加日志.xlsx: 请确认网络盘(P:/Z:)已挂载, "
            "或用环境变量 BUNKER_XLSX 指定完整路径")
    print("source:", xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["燃油添加日志"]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3:
        print("[WARN] 燃油添加日志.xlsx 行数过少")
        return {}
    header = rows[0]                       # 油种名行
    # 子表头在 rows[1]: 价格/价格单位/数量/数量单位/合计
    qty_cols = find_qty_cols(header)
    print("抓取体积列:", {k: (v, v + 2) for k, v in qty_cols.items()})

    agg = defaultdict(lambda: defaultdict(float))   # vessel -> date -> MT(合计)
    # vessel -> date -> {"ls","hs","mg","us"} 分油种体积
    tagg = defaultdict(lambda: defaultdict(lambda: {"ls": 0.0, "hs": 0.0, "mg": 0.0, "us": 0.0}))
    detail = []                              # 明细 (船名, 航次, 地点, 日期, 体积)
    n_rows = 0
    for r in rows[2:]:
        vessel = (r[1] or "").strip() if len(r) > 1 else ""
        voyage = (r[2] or "").strip() if len(r) > 2 else ""
        location = (r[3] or "").strip() if len(r) > 3 else ""
        date = (r[4] or "").strip() if len(r) > 4 else ""
        if not vessel or not date:
            continue
        if is_offline(vessel):
            continue
        dd = date[:10]                       # 日期归一化为 YYYY-MM-DD
        total = 0.0
        for label, ci in qty_cols.items():
            qty = r[ci + 2] if len(r) > ci + 2 else None
            try:
                if qty not in (None, ""):
                    qv = float(qty)
                    total += qv
                    k = OIL_KEY.get(label)
                    if k:
                        tagg[vessel][dd][k] += qv
            except (TypeError, ValueError):
                pass
        if total <= 0:
            continue
        agg[vessel][dd] += round(total, 3)
        detail.append({"vessel": vessel, "voyage": voyage, "location": location,
                       "date": dd, "mt": round(total, 3)})
        n_rows += 1
    # 收敛为普通 dict
    out = {}
    for v, dmap in agg.items():
        out[v] = {d: round(m, 2) for d, m in dmap.items()}
    # 分油种收敛: {"船名": {"YYYY-MM-DD": {"ls":x,"hs":y,"mg":z,"us":w}}}
    types_out = {}
    for v, dmap in tagg.items():
        vd = {}
        for d, tm in dmap.items():
            tm2 = {k: round(val, 2) for k, val in tm.items() if val}
            if tm2:
                vd[d] = tm2
        if vd:
            types_out[v] = vd
    print("加油事件(明细行):", n_rows, " 涉及船:", len(out))
    return out, detail, types_out


def rob_vessels(history_csv=HISTORY_CSV):
    vs = set()
    if os.path.exists(history_csv):
        with open(history_csv, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                v = (row.get("vessel") or "").strip()
                if v and not is_offline(v):
                    vs.add(v)
    return vs


if __name__ == "__main__":
    out, detail, types_out = build_bunkering()
    os.makedirs(ROB_DIR, exist_ok=True)
    with open(BUNKER_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=0)
    with open(BUNKER_TYPES_JSON, "w", encoding="utf-8") as f:
        json.dump(types_out, f, ensure_ascii=False, indent=0)
    print("wrote", BUNKER_JSON, "(%d vessels)" % len(out))
    print("wrote", BUNKER_TYPES_JSON, "(%d vessels, split by oil type)" % len(types_out))
    rv = rob_vessels()
    overlap = [v for v in out if v in rv]
    print("与 ROB 历史船名重合:", len(overlap), "->", overlap[:20])
    if not overlap:
        print("[INFO] 无重合(船名命名可能不一致); 前 10 个 bunkering 船名:", list(out.keys())[:10])
        print("       前 10 个 ROB 船名:", list(rv)[:10])
