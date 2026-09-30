#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_fleet_movement.py  —— 重建版 CUL DAILY MOVEMENT 大Excel
=====================================================================
按当前 2026/ 各船文件夹(=当前船队, 增删自动反映)重建大Excel。
【完全不依赖旧大Excel】——所有权威信息来自:
  - vessel.csv   (GitHub 仓库内, 船名<->代码<->显示名; 两机 git pull 同步, 用户手动维护)
  - P盘 PIC汇总.xlsx (人工维护的 PIC 对照表)

规则(用户2026-07-30确认):
  1. 港口行: 仅显示 ETB 在 今日±WINDOW天 窗口内的行 (默认±30)。
  2. 船名代码(code, 块头C9) / 显示名(C4): 从 vessel.csv 取(权威);
     仅当 vessel.csv 缺该船时, code 回退 源R1C9、显示名回退 文件夹名(并报警提示补登)。
  3. PIC(块头C16): 从 P盘 PIC汇总.xlsx 取(按文件夹名); 取不到则留空待补。
  4. 船块顺序: 按 ROUTE_ORDER 常量(固化20组, 可编辑) 分组; 组内按船名(文件夹名)排序;
     新航线追加到末尾(字母序)。
  5. Remark: 取源第一个sheet中 ETB 最接近今日 那行的 C18, 重建成 "Remark:航次 港口 原文";
     若该行为空(无remark)则【整行不写】。

重要事实(已核查):
  - 源 R1 C4(船名) 经常为空/错 -> 显示名以 vessel.csv 的 ship 列为准, 缺则文件夹名。
  - 源 R1 C1(航线码) 与规范航线码大量改名 -> 用 ROUTE_OVERRIDE / ROUTE_ALIAS 校正。
  - 源 R1 C1 可能为空(ZBM) -> 回退 ROUTE_FALLBACK。
  - 源 R1 C9(船代码) 布局不统一(ASR代码在C12) -> 以 vessel.csv 为准。


用法:
  python build_fleet_movement.py --src "P:\\...\\2026" --output "CUL DAILY MOVEMENT.rebuilt.xlsx"
  (culadmin 那台默认 Z: 盘, 直接 python build_fleet_movement.py)
"""
import argparse, os, glob, shutil, csv, re
from datetime import datetime, date, timedelta
import openpyxl

# 数据更新目录: 本机(leahliu)=P:, 另一台(culadmin)=Z:。自动探测存在的盘符, 两机通用, 无需传参。
_BASES = [
    r"Z:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新",
    r"P:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新",
]
UPD_DIR = next((b for b in _BASES if os.path.isdir(b)), _BASES[0])
# 脚本生成的输出统一放这个子文件夹(本机P:/另一台Z: 自动切换), 与源数据 2026/ 分开。
GEN_SUBDIR = "生成结果"
GEN_DIR = os.path.join(UPD_DIR, GEN_SUBDIR)
DEFAULT_SRC = os.path.join(UPD_DIR, "2026")
DEFAULT_OUT = os.path.join(GEN_DIR, "CUL DAILY MOVEMENT.rebuilt.xlsx")
# 船名<->代码<->显示名 权威表(GitHub 仓库内, 两机 git pull 同步); 用户在此手动维护。
VESSEL_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vessel.csv")
# PIC 权威表: 生成结果子文件夹内的 PIC汇总.xlsx (人工编辑主文件, 仅 xlsx)。两机通用。
DEFAULT_PIC = os.path.join(GEN_DIR, "PIC汇总.xlsx")

# 航线分组顺序(固化常量, 与当前网页/大Excel展示一致; 新增航线追加到末尾)。
# 组内按船名(文件夹名)排序。如要调整顺序, 改这里即可。
ROUTE_ORDER = ["ST3","CHT","HDT","CST","CCT","NP2","REX","CGS","AEM","EVHA",
               "SGX","SJA","JPS","SHTG","RTS","WAT","KCI","GTS","NSX","SL1","CGX","HLX","IMR","NAX","RES","NSCT1"]
# 注: WAT(KR CELEBES)/KCI(RACINE)/GTS(TB JINJIANG) 为用户确认的真实服务名, 已并入上表。
# 未来出现的新 lane 由"段标题结构位置"(PIC 块头/含船名船代码的段头)识别并原样保留,
# 并经 build_route_order / 前端 _buildLaneList 自动追加到排序末尾。
ROUTE_FALLBACK = ""      # 源无任何段标题行时才留空(不再静默猜 RTS; 空航线由用户补)
WINDOW_DAYS = 30

# 航线修正(用户2026-07-30确认): 文件夹 -> 规范航线码(覆盖源R1C1的改名/错误)
ROUTE_OVERRIDE = {"CUL NANSHA": "CCT"}     # 源R1C1误为HDT, 实为CCT
# 航线合并: 源航线码 -> 规范航线码(同一条航线在源里有不同叫法)
ROUTE_ALIAS = {"AM1": "AEM"}                # AEM 与 AM1 是同一航线

# ── 已下线船舶(退租/下线) ──
# 源目录里有专门的子文件夹存放下线船船期(历史数据, 其他模块的统计需要用到)。
# 命名不统一: "CUL JAKARTA-已下线" / "已下线-CUL JAKARTA" 两种写法可能是同一艘船,
# 故按【去掉已下线标记的规范船名】去重, 保留 xlsx 修改时间最新的那份。
# 显示名统一补 "-已下线" 后缀 —— 网页端 DECOM_MARK 靠这个标记识别下线船。
RETIRED_DIR_NAME = "已下线船舶"
RETIRED_SUFFIX = "-已下线"
# 下线船代码兜底表(vessel.csv / PIC汇总 均未登记时使用);
# 表里没有的仍回退 read_source 取源 R1C9, 取不到会在扫描时 WARN 提示补登。
RETIRED_VESSEL_CODES = {
    "CULJAKARTA": "CUJK",
    "ZHONGGUCHENGDU": "ZGCD",
    "GUOFUMINQIANG": "GFMQ",
    "TBFENGZE": "TBFE",
}

# 大Excel列头(沿用旧文件标签, 与源C1..C16位置一一对应)
COL_HEADERS = ["PORT","man in","wait","Proforma","ltm eta","ltm etd","VOY. NO",
               "date","ETA","ETB","ETD","run","Port Stay(hr)","fsp distance",
               "speed","ETA Delay/Ahead"]

SEGMENT_TITLE = "CUL VESSEL DAILY MOVEMENT  "

def norm(s): return str(s or "").strip().upper().replace(" ", "")
def norm_voy(s):
    if s is None: return ""
    return str(s).strip().upper().replace(" ", "")

# ── 表头对齐: 大Excel目标列 -> 源文件表头候选(归一化名) ──
# 源文件列顺序不统一(如ASR多一列TERMINAL把VOY.NO推到C8), 故按"表头名"映射而非固定列位。
def norm_h(s):
    return (s or "").strip().upper().replace(" ", "")

def excel_serial_to_dt(v):
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(v))
    except Exception:
        return None

def norm_date_value(v):
    """真实日期 -> datetime(统一); 文本标记(OMIT/Mon/...) -> 原字符串; 空 -> None。"""
    if v is None: return None
    if isinstance(v, datetime): return v
    if isinstance(v, (int, float)):
        d = excel_serial_to_dt(v)
        return d if d else None
    s = str(v).strip()
    if s == "": return None
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d", "%m/%d %H:%M", "%m/%d", "%d/%m %H:%M", "%d/%m"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return s

# 已知航线码集合(用于段标题检测: C1 或 C9 匹配已知航线码 = 段标题行, 不是数据行)
# 仅含"服务/lane 码", 不含港口码 —— 港口码见下方 PORT_CODES 黑名单。
KNOWN_LANES = set(norm(r) for r in ROUTE_ORDER) | {norm(k) for k in ROUTE_ALIAS} | {norm(v) for v in ROUTE_ALIAS.values()}

# 港口码黑名单(UN/LOCODE 等): 这些码即使出现在标题样位置也【绝不】当作 lane。
# 这是"确认是 lane"的硬闸, 防止 CNNGB/CNTAO/TRALI/DJJIB 等港口被误判为航线。
PORT_CODES = {
    # 中国
    "CNNGB","CNTAO","CNNAS","CNSHA","CNNSA","CNXMN","CNXGG","CNDLC","CNTXG","CNYTN","CNHKG",
    # 新加坡/马来西亚
    "SGSIN","SGTPP","SGMAL","MYTPP","MYPGU","MYPKG",
    # 红海/中东 港口
    "EGSOK","EGSUZ","EGSPS","EGALY","SAJED","SADMM","SAUQD","AEJEA","AEDXB","AEAUH","OMSAL","OMPSS","OMKHS",
    # 斯里兰卡/印度
    "LKCMB","LKTUT","INNSA","INMAA","INRTC","INIXY","INBLR","INBOM",
    # 欧洲/地中海
    "NLRTM","BEANR","DEHAM","GBFXT","GBSOU","ESVLC","ESALG","ITGOA","ITTPS","GRPIR","TRMER","TRIST",
    # 美湾/北美
    "USNYC","USLGB","USSAV","USLAX","USLGB","USHOU","USCHS",
}

def is_port_code(code):
    """端口码硬判: 这些码即使出现在标题样位置也绝不当作 lane。"""
    return norm(code) in PORT_CODES

def is_lane_code(code):
    """确认某文本是否为【已知航线服务码】。
    仅接受 KNOWN_LANES(ROUTE_ORDER + 别名 + 并入的真实服务名)中的码;
    港口码(PORT_CODES)与任何未知/杂文本一律否决 —— 绝不把端口当 lane。
    未来出现的新 lane 由"段标题结构位置"(PIC 块头 / 含船名船代码的段头)识别并原样保留,
    而非靠形态猜测, 因此此处只需严格白名单即可保证"确认是 lane"。
    """
    n = norm(code)
    if not n:
        return False
    if n in PORT_CODES:
        return False
    return n in KNOWN_LANES

def detect_route(vessel_code, c1_val, c9_val):
    """从段标题行检测真正的航线码(仅认已知 lane, 港口码与未知码一律否决)。
    源布局不统一: 有的航线在C1、代码在C9, 有的反过来(EVERLASTING HARVEST: C1=EVHA=代码, C9=REX=航线)。
    若候选等于本船代码(船名缩写)则跳过(那是代码不是航线)。
    返回航线码字符串, 或None(未检测到)。新 lane 由 read_source 的结构位置(段头含船名)兜底识别。
    """
    vcode = norm(vessel_code) if vessel_code else None
    for col, val in ((1, c1_val), (9, c9_val)):
        if val is None or isinstance(val, datetime):
            continue
        n = norm(val)
        if not is_lane_code(val):   # 仅已知 lane(白名单, 已排除港口码)
            continue
        if vcode and n == vcode:
            continue
        return str(val).strip()
    return None

TARGET_HEADERS = {
    1:  ["PORT"],
    2:  ["MANIN", "MAN IN"],
    3:  ["WAIT"],
    4:  ["PROFORMA"],
    5:  ["LTS ETB", "LTM ETB"],          # 源LTS/LTM ETB -> 大Excel C5(ltm eta位置)
    6:  ["LTS ETD", "LTM ETD"],
    7:  ["VOY.NO", "VOYNO.", "VOYNO", "VOY. NO"],
    8:  ["DATE"],
    9:  ["ETA"],
    10: ["ETB"],
    11: ["ETD"],
    12: ["RUN"],
    13: ["PORTSTAY(HR)", "PORTSTAY"],
    14: ["FSPDISTANCE", "FSP DISTANCE"],
    15: ["SPEED"],
    16: ["ETADELAY/AHEAD", "ETA DELAY/AHEAD"],
}

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

_RET_RE = re.compile(r"[\s\-_]*已下线[\s\-_]*", re.I)

def strip_retired(name):
    """去掉船名里的下线标记: 'CUL JAKARTA-已下线' / '已下线-CUL JAKARTA' -> 'CUL JAKARTA'。
    用于查 vessel.csv / PIC汇总(权威表里的登记名不带后缀)。"""
    s = _RET_RE.sub("", str(name or "")).strip().strip("-_ ").strip()
    return s

def collect_vessel_folders(src):
    """扫描船队目录(含 "已下线船舶/" 子目录), 返回 [(folder_path, folder_name, base_name, is_retired), ...]
    - folder_name: 磁盘上的原始目录名
    - base_name  : 去掉 '已下线' 标记的规范船名(查权威表/显示名用)
    - is_retired : 是否来自 "已下线船舶/" 目录
    同一艘船存在两份目录(X-已下线 与 已下线-X)时按规范名去重, 保留 xlsx 最新的一份。
    """
    cand = []
    if os.path.isdir(src):
        for d in sorted(os.listdir(src)):
            full = os.path.join(src, d)
            if not os.path.isdir(full):
                continue
            if d == RETIRED_DIR_NAME:
                for sub in sorted(os.listdir(full)):
                    subp = os.path.join(full, sub)
                    if os.path.isdir(subp):
                        cand.append((subp, sub, True))
            else:
                cand.append((full, d, False))
    best, order = {}, []
    for path, name, ret in cand:
        xlsx = latest_xlsx(path)
        if not xlsx:
            print(f"  [WARN] 无xlsx跳过: {path}"); continue
        base = strip_retired(name)
        key = norm(base)
        mt = os.path.getmtime(xlsx)
        prev = best.get(key)
        if prev is None:
            order.append(key)
        elif mt <= prev["mt"]:
            print(f"  [去重] {base}: 保留较新副本 -> {os.path.basename(prev['xlsx'])} (跳过 {os.path.basename(xlsx)})")
            continue
        else:
            print(f"  [去重] {base}: 采用较新副本 -> {os.path.basename(xlsx)} (替代 {os.path.basename(prev['xlsx'])})")
        best[key] = {"path": path, "folder": name, "base": base, "ret": ret, "xlsx": xlsx, "mt": mt}
    return [(best[k]["path"], best[k]["folder"], best[k]["base"], best[k]["ret"]) for k in order]

def _has_port_header(ws):
    """判断某 sheet 是否含 PORT 表头(即真正的船期数据表)。"""
    for r in range(1, min(ws.max_row, 200) + 1):
        v = ws.cell(r, 1).value
        if v and str(v).strip().upper() == "PORT":
            return True
    return False

def _pick_data_sheet(wb):
    """多 sheet 工作簿中挑选正确的船期数据表。

    规则(用户 2026-09-22 确认): 始终读取【第一个 sheet】= wb.sheetnames[0]。
    各船源文件由用户自行保证首张表即当前维护的船期数据表(如 M.MARINER 首表=MMRN SGX)。
    不再优先激活表, 也不做 simulation/PORT 表头兜底——若首表非数据表, 由用户整理目录。
    单 sheet 工作簿 sheetnames[0]==唯一表, 行为不变。
    """
    return wb[wb.sheetnames[0]]

def read_source(path, vessel_code=None, folder_name=None):
    """读源数据表(自动挑选正确的 sheet)。按表头名(非固定列位)映射到大Excel列, 兼容源列序差异(如ASR多TERMINAL列)。
    返回 dict: route, code, rows[ {display:{col:val}, etb:datetime|None, voy, port, remark} ]。
    - 日期列(C5/C6/C8/C9/C10/C11)统一为 datetime(真实日期) 或 文本标记(OMIT/Mon..)。
    - Voy.No 若空, 向上就近取最近的有航次号的行。
    - 段标题检测(两层):
      1. KNOWN_LANES 匹配 C1/C9 (已知航线码 → 段标题, 港口名不会误判)
      2. 回退: C4 或 C9 匹配船代码/文件夹名 (未知航线码如 NSCT1, 但段标题行总有船名/代码)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = _pick_data_sheet(wb)
    # 找表头行(PORT)
    hr = None
    for r in range(1, min(ws.max_row, 200) + 1):
        if norm(ws.cell(r, 1).value) == "PORT":
            hr = r; break
    if hr is None:
        print(f"  [WARN] 首张表 '{ws.title}' 无 PORT 表头, 该船源文件未产出数据(请确认首表是否为船期数据)")
        return {"route": "", "code": None, "rows": []}
    # 初始航线: 从 PORT 表头往上找【段标题行】(C1 非空 且 非 PORT 且 C9 非日期).
    # 段标题行的航线码严格按源文件原样采用 —— 未知码(WAT/KCI/GTS…)也保留,
    # 不再因不在白名单(KNOWN_LANES)而被丢弃/误判; 端口数据行 C9 是 ETA 日期, 自动跳过.
    # 兼容两种源布局: 有的 R1 直接是航线码; 有的是合并大标题行(真正航线码在 R2C1).
    route = ""
    code = None
    # 船名/代码归一化集合(用于段标题回退检测) —— 段头那行往往是船代码本身, 不能当航线
    vessel_markers = set()
    if vessel_code: vessel_markers.add(norm(vessel_code))
    if folder_name: vessel_markers.add(norm(folder_name))
    for r in range(hr - 1, 0, -1):
        c1v = ws.cell(r, 1).value
        c9v = ws.cell(r, 9).value
        if code is None and c9v is not None and not isinstance(c9v, datetime):
            code = c9v
        if c1v is None:
            continue
        s1 = norm(c1v)
        if s1 == "PORT":
            continue
        if isinstance(c9v, datetime):
            continue   # 港口数据行(C9=ETA日期), 非段标题
        det = detect_route(vessel_code, c1v, c9v)
        if det:
            route = det
            break
        # 该行 C1 就是本船的船名/船代码(段头的船名单元格), 不是航线 —— 继续向上找
        if vessel_markers and s1 in vessel_markers:
            continue
        # 新 lane: 段标题行(C1 非港口码)原样保留为航线; 港口码不当航线, 继续向上找
        if not is_port_code(c1v):
            route = str(c1v).strip()
            break
    if not route:
        # 兜底: R1C1 原文(仅当上方真的无任何段标题行时); 港口码会被 is_port_code 否决。
        # R1C1 若就是本船名/船代码(如 ZGCD), 也不当航线 —— 交给下面的 lanes_in_file 继承真实航线。
        c1_r1 = ws.cell(1, 1).value
        if vessel_markers and norm(c1_r1) in vessel_markers:
            route = ""
        else:
            route = str(c1_r1).strip() if (c1_r1 and not is_port_code(c1_r1)) else ""
    if code is None:
        code = ws.cell(1, 9).value
    # 源表头归一名 -> 源列号
    src_hdr = {}
    for c in range(1, ws.max_column + 1):
        h = norm_h(ws.cell(hr, c).value)
        if h and h not in src_hdr:
            src_hdr[h] = c
    # 目标列 -> 源列号
    col_map = {}
    for tcol, cands in TARGET_HEADERS.items():
        for cand in cands:
            sc = src_hdr.get(norm_h(cand))
            if sc:
                col_map[tcol] = sc; break
    # 船名/代码归一化集合已在上方定义(初始航线检测也要用)
    raw = []
    current_route = route   # 初始航线, 遇到中间段标题行会切换
    for r in range(hr + 1, ws.max_row + 1):
        c1 = ws.cell(r, 1).value
        if c1 is None: continue
        s1 = norm(c1)
        if s1 == "PORT": continue          # 多段航次表, 跳过重复列头继续读
        # ── 段标题检测(两层) ──
        # 1. KNOWN_LANES 匹配 C1/C9 (已知航线码)
        c9_val = ws.cell(r, 9).value
        seg_route = detect_route(vessel_code, c1, c9_val)
        if seg_route:
            current_route = seg_route
            continue    # 段标题行本身不作为数据行
        # 2. 回退: C4 或 C9 匹配船代码/文件夹名 (未知航线码, 但段标题行总有船名)
        c4_val = ws.cell(r, 4).value
        c4_str = norm(c4_val) if isinstance(c4_val, str) else ""
        c9_str = norm(c9_val) if isinstance(c9_val, str) and not isinstance(c9_val, datetime) else ""
        # 该行 C1 就是本船的船名/船代码(段头的船名单元格) -> 不是航线, 也不改变当前航线。
        # 例: ZHONG GU CHENG DU 段头有一行 C1='ZGCD'(就是船代码),
        # 不排除的话会被当成独立航线 'ZGCD', 多出一个假的航线分组。
        if vessel_markers and norm(c1) in vessel_markers:
            continue
        if vessel_markers and (c4_str in vessel_markers or c9_str in vessel_markers) and not is_port_code(c1):
            # 段标题行(C4/C9 含船名/代码), 更新航线为 C1; 港口码不当 lane
            current_route = str(c1).strip() if c1 else current_route
            continue
        display = {}
        for tcol in range(1, 17):
            sc = col_map.get(tcol)
            display[tcol] = ws.cell(r, sc).value if sc else None
        etb = display.get(10)
        etb = etb if isinstance(etb, datetime) else None
        raw.append({
            "display": display,
            "etb": etb,
            "voy_raw": norm_voy(display.get(7)),
            "port": s1,
            "remark": ws.cell(r, 18).value or ws.cell(r, 19).value,
            "row_route": current_route,      # ← 该行所属的实际航线(支持一船多段)
        })
    # 未标航线的首段(段头只有船代码、无 lane): 继承本文件里第一个已知航线,
    # 避免产生空 lane / 假 lane 分组(如 ZHONG GU CHENG DU 首段 -> AEM)。
    lanes_in_file = [rr["row_route"] for rr in raw if is_lane_code(rr["row_route"])]
    if lanes_in_file:
        for rr in raw:
            if not rr["row_route"]:
                rr["row_route"] = lanes_in_file[0]
    # Voy.No 向上就近: 每行若空, 向上(表中更靠上)找最近的有航次号的行
    for i, rr in enumerate(raw):
        if rr["voy_raw"]:
            continue
        for j in range(i - 1, -1, -1):
            if raw[j]["voy_raw"]:
                rr["voy_raw"] = raw[j]["voy_raw"]; break
    # 统一日期格式 + 写入display
    for rr in raw:
        for tcol in (5, 6, 8, 9, 10, 11):
            rr["display"][tcol] = norm_date_value(rr["display"].get(tcol))
        rr["display"][7] = rr["voy_raw"]
        rr["voy"] = rr["voy_raw"]
    return {"route": route, "code": code, "rows": raw}

def load_vessel_csv(path):
    """GitHub 仓库内 vessel.csv: 船名(或文件夹名)-> {code, display}。用户手动维护。
    显示名取 ship 列原始值; 代码取 code 列。两机 git pull 同步。"""
    d = {}
    if not os.path.exists(path):
        return d
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            ship = (row.get("ship") or "").strip()
            code = (row.get("code") or "").strip()
            if ship:
                d[norm(ship)] = {"code": code, "display": ship}
    return d

def load_pic(path):
    """PIC 权威表: 优先 xlsx(人工编辑主文件), 回退 csv(镜像)。按文件夹名取PIC。
    取不到返回空dict(回退旧大Excel)。"""
    if str(path).lower().endswith(".xlsx"):
        return _load_pic_xlsx(path)
    return _load_pic_csv(path)

def _load_pic_xlsx(path):
    d = {}
    if not os.path.exists(path):
        return d
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return d
    ws = wb[wb.sheetnames[0]]
    # 表头: 航线|船名(显示)|文件夹名|船代码|PIC|状态 -> C3=文件夹名, C5=PIC
    for r in range(2, ws.max_row + 1):
        fol = ws.cell(r, 3).value
        pic = ws.cell(r, 5).value
        if fol is not None and str(fol).strip():
            d[norm(str(fol).strip())] = str(pic or "").strip()
    return d

def _load_pic_csv(path):
    d = {}
    if not os.path.exists(path):
        return d
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            fol = (row.get("文件夹名") or row.get("folder") or "").strip()
            pic = (row.get("PIC") or row.get("pic") or "").strip()
            if fol:
                d[norm(fol)] = pic
    return d

def load_pic_code_map(path):
    """PIC汇总.xlsx: A=航线 B=船名(显示) C=文件夹名 D=船代码 E=PIC F=状态。
    返回 {规范化文件夹名: 船代码}。文件不可达/缺列时返回空 dict(回退 vessel.csv/源R1C9)。
    这是 detect_route 识别 lane vs 船代码 的权威依据(用户维护的 PIC汇总 优先于 vessel.csv)。"""
    d = {}
    if not os.path.exists(path):
        return d
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return d
    ws = wb[wb.sheetnames[0]]
    for r in range(2, ws.max_row + 1):
        fol = ws.cell(r, 3).value   # C=文件夹名
        code = ws.cell(r, 4).value  # D=船代码
        if fol and str(fol).strip() and code and str(code).strip():
            d[norm(str(fol).strip())] = str(code).strip()
    return d

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

def build_route_order(groups):
    """分组顺序: 跟随 ROUTE_ORDER 常量(固化, 可编辑); 常量里没有的新航线追加到末尾(按字母序)。"""
    seen = [r for r in ROUTE_ORDER if norm(r) in groups]
    extra = sorted([r for r in groups if norm(r) not in [norm(x) for x in ROUTE_ORDER]])
    return seen + extra

def main():
    ap = argparse.ArgumentParser(description="按当前船队重建 CUL DAILY MOVEMENT 大Excel")
    ap.add_argument("--src", default=DEFAULT_SRC)
    ap.add_argument("--vessel", default=VESSEL_CSV, help="GitHub vessel.csv (船名->代码/显示名)")
    ap.add_argument("--pic", default=DEFAULT_PIC, help="PIC 汇总.xlsx (人工编辑主文件, 文件夹名->PIC)")
    ap.add_argument("--output", default=DEFAULT_OUT)
    ap.add_argument("--today", default=None, help="基准日 YYYY-MM-DD, 默认今天(用于选最近船期行)")
    ap.add_argument("--window", type=int, default=WINDOW_DAYS, help="兼容旧参数, 已被 --year 取代")
    ap.add_argument("--year", type=int, default=2026, help="船期年份窗口(默认2026): 加载该年第一个 sheet 的全部船期")
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()

    today = datetime.strptime(args.today, "%Y-%m-%d").date() if args.today else date.today()
    # 2026 全年窗口: 加载每个源表第一个 sheet 的全部 {year} 船期, 不再按 today±window 截断
    lo = date(args.year, 1, 1)
    hi = date(args.year, 12, 31)

    print(f"[基准日] {today}  年份窗口 {args.year} -> [{lo} ~ {hi}]")
    print("=== 1/4 读权威表(vessel.csv / P盘PIC) ===")
    vessel = load_vessel_csv(args.vessel)
    pic_tbl = load_pic(args.pic)
    pic_code_tbl = load_pic_code_map(args.pic)   # 船代码(权威): PIC汇总 D列 -> 文件夹名
    print(f"  vessel.csv: {len(vessel)} 条 | P盘PIC表: {len(pic_tbl)} 条 | PIC船代码: {len(pic_code_tbl)} 条")

    print("=== 2/4 扫描当前船队(2026/ 含 已下线船舶/) ===")
    fleet = collect_vessel_folders(args.src)   # [(path, folder, base, retired)]
    ships = []   # {folder, route, code, display, rows, retired}
    retired_cnt = sum(1 for _f in fleet if _f[3])
    print(f"  扫描到船目录 {len(fleet)} 个 (在营 {len(fleet)-retired_cnt} / 已下线 {retired_cnt})")
    for folder_path, fol, base, retired in fleet:
        p = latest_xlsx(folder_path)
        if not p:
            print(f"  [WARN] 无xlsx跳过: {fol}"); continue
        key = norm(base)      # 权威表按【去除下线标记的规范船名】查找
        # 先查 vessel.csv 取船代码, 传给 read_source 做段标题检测(区分航线码 vs 船代码)
        vent = vessel.get(key)
        # 船代码优先级: PIC汇总(权威, 用户维护) -> vessel.csv -> 下线船兜底表 -> (read_source 内回退 源R1C9)
        vcode = (pic_code_tbl.get(key)
                 or (vent.get("code") if vent else None)
                 or (RETIRED_VESSEL_CODES.get(key) if retired else None))
        d = read_source(p, vessel_code=vcode, folder_name=base)
        route = canon_route(base, d["route"])   # 应用覆盖+合并(覆盖表按规范船名)
        # 显示名/船代码: PIC汇总优先, 回退 vessel.csv, 再回退 源R1C9/文件夹名
        code = vcode or d["code"]
        disp = (vent.get("display") if vent else None) or base
        if retired:
            disp = strip_retired(disp) + RETIRED_SUFFIX   # 网页端靠此后缀识别已下线船
        if not vcode and not vent:
            print(f"  [WARN] 船未在 PIC汇总/vessel.csv 登记: {fol} (code 回退 源R1C9/文件夹名, 建议补登)")
        # ── 按逐行航线拆分子块(支持一船多段, 如 ZYHS SGX→NP2)
        #     但拆之前先用 ±30天窗口过滤: 只有多段同时有窗口内数据才拆;
        #     历史航次(如 CUL HUANGPU 的 CHT/SL1/CST)无窗口内数据则自动忽略。 ──
        sub_routes = {}   # route -> [rows with row_route]
        for rr in d["rows"]:
            sr = canon_route(base, rr.get("row_route", route))
            sub_routes.setdefault(sr, []).append(rr)
        # 对每段检查是否有 窗口内数据
        active_segments = {}
        for sub_r, sub_rows in sub_routes.items():
            has_data = any(rr.get("etb") is not None and lo <= rr["etb"].date() <= hi for rr in sub_rows)
            if has_data:
                active_segments[sub_r] = sub_rows
        if len(active_segments) > 1:
            print(f"  [{base}] 拆为 {len(active_segments)} 个航线段(仅窗口内有数据): {', '.join(active_segments.keys())}")
        elif len(active_segments) == 0:
            # 整个文件都没有窗口内数据 — 保留第一段(船仍显示, 只是无港口行)
            first_r = next(iter(sub_routes))
            active_segments = {first_r: sub_routes[first_r]}
        for sub_r, sub_rows in active_segments.items():
            ships.append({"folder": base, "route": sub_r, "code": code,
                          "display": disp, "rows": sub_rows, "retired": retired})
    _ret_blocks = sum(1 for s in ships if s.get("retired"))
    print(f"  写出船块数: {len(ships)} (在营 {len(ships)-_ret_blocks} / 已下线 {_ret_blocks})")

    print("=== 3/4 分组排序 + 写表 ===")
    # 分组
    groups = {}
    for s in ships:
        groups.setdefault(norm(s["route"]), []).append(s)
    # 顺序: 跟随 ROUTE_ORDER 常量(固化), 新航线追加末尾
    ordered_routes = build_route_order(groups)

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
            disp = s["display"]
            # PIC: P盘PIC汇总.xlsx(权威); 取不到则留空待补
            pic_raw = pic_tbl.get(norm(fol))
            pic_clean = str(pic_raw).replace("PIC:", "").replace("PIC :", "").strip() if pic_raw else ""
            # 块头
            setc(row, 1, route, font=bold)
            setc(row, 4, disp, font=bold)
            setc(row, 8, "DATE")
            setc(row, 9, s["code"])
            setc(row, 16, "PIC: " + pic_clean)   # 始终带PIC:前缀, 保证网页解析识别块(含无PIC新船)
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
                for c in range(1, 17):
                    setc(row, c, rr["display"].get(c), bd=True)
                row += 1; shown += 1
            # Remark: 源文件所有行 remark, 不限于最接近今天
            for i, rr in enumerate(s["rows"]):
                rem = rr["remark"]
                if not rem: continue
                voy = rr["voy"] or nearest_voy_upward(s["rows"], i)
                txt = f"Remark:{voy} {rr['port']} {rem}"
                setc(row, 1, txt)
                row += 1
            row += 1   # 块间空行
            blocks_written += 1
        row += 1   # 段间空行

    print("=== 4/4 保存 ===")
    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    if os.path.exists(args.output) and not args.no_backup:
        bak = args.output + ".bak_" + datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(args.output, bak)
        print(f"  已备份 -> {bak}")
    wb.save(args.output)
    print(f"  航线组数: {len(ordered_routes)}  写出船块数: {blocks_written}")
    print(f"  输出: {args.output}")

if __name__ == "__main__":
    raise SystemExit(main())
