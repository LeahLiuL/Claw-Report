"""航次油耗计算 (voyage fuel consumption).

思路 (与用户定义一致):
  一个航次 = 本航次「第一个挂靠港的 Berth(ETB)时间」→ 下一航次「第一个挂靠港的 Berth 时间」。

  **航次号归一化(重要)**: 2630W 与 2630E 只是方向不同, 属于**同一个航次**。
  因此航次号去掉尾部的方向字母(W/E/N/S...), 按 (船名, 数字航次) 归组, 该组下
  所有方向(W+E)的挂靠港合并成一条完整 rotation, 按 ETB 排序。
  例: CUHP 2630 -> start = 2630W 首港 TWKEL Berth(08-22 20:36)
                   end   = 2631W 首港 TWKEL Berth(09-02 00:30)
       rotation = TWKEL → TXTG → TWKHH (2630W) → CNSWA → CNNAS → CNSHK (2630E)

  非标准航次值(OMIT / BUNKER / BUNKERONLY / ADHOC / PRIVATECALL / 空 / '00:00:00')
  不是真实航次, 直接跳过。

数据来源:
  - 航次边界: 本仓库 cul_daily_movement.html 的 TODAY_DATA.fullSchedule
    (每条记录是一个挂靠港, 含 vessel/voy/port/etb/etbRaw; 同 vessel+voy 取最早 etb 作为首港)。
  - 油耗: rob_data/rob_history.csv 里该船的 ROB(LSFO/HSFO/MGO/ULSFO)时间序列。
    **消耗 = 期初 ROB - 期末 ROB + 期间加油量**; 加油量来自《燃油添加日志》(只含体积, 不含价格)。
    例如期初 ROB 1000, 加了 500, 期末 ROB 1200, 则真实消耗 = 1000 - 1200 + 500 = 300。
    航次页展示的 B.LSFO/B.HSFO/B.MGO/B.ULSFO 为期间各油种加油量(体积, MT)。

不依赖 Outlook / Excel, 可独立测试。
"""
import os
import re
import json
import csv
import datetime
from collections import defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
DM_HTML = os.path.join(BASE, "cul_daily_movement.html")
ROB_DIR = os.path.join(BASE, "rob_data")
HISTORY_CSV = os.path.join(ROB_DIR, "rob_history.csv")
# 报告级归档(harvest_reports.py 产出): 一封报告一行。日快照每船每天只有一条,
# 而船长每天实际会发 NOON+BERTH+SAILING 多封, BERTH 恰恰是航次边界的取数来源。
REPORTS_JSONL = os.path.join(ROB_DIR, "rob_reports.jsonl")
OFFLINE_MARK = "已下线"


def is_offline(name):
    return OFFLINE_MARK in (name or "")


# 标准航次号: 3 位以上数字 + 1 个方向字母(如 2630W / 2630E)
VOY_RE = re.compile(r"^(\d{3,})([A-Za-z])$")


def split_voy(voy):
    """'2630W' -> ('2630', 'W'); 非标准航次值 -> None (应跳过)。"""
    m = VOY_RE.match((voy or "").strip())
    if not m:
        return None
    return m.group(1), m.group(2).upper()


def _num(x):
    try:
        if x is None or x == "":
            return None
        return float(x)
    except Exception:
        return None


def _dt_from_raw(raw, hm):
    """raw='YYYY-MM-DD' + hm='MM/DD HH:MM' -> datetime。"""
    if not raw:
        return None
    try:
        y, mo, da = raw.split("-")
        hh, mm = 0, 0
        if hm:
            m = re.match(r"(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2})", hm or "")
            if m:
                hh, mm = int(m.group(3)), int(m.group(4))
        return datetime.datetime(int(y), int(mo), int(da), hh, mm)
    except Exception:
        return None


def _parse_etb(rec):
    return _dt_from_raw(rec.get("etbRaw"), rec.get("etb"))


def _parse_etd(rec):
    return _dt_from_raw(rec.get("etdRaw"), rec.get("etd"))


def _parse_eta(rec):
    return _dt_from_raw(rec.get("etaRaw"), rec.get("eta"))


def _parse_rt(rt):
    """'YYYY-MM-DD HH:MM' -> datetime。"""
    try:
        return datetime.datetime.strptime(rt, "%Y-%m-%d %H:%M")
    except Exception:
        try:
            return datetime.datetime.strptime(rt[:10], "%Y-%m-%d")
        except Exception:
            return None


def load_full_schedule(dm_html=DM_HTML):
    if not os.path.exists(dm_html):
        return None
    try:
        html = open(dm_html, encoding="utf-8", errors="replace").read()
        m = re.search(r"const\s+TODAY_DATA\s*=\s*", html)
        if not m:
            return None
        data, _ = json.JSONDecoder().raw_decode(html[m.end():])
        return data.get("fullSchedule", [])
    except Exception as e:
        print("[WARN] voyage: parse fullSchedule failed:", e)
        return None


def load_history(history_csv=HISTORY_CSV, berth_only=False):
    """读取 ROB 历史, 返回 vessel -> [(dt, (ls,hs,mg,us), rtype), ...]。

    berth_only=True: 仅返回 report_type=='BERTH' 的行(兼容旧调用)。
    默认 False: 返回全部类型, 由调用方按类型拆分 —— 航次油耗用 BERTH 优先,
    找不到 BERTH 时按需求用当天 NOON 兜底(见 _window_consumption)。
    依赖 rob_history.csv 的 report_type 列已正确标注; 该列由 Outlook 历史回填
    (backfill_history.py)基于附件正文内容补全。

    另会合并 rob_data/rob_reports.jsonl(报告级归档, 若存在)。日快照是「船×日=1 行」,
    同一船的 BERTH/SAILING 报告大量被当天最新一封覆盖; 归档补齐这些时刻后, 航次边界
    才能取到真正的靠泊读数。**只补日快照里没有的时刻**(同船同分钟视为同一份报告),
    不覆盖既有口径 —— 保证既有航次数字不因接入归档而突变。"""
    hist = defaultdict(list)  # vessel -> [(dt, oils, rtype), ...]
    seen = set()              # (vessel, 'YYYY-MM-DD HH:MM')

    def _push(v, dt, oils, rt):
        k = (v, dt.strftime("%Y-%m-%d %H:%M"))
        if k in seen:
            return False
        seen.add(k)
        hist[v].append((dt, oils, rt))
        return True

    if os.path.exists(history_csv):
        try:
            with open(history_csv, encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    v = (row.get("vessel") or "").strip()
                    if is_offline(v):
                        continue
                    rt = (row.get("report_type") or "").strip().upper()
                    if berth_only and rt != "BERTH":
                        continue
                    dt = _parse_rt((row.get("report_time") or "")[:16])
                    if dt is None:
                        continue
                    oils = (_num(row.get("lsfo")), _num(row.get("hsfo")),
                            _num(row.get("mgo")), _num(row.get("ulsfo")))
                    if all(o is None for o in oils):
                        continue
                    _push(v, dt, oils, rt)
        except Exception as e:
            print("[WARN] voyage: read history failed:", e)

    n_arch = 0
    if os.path.exists(REPORTS_JSONL):
        try:
            with open(REPORTS_JSONL, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        r = json.loads(line)
                    except Exception:
                        continue
                    v = (r.get("vessel") or "").strip()
                    if not v or is_offline(v):
                        continue
                    rt = (r.get("report_type") or "").strip().upper()
                    if berth_only and rt != "BERTH":
                        continue
                    dt = _parse_rt((r.get("received") or "")[:16])
                    if dt is None:
                        continue
                    oils = (_num(r.get("lsfo")), _num(r.get("hsfo")),
                            _num(r.get("mgo")), _num(r.get("ulsfo")))
                    if all(o is None for o in oils):
                        continue
                    if _push(v, dt, oils, rt):
                        n_arch += 1
        except Exception as e:
            print("[WARN] voyage: read rob_reports failed:", e)
    if n_arch:
        print("[voyage] +%d 条报告级读数 (rob_reports.jsonl)" % n_arch)
    return hist


def _window_consumption(berth_series, noon_series, start_dt, end_dt, bunker_by_date,
                        guard_negative=True):
    """计算航次 [本航次首港berth=start_dt, 下一航次首港berth=end_dt] 内各油种真实消耗。

    公式: 消耗 = 期初 ROB - 期末 ROB + 【两条读数之间】的加油量

    端点(2026-11-12 修正): 期初 = 开航时刻 start_dt 对应的读数, 期末 = 到港时刻 end_dt
    对应的读数, 各自取边界【前后最近】的一条(BERTH/NOON 同权重, 同时刻 BERTH 优先)。
    同一个边界永远映射到同一条读数, 因此上一航次的期末 == 下一航次的期初, 相邻航次
    窗口首尾相接, 既不重叠也不留缝。旧口径只允许往后取, 于是:
      * CUHP 2640 于 10-02 03:06 开航, 却只能取到 10-09 的 310.83(已含 10-03 加的
        297.5 MT), 再按 [start,end] 计入该笔加油 -> 393.08, 真实约 178.5;
      * 边界前刚投过的读数比边界后那份更近时会被跳过, 白丢一段消耗。
    进行中航次(end_dt=None): 期末取「截至 now 最近」的读数, 显示至今的部分消耗。
    加油量(2026-11-12 修正): 窗口取【期初读数时刻, 期末读数时刻】, 逐油种各算各的,
      不再用 [start_dt, end_dt]。这样 ROB(a) - ROB(b) + bunker(a..b) 恒等于 a→b 的真实
      消耗, 读数落在边界之外也不会再把同一笔加油重复计入。端点缺失才退回航次边界。
      按日期比较(加油日志只有日期, 没有时分)。
    BERTH/NOON 不做全局 BERTH 优先 —— 否则航次中间挂靠港的 BERTH 报告会被误当边界读数
    (CUHP 2629/2630 实测错位 40MT)。
    返回 (cons[ls,hs,mg,us], bunker_total, bunker_oil[ls,hs,mg,us], basis),
    basis∈{'BERTH','NOON','MIXED',''} 标明本段消耗基于哪种报告。"""
    cons = [None, None, None, None]
    first = [None, None, None, None]
    last = [None, None, None, None]
    f_basis = [None, None, None, None]
    l_basis = [None, None, None, None]
    f_dt = [None, None, None, None]   # 期初读数时刻(逐油种, 用于确定加油窗口)
    l_dt = [None, None, None, None]   # 期末读数时刻
    oils = ["ls", "hs", "mg", "us"]
    bunker_oil = [0.0, 0.0, 0.0, 0.0]
    now = datetime.datetime.now()
    # 零值脏数据防护: 若该船该油种在整条读数序列里存在 >0 值, 则 0.0 视为脏数据
    # (如 RACINE 多条 BERTH 报 lsfo=0 而前后都是 1800+ MT), 取边界读数时跳过;
    # 若该油种全程为 0(船本就不烧该油), 0 是真实值, 保留。
    has_pos = [False, False, False, False]
    for _src in (berth_series, noon_series):
        for _p in _src:
            for _k in range(4):
                if (_p[1][_k] is not None and 0 < _p[1][_k] <= 5000.0):
                    has_pos[_k] = True
    # 边界读数搜索窗口: 边界前后各容忍一段, 超出则认为该边界确实没有可用读数。
    # 两侧都不能放太松, 否则会把隔壁航次刚用过的读数借来, 两段就重叠了。
    PREV_MAX = datetime.timedelta(days=5)
    NEXT_MAX = datetime.timedelta(days=10)
    # 单油种 ROB 物理上限: 中联船队最大船单一油种存油也在 3000 MT 量级,
    # 超过此值只可能是解析错位(读到总舱容/累计消耗/单位错乱的列)。这类读数一旦被
    # 当成端点, 会把整船数据污染成几百上千吨的假消耗, 因此直接排除, 不参与端点选取。
    ROB_ABS_MAX = 5000.0
    _prev_s = PREV_MAX.total_seconds()
    _next_s = NEXT_MAX.total_seconds()

    def _pick(t, k):
        """边界时刻 t -> 唯一一条 ROB 读数。

        取 [t-PREV_MAX, t+NEXT_MAX] 内【离 t 最近】的那条; 同时刻 BERTH 优先,
        等距时偏向 t 之后。同一个 t 永远选出同一条读数, 于是「上一航次的期末」与
        「下一航次的期初」必然是同一读数 —— 相邻航次窗口首尾相接、绝不重叠,
        同一笔加油不会在两个航次里各算一次, 全船合计守恒。

        这条「双侧最近」规则同时修掉旧口径的两个病(旧口径只允许取边界之后):
          * CUHP 2640: 10-02 03:06 开航, 之后 7 天无报告, 只能取 10-09 的 310.83,
            而该值已包含 10-03 加的 297.5, 再按区间计入加油 -> 重复计, 393.08(真值约 178.5);
          * 反过来, 边界前刚投过一份读数、比边界后那份更近时, 旧口径会跳过它,
            白丢一段消耗。
        """
        best, best_key = None, None
        for src, tag in ((berth_series, "B"), (noon_series, "N")):
            for p in src:
                dt, val = p[0], p[1][k]
                if val is None or val > ROB_ABS_MAX or (has_pos[k] and val <= 0):
                    continue
                gap = (dt - t).total_seconds()
                if gap < -_prev_s or gap > _next_s:
                    continue
                # 排序键: 绝对时距 -> 同为 BERTH 优先 -> 偏向边界之后
                key = (abs(gap), 0 if tag == "B" else 1, 0 if gap >= 0 else 1)
                if best_key is None or key < best_key:
                    best_key, best = key, (dt, val, tag)
        return best

    def _pick_after(t, k, limit):
        """兜底: t 之后最近的一条读数(不限距离, 但不超过 limit)。等价于旧口径。"""
        best, best_key = None, None
        for src, tag in ((berth_series, "B"), (noon_series, "N")):
            for p in src:
                dt, val = p[0], p[1][k]
                if val is None or val > ROB_ABS_MAX or (has_pos[k] and val <= 0):
                    continue
                if dt < t or (limit is not None and dt > limit):
                    continue
                key = (dt, 0 if tag == "B" else 1)
                if best_key is None or key < best_key:
                    best_key, best = key, (dt, val, tag)
        return best

    for k in range(4):
        # 该油种在整条序列里全程无正读数 -> 本船不烧这类油, 消耗应显示为缺失("-"),
        # 而不是误导性的 "0.0 MT"(会让人以为航次真的零消耗)。
        if not has_pos[k]:
            cons[k] = None
            continue
        # 期初 / 期末 = 边界时刻各自唯一映射的那条读数(见 _pick)
        pick = _pick(start_dt, k)
        if pick is None:
            # 兜底(旧口径): 边界之后最近的一条, 不限距离。只在 ±5/10 天内一条读数都没有时
            # 触发 —— 这类超稀疏航次本来也只能算个下限, 但至少不让整行变空白。
            pick = _pick_after(start_dt, k, end_dt)
        if pick:
            f_dt[k], first[k], f_basis[k] = pick[0], pick[1], pick[2]
        if end_dt is not None:
            pick = _pick(end_dt, k)
            if pick is None:
                # 同上, 期末也留旧口径兜底, 免得一条读数都没有的超稀疏航次整行空白。
                pick = _pick_after(end_dt, k, end_dt + NEXT_MAX)
            if pick:
                l_dt[k], last[k], l_basis[k] = pick[0], pick[1], pick[2]
        else:
            # 进行中航次: 期末取"截至 now 时间上最近"的 BERTH/NOON 读数。
            # ⚠️ 必须与已结束航次同口径(2026-10-20 定稿: BERTH/NOON 同权重按时间排序,
            # 同时刻 BERTH 优先)。旧写法 "BERTH 序列先扫, 找到即 break" 是全局 BERTH 优先,
            # 只要存在任意一条历史 BERTH 就会拿它当期末 —— 归档补回大量 BERTH 后暴露:
            # SHENGTANG 2634 期末取到 09-14 的旧 BERTH, 而 11-10 NOON 就在手里,
            # MGO 由 7.13 虚增到 110.17; HONG DA XIN 728 2635 LSFO 171.28 -> 89.14 同因。
            now_cands = []
            for src, tag in ((berth_series, "B"), (noon_series, "N")):
                for p in src:
                    if (p[0] <= now and p[1][k] is not None
                            and p[1][k] <= ROB_ABS_MAX
                            and not (has_pos[k] and p[1][k] <= 0)):
                        now_cands.append((p[0], tag, p[1][k]))
            if now_cands:
                # 时间最近者胜; 同一时刻 (time, True) > (time, False) -> BERTH 胜
                best = max(now_cands, key=lambda x: (x[0], x[1] == "B"))
                last[k] = best[2]; l_basis[k] = best[1]; l_dt[k] = best[0]
    # 加油量: 按【航次日期区间】归属, 左闭右开 —— 起航日当天的加油算本航次, 到港日当天的
    # 算下一航次。这样每个日期恰好落在一个航次里, 全船合计守恒, 不会出现漏计或双计。
    # 旧口径用 datetime 直接比 (start_dt <= bd <= end_dt), 而日志日期是 00:00、航次起点
    # 是 ETB 时分, 于是「开航当天」的加油永远被算到上一航次 —— 实测 CUL LAEMCHABANG 2642
    # 开航日 10-02 加的 790.2 MT 被推给上一航次, 本航次期初取到加油前的 155.52 后直接变负数。
    if bunker_by_date:
        lo_d = start_dt.date()
        hi_d = (end_dt or now).date()
        for dstr, tmap in bunker_by_date.items():
            try:
                bd = datetime.datetime.strptime(dstr[:10], "%Y-%m-%d").date()
            except Exception:
                continue
            if lo_d <= bd < hi_d:
                for k, t in enumerate(oils):
                    bunker_oil[k] += float((tmap or {}).get(t) or 0)
    for k in range(4):
        if not has_pos[k]:
            continue
        # 端点顺序异常(期初读数晚于期末读数, 只可能出现在极稀疏的航次)时不计算,
        # 否则会得到一个被颠倒的负数。
        if first[k] is not None and last[k] is not None and f_dt[k] <= l_dt[k]:
            raw = first[k] - last[k] + bunker_oil[k]
            # 航次总消耗默认护栏: 油耗为负通常意味期间有加油未登记 / ROB 读数异常, 不可信 -> None。
            # 但 legs 子段累加需保留原始负值(guard_negative=False): 多段净变化可 telescoping 相加,
            # 负值段即"加油导致 ROB 上升", 累加后仍得正确总消耗。
            if raw < 0 and guard_negative:
                cons[k] = None
            else:
                cons[k] = round(raw, 2)
    bunker_total = round(sum(bunker_oil), 2)
    used = [b for b in f_basis + l_basis if b]
    basis = ""
    if used:
        basis = "BERTH" if all(b == "B" for b in used) else ("NOON" if all(b == "N" for b in used) else "MIXED")
    return cons, bunker_total, bunker_oil, basis


def compute_voyages(dm_html=DM_HTML, history_csv=HISTORY_CSV, bunkering=None, bunker_types=None):
    fs = load_full_schedule(dm_html)
    if not fs:
        return []
    hist = load_history(history_csv, berth_only=False)

    # 每个 (vessel, 航次号[去方向]) 取最早 etb 作为首港; 同时记录涉及的方向
    first = {}
    dirs = defaultdict(set)
    for r in fs:
        v = (r.get("vessel") or "").strip()
        sv = split_voy(r.get("voy"))
        if not v or not sv or is_offline(v):
            continue
        base, d = sv
        dt = _parse_etb(r)
        if dt is None:
            continue
        key = (v, base)
        dirs[key].add(d)
        if key not in first or dt < first[key][0]:
            first[key] = (dt, (r.get("port") or "").strip(),
                          (r.get("route") or "").strip(), (r.get("code") or "").strip())

    bv = defaultdict(list)
    for (v, base), (dt, port, route, code) in first.items():
        bv[v].append((dt, base, port, route, code, "+".join(sorted(dirs[(v, base)]))))

    # ---- rotation: 每个 (vessel, 航次号) 的挂靠港明细(跨方向合并, 按 ETB 排序, 去重) ----
    rot_raw = defaultdict(list)   # (v,base) -> [(dt, leg), ...]
    seen_leg = set()
    for r in fs:
        v = (r.get("vessel") or "").strip()
        sv = split_voy(r.get("voy"))
        if not v or not sv or is_offline(v):
            continue
        base, d = sv
        voy = (r.get("voy") or "").strip()
        dt = _parse_etb(r)
        if dt is None:
            continue
        port = (r.get("port") or "").strip()
        sig = (v, voy, port, dt)
        if sig in seen_leg:
            continue
        seen_leg.add(sig)
        etd, eta = _parse_etd(r), _parse_eta(r)
        rot_raw[(v, base)].append((dt, {
            "voy": voy,                       # 保留原始方向航次号, 便于分辨 W / E
            "port": port,
            "eta": eta.strftime("%m-%d %H:%M") if eta else "",
            "etb": dt.strftime("%m-%d %H:%M"),
            "etd": etd.strftime("%m-%d %H:%M") if etd else "",
            "stay": (r.get("portStay") or "").strip(),
            "remark": (r.get("remark") or "").strip(),
        }))
    rot = {}
    for k, lst in rot_raw.items():
        lst.sort(key=lambda x: x[0])          # 按真实 datetime 排序(避免跨年字符串误排)
        rot[k] = [leg for _, leg in lst]

    now = datetime.datetime.now()
    voyages = []
    for v, items in bv.items():
        items.sort()
        # dt-only key: 同秒两行(不同秒被解析折叠/真实同秒)若按整 tuple 比较,
        # 会回退比较 oils 列表 -> None < float 崩溃(2026-10-10 实测 45 行回补后触发)
        allpts = sorted(hist.get(v, []), key=lambda p: p[0])   # (dt, oils, rtype)
        berth_series = [(dt, o) for dt, o, rt in allpts if rt == "BERTH"]
        # NOON 优先; report_type 为空的历史孤儿行(无法从附件重新判定)按需求3 兜底当 NOON 用,
        # 避免"窗口内无 BERTH 也无 NOON 标签"的航次因类型缺失而整体无数据。
        noon_series = [(dt, o) for dt, o, rt in allpts if rt in ("NOON", "", None)]
        for i in range(len(items)):
            start_dt, voy, port, route, code, vdirs = items[i]
            if i + 1 < len(items):
                end_dt = items[i + 1][0]
                next_voy = items[i + 1][1]
            else:
                end_dt = None
                next_voy = None
            # 状态判定(避免把未来航次当作进行中)
            if start_dt > now:
                continue                     # upcoming: 尚未开始, 无数据可算, 跳过
            elif end_dt is None or end_dt > now:
                status = "ongoing"           # 进行中(部分消耗)
            else:
                status = "completed"
            end_eff = end_dt if (end_dt and end_dt <= now) else now
            v_bunker = (bunker_types or {}).get(v)
            # berth_series / noon_series 为该船 BERTH / NOON 的 ROB 点; _window_consumption
            # 内部按本航次首港berth(start_dt) / 下一航次首港berth(end_dt) 取期初/期末 ROB,
            # 期末优先取 end_dt 之后最近读数(下一航次首港 berth 对应); 找不到 BERTH 时用 NOON 兜底。
            # 进行中航次(end 在未来或缺失): 传 None 让 _window_consumption 走"至今最近读数"
            # 分支, 显示至今的部分消耗; 否则期末窗口落在未来, 永远取不到读数 -> 整段空白。
            end_for_calc = end_dt if (end_dt is not None and end_dt <= now) else None
            cons, bunker_total, bunker_oil, basis = _window_consumption(
                berth_series, noon_series, start_dt, end_for_calc, v_bunker)
            # ---- legs: 相邻 berth 段油耗(供前端 port->port 任意区间累加) ----
            # 优先用 BERTH 报告 ROB(berth_series), 找不到时同款 NOON 兜底; 每段 from=上一港
            # berth 之后最近 ROB, to=下一港 berth 之后最近 ROB; 与航次总油耗同一公式,
            # 任意 From->To = 中间 leg 之和(保留负段以便 telescoping 相加)。
            legs = []
            raw_rot = rot_raw.get((v, voy), [])
            for li in range(len(raw_rot) - 1):
                a_dt, a_leg = raw_rot[li]
                b_dt, b_leg = raw_rot[li + 1]
                if a_dt is None or b_dt is None:
                    continue
                l_cons, _l_tot, l_oil, _lb = _window_consumption(
                    berth_series, noon_series, a_dt, b_dt, v_bunker,
                    guard_negative=False)
                legs.append({
                    "from_port": a_leg.get("port"),
                    "from_etb": a_leg.get("etb"),
                    "to_port": b_leg.get("port"),
                    "to_etb": b_leg.get("etb"),
                    "days": round((b_dt - a_dt).total_seconds() / 86400.0, 1),
                    "ls": l_cons[0], "hs": l_cons[1], "mg": l_cons[2], "us": l_cons[3],
                    "bunker_ls": round(l_oil[0], 2), "bunker_hs": round(l_oil[1], 2),
                    "bunker_mg": round(l_oil[2], 2), "bunker_us": round(l_oil[3], 2),
                    "bunker": round(sum(l_oil), 2),
                })
            # 用小数天, 避免进行中航次(start 距今不足 1 天)显示 0 天
            days = round((end_eff - start_dt).total_seconds() / 86400.0, 1)
            # 加油量(分油种) 来自《燃油添加日志》 -> bunkering_types.json
            #   已在 _window_consumption 内按航次窗口合计; 此处仅做展示四舍五入
            b_ls, b_hs, b_mg, b_us = (round(x, 2) for x in bunker_oil)
            # Bunker 合计 = 各油种之和(展示一致); 源数据为官方加油日志体积, 不含价格
            bunker = round(b_ls + b_hs + b_mg + b_us, 2)
            voyages.append({
                "vessel": v, "code": code, "lane": route, "voy": voy,
                "voy_dirs": vdirs,
                "first_port": port,
                "start": start_dt.strftime("%Y-%m-%d %H:%M"),
                "next_voy": next_voy,
                "end": (end_eff.strftime("%Y-%m-%d %H:%M") if status != "ongoing" else ""),
                "days": days, "status": status,
                "ls": cons[0], "hs": cons[1], "mg": cons[2], "us": cons[3],
                "bunker": bunker,
                "bunker_ls": b_ls, "bunker_hs": b_hs,
                "bunker_mg": b_mg, "bunker_us": b_us,
                "rotation": rot.get((v, voy), []),
                "hasData": any(c is not None for c in cons),
                "basis": basis,
                "legs": legs,
            })
    # 排序: 航线(Lane)升序 -> 同航线内 start 倒序(最新航次在最前); 无航线的排在最后
    voyages.sort(key=lambda x: x["start"], reverse=True)
    voyages.sort(key=lambda x: (
        (x.get("lane") or "").strip() == "",
        (x.get("lane") or "").strip().upper(),
    ))
    return voyages


if __name__ == "__main__":
    import pprint
    vs = compute_voyages()
    print("total voyages:", len(vs))
    withdata = [x for x in vs if x["hasData"]]
    print("with ROB data:", len(withdata))
    pprint.pprint(withdata[:12])
