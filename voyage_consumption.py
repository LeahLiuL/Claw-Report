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
    (backfill_history.py)基于附件正文内容补全。"""
    hist = defaultdict(list)  # vessel -> [(dt, oils, rtype), ...]
    if not os.path.exists(history_csv):
        return hist
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
                hist[v].append((dt, oils, rt))
    except Exception as e:
        print("[WARN] voyage: read history failed:", e)
    return hist


def _window_consumption(berth_series, noon_series, start_dt, end_dt, bunker_by_date, guard_negative=True):
    """计算航次 [本航次首港berth=start_dt, 下一航次首港berth=end_dt] 内各油种真实消耗。

    公式: 消耗 = 期初 ROB - 期末 ROB + 期间加油量
      - 期初 ROB: 优先 BERTH 中 >= start_dt 最近读数; 无则 NOON 中 >= start_dt 最近;
                  再无则退化取 start 之前最近(BERTH 优先, 再 NOON)。
      - 期末 ROB: 优先 BERTH 中 >= end_dt 最近读数(下一航次首港 berth 对应);
                  无则 NOON 中 >= end_dt 最近; 再无则退化取 <= end_dt(或 now) 最近。
      - 加油量  : 窗口 [start_dt, end_dt] 内《燃油添加日志》各油种体积合计。
    优先用 BERTH(用户需求2: 只用本航次首港 berth -> 下一航次首港 berth 的读数, 不乱用其他日期);
    找不到 BERTH 时按需求3 用当天 NOON 兜底。
    返回 (cons[ls,hs,mg,us], bunker_total, bunker_oil[ls,hs,mg,us], basis),
    basis∈{'BERTH','NOON','MIXED',''} 标明本段消耗基于哪种报告。"""
    cons = [None, None, None, None]
    first = [None, None, None, None]
    last = [None, None, None, None]
    f_basis = [None, None, None, None]
    l_basis = [None, None, None, None]
    oils = ["ls", "hs", "mg", "us"]
    bunker_oil = [0.0, 0.0, 0.0, 0.0]
    now = datetime.datetime.now()
    if bunker_by_date:
        for dstr, tmap in bunker_by_date.items():
            try:
                bd = datetime.datetime.strptime(dstr[:10], "%Y-%m-%d")
            except Exception:
                continue
            if start_dt <= bd <= (end_dt or now):
                for k, t in enumerate(oils):
                    bunker_oil[k] += float((tmap or {}).get(t) or 0)
    # 期末允许紧邻下一航次首港 berth 的最大滞后(超过则视为更晚航次的读数, 不应用于本航次)
    END_GAP = datetime.timedelta(days=10)
    for k in range(4):
        # 期初: 窗口 [start, 本航次end] 内 BERTH 优先 / NOON 兜底(需求2: 只用本航次首港之后读数,
        #       不乱用更早或更晚日期); 窗口内无点则期初缺失 -> 本航次无 BERTH/NOON 数据。
        start_upper = end_dt or now
        for src, tag in ((berth_series, "B"), (noon_series, "N")):
            for p in src:
                if start_dt <= p[0] <= start_upper and p[1][k] is not None:
                    first[k] = p[1][k]; f_basis[k] = tag; break
            if first[k] is not None:
                break
        # 期末: 下一航次首港 berth(end) 之后、且紧邻(<= end+END_GAP)的 BERTH/NOON(需求3 兜底);
        #       越界(实为更晚航次的读数)则缺失, 避免把后续航次 berth 误当本航次期末。
        if end_dt is not None:
            el = end_dt + END_GAP
            for src, tag in ((berth_series, "B"), (noon_series, "N")):
                for p in src:
                    if end_dt <= p[0] <= el and p[1][k] is not None:
                        last[k] = p[1][k]; l_basis[k] = tag; break
                if last[k] is not None:
                    break
        else:
            for src, tag in ((berth_series, "B"), (noon_series, "N")):
                for p in reversed(src):
                    if p[0] <= now and p[1][k] is not None:
                        last[k] = p[1][k]; l_basis[k] = tag; break
                if last[k] is not None:
                    break
        if first[k] is not None and last[k] is not None:
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
            cons, bunker_total, bunker_oil, basis = _window_consumption(
                berth_series, noon_series, start_dt, end_dt, v_bunker)
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
                    berth_series, noon_series, a_dt, b_dt, v_bunker, guard_negative=False)
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
