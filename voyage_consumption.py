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
    **只累加窗口内的 ROB 下降量** → 自然排除加油(加油表现为 ROB 上升, 不计入消耗),
    因此即便尚未拿到加油明细, 航次油耗也是准确的; 加油量另行作为 bunker 字段展示。

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


def load_history(history_csv=HISTORY_CSV):
    hist = defaultdict(list)  # vessel -> [(dt, (ls,hs,mg,us)), ...]
    if not os.path.exists(history_csv):
        return hist
    try:
        with open(history_csv, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                v = (row.get("vessel") or "").strip()
                if is_offline(v):
                    continue
                dt = _parse_rt((row.get("report_time") or "")[:16])
                if dt is None:
                    continue
                oils = (_num(row.get("lsfo")), _num(row.get("hsfo")),
                        _num(row.get("mgo")), _num(row.get("ulsfo")))
                if all(o is None for o in oils):
                    continue
                hist[v].append((dt, oils))
    except Exception as e:
        print("[WARN] voyage: read history failed:", e)
    return hist


def _window_consumption(pts):
    """pts 已按时间排序; 只累加各油种相邻下降量 = 真实消耗(加油=上升被排除)。
    返回 ( [ls,hs,mg,us], bunker_total ) 。"""
    cons = [0.0, 0.0, 0.0, 0.0]
    bunker = 0.0
    for i in range(1, len(pts)):
        prev, cur = pts[i - 1][1], pts[i][1]
        for k in range(4):
            a, b = prev[k], cur[k]
            if a is None or b is None:
                continue
            d = a - b
            if d > 0:
                cons[k] += d
            elif d < 0:
                bunker += (-d)
    return [round(c, 2) for c in cons], round(bunker, 2)


def compute_voyages(dm_html=DM_HTML, history_csv=HISTORY_CSV, bunkering=None, bunker_types=None):
    fs = load_full_schedule(dm_html)
    if not fs:
        return []
    hist = load_history(history_csv)

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
        series = sorted(hist.get(v, []))
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
            pts = [p for p in series
                   if p[0] >= start_dt and p[0] <= end_eff]
            # 若窗口内无点, 退而求其次: 取窗口前后最近的两个点估算
            if len(pts) < 2 and series:
                before = [p for p in series if p[0] < start_dt]
                after = [p for p in series if p[0] > end_eff]
                if before and after:
                    pts = [before[-1], after[0]]
            cons, bunker = _window_consumption(pts)
            # 用小数天, 避免进行中航次(start 距今不足 1 天)显示 0 天
            days = round((end_eff - start_dt).total_seconds() / 86400.0, 1)
            # ---- 加油量(分油种) 来自《燃油添加日志》 -> bunkering_types.json ----
            b_ls = b_hs = b_mg = b_us = 0.0
            if bunker_types and v in bunker_types:
                for dstr, tmap in bunker_types[v].items():
                    try:
                        bd = datetime.datetime.strptime(dstr[:10], "%Y-%m-%d")
                    except Exception:
                        continue
                    if start_dt <= bd <= end_eff:
                        b_ls += float(tmap.get("ls") or 0)
                        b_hs += float(tmap.get("hs") or 0)
                        b_mg += float(tmap.get("mg") or 0)
                        b_us += float(tmap.get("us") or 0)
            b_ls, b_hs = round(b_ls, 2), round(b_hs, 2)
            b_mg, b_us = round(b_mg, 2), round(b_us, 2)
            bunker_log = round(b_ls + b_hs + b_mg + b_us, 2)
            # 官方加油日志优先; 无记录时退回窗口内 ROB 增幅推断值(避免重复计数)
            if bunker_log > 0:
                bunker = bunker_log
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
                "hasData": len(pts) >= 2,
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
