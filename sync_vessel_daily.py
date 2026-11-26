#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vessel Daily Report 数据同步（Fuel Report 新 tab）.

数据源:
  1) 《Vessel Daily Consumption.xlsx》 -> TCD Daily Consumption / Actual Daily Consumption
     两个 sheet 的 speed -> 各油种每日消耗曲线 + 港口停留消耗.
  2) Full Schedule (bi_vessel_departure.csv) 的默认口径船清单:
     operator=CUL, Service Type O lanes 隐藏, ETA 窗口 = 今天-15天 ~ 下月同日.
     船代码 = vessel_voyage 前 4 位; 船名来自 vessel_names.json.

输出:
  rob_data/vessel_daily_report.json
    {"source","parsed","window","full_schedule":[{code,name,lane,calls}],
     "tcd":{code:{code,name,speeds,port_stay}}, "actual":{...}}

设计: rob_refresh.py 渲染前调用 build_daily_report() 尝试刷新; P: 盘或 GitHub
      不可达时沿用仓库里已有的 JSON, 不阻塞主流程.
"""
import os
import csv
import json
import datetime
import calendar
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(HERE, "rob_data")
OUT_JSON = os.path.join(ROB_DIR, "vessel_daily_report.json")

XLSX_DEFAULT = r"P:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新\Vessel Daily Consumption.xlsx"
LOCAL_SCHED_DIR = r"C:\Users\leahliu\WorkBuddy\Schedule-Simulation\shipping_data"
GH_RAW = "https://raw.githubusercontent.com/LeahLiuL/Schedule-Simulation/main/shipping_data/"
LANE_TYPE_O_FALLBACK = set('AE1 CNN CNT CTT CV6 EUF EXS FDR ISS LHT MCT ME2 MED NCT PRD PRF RE1 SEA SHT TCX TWS TWT WAT XRF YRD YRF'.split())


def _num(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_sheet(ws):
    """按 '新船行有 code / 后续行只有速度档' 的结构解析一个 sheet."""
    vessels, order, cur = {}, [], None
    started = False
    for row in ws.iter_rows(values_only=True):
        cols = list(row) + [None] * 8
        code = str(cols[0]).strip() if cols[0] is not None else ""
        name = str(cols[1]).strip() if cols[1] is not None else ""
        low = code.lower()
        if (not started and ("vessel code" in low or low == "code")
                and "vessel name" in str(cols[1]).lower()):
            started = True          # 表头行(Actual sheet 有两行表头)
            continue
        if not started:
            continue
        if low in ("vessel code", "code"):
            continue
        speed = _num(cols[2])
        if code:
            cur = vessels.get(code)
            if cur is None:
                cur = {"code": code, "name": name, "speeds": [],
                       "port_stay": {"lsfo": _num(cols[6]), "mgo": _num(cols[7])}}
                vessels[code] = cur
                order.append(code)
            else:
                cur["name"] = cur["name"] or name
                if _num(cols[6]) is not None:
                    cur["port_stay"]["lsfo"] = _num(cols[6])
                if _num(cols[7]) is not None:
                    cur["port_stay"]["mgo"] = _num(cols[7])
        elif cur is None:
            continue
        if speed is not None:
            cur["speeds"].append({"speed": speed, "lsfo": _num(cols[3]),
                                  "hsfo": _num(cols[4]), "mgo": _num(cols[5])})
    for v in vessels.values():
        v["speeds"].sort(key=lambda d: d["speed"])
    return {k: vessels[k] for k in order}


def _read_text(name, url):
    """本地优先, 否则 GitHub raw; 返回文本或 None."""
    local = os.path.join(LOCAL_SCHED_DIR, name)
    if os.path.isfile(local):
        try:
            with open(local, encoding="utf-8-sig") as f:
                return f.read()
        except Exception:
            pass
    try:
        with urllib.request.urlopen(url + name, timeout=15) as r:
            return r.read().decode("utf-8-sig")
    except Exception:
        return None


def full_schedule_vessels():
    """返回 (vessels, window); 失败返回 ([], None)."""
    txt = _read_text("bi_vessel_departure.csv", GH_RAW)
    if not txt:
        return [], None
    names_txt = _read_text("vessel_names.json", GH_RAW)
    lanes_txt = _read_text("lanes.json", GH_RAW)
    names = {}
    try:
        names = json.loads(names_txt) if names_txt else {}
    except Exception:
        pass
    hidden = set(LANE_TYPE_O_FALLBACK)
    if lanes_txt:
        try:
            lanes = json.loads(lanes_txt)
            hidden = {k for k, v in lanes.items()
                      if str((v or {}).get("type", "")).upper() == "O"} or hidden
        except Exception:
            pass

    t = datetime.date.today()
    year, month = (t.year + 1, 1) if t.month == 12 else (t.year, t.month + 1)
    to = datetime.date(year, month, min(t.day, calendar.monthrange(year, month)[1]))
    frm = t - datetime.timedelta(days=15)

    calls = {}
    reader = csv.DictReader(txt.splitlines())
    for r in reader:
        if (r.get("operator") or "").strip().upper() != "CUL":
            continue
        svc = (r.get("svc") or "").strip().upper()
        if svc in hidden:
            continue
        d = (r.get("arrived_anchorage") or "")[:10]
        try:
            dd = datetime.date.fromisoformat(d)
        except Exception:
            continue
        if not (frm <= dd <= to):
            continue
        vv = (r.get("vessel_voyage") or "").strip().upper()
        code = vv[:4]
        if not code:
            continue
        c = calls.setdefault(code, {"code": code,
                                    "name": names.get(code) or code,
                                    "lane": svc, "calls": 0})
        c["calls"] += 1
    out = sorted(calls.values(), key=lambda x: (x["name"], x["code"]))
    return out, {"from": frm.isoformat(), "to": to.isoformat()}


def build_daily_report(xlsx=None):
    """刷新并返回 payload 片段; 任何异常都抛给调用方兜底."""
    src = xlsx or os.environ.get("VESSEL_DAILY_XLSX", XLSX_DEFAULT)
    import openpyxl
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    tcd = parse_sheet(wb["TCD Daily Consumption"])
    actual = parse_sheet(wb["Actual Daily Consumption"])
    fs, window = full_schedule_vessels()
    out = {
        "source": os.path.basename(src),
        "parsed": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "window": window,
        "full_schedule": fs,
        "tcd": tcd,
        "actual": actual,
    }
    os.makedirs(ROB_DIR, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print("[daily] xlsx TCD=%d Actual=%d | Full Schedule=%d vessels (%s) -> %s"
          % (len(tcd), len(actual), len(fs), window, OUT_JSON))
    return out


if __name__ == "__main__":
    r = build_daily_report()
    codes = [v["code"] for v in r["full_schedule"]]
    miss = [c for c in codes if c not in r["tcd"] and c not in r["actual"]]
    print("[daily] FS codes:", ", ".join(codes))
    print("[daily] no consumption:", len(miss), ", ".join(miss))
