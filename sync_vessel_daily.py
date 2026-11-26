#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vessel Daily Report 数据同步（Fuel Report 新 tab）.

数据源（2026-11-26 二改: 用户已把 TCD/Actual 拆成独立 CSV）:
  1) 油耗数据(首选, 真分流):
     Schedule-Simulation/shipping_data/vessels_actual.csv  -> actual 组
     Schedule-Simulation/shipping_data/vessels_tcd.csv     -> tcd 组
     Schedule-Simulation/shipping_data/vessels_legacy.csv  -> 补档(见下)
     GitHub raw 优先(用户可能直接在浏览器改), 本地文件兜底。
     列同旧 vessels.csv: vessel_code / vessel_name / speed_knots /
         lsfo_mt_day / hsfo_mt_day / mgo_mt_day /
         portstay_lsfo_mt_day / portstay_mgo_mt_day / hire_daily
     legacy 补档规则: 按船归属(该船在哪组有档就补到哪组, 都有则补 ACT;
     都没有的按 ACT), 只补该组缺失的 speed, 已有档不覆盖; 全 0 行忽略。
  2) 兜底: 三份拆分 CSV 任一份读不到时, 回退读旧 vessels.csv
     (按行内标注列分流; 无标注则同份数据放两组 + split_by_flag=false)。
  3) 船清单: Full Schedule (bi_vessel_departure.csv) 默认口径:
     operator=CUL, Service Type O lanes 隐藏, ETA 窗口 = 今天-15天 ~ 下月同日.
     船代码 = vessel_voyage 前 4 位; 船名来自 vessel_names.json.

TCD / Actual 标注:
  拆分文件可用时 split_by_flag=true, 前端"同一份数据"提示自动消失。

输出:
  rob_data/vessel_daily_report.json
    {"source","parsed","window","split_by_flag","full_schedule",
     "tcd":{code:{code,name,speeds,port_stay}}, "actual":{...}}

设计: rob_refresh.py 渲染前调用 build_daily_report() 尝试刷新; GitHub 与本地
      都不可达时沿用仓库里已有的 JSON, 不阻塞主流程.
"""
import os
import csv
import io
import json
import datetime
import calendar
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROB_DIR = os.path.join(HERE, "rob_data")
OUT_JSON = os.path.join(ROB_DIR, "vessel_daily_report.json")

VESSELS_CSV = "vessels.csv"
VESSELS_ACTUAL_CSV = "vessels_actual.csv"
VESSELS_TCD_CSV = "vessels_tcd.csv"
VESSELS_LEGACY_CSV = "vessels_legacy.csv"
LOCAL_SCHED_DIR = r"C:\Users\leahliu\WorkBuddy\Schedule-Simulation\shipping_data"
GH_RAW = "https://raw.githubusercontent.com/LeahLiuL/Schedule-Simulation/main/shipping_data/"
LANE_TYPE_O_FALLBACK = set('AE1 CNN CNT CTT CV6 EUF EXS FDR ISS LHT MCT ME2 MED NCT PRD PRF RE1 SEA SHT TCX TWS TWT WAT XRF YRD YRF'.split())

# 可能的标注列名 / 取值(小写匹配)
FLAG_COLS = ('data_type', 'datatype', 'type', 'basis', 'category', 'kind',
             'source', 'tcd_actual', 'tcd/actual', 'tcd_or_actual',
             '数据类型', '类型', '口径', '数据口径')
TCD_WORDS = ('tcd', 'theory', 'theoretical', 'theo', 'design', 'plan', '设计', '理论', '计划')
ACT_WORDS = ('actual', 'act', 'real', '实际', '实绩')


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


def _read_text(name, url=None, remote_first=False):
    """读 shipping_data 下的文本。

    name: 文件名; url: 完整 raw URL(不给则用 GH_RAW+name).
    remote_first=True 时 GitHub 优先(用户可能浏览器直接改), 本地兜底;
     otherwise 本地优先 (兼容旧的离线工作区)。
    """
    local = os.path.join(LOCAL_SCHED_DIR, name)
    full_url = url or (GH_RAW + name)

    def _local():
        if os.path.isfile(local):
            try:
                with open(local, encoding="utf-8-sig") as f:
                    return f.read()
            except Exception:
                pass
        return None

    def _remote():
        try:
            with urllib.request.urlopen(full_url, timeout=15) as r:
                return r.read().decode("utf-8-sig")
        except Exception:
            return None

    if remote_first:
        return _remote() or _local()
    return _local() or _remote()


def _flag_bucket(row):
    """返回 'tcd' / 'actual' / None(未标注或无法识别)."""
    for col in FLAG_COLS:
        v = row.get(col)
        if v is None or not str(v).strip():
            continue
        s = str(v).strip().lower()
        if any(w in s for w in TCD_WORDS):
            return "tcd"
        if any(w in s for w in ACT_WORDS):
            return "actual"
    return None


def parse_vessels_csv(txt):
    """按 vessel_code 分组解析; 返回 (tcd, actual, split_by_flag).

    同一船同一 speed 若同时出现在两档(TCD/Actual), 分别保留各自的油种值。
    port_stay 按船取(同船多行通常相同, 有值覆盖空值)。
    """
    reader = csv.DictReader(io.StringIO(txt))
    codes = list(reader.fieldnames or [])
    col = {c: c for c in codes}
    # 容错列名大小写/空格
    def g(row, *names):
        for n in names:
            if n in row and row[n] is not None:
                return row[n]
        return None

    has_flag = any(c.strip().lower() in FLAG_COLS for c in codes)
    tcd, actual = {}, {}
    seen_flag = set()

    def _bucket_for(dst, code, name):
        d = dst.get(code)
        if d is None:
            d = {"code": code, "name": name, "speeds": [],
                 "port_stay": {"lsfo": None, "mgo": None}}
            dst[code] = d
        if name and not d["name"]:
            d["name"] = name
        return d

    for row in reader:
        code = (g(row, 'vessel_code', 'Vessel Code', 'Code') or '').strip()
        if not code or code.lower() in ('vessel code', 'code'):
            continue
        name = (g(row, 'vessel_name', 'Vessel Name') or '').strip()
        speed = _num(g(row, 'speed_knots', 'Speed (knots)', 'Speed'))
        if speed is None:
            continue
        band = {"speed": speed,
                "lsfo": _num(g(row, 'lsfo_mt_day', 'LSFO (mt/day)', 'LSFO MT/D')),
                "hsfo": _num(g(row, 'hsfo_mt_day', 'HSFO (mt/day)', 'HSFO MT/D')),
                "mgo": _num(g(row, 'mgo_mt_day', 'MGO (mt/day)', 'MGO MT/D'))}
        ps_lsfo = _num(g(row, 'portstay_lsfo_mt_day', 'PS LSFO (mt/day)', 'Port Stay LSFO'))
        ps_mgo = _num(g(row, 'portstay_mgo_mt_day', 'PS MGO (mt/day)', 'PS MGO'))

        bucket = _flag_bucket(row) if has_flag else None
        if bucket:
            seen_flag.add(bucket)
            targets = [tcd] if bucket == 'tcd' else [actual]
        elif has_flag:
            # 有标注列但该行没填: 默认归 Actual
            targets = [actual]
        else:
            # 无标注列(当前状态): 两份都放, 前端提示是同一份数据
            targets = [tcd, actual]

        for dst in targets:
            d = _bucket_for(dst, code, name)
            d["speeds"].append(dict(band))
            if ps_lsfo is not None:
                d["port_stay"]["lsfo"] = ps_lsfo
            if ps_mgo is not None:
                d["port_stay"]["mgo"] = ps_mgo

    for d in list(tcd.values()) + list(actual.values()):
        d["speeds"].sort(key=lambda b: b["speed"])
    split = has_flag and len(seen_flag) >= 2
    return tcd, actual, split


def parse_one_group(txt):
    """把单份已分流的 CSV(actual/tcd/legacy)解析成 {code:{code,name,speeds,port_stay}}.

    所有行归入同一组; 全 0 档(speed>0 但三个油种都<=0 且 PS 也<=0)视为占位忽略。
    """
    out = {}
    reader = csv.DictReader(io.StringIO(txt))
    for row in reader:
        code = (row.get('vessel_code') or '').strip()
        if not code or code.lower() in ('vessel code', 'code', 'vessel_code'):
            continue
        name = (row.get('vessel_name') or '').strip()
        speed = _num(row.get('speed_knots'))
        if speed is None or speed <= 0:
            continue
        lsfo = _num(row.get('lsfo_mt_day'))
        hsfo = _num(row.get('hsfo_mt_day'))
        mgo = _num(row.get('mgo_mt_day'))
        ps_lsfo = _num(row.get('portstay_lsfo_mt_day'))
        ps_mgo = _num(row.get('portstay_mgo_mt_day'))
        nz = lambda v: (v or 0) > 0
        if not any(nz(v) for v in (lsfo, hsfo, mgo, ps_lsfo, ps_mgo)):
            continue  # 全 0 占位船, 页面留空
        d = out.get(code)
        if d is None:
            d = {"code": code, "name": name, "speeds": [],
                 "port_stay": {"lsfo": None, "mgo": None}}
            out[code] = d
        if name and not d["name"]:
            d["name"] = name
        d["speeds"].append({"speed": speed, "lsfo": lsfo,
                            "hsfo": hsfo, "mgo": mgo})
        if ps_lsfo is not None and nz(ps_lsfo):
            d["port_stay"]["lsfo"] = ps_lsfo
        if ps_mgo is not None and nz(ps_mgo):
            d["port_stay"]["mgo"] = ps_mgo
    for d in out.values():
        d["speeds"].sort(key=lambda b: b["speed"])
    return out


def merge_legacy(tcd, actual, legacy):
    """legacy 档按船归属补进实际组: 该船在哪组有档就补到哪组(都有→ACT, 都没有→ACT).

    只补该组缺失的 speed; 已有档不覆盖(优先信任 ACT/TCD 拆分源)。
    """
    added = {"tcd": 0, "actual": 0}
    for code, ld in legacy.items():
        in_t, in_a = code in tcd, code in actual
        dst = tcd if (in_t and not in_a) else actual
        bucket = "tcd" if dst is tcd else "actual"
        d = dst.get(code)
        if d is None:
            d = {"code": code, "name": ld.get("name", code), "speeds": [],
                 "port_stay": {"lsfo": None, "mgo": None}}
            dst[code] = d
        have = {b["speed"] for b in d["speeds"]}
        for b in ld["speeds"]:
            if b["speed"] not in have:
                d["speeds"].append(dict(b))
                have.add(b["speed"])
                added[bucket] += 1
        for k in ("lsfo", "mgo"):
            if d["port_stay"].get(k) is None and ld["port_stay"].get(k) is not None:
                d["port_stay"][k] = ld["port_stay"][k]
    for d in list(tcd.values()) + list(actual.values()):
        d["speeds"].sort(key=lambda b: b["speed"])
    return added


def full_schedule_vessels():
    """返回 (vessels, window); 失败返回 ([], None)."""
    txt = _read_text("bi_vessel_departure.csv")
    if not txt:
        return [], None
    names_txt = _read_text("vessel_names.json")
    lanes_txt = _read_text("lanes.json")
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


def build_daily_report():
    """刷新并返回 payload 片段; 任何异常都抛给调用方兜底."""
    atxt = _read_text(VESSELS_ACTUAL_CSV, remote_first=True)
    ttxt = _read_text(VESSELS_TCD_CSV, remote_first=True)
    ltxt = _read_text(VESSELS_LEGACY_CSV, remote_first=True)
    legacy_note = ""
    if atxt and ttxt:
        # 首选: 用户已拆分的 ACT/TCD CSV, 真分流
        actual = parse_one_group(atxt)
        tcd = parse_one_group(ttxt)
        split = True
        source = "vessels_actual.csv + vessels_tcd.csv (Schedule-Simulation/shipping_data)"
        if ltxt:
            added = merge_legacy(tcd, actual, parse_one_group(ltxt))
            legacy_note = " legacy补档 TCD+%d ACT+%d" % (added["tcd"], added["actual"])
    else:
        # 兜底: 旧合并 CSV + 行内标注识别
        txt = _read_text(VESSELS_CSV, remote_first=True)
        if not txt:
            raise RuntimeError("vessels_actual.csv / vessels_tcd.csv / vessels.csv 均无法读取")
        tcd, actual, split = parse_vessels_csv(txt)
        source = "vessels.csv (Schedule-Simulation/shipping_data, fallback)"
    fs, window = full_schedule_vessels()
    out = {
        "source": source,
        "parsed": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "window": window,
        "split_by_flag": split,
        "full_schedule": fs,
        "tcd": tcd,
        "actual": actual,
    }
    os.makedirs(ROB_DIR, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print("[daily] %s TCD=%d Actual=%d split_by_flag=%s%s | Full Schedule=%d (%s) -> %s"
          % (source.split(' ')[0], len(tcd), len(actual), split,
             legacy_note, len(fs), window, OUT_JSON))
    return out


if __name__ == "__main__":
    r = build_daily_report()
    codes = [v["code"] for v in r["full_schedule"]]
    miss = [c for c in codes if c not in r["tcd"] and c not in r["actual"]]
    print("[daily] FS codes:", ", ".join(codes))
    print("[daily] no consumption:", len(miss), ", ".join(miss))
