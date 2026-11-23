# -*- coding: utf-8 -*-
"""
Abstract Log (航次报) 全字段解析入库.

数据源: reports.db 里 subject/attach_name 含 ABSTRACT 的报告原件 (归档 xlsx).
模板: CUL 标准 'rep form' sheet (46x55), 每个港口一对 A(到港)/D(离港) 行,
      D 行的 SEA VOY 段 (GPS距离/平均SOG/主机辅机锅炉油耗) 描述 "本港->下一港" 航段.

输出独立库 rob_data/abstract_log.db (不碰 reports.db):
  voyages  航次级一封一行
  legs     航段级每港->下一港一行
  meta     解析元信息

用法:
  python abstract_log_db.py                 # dry-run, 只解析打印统计, 不写库
  python abstract_log_db.py --apply         # 写库
  python abstract_log_db.py --limit 5       # 只处理前5封(验证)
  python abstract_log_db.py --apply --limit 5
"""
import os
import re
import sys
import json
import sqlite3
import argparse
import datetime
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_DB = os.path.join(HERE, "rob_data", "reports.db")
OUT_DB = os.path.join(HERE, "rob_data", "abstract_log.db")
OVERRIDE_JSON = os.path.join(HERE, "rob_data", "abstract_file_override.json")

# Abstract Log 原件 override: 当初 collect 选错附件(CII/NOON)的邮件, 补抓后走这里
try:
    FILE_OVERRIDE = json.load(open(OVERRIDE_JSON, encoding="utf-8")) if os.path.exists(OVERRIDE_JSON) else {}
except Exception:
    FILE_OVERRIDE = {}

# ---------- 值清洗 ----------
BAD = {"", "-", "X", "NA", "N/A", "#DIV/0!", "#VALUE!", "#REF!", "None"}


def _num(v):
    """转 float, 错误值/占位符 -> None"""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s.upper() in BAD:
        return None
    try:
        return float(s.replace(",", ""))
    except Exception:
        return None


def _text(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.upper() in BAD or s.startswith("#") else s


def _date(v):
    """datetime / Excel serial / 字符串 -> ISO 字符串"""
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and 20000 < float(v) < 80000:
        try:
            return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))).strftime("%Y-%m-%d")
        except Exception:
            return ""
    return _text(v)


# ---------- 列映射 (rep form 标准模板) ----------
# 航次级: 直接坐标
VOY_CELLS = {
    "voyage_no": "B2",
    "voy_commenced": "B3",
    "voy_completed": "B4",
    "ship_name": "M4",
    "master": "AO2",
    "chief_engineer": "AO3",
    "prepared_by": "AO4",
    "pitch": "W4",
    "voy_ttl_distance": "AN33",       # NM
    "voy_ttl_days": "AN34",
    "voy_ttl_hours": "AP34",
    "voy_ttl_mins": "AS34",
    "avg24_me_fo": "AA31",            # Ave/24h 主机燃油
    "ttl_me_fo": "Z32",               # Total 主机燃油
    "ttl_ae_fo": "AB32",
    "ttl_sea_hrs": "M32",
    "ttl_harb_hrs": "K32",
    "ttl_inport_hrs": "I32",
    "ttl_sea_dist_gps": "Q32",
    "ttl_prop_consumption": "AX32",
    # 油水结存表 (行36 COMMENCED / 行41 BALANCE / 行42 VOY TOTAL CONSUMPTION)
    "bunk_start_port": "D36",
    "bunk_start_lsfo": "I36",
    "bunk_start_mgo": "K36",
    "bunk_start_hsfo": "G36",
    "bunk_start_me_cyl": "M36",
    "bunk_end_port": "D41",
    "bunk_end_lsfo": "I41",
    "bunk_end_mgo": "K41",
    "bunk_end_hsfo": "G41",
    "tot_cons_lsfo": "I42",
    "tot_cons_mgo": "K42",
    "tot_cons_hsfo": "G42",
}

# 航段级: D 行(离港行)的列 -> 字段. A 行(到港行)的列单独列.
LEG_D_COLS = {
    "sea_hrs": "M",              # SEA VOYAGE hours
    "dist_sea_log": "P",         # SEA VOY LOG
    "dist_sea_gps": "Q",         # SEA VOY GPS  <-- 关键 distance
    "avg_sog": "R",              # Ave.SOG  <-- 关键 speed
    "avg_rpm": "S",
    "pitch": "T",
    "theory_dist": "U",
    "slip": "V",
    "wind_dir": "W",
    "wind_force": "X",
    "sea_state": "Y",
    "fo_me_sea": "Z",            # 燃油 主机 海上消耗 M/T
    "fo_me_24h": "AA",
    "fo_ae_sea": "AB",
    "fo_blr_sea": "AC",
    "do_me_sea": "AL",           # 柴油 主机 海上
    "do_ae_sea": "AM",
    "do_blr_sea": "AN",
    "total_current": "AV",
    "avg_current_spd": "AW",
    "total_cons_fo": "AX",
    "total_cons_do": "AY",
    "current_cons": "AZ",
    "sea_cond_cons": "BA",
    "fuel_index": "BB",
}
LEG_A_COLS = {
    "arr_draft": "G",            # 到港吃水
    "hours_in_port": "I",
    "harbour_hrs": "K",
    "dist_harbour": "O",         # HARB VOY distance
    "fo_me_harb": "AD",          # 燃油 主机 港内
    "fo_ae_harb": "AF",
    "fo_blr_harb": "AG",
    "fo_me_inport": "AH",
    "fo_ae_inport": "AI",
    "do_me_harb": "AO",
    "dep_draft": None,           # D 行 G
}


def _cell(ws, ref):
    try:
        return ws[ref].value
    except Exception:
        return None


def _col(ws, letter, row):
    try:
        return ws["%s%d" % (letter, row)].value
    except Exception:
        return None


def parse_report(path):
    """解析一封 Abstract Log -> (voy_dict, legs_list) 或 (None, reason)"""
    import openpyxl
    try:
        # 用非 read_only: 文件小(<150KB), 但需大量随机单元格访问,
        # read_only 模式随机访问是 O(n) 逐行扫描(实测 6.7s vs 0.44s/封)
        wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    except Exception as e:
        return None, "openpyxl fail: %s" % e
    # 找 rep form sheet
    ws = None
    for name in wb.sheetnames:
        if "rep form" in name.lower() or "abstract" in name.lower():
            ws = wb[name]
            break
    if ws is None:
        ws = wb[wb.sheetnames[0]]
    # 校验是否 Abstract Log 模板 (G2 或 A2 附近有 ABSTRACT LOG 标记)
    g2 = _text(_cell(ws, "G2")).upper()
    a2 = _text(_cell(ws, "A2")).upper()
    if "ABSTRACT" not in g2 and "ABSTRACT" not in a2:
        wb.close()
        return None, "not abstract template (G2=%r A2=%r)" % (g2, a2)
    if "VOY" not in a2 and "VOY" not in g2:
        wb.close()
        return None, "no VOY marker"

    # ---- 航次级 ----
    voy = {}
    for k, ref in VOY_CELLS.items():
        v = _cell(ws, ref)
        if k in ("voy_commenced", "voy_completed"):
            voy[k] = _date(v)
        elif k in ("ship_name", "master", "chief_engineer", "prepared_by",
                   "bunk_start_port", "bunk_end_port", "voyage_no"):
            voy[k] = _text(v)
        else:
            voy[k] = _num(v)

    # ---- 航段级: 扫描港口 A/D 对 ----
    max_row = ws.max_row or 46
    # 找汇总起始行 (A列 = 'Ave/24h' 或 'Total or Average')
    stop = max_row + 1
    for r in range(10, min(max_row, 45) + 1):
        a = _text(_col(ws, "A", r)).upper()
        if "AVE/24H" in a or "TOTAL OR AVERAGE" in a:
            stop = r
            break

    legs = []
    pending = None  # 当前 A 行(到港), 等待配对 D 行
    for r in range(10, stop):
        port = _text(_col(ws, "A", r))
        ad = _text(_col(ws, "B", r)).upper()
        if ad == "A" and port:
            pending = {
                "port": port,
                "row_a": r,
                "arr_time": _date(_col(ws, "C", r)),
            }
            for k, letter in LEG_A_COLS.items():
                if letter:
                    pending[k] = _num(_col(ws, letter, r))
        elif ad == "D" and pending is not None:
            leg = dict(pending)
            leg["dep_time"] = _date(_col(ws, "C", r))
            leg["dep_draft"] = _num(_col(ws, "G", r))
            for k, letter in LEG_D_COLS.items():
                leg[k] = _num(_col(ws, letter, r))
            legs.append(leg)
            pending = None
        # 空模板行(A/D 但无港口代码): 忽略; 保留 pending 等下一个 D
    wb.close()
    if not legs:
        return None, "no legs parsed (stop=%d)" % stop
    # from_port = 本行港口; to_port = 下一 leg 的港口 (本港->下一港航段)
    for i, leg in enumerate(legs):
        leg["from_port"] = leg.get("port", "")
        leg["to_port"] = legs[i + 1]["port"] if i + 1 < len(legs) else ""
    return voy, legs


def init_db(out=OUT_DB):
    c = sqlite3.connect(out)
    c.execute("PRAGMA journal_mode=WAL")
    c.execute("""CREATE TABLE IF NOT EXISTS voyages(
        mail_id TEXT PRIMARY KEY, vessel TEXT, received TEXT, subject TEXT,
        attach_name TEXT, file_path TEXT,
        voyage_no TEXT, voy_commenced TEXT, voy_completed TEXT, ship_name TEXT,
        master TEXT, chief_engineer TEXT, prepared_by TEXT, pitch REAL,
        voy_ttl_distance REAL, voy_ttl_days REAL, voy_ttl_hours REAL, voy_ttl_mins REAL,
        avg24_me_fo REAL, ttl_me_fo REAL, ttl_ae_fo REAL,
        ttl_sea_hrs REAL, ttl_harb_hrs REAL, ttl_inport_hrs REAL,
        ttl_sea_dist_gps REAL, ttl_prop_consumption REAL,
        bunk_start_port TEXT, bunk_start_hsfo REAL, bunk_start_lsfo REAL, bunk_start_mgo REAL,
        bunk_start_me_cyl REAL, bunk_end_port TEXT, bunk_end_hsfo REAL,
        bunk_end_lsfo REAL, bunk_end_mgo REAL,
        tot_cons_hsfo REAL, tot_cons_lsfo REAL, tot_cons_mgo REAL,
        leg_count INT, parsed_at TEXT
    )""")
    leg_cols = ["mail_id", "vessel", "voyage_no", "received", "seq",
                "from_port", "to_port", "arr_time", "dep_time",
                "arr_draft", "dep_draft", "hours_in_port", "harbour_hrs", "sea_hrs",
                "dist_harbour", "dist_sea_log", "dist_sea_gps",
                "avg_sog", "avg_rpm", "pitch", "theory_dist", "slip",
                "wind_dir", "wind_force", "sea_state",
                "fo_me_sea", "fo_me_24h", "fo_ae_sea", "fo_blr_sea",
                "fo_me_harb", "fo_ae_harb", "fo_blr_harb", "fo_me_inport", "fo_ae_inport",
                "do_me_sea", "do_ae_sea", "do_blr_sea", "do_me_harb",
                "total_current", "avg_current_spd", "total_cons_fo", "total_cons_do",
                "current_cons", "sea_cond_cons", "fuel_index"]
    c.execute("CREATE TABLE IF NOT EXISTS legs(%s)" % ", ".join(
        ["%s %s" % (col, "TEXT" if col in ("mail_id", "vessel", "voyage_no", "received",
                                          "from_port", "to_port", "arr_time", "dep_time",
                                          "wind_dir", "wind_force", "sea_state") else "REAL")
         for col in leg_cols] + ["PRIMARY KEY(mail_id, seq)"]))
    c.execute("CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT)")
    c.commit()
    return c, leg_cols


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true", help="写库(默认 dry-run)")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--out", default=OUT_DB)
    args = ap.parse_args()

    src = sqlite3.connect(SRC_DB)
    rows = src.execute("""SELECT mail_id,received,vessel,subject,attach_name,file_path FROM reports
        WHERE UPPER(COALESCE(subject,'')) LIKE '%ABSTRACT%'
           OR UPPER(COALESCE(attach_name,'')) LIKE '%ABSTRACT%'
        ORDER BY received""").fetchall()
    if args.limit:
        rows = rows[:args.limit]
    print("候选 Abstract Log 报告: %d" % len(rows))

    c, leg_cols = init_db(args.out)
    ok, fail = [], []
    ins_v = ins_l = 0
    reasons = Counter()
    for i, (mid, received, vessel, subject, attach, fpath) in enumerate(rows, 1):
        # 优先用补抓的 Abstract Log 原件
        if mid in FILE_OVERRIDE and os.path.exists(FILE_OVERRIDE[mid]):
            fpath = FILE_OVERRIDE[mid]
        if not fpath or not os.path.exists(fpath):
            fail.append((mid, "file missing")); reasons["file missing"] += 1
            continue
        res, legs = parse_report(fpath)
        if res is None:
            fail.append((mid, legs)); reasons[str(legs)[:40]] += 1
            if i <= 8:
                print("  [skip] %s %s" % (mid, legs))
            continue
        voy = res
        # 填元数据
        voy.update(mail_id=mid, received=received, vessel=vessel or "",
                   subject=subject or "", attach_name=attach or "", file_path=fpath,
                   leg_count=len(legs))
        ok.append((mid, vessel, voy.get("voyage_no"), len(legs)))
        if args.apply:
            cols = list(voy.keys())
            c.execute("INSERT OR REPLACE INTO voyages(%s) VALUES(%s)" % (
                ",".join(cols), ",".join("?" * len(cols))), [voy[k] for k in cols])
            ins_v += 1
            for seq, leg in enumerate(legs, 1):
                leg.update(mail_id=mid, vessel=vessel or "", received=received,
                           voyage_no=voy.get("voyage_no", ""), seq=seq)
                vals = [leg.get(k) for k in leg_cols]
                c.execute("INSERT OR REPLACE INTO legs(%s) VALUES(%s)" % (
                    ",".join(leg_cols), ",".join("?" * len(leg_cols))), vals)
                ins_l += 1
        if i <= 5:
            print("  [ok] %s %-16s voy=%-8s legs=%d" % (
                mid[:10], (vessel or "")[:16], voy.get("voyage_no"), len(legs)))

    if args.apply:
        c.execute("INSERT OR REPLACE INTO meta(k,v) VALUES(?,?)",
                  ("last_parse", datetime.datetime.now().isoformat(timespec="seconds")))
        c.execute("INSERT OR REPLACE INTO meta(k,v) VALUES(?,?)",
                  ("source_count", str(len(rows))))
        c.commit()
        print("\n已写库: voyages=%d legs=%d -> %s" % (ins_v, ins_l, args.out))
    else:
        print("\nDRY-RUN (未写库). 成功 %d / 失败 %d" % (len(ok), len(fail)))
    print("解析成功: %d 封, 总航段: %d" % (len(ok), sum(x[3] for x in ok)))
    print("失败原因:", dict(reasons))
    if fail and len(fail) <= 20:
        print("失败清单:", fail)
    c.close()
    src.close()


if __name__ == "__main__":
    main()
