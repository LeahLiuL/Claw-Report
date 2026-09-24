#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
解析《Vessel Daily Consumption - TCD Daily Consumption.csv》为结构化 JSON。

源文件字段（单区块，无 sheet）：
    Vessel Code, Vessel Name, Speed, LSFO MT/D, HSFO MT/D, MGO MT/D,
    Port Stay LSFO, PS MGO, ...(空列)

含义：每张船一列 (Speed -> 各油种每日消耗 MT/D) 的设计/标称油耗曲线，
外加港口停留消耗 (Port Stay LSFO / PS MGO)。

输出：
    rob_data/tdr_consumption.json
    {
      "source": "<csv 文件名>",
      "parsed": "2026-09-24",
      "vessels": {
         "<VESSEL CODE>": {
             "code": "...", "name": "...",
             "speeds": [ {"speed":11.0,"lsfo":11.5,"hsfo":11.5,"mgo":null}, ... ],
             "port_stay": {"lsfo":4.5,"mgo":null}
         }, ...
      }
    }

同时把源 CSV 复制进仓库（rob_data/TDR_Daily_Consumption.csv）作为留档。
"""
import os
import csv
import json
import shutil
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)  # Claw-Report/

# 默认源路径（可经环境变量覆盖）
SRC_DEFAULT = r"C:\CULINES\Claw Report\Vessel Daily Consumption-TCD Daily Consumption.csv"
OUT_JSON = os.path.join(HERE, "tdr_consumption.json")
OUT_CSV = os.path.join(HERE, "TDR_Daily_Consumption.csv")


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


def parse(src):
    if not os.path.isfile(src):
        raise FileNotFoundError(src)
    vessels = {}
    order = []
    cur = None
    with open(src, encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        rows = [row for row in r]
    # 跳过表头（首个含 'Vessel Code' 的行），从下一行开始
    start = 0
    for i, row in enumerate(rows):
        if row and row[0].strip().lower().startswith("vessel code"):
            start = i + 1
            break
    for row in rows[start:]:
        # 取前 8 列（其余为空列忽略）
        cols = (row + [""] * 8)[:8]
        code = cols[0].strip()
        name = cols[1].strip()
        speed = _num(cols[2])
        lsfo = _num(cols[3])
        hsfo = _num(cols[4])
        mgo = _num(cols[5])
        ps_lsfo = _num(cols[6])
        ps_mgo = _num(cols[7])

        if code:  # 新船开始
            cur = {
                "code": code,
                "name": name,
                "speeds": [],
                "port_stay": {"lsfo": ps_lsfo, "mgo": ps_mgo},
            }
            if code not in vessels:
                vessels[code] = cur
                order.append(code)
            else:
                # 同 code 再次出现（如 MEDKON DON 在文件末尾重复），合并 port_stay
                vessels[code]["name"] = vessels[code]["name"] or name
                if ps_lsfo is not None:
                    vessels[code]["port_stay"]["lsfo"] = ps_lsfo
                if ps_mgo is not None:
                    vessels[code]["port_stay"]["mgo"] = ps_mgo
                cur = vessels[code]
        if cur is None:
            continue
        if speed is not None:
            cur["speeds"].append({
                "speed": speed,
                "lsfo": lsfo,
                "hsfo": hsfo,
                "mgo": mgo,
            })
    # 排序 speed 升序，便于绘图
    for v in vessels.values():
        v["speeds"].sort(key=lambda d: d["speed"])
    return vessels, order


def main():
    src = os.environ.get("TDR_SRC", SRC_DEFAULT)
    vessels, order = parse(src)
    out = {
        "source": os.path.basename(src),
        "parsed": datetime.date.today().isoformat(),
        "count": len(vessels),
        "vessels": {k: vessels[k] for k in order},
    }
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    # 留档源 CSV
    if os.path.isfile(src) and src != OUT_CSV:
        shutil.copy2(src, OUT_CSV)
    print("TDR parsed: %d vessels -> %s" % (len(vessels), OUT_JSON))
    # 打印摘要
    for k in order[:6]:
        v = vessels[k]
        print("  %-6s %-22s speeds=%d port_stay=%s" % (
            v["code"], v["name"], len(v["speeds"]), v["port_stay"]))
    if len(order) > 6:
        print("  ... +%d more" % (len(order) - 6))


if __name__ == "__main__":
    main()
