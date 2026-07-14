#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Full processing for ALL 5 sheets of Vessel Bapfile.xlsx:
  1) Streaming aggregation -> agg.json  (feeds the dashboard)
  2) Build SQLite detail DB (bapfile.db) with the 14 display columns +
     rev_month + sheet, indexed for fast VVD / LANE / time filtering.

Uses a chunked regex (C-backed) over the decompressed worksheet XML.
"""
import json, os, re, sqlite3, datetime, zipfile
from collections import Counter

SRC = r"C:\CULINES\Claw Report\Vessel Bapfile.xlsx"
HERE = os.path.dirname(os.path.abspath(__file__))
OUT_AGG = os.path.join(HERE, "agg.json")
OUT_DB = os.path.join(HERE, "bapfile.db")

COLS = ["VSL_CD","SCH_VOY_NR","SCH_DIR_CD","VVD","REVENUE_MONTH","LANE","CONT_NR",
        "VSL_SCH_PORT_CD","CARRIER_ID","FE_FLG","POR_CD","POL_CD","POD_CD","DEL_CD",
        "CONT_TP_SIZE_CD","CONT_WT","EDI_COC_SOC","BKG_COC_SOC","FIXED_FLG","AWK_FLG",
        "DG_FLG","RF_FLG","BB_FLG","TEU","BL_NO","UNIT","SLOT_OWN_PTR_ID","CONT_OPR_PTR_ID"]

def col_to_idx(letters):
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

EXCEL_EPOCH = datetime.datetime(1899, 12, 30)
_MONTH_RE = re.compile(r'^(\d{4})[-/](\d{1,2})')
def to_month(val):
    if not val:
        return None
    val = str(val).strip()
    m = _MONTH_RE.match(val)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    try:
        d = EXCEL_EPOCH + datetime.timedelta(days=int(round(float(val))))
        return d.strftime("%Y-%m")
    except Exception:
        return None

# ---- aggregate accumulators ----
total_rows = 0
rows_by_sheet = {}
month_rows = Counter(); month_teu = Counter(); month_wt = Counter(); month_unit = Counter()
lane_rows = Counter(); lane_teu = Counter(); lane_wt = Counter()
carrier_rows = Counter(); carrier_teu = Counter()
vessel_rows = Counter(); vessel_teu = Counter()
dir_rows = Counter()
ctype_rows = Counter(); ctype_teu = Counter()
fe_rows = Counter(); fe_teu = Counter()
reefer_rows = 0; dg_rows = 0; awk_rows = 0; bb_rows = 0
edisoc = Counter(); bkgsoc = Counter()
pol_rows = Counter(); pod_rows = Counter()
slot_rows = Counter(); opr_rows = Counter()
tot_teu = 0.0; tot_wt = 0.0; tot_unit = 0.0
n_full = 0; n_empty = 0
distinct_vvd = set()
samples = []

CELL_RE = re.compile(r'<c r="([A-Z]+)\d+"[^>]*?(?:>(?:<is><t>(.*?)</t></is>|<v>(.*?)</v>)?</c>|/>)')
ROW_SPLIT = b'</row>'

# ---- SQLite ----
con = sqlite3.connect(OUT_DB)
con.execute("PRAGMA synchronous=OFF")
con.execute("PRAGMA journal_mode=OFF")
con.execute("""CREATE TABLE IF NOT EXISTS bapfile (
    id INTEGER PRIMARY KEY,
    vvd TEXT, lane TEXT, container_no TEXT, fe TEXT,
    pol TEXT, pod TEXT, type_size TEXT, weight REAL,
    awk TEXT, dg TEXT, rf TEXT, bb TEXT,
    slot_opr TEXT, cont_opr TEXT, rev_month TEXT, sheet INTEGER,
    target_port TEXT
)""")
batch = []
BATCH_SIZE = 100000

def flush_batch():
    global batch
    if batch:
        con.executemany(
            "INSERT INTO bapfile (vvd,lane,container_no,fe,pol,pod,type_size,weight,"
            "awk,dg,rf,bb,slot_opr,cont_opr,rev_month,sheet,target_port) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            batch)
        batch = []

def process_row(d, sheet_idx, sheet_name):
    global total_rows, reefer_rows, dg_rows, awk_rows, bb_rows, tot_teu, tot_wt, tot_unit
    global n_full, n_empty
    total_rows += 1

    mon = to_month(d.get("REVENUE_MONTH","")) if d.get("REVENUE_MONTH") else None
    try: teu = float(d.get("TEU") or 0)
    except Exception: teu = 0.0
    try: wt = float(d.get("CONT_WT") or 0)
    except Exception: wt = 0.0
    try: unit = float(d.get("UNIT") or 0)
    except Exception: unit = 0.0

    if mon:
        month_rows[mon]+=1; month_teu[mon]+=teu; month_wt[mon]+=wt; month_unit[mon]+=unit

    lane = d.get("LANE") or "(空)"
    lane_rows[lane]+=1; lane_teu[lane]+=teu; lane_wt[lane]+=wt

    carrier = d.get("CARRIER_ID") or "(空)"
    carrier_rows[carrier]+=1; carrier_teu[carrier]+=teu

    vsl = d.get("VSL_CD") or "(空)"
    vessel_rows[vsl]+=1; vessel_teu[vsl]+=teu

    dir_rows[d.get("SCH_DIR_CD") or "(空)"]+=1

    ctype = d.get("CONT_TP_SIZE_CD") or "(空)"
    ctype_rows[ctype]+=1; ctype_teu[ctype]+=teu

    fe = d.get("FE_FLG") or "(空)"
    fe_rows[fe]+=1; fe_teu[fe]+=teu
    if fe=="Full": n_full+=1
    elif fe=="Empty": n_empty+=1

    if (d.get("RF_FLG") or "N")=="Y": reefer_rows+=1
    if (d.get("DG_FLG") or "N")=="Y": dg_rows+=1
    if (d.get("AWK_FLG") or "N")=="Y": awk_rows+=1
    if (d.get("BB_FLG") or "N")=="Y": bb_rows+=1

    edisoc[d.get("EDI_COC_SOC") or "(空)"]+=1
    bkgsoc[d.get("BKG_COC_SOC") or "(空)"]+=1
    pol_rows[d.get("POL_CD") or "(空)"]+=1
    pod_rows[d.get("POD_CD") or "(空)"]+=1
    slot_rows[d.get("SLOT_OWN_PTR_ID") or "(空)"]+=1
    opr_rows[d.get("CONT_OPR_PTR_ID") or "(空)"]+=1

    tot_teu+=teu; tot_wt+=wt; tot_unit+=unit
    vvd = d.get("VVD")
    if vvd: distinct_vvd.add(vvd)
    if sheet_idx == 0 and total_rows <= 200 and len(samples) < 8:
        samples.append({c: d.get(c,"") for c in COLS})

    # detail row
    batch.append((
        vvd, lane, d.get("CONT_NR") or "", fe,
        d.get("POL_CD") or "", d.get("POD_CD") or "", d.get("CONT_TP_SIZE_CD") or "", wt,
        d.get("AWK_FLG") or "", d.get("DG_FLG") or "", d.get("RF_FLG") or "", d.get("BB_FLG") or "",
        d.get("SLOT_OWN_PTR_ID") or "", d.get("CONT_OPR_PTR_ID") or "", mon, sheet_idx + 1,
        d.get("VSL_SCH_PORT_CD") or ""
    ))
    if len(batch) >= BATCH_SIZE:
        flush_batch()

def parse_sheet(stream, sheet_idx, sheet_name):
    global total_rows
    before = total_rows
    pending = b""
    chunk = 1 << 24
    while True:
        data = stream.read(chunk)
        if not data:
            break
        buf = pending + data
        parts = buf.split(ROW_SPLIT)
        pending = parts[-1]
        for seg in parts[:-1]:
            seg = seg + ROW_SPLIT
            if b"VSL_CD" in seg:
                continue
            try:
                s = seg.decode("utf-8")
            except Exception:
                s = seg.decode("utf-8", "replace")
            cells = []
            for m in CELL_RE.finditer(s):
                col = m.group(1)
                val = m.group(2) if m.group(2) is not None else (m.group(3) or "")
                cells.append((col, val))
            if cells:
                d = {}
                for col, val in cells:
                    idx = col_to_idx(col)
                    if 0 <= idx < len(COLS):
                        d[COLS[idx]] = val
                process_row(d, sheet_idx, sheet_name)
    if pending.strip():
        seg = pending + ROW_SPLIT
        if b"VSL_CD" not in seg:
            try:
                s = seg.decode("utf-8")
            except Exception:
                s = seg.decode("utf-8", "replace")
            cells = []
            for m in CELL_RE.finditer(s):
                col = m.group(1)
                val = m.group(2) if m.group(2) is not None else (m.group(3) or "")
                cells.append((col, val))
            if cells:
                d = {}
                for col, val in cells:
                    idx = col_to_idx(col)
                    if 0 <= idx < len(COLS):
                        d[COLS[idx]] = val
                process_row(d, sheet_idx, sheet_name)
    rows_by_sheet[sheet_name] = {"data_rows": total_rows - before}
    print(f"[OK] {sheet_name}: {total_rows-before} data rows", flush=True)

def main():
    zf = zipfile.ZipFile(SRC)
    # sheet display names from workbook.xml
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    names = re.findall(r'<sheet[^>]*name="([^"]+)"', wb)
    sheet_entries = [n for n in zf.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
    sheet_entries.sort(key=lambda x: int(re.search(r"sheet(\d+)\.xml$", x).group(1)))
    for i, entry in enumerate(sheet_entries):
        name = names[i] if i < len(names) else entry
        with zf.open(entry) as s:
            parse_sheet(s, i, name)
    flush_batch()
    con.commit()
    print("[INDEX] creating indexes ...", flush=True)
    con.execute("CREATE INDEX IF NOT EXISTS ix_vvd ON bapfile(vvd)")
    con.execute("CREATE INDEX IF NOT EXISTS ix_lane ON bapfile(lane)")
    con.execute("CREATE INDEX IF NOT EXISTS ix_month ON bapfile(rev_month)")
    con.execute("CREATE INDEX IF NOT EXISTS ix_lane_month ON bapfile(lane, rev_month)")
    con.commit()
    con.close()

    result = {
        "total_rows": total_rows,
        "distinct_vvd": len(distinct_vvd),
        "rows_by_sheet": rows_by_sheet,
        "totals": {"teu": round(tot_teu,1), "weight_tons": round(tot_wt,1),
                   "unit": round(tot_unit,1), "n_full": n_full, "n_empty": n_empty},
        "month": {"rows": dict(month_rows), "teu": {k: round(v,1) for k,v in month_teu.items()},
                  "weight_tons": {k: round(v,1) for k,v in month_wt.items()},
                  "unit": {k: round(v,1) for k,v in month_unit.items()}},
        "lane": {"rows": dict(lane_rows.most_common()), "teu": {k: round(v,1) for k,v in lane_teu.most_common()},
                 "weight_tons": {k: round(v,1) for k,v in lane_wt.most_common()}},
        "carrier": {"rows": dict(carrier_rows.most_common()), "teu": {k: round(v,1) for k,v in carrier_teu.most_common()}},
        "vessel": {"rows": dict(vessel_rows.most_common()), "teu": {k: round(v,1) for k,v in vessel_teu.most_common()}},
        "dir": dict(dir_rows),
        "cont_type": {"rows": dict(ctype_rows.most_common()), "teu": {k: round(v,1) for k,v in ctype_teu.most_common()}},
        "fe": {"rows": dict(fe_rows), "teu": {k: round(v,1) for k,v in fe_teu.items()}},
        "flags": {"reefer": reefer_rows, "dg": dg_rows, "awk": awk_rows, "bb": bb_rows},
        "coc_soc": {"edi": dict(edisoc.most_common()), "bkg": dict(bkgsoc.most_common())},
        "ports": {"pol_top": dict(pol_rows.most_common(25)), "pod_top": dict(pod_rows.most_common(25))},
        "slot_owner": dict(slot_rows.most_common(20)),
        "cont_operator": dict(opr_rows.most_common(20)),
        "samples": samples,
    }
    with open(OUT_AGG, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n[DONE] total_rows={total_rows} distinct_vvd={len(distinct_vvd)} -> {OUT_AGG}", flush=True)
    print(f"[DONE] sqlite db -> {OUT_DB}", flush=True)

if __name__ == "__main__":
    main()
