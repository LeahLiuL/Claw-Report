#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate static, GitHub-Pages-friendly sharded data from bapfile.db.

Two indexes:
  data/cont/<PREFIX>.json.gz   -- one row per container_no, deduped by Target Port
                                  (keyed by first 4 chars of container_no = operator prefix).
                                  Used for container-number VLOOKUP.
  data/month/<YYYY-MM>.json.gz -- all detail rows of that revenue month (for browsing
                                  by VVD / Lane / Revenue Month / special-flag filters).

Plus manifest.json (list of prefixes / months, column defs, generated_at, totals).

Output layout is ready to be served as a static site (GitHub Pages).
No Python backend required at view time.
"""
import os, json, gzip, sqlite3, time

DB = r"C:/Users/leahliu/Claw-Report/bapfile.db"
OUT = r"C:/Users/leahliu/Claw-Report/site"

BASE_COLS = ["vvd", "lane", "container_no", "fe", "pol", "pod", "type_size",
             "weight", "awk", "dg", "rf", "bb", "slot_opr", "cont_opr", "rev_month"]
CONT_COLS = BASE_COLS + ["target_port", "mismatch"]
MONTH_COLS = BASE_COLS + ["target_port"]


def gen_cont_index(con, out_dir):
    """One row per container_no, deduped by Target Port. Stream ordered by prefix
    so only one gzip file handle is open at a time."""
    d = os.path.join(out_dir, "data", "cont")
    os.makedirs(d, exist_ok=True)
    cols = CONT_COLS
    sql = (
        "WITH ranked AS ("
        " SELECT id," + ",".join(BASE_COLS) + ",target_port,"
        " ROW_NUMBER() OVER (PARTITION BY container_no"
        "   ORDER BY (CASE WHEN target_port IS NOT NULL AND pol=target_port THEN 0 ELSE 1 END),"
        "   rev_month DESC, id DESC) AS _rn"
        " FROM bapfile)"
        " SELECT " + ",".join(BASE_COLS) + ",target_port,"
        " CASE WHEN target_port IS NOT NULL AND pol=target_port THEN '' ELSE 'Y' END AS mismatch"
        " FROM ranked WHERE _rn=1"
        " ORDER BY substr(container_no,1,4), container_no"
    )
    cur = con.execute(sql)
    prefixes = []
    cur_prefix = None
    f = None
    first = True
    n = 0
    t0 = time.time()
    while True:
        row = cur.fetchone()
        if row is None:
            break
        container_no = row[cols.index("container_no")]
        prefix = container_no[:4]
        if prefix != cur_prefix:
            if f is not None:
                f.write("]}")
                f.close()
            cur_prefix = prefix
            prefixes.append(prefix)
            path = os.path.join(d, prefix + ".json.gz")
            f = gzip.open(path, "wt", encoding="utf-8", compresslevel=9)
            f.write('{"cols":' + json.dumps(cols, ensure_ascii=False) + ',"rows":[')
            first = True
        vals = [("" if v is None else v) for v in row]
        if not first:
            f.write(",")
        f.write(json.dumps(vals, ensure_ascii=False))
        first = False
        n += 1
        if n % 100000 == 0:
            print(f"  cont: {n} rows, {len(prefixes)} prefixes, {time.time()-t0:.0f}s")
    if f is not None:
        f.write("]}")
        f.close()
    print(f"cont index done: {n} containers, {len(prefixes)} prefixes, {time.time()-t0:.0f}s")
    return prefixes


def gen_month_index(con, out_dir):
    d = os.path.join(out_dir, "data", "month")
    os.makedirs(d, exist_ok=True)
    cols = MONTH_COLS
    months = [r[0] for r in con.execute(
        "SELECT DISTINCT rev_month FROM bapfile WHERE rev_month IS NOT NULL ORDER BY rev_month")]
    month_files = {}
    for m in months:
        path = os.path.join(d, m + ".json.gz")
        f = gzip.open(path, "wt", encoding="utf-8", compresslevel=9)
        f.write('{"cols":' + json.dumps(cols, ensure_ascii=False) + ',"rows":[')
        cur = con.execute(
            "SELECT " + ",".join(BASE_COLS) + ",target_port FROM bapfile WHERE rev_month=?", (m,))
        first = True
        cnt = 0
        while True:
            row = cur.fetchone()
            if row is None:
                break
            vals = [("" if v is None else v) for v in row]
            if not first:
                f.write(",")
            f.write(json.dumps(vals, ensure_ascii=False))
            first = False
            cnt += 1
        f.write("]}")
        f.close()
        month_files[m] = "data/month/" + m + ".json.gz"
        sz = os.path.getsize(path)
        print(f"  month {m}: {cnt} rows, {sz/1024/1024:.1f} MB gzip")
    return months, month_files


def gen_manifest(con, out_dir, prefixes, months, month_files):
    lanes = [r[0] for r in con.execute(
        "SELECT DISTINCT lane FROM bapfile WHERE lane IS NOT NULL AND lane<>'' ORDER BY lane")]
    cont_files = {p: "data/cont/" + p + ".json.gz" for p in prefixes}
    total_containers = con.execute(
        "SELECT COUNT(DISTINCT container_no) FROM bapfile").fetchone()[0]
    total_rows = con.execute("SELECT COUNT(*) FROM bapfile").fetchone()[0]
    manifest = {
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "base_cols": BASE_COLS,
        "cont_cols": CONT_COLS,
        "month_cols": MONTH_COLS,
        "rev_months": months,
        "lanes": lanes,
        "cont_prefixes": prefixes,
        "cont_files": cont_files,
        "month_files": month_files,
        "total_containers": total_containers,
        "total_rows": total_rows,
    }
    with open(os.path.join(out_dir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    print(f"manifest.json written: {total_containers} containers, {total_rows} rows, "
          f"{len(prefixes)} prefixes, {len(months)} months")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Generate static sharded data for GitHub Pages")
    ap.add_argument("--db", default=DB, help="path to bapfile.db")
    ap.add_argument("--out", default=OUT, help="output site directory")
    args = ap.parse_args()
    t0 = time.time()
    con = sqlite3.connect(args.db)
    os.makedirs(args.out, exist_ok=True)
    print("Generating container index...")
    prefixes = gen_cont_index(con, args.out)
    print("Generating month index...")
    months, month_files = gen_month_index(con, args.out)
    print("Writing manifest...")
    gen_manifest(con, args.out, prefixes, months, month_files)
    con.close()
    print(f"ALL DONE in {time.time()-t0:.0f}s -> {args.out}")
