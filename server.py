#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Local query server for Vessel Bapfile detail table.
Serves bapfile_detail.html and JSON APIs backed by bapfile.db (SQLite).

Filterable by:
  - VVD (船名航次)         : ?vvd=  (contains, LIKE %x%)
  - LANE (航线)            : ?lane= (exact)
  - time (REVENUE_MONTH)   : ?m_from=YYYY-MM & m_to=YYYY-MM
  - optional flag toggles  : ?flags=DG&flags=RF  (AWK/DG/RF/BB -> = 'Y')

Endpoints:
  GET /                     -> detail HTML
  GET /api/months           -> ["2024-01", ...]
  GET /api/lanes            -> ["REX", ...]
  GET /api/query?vvd=&lane=&m_from=&m_to=&flags=&sort=&dir=&page=&page_size=
  GET /api/export?...       -> CSV (capped at 200000 rows)
"""
import os, json, re, sqlite3, urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(HERE, "bapfile.db")
HTML = os.path.join(HERE, "bapfile_detail.html")
PORT = 8765

COLUMNS = ["vvd","lane","container_no","fe","pol","pod","type_size","weight",
           "awk","dg","rf","bb","slot_opr","cont_opr","rev_month"]
EXPORT_LIMIT = 200000

def conn():
    c = sqlite3.connect(DB)
    c.row_factory = None
    return c

def build_where(params):
    conds = []
    args = []
    vvd = (params.get("vvd", [""])[0] or "").strip()
    lane = (params.get("lane", [""])[0] or "").strip()
    m_from = (params.get("m_from", [""])[0] or "").strip()
    m_to = (params.get("m_to", [""])[0] or "").strip()
    flags = params.get("flags", [])
    if vvd:
        conds.append("vvd LIKE ?"); args.append(f"%{vvd}%")
    if lane:
        conds.append("lane = ?"); args.append(lane)
    if m_from:
        conds.append("rev_month >= ?"); args.append(m_from)
    if m_to:
        conds.append("rev_month <= ?"); args.append(m_to)
    for f in flags:
        col = str(f).lower()
        if col in ("awk", "dg", "rf", "bb"):
            conds.append(f"{col} = ?"); args.append("Y")
    where = ("WHERE " + " AND ".join(conds)) if conds else ""
    return where, args

def parse_container_list(params):
    """Split a pasted container-number list (newline/comma/space/semicolon) into a deduped, uppercased list."""
    raw = params.get("cont_nrs", [""])[0] or ""
    out, seen = [], set()
    for x in re.split(r"[\s,;]+", raw):
        x = x.strip().upper()
        if x and x not in seen:
            seen.add(x); out.append(x)
    return out

def apply_container_join(c, nums):
    """If a container list is given, load it into a TEMP table and return a JOIN clause.
    Using a temp table avoids the SQLite 999-variable limit for large IN() lists."""
    if not nums:
        return ""
    c.execute("DROP TABLE IF EXISTS tmp_cont")
    c.execute("CREATE TEMP TABLE tmp_cont (n TEXT)")
    c.executemany("INSERT INTO tmp_cont (n) VALUES (?)", [(x,) for x in nums])
    return " JOIN tmp_cont ON bapfile.container_no = tmp_cont.n"

SORT_MAP = {
    "vvd": "vvd", "lane": "lane", "pol": "pol", "pod": "pod",
    "weight": "weight", "rev_month": "rev_month", "container_no": "container_no",
}

# Key used to decide whether two latest-month rows of the same container are "the same":
# SLOT_OPR + CONT_OPR + POL + POD (NULL-safe).
KEYEXPR = ("COALESCE(slot_opr,'~')||'|'||COALESCE(cont_opr,'~')||'|'||"
           "COALESCE(pol,'~')||'|'||COALESCE(pod,'~')")

def latest_merge_sql(cont_join, where):
    """For each container_no keep ONE row:
       - Prefer rows where POL == target_port (VSL_SCH_PORT_CD); among those take the
         newest by Revenue Month (tie-break: highest id).
       - If the container has NO row with POL == target_port, fall back to the newest
         row overall and flag it mismatch='Y'.

    Returns (ctes, body); `body` is a SELECT that can be wrapped as a subquery.
    Output column order == COLUMNS + ['target_port','mismatch']."""
    cols = ",".join(COLUMNS)
    ctes = (
        "WITH ranked AS ("
        f" SELECT id, {cols}, target_port,"
        f" ROW_NUMBER() OVER (PARTITION BY container_no"
        f"   ORDER BY (CASE WHEN target_port IS NOT NULL AND pol = target_port THEN 0 ELSE 1 END),"
        f"   rev_month DESC, id DESC) AS _rn"
        f" FROM bapfile {cont_join} {where}"
        ")"
    )
    body = (
        f" SELECT {cols}, target_port,"
        " CASE WHEN target_port IS NOT NULL AND pol = target_port THEN '' ELSE 'Y' END AS mismatch"
        " FROM ranked WHERE _rn = 1"
        " ORDER BY container_no"
    )
    return ctes, body

class H(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.wfile.write(body)

    def do_GET(self):
        p = urllib.parse.urlparse(self.path)
        if p.path in ("/", "/index.html"):
            try:
                with open(HTML, "r", encoding="utf-8") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            except Exception as e:
                self._send(500, json.dumps({"error": str(e)}, ensure_ascii=False))
            return
        if p.path == "/api/months":
            c = conn()
            rows = c.execute("SELECT DISTINCT rev_month FROM bapfile WHERE rev_month IS NOT NULL ORDER BY rev_month").fetchall()
            c.close()
            self._send(200, json.dumps([r[0] for r in rows], ensure_ascii=False))
            return
        if p.path == "/api/lanes":
            c = conn()
            rows = c.execute("SELECT DISTINCT lane FROM bapfile WHERE lane IS NOT NULL AND lane<>'' ORDER BY lane").fetchall()
            c.close()
            self._send(200, json.dumps([r[0] for r in rows], ensure_ascii=False))
            return
        if p.path == "/api/query":
            self._query(p.query)
            return
        if p.path == "/api/export":
            self._export(p.query)
            return
        self._send(404, json.dumps({"error": "not found"}, ensure_ascii=False))

    def _query(self, query):
        params = urllib.parse.parse_qs(query)
        where, args = build_where(params)
        try:
            page = max(1, int(params.get("page", ["1"])[0]))
        except Exception:
            page = 1
        try:
            page_size = min(1000, max(10, int(params.get("page_size", ["100"])[0])))
        except Exception:
            page_size = 100
        sort = params.get("sort", ["rev_month"])[0]
        direction = params.get("dir", ["desc"])[0].lower()
        sort_col = SORT_MAP.get(sort, "rev_month")
        sort_sql = f"{sort_col} {'DESC' if direction == 'desc' else 'ASC'}"
        c = conn()
        cont_join = apply_container_join(c, parse_container_list(params))
        latest = params.get("latest", [""])[0] in ("1", "Y", "y", "true")
        if latest:
            # One row per container: prefer POL==target_port, else latest month + flag mismatch
            ctes, body = latest_merge_sql(cont_join, where)
            cols_out = COLUMNS + ["target_port", "mismatch"]
            total = c.execute(f"{ctes} SELECT COUNT(*) FROM ({body})", args).fetchone()[0]
            offset = (page - 1) * page_size
            sql = (f"{ctes} SELECT * FROM ({body}) "
                   f"ORDER BY {sort_sql}, container_no LIMIT ? OFFSET ?")
        else:
            cols_out = COLUMNS
            total = c.execute(f"SELECT COUNT(*) FROM bapfile {cont_join} {where}", args).fetchone()[0]
            offset = (page - 1) * page_size
            sql = (f"SELECT {','.join(COLUMNS)} FROM bapfile {cont_join} {where} "
                   f"ORDER BY {sort_sql}, vvd LIMIT ? OFFSET ?")
        rows = c.execute(sql, args + [page_size, offset]).fetchall()
        c.close()
        out = [dict(zip(cols_out, r)) for r in rows]
        self._send(200, json.dumps({
            "total": total, "page": page, "page_size": page_size,
            "columns": cols_out, "rows": out,
        }, ensure_ascii=False))

    def _export(self, query):
        params = urllib.parse.parse_qs(query)
        where, args = build_where(params)
        c = conn()
        cont_join = apply_container_join(c, parse_container_list(params))
        latest = params.get("latest", [""])[0] in ("1", "Y", "y", "true")
        if latest:
            ctes, body = latest_merge_sql(cont_join, where)
            header = COLUMNS + ["target_port", "mismatch"]
            sql = (f"{ctes} SELECT * FROM ({body}) "
                   f"ORDER BY rev_month DESC, container_no LIMIT ?")
        else:
            sql = (f"SELECT {','.join(COLUMNS)} FROM bapfile {cont_join} {where} "
                   f"ORDER BY rev_month DESC, vvd LIMIT ?")
            header = COLUMNS
        rows = c.execute(sql, args + [EXPORT_LIMIT]).fetchall()
        c.close()
        import io, csv
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(header)
        for r in rows:
            r = list(r)
            w.writerow(["" if x is None else x for x in r])
        body = buf.getvalue().encode("utf-8-sig")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Disposition", "attachment; filename=bapfile_export.csv")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *a):
        pass

def main():
    if not os.path.exists(DB):
        print("ERROR: bapfile.db not found. Run process_all.py first.")
        return
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), H)
    srv.daemon_threads = True
    print(f"[OK] Vessel Bapfile detail server -> http://127.0.0.1:{PORT}")
    print("      (Ctrl+C to stop)")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n[STOP] server stopped")

if __name__ == "__main__":
    main()
