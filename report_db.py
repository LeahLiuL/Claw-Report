# -*- coding: utf-8 -*-
"""
Fleet Report Warehouse - 全字段报告库 (SQLite + xlsx 原件归档)

背景: rob_history.csv 只有 16 列采纳值, POB/FWE/ANCHOR 各阶段值和大量字段
(吃水/消耗/加油/天气/港口/航次...) 全被丢弃, 导致每次诊断都要重连 Outlook
打开原始附件。本模块一次性把所有报告原件抓下来归档, 再离线抽取全字段入库,
之后统计/回刷/诊断全部走本地 SQL, 脱离 Outlook。

两阶段:
  --collect   连 Outlook, 扫 vesselreport 全部文件夹, 把报告 xlsx 原件 SaveAsFile
              归档到 rob_data/report_archive/YYYY/MM/<mail_id>.xlsx, 元数据写 reports 表。
              只读邮件 + 存附件, 不解析、不改已读状态。断点续传(已有 mail_id 跳过)。
  --parse     离线读归档 xlsx, 全字段 key-value 抽取写 fields 表, 可 --force 重跑。
              改解析逻辑后只需重跑 parse, 不再碰邮箱。
  --stats     打印库内统计。

用法:
  python report_db.py --collect                 # 全量 2026 至今
  python report_db.py --collect --limit 200     # 先在 200 封上试
  python report_db.py --parse
  python report_db.py --parse --force
  python report_db.py --stats
"""
import os
import sys
import re
import csv
import json
import time
import hashlib
import tempfile
import argparse
import sqlite3
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import rob_refresh as R

DB_PATH = os.path.join(BASE, "rob_data", "reports.db")
ARCHIVE = os.path.join(BASE, "rob_data", "report_archive")
SINCE = "01/01/2026 12:00 AM"   # Outlook Restrict 格式

# -------------------------- 工具 --------------------------
def _mail_id(it):
    try:
        eid = str(it.EntryID)
    except Exception:
        eid = (str(getattr(it, "Subject", "")) + "|" +
               getattr(it, "ReceivedTime", datetime.min).strftime("%Y%m%d%H%M%S"))
    return hashlib.md5(eid.encode("utf-8", "ignore")).hexdigest()[:20]


def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def init_db(conn):
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS reports (
        mail_id      TEXT PRIMARY KEY,
        received     TEXT,
        subject      TEXT,
        sender       TEXT,
        folder       TEXT,
        attach_name  TEXT,
        file_path    TEXT,
        vessel       TEXT,
        code         TEXT,
        report_type  TEXT,
        field_count  INTEGER DEFAULT 0,
        parsed_at    TEXT,
        collected_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_reports_vessel ON reports(vessel);
    CREATE INDEX IF NOT EXISTS idx_reports_type   ON reports(report_type);
    CREATE INDEX IF NOT EXISTS idx_reports_recv   ON reports(received);

    CREATE TABLE IF NOT EXISTS fields (
        mail_id    TEXT,
        sheet      TEXT,
        cell       TEXT,
        stage      TEXT,
        label      TEXT,
        field      TEXT,
        value_num  REAL,
        value_text TEXT,
        unit       TEXT,
        PRIMARY KEY (mail_id, sheet, cell, field)
    );
    CREATE INDEX IF NOT EXISTS idx_fields_mail  ON fields(mail_id);
    CREATE INDEX IF NOT EXISTS idx_fields_field ON fields(field);
    CREATE INDEX IF NOT EXISTS idx_fields_stage ON fields(stage);

    CREATE TABLE IF NOT EXISTS meta (
        k TEXT PRIMARY KEY,
        v TEXT
    );
    """)
    conn.commit()


# -------------------------- 船名判定 --------------------------
def _build_vessel_index():
    """两条反查链: 发件邮箱 -> 船; 主题/正文 token -> 船。"""
    inv_mail = {}
    try:
        sm = R.load_sender_map()
        for norm_v, mail in sm.items():
            inv_mail[(mail or "").lower()] = norm_v
    except Exception:
        pass
    fleet = []
    try:
        fleet = R.load_fleet()
    except Exception:
        pass
    toks = []
    for r in fleet:
        v = (r.get("vessel") or "").strip()
        if not v:
            continue
        nv = R.norm(v)
        toks.append((v, nv, r.get("code", "")))
    # 长名优先, 避免 "CUL HOCHIMINH" 被 "CUL" 抢先匹配
    toks.sort(key=lambda x: -len(x[0]))
    return inv_mail, toks


def _guess_vessel(subject, sender, all_text, inv_mail, toks):
    mail = (sender or "").lower()
    if mail in inv_mail:
        nv = inv_mail[mail]
        for v, n, code in toks:
            if n == nv:
                return v, code
        return nv, ""
    hay = ((subject or "") + " " + (all_text or "")).upper()
    for v, nv, code in toks:
        if R.norm(v).upper() in hay or v.upper() in hay:
            return v, code
    return "", ""


# -------------------------- 阶段①: 抓取归档 --------------------------
def collect(limit=None, since=SINCE):
    import win32com.client as win32
    os.makedirs(ARCHIVE, exist_ok=True)
    conn = _conn()
    init_db(conn)
    have = {r[0] for r in conn.execute("SELECT mail_id FROM reports")}
    print("[collect] 已有 %d 封归档, 跳过" % len(have))

    ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    st = next((ns.Stores(i) for i in range(1, ns.Stores.Count + 1)
               if "vesselreport" in (ns.Stores(i).DisplayName or "").lower()), None)
    if st is None:
        print("[collect] 找不到 vesselreport 邮箱"); return
    root = st.GetRootFolder()
    folders = [root]
    stack = [(root, 0)]
    while stack:
        f, d = stack.pop()
        if d >= 4:
            continue
        try:
            for ch in f.Folders:
                folders.append(ch)
                stack.append((ch, d + 1))
        except Exception:
            pass
    print("[collect] 扫描 %d 个文件夹" % len(folders))

    inv_mail, toks = _build_vessel_index()
    total = 0
    added = 0
    t0 = time.time()
    for fld in folders:
        try:
            items = fld.Items
            items.Sort("[ReceivedTime]", True)
            sub = items.Restrict("[ReceivedTime] >= '%s'" % since)
        except Exception:
            continue
        for it in sub:
            total += 1
            if limit and added >= limit:
                break
            mid = _mail_id(it)
            if mid in have:
                continue
            try:
                att = R.pick_report_attachment(it)
            except Exception:
                continue
            if att is None:
                continue
            try:
                rt = it.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                rt = ""
            try:
                subj = str(it.Subject or "")
            except Exception:
                subj = ""
            sender = R.get_sender(it)
            try:
                fn = str(att.FileName or "x.xlsx")
            except Exception:
                fn = "x.xlsx"
            # 归档路径 YYYY/MM
            ym = rt[:7].replace("-", "/") if len(rt) >= 7 else "unknown/00"
            dest_dir = os.path.join(ARCHIVE, ym)
            os.makedirs(dest_dir, exist_ok=True)
            dest = os.path.join(dest_dir, mid + ".xlsx")
            try:
                att.SaveAsFile(dest)
            except Exception as e:
                print("  [SAVE-ERR] %s %s" % (mid, str(e)[:50]))
                continue
            vessel, code = _guess_vessel(subj, sender, "", inv_mail, toks)
            conn.execute(
                "INSERT OR REPLACE INTO reports(mail_id,received,subject,sender,folder,"
                "attach_name,file_path,vessel,code,collected_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                (mid, rt, subj[:300], sender[:120], fld.Name[:60], fn[:200], dest,
                 vessel, code, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            have.add(mid)
            added += 1
            if added % 200 == 0:
                print("  [collect] 新增 %d / 扫描 %d (%.0fs)" %
                      (added, total, time.time() - t0))
        if limit and added >= limit:
            break
    conn.close()
    print("[collect] DONE 新增 %d, 扫描 %d, 耗时 %.0fs" % (added, total, time.time() - t0))


# -------------------------- 阶段②: 全字段抽取 --------------------------
# 需要跳过的参考 sheet(港口代码表/船舶代码表, 每份几千行, 非报告数据)
SKIP_SHEET = re.compile(r"code\s*list|port\s*list|vessel\s*list|sheet\d*$", re.I)

STAGES = ["ANCHOR AWEIGH", "DROP ANCHOR", "RESUME SAILING",
          "STARTING DRIFTING", "DRIFTING", "POB", "FWE"]

# 右边值若是"另一个标签", 不能当值抽走
LABEL_HINT = re.compile(
    r"ROB|DRAFT|CONSUMPTION|ADD |TIME|PORT|NAME|LATITUDE|LONGITUDE|REMARK|"
    r"DISTANCE|SPEED|WEATHER|WIND|SEAS|ETA|GM|BM|SF|TUG|TERMINAL|BERTH|VOYAGE|"
    r"TOTAL|DATE|HOURS|MASTER|ENGINEER|PREPARED|PITCH|RPM|CONSUME|STATEMENT|"
    r"QUANTITY|BUNKERS|BALANCE|RECEIVED|COMMENCED", re.I)

UNIT_OK = re.compile(r"^(MT|M|NM|HRS?|KTS?|%|UNIT|UNITS|DAY|DAYS|°|′|\"|KN|RPM|PRM|HPA|LTR|LTRS|TON|TONS)$", re.I)


def _norm_field(label, stage):
    up = label.strip().upper()
    # 括号里若是油种, 保留到字段名里, 避免 BOILER(LSFO)/BOILER(MGO) 撞成同名丢值
    m = re.search(r"(ME|AE|BOILER)?\s*(CONSUMPTION)?\s*\(\s*(LSFO|HSFO|ULSFO|MGO)\s*\)", up)
    paren_oil = m.group(3) if m else ""
    head = re.sub(r"\([^)]*\)", " ", up)
    if stage:
        head = head.replace(stage, " ", 1)
    s = re.sub(r"[^A-Z0-9]+", "_", head).strip("_")
    if paren_oil:
        # 去掉 head 尾部的 CONSUMPTION 重复, 统一 <HEAD>_CONSUMPTION_<OIL>
        s = re.sub(r"_CONSUMPTION$", "", s)
        s = s + "_CONSUMPTION_" + paren_oil
    return s


def _stage_of(label):
    up = label.upper()
    for st in STAGES:
        if st in up:
            return st
    return ""


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if re.match(r"^-?\d+(\.\d+)?$", s):
        return float(s)
    return None


def extract_fields(path):
    """从一个归档 xlsx 抽取所有 key->value 字段, 返回 (rows, all_text, report_type)。"""
    import openpyxl
    import warnings
    warnings.filterwarnings("ignore")
    rows = []
    all_text = []
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return rows, "", ""
    for ws in wb.worksheets:
        if SKIP_SHEET.search(ws.title or ""):
            continue
        # 先按行取出值列表
        grid = []
        for r in ws.iter_rows(values_only=True):
            grid.append(list(r))
        for ri, row in enumerate(grid):
            for ci, cv in enumerate(row):
                if cv is None:
                    continue
                cs = str(cv).strip()
                if not cs:
                    continue
                all_text.append(cs.upper())
                # 只把"字符串单元格"当标签
                if not isinstance(cv, str):
                    continue
                label = cs
                # --- 噪声标签过滤: 单位/占位符/错误值, 不当字段 ---
                lup = label.strip().upper()
                if UNIT_OK.match(lup) or lup in ("X", "#DIV/0!", "#VALUE!", "#REF!", "N/A", "NA"):
                    continue
                if lup.startswith("#") or re.match(r"^DIV_?0", lup):
                    continue
                # 向右找第一个值(数字 或 短文本), 遇到下一个标签立即停
                val = None
                unit = ""
                for j in range(ci + 1, min(ci + 6, len(row))):
                    v = row[j]
                    if v is None or (isinstance(v, str) and not v.strip()):
                        continue
                    n = _num(v)
                    if n is not None:
                        val = n
                        # 单位在下一个非空短文本里
                        for k in range(j + 1, min(j + 3, len(row))):
                            u = row[k]
                            if u is None or (isinstance(u, str) and not u.strip()):
                                continue
                            us = str(u).strip()
                            if UNIT_OK.match(us):
                                unit = us.upper()
                            break
                        break
                    vs = str(v).strip()
                    # 文本值: 短、且不是另一个标签
                    if len(vs) <= 40 and not LABEL_HINT.search(vs) and vs not in ("MT", "M", "NM"):
                        val = vs
                        break
                    break
                if val is None:
                    continue
                stage = _stage_of(label)
                field = _norm_field(label, stage)
                if not field:
                    continue
                cell = "%s%d" % (_col_letter(ci + 1), ri + 1)
                rows.append((ws.title, cell, stage, label[:120], field,
                             val if isinstance(val, float) else None,
                             val if isinstance(val, str) else None, unit))
    try:
        wb.close()
    except Exception:
        pass
    kind = R.detect_report_kind(" ".join(all_text))
    return rows, " ".join(all_text), kind


def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_all(force=False, limit=None):
    conn = _conn()
    init_db(conn)
    inv_mail, toks = _build_vessel_index()
    # 待解析列表
    if force:
        todo = list(conn.execute("SELECT mail_id,file_path,subject,sender FROM reports"))
    else:
        todo = list(conn.execute(
            "SELECT mail_id,file_path,subject,sender FROM reports WHERE parsed_at IS NULL"))
    print("[parse] 待解析 %d 封 (force=%s)" % (len(todo), force))
    done = 0
    t0 = time.time()
    for mid, path, subj, sender in todo:
        if limit and done >= limit:
            break
        if not path or not os.path.exists(path):
            continue
        rows, all_text, kind = extract_fields(path)
        # 船名: 已有则用, 否则从正文补判
        rec = conn.execute("SELECT vessel,code FROM reports WHERE mail_id=?", (mid,)).fetchone()
        vessel, code = (rec or ["", ""])
        if not vessel:
            vessel, code = _guess_vessel(subj, sender, all_text, inv_mail, toks)
        conn.execute("DELETE FROM fields WHERE mail_id=?", (mid,))
        conn.executemany(
            "INSERT OR IGNORE INTO fields(mail_id,sheet,cell,stage,label,field,value_num,value_text,unit)"
            " VALUES(?,?,?,?,?,?,?,?,?)",
            [(mid, s, ce, st, lb, fd, vn, vt, un)
             for (s, ce, st, lb, fd, vn, vt, un) in rows])
        conn.execute(
            "UPDATE reports SET report_type=?, vessel=?, code=?, field_count=?, parsed_at=? "
            "WHERE mail_id=?",
            (kind, vessel, code, len(rows), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), mid))
        conn.commit()
        done += 1
        if done % 200 == 0:
            print("  [parse] %d / %d (%.0fs)" % (done, len(todo), time.time() - t0))
    conn.close()
    print("[parse] DONE 解析 %d 封, 耗时 %.0fs" % (done, time.time() - t0))


# -------------------------- 统计 --------------------------
def stats():
    conn = _conn()
    init_db(conn)
    print("=== 报告库统计 ===")
    for label, sql in [
        ("reports 总数", "SELECT COUNT(*) FROM reports"),
        ("已解析", "SELECT COUNT(*) FROM reports WHERE parsed_at IS NOT NULL"),
        ("未解析", "SELECT COUNT(*) FROM reports WHERE parsed_at IS NULL"),
        ("fields 总行数", "SELECT COUNT(*) FROM fields"),
        ("不同字段数", "SELECT COUNT(DISTINCT field) FROM fields"),
    ]:
        try:
            print("  %-16s: %s" % (label, conn.execute(sql).fetchone()[0]))
        except Exception as e:
            print("  %-16s: ERR %s" % (label, e))
    print("  --- 报告类型分布 ---")
    try:
        for r in conn.execute("SELECT report_type,COUNT(*) FROM reports GROUP BY report_type ORDER BY 2 DESC"):
            print("    %-12s %s" % (str(r[0] or "(空)"), r[1]))
    except Exception:
        pass
    print("  --- 常见字段 TOP 30 ---")
    try:
        for r in conn.execute("SELECT field,COUNT(*) FROM fields GROUP BY field ORDER BY 2 DESC LIMIT 30"):
            print("    %-32s %s" % (r[0], r[1]))
    except Exception:
        pass
    conn.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--collect", action="store_true")
    ap.add_argument("--parse", action="store_true")
    ap.add_argument("--stats", action="store_true")
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--since", default=SINCE)
    a = ap.parse_args()
    if a.collect:
        collect(a.limit, a.since)
    elif a.parse:
        parse_all(a.force, a.limit)
    elif a.stats:
        stats()
    else:
        ap.print_help()
