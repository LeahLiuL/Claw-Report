#!/usr/bin/env python3
"""Generate DATABASES.html - a public 'database announcement center' hosted on
GitHub Pages. Contains: two databases intro, online viewers, download links,
table schemas, query examples, rebuild steps, confidentiality note.
Data stats are read live from the two SQLite DBs so numbers stay accurate.
"""
import os
import sqlite3
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
REPORTS_DB = os.path.join(HERE, "rob_data", "reports.db")
ABSTRACT_DB = os.path.join(HERE, "rob_data", "abstract_log.db")
OUT = os.path.join(HERE, "DATABASES.html")

PAGES = "https://leahliul.github.io/Claw-Report"
REL = "https://github.com/LeahLiuL/Claw-Report/releases/download/db-snapshot-2026-11-23"


def stats():
    s = sqlite3.connect(REPORTS_DB)
    n_reports = s.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
    n_fields = s.execute("SELECT COUNT(*) FROM fields").fetchone()[0]
    n_vess_s = s.execute("SELECT COUNT(DISTINCT vessel) FROM reports WHERE vessel!=''").fetchone()[0]
    dmin, dmax = s.execute("SELECT MIN(received),MAX(received) FROM reports").fetchone()
    types = s.execute("SELECT report_type,COUNT(*) FROM reports GROUP BY report_type ORDER BY 2 DESC").fetchall()
    s.close()
    a = sqlite3.connect(ABSTRACT_DB)
    n_voy = a.execute("SELECT COUNT(*) FROM voyages").fetchone()[0]
    n_legs = a.execute("SELECT COUNT(*) FROM legs").fetchone()[0]
    n_vess_a = a.execute("SELECT COUNT(DISTINCT vessel) FROM voyages WHERE vessel!=''").fetchone()[0]
    amin, amax = a.execute("SELECT MIN(voy_completed),MAX(voy_completed) FROM voyages WHERE voy_completed!=''").fetchone()
    try:
        last_parse = a.execute("SELECT v FROM meta WHERE k='last_parse'").fetchone()[0]
    except Exception:
        last_parse = ""
    leg_cols = [r[1] for r in a.execute("PRAGMA table_info(legs)")]
    voy_cols = [r[1] for r in a.execute("PRAGMA table_info(voyages)")]
    a.close()
    return dict(
        n_reports=n_reports, n_fields=n_fields, n_vess_s=n_vess_s,
        s_from=(dmin or "")[:10], s_to=(dmax or "")[:10], types=types,
        n_voy=n_voy, n_legs=n_legs, n_vess_a=n_vess_a,
        a_from=(amin or "")[:10], a_to=(amax or "")[:10], last_parse=last_parse,
        leg_cols=leg_cols, voy_cols=voy_cols,
        sz_s=os.path.getsize(REPORTS_DB) / 1e6, sz_a=os.path.getsize(ABSTRACT_DB) / 1e6,
    )


def main():
    d = stats()
    type_str = " / ".join("%s %s" % (t or "未分类", "{:,}".format(n)) for t, n in d["types"])
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    leg_cols_js = ", ".join("'%s'" % c for c in d["leg_cols"])
    voy_cols_js = ", ".join("'%s'" % c for c in d["voy_cols"])

    html = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CUL 数据库公告中心</title>
<style>
:root{
  --bg:#0d1117; --panel:#161b22; --panel2:#1c2330; --border:#2d3748;
  --txt:#e6edf3; --muted:#8b98a9; --accent:#4fc3f7; --accent2:#29b6f6;
  --up:#f85149; --down:#3fb950; --warn:#d29922; --gold:#e3b341;
  --purple:#a371f7;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI','PingFang SC','Microsoft YaHei',Arial,sans-serif;background:var(--bg);color:var(--txt);font-size:13px;line-height:1.6;}
a{color:var(--accent);text-decoration:none;}
a:hover{text-decoration:underline;}
.header{background:linear-gradient(135deg,#1a3a5c 0%,#0d5a7a 100%);padding:22px 30px;border-bottom:1px solid var(--border);}
.header h1{font-size:22px;font-weight:700;letter-spacing:.5px;}
.header .sub{font-size:12px;color:#a8c8e0;margin-top:5px;}
.header .stamp{display:inline-block;margin-top:10px;font-size:11px;background:rgba(255,255,255,.1);padding:5px 12px;border-radius:6px;border:1px solid rgba(255,255,255,.18);}
.wrap{max-width:1180px;margin:0 auto;padding:22px 24px;}
.notice{background:rgba(210,153,34,.1);border:1px solid var(--warn);border-radius:9px;padding:13px 18px;margin-bottom:22px;font-size:12px;color:#e8c877;}
.notice b{color:var(--gold);}
.intro{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:16px 20px;margin-bottom:22px;font-size:13px;color:var(--muted);}
.intro b{color:var(--txt);}
.db{background:var(--panel);border:1px solid var(--border);border-radius:11px;padding:20px 22px;margin-bottom:22px;}
.db-head{display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px;border-bottom:1px solid var(--border);padding-bottom:14px;margin-bottom:16px;}
.db-title{font-size:17px;font-weight:700;color:var(--accent);display:flex;align-items:center;gap:9px;}
.badge{font-size:10px;padding:3px 9px;border-radius:20px;font-weight:600;}
.badge.online{background:rgba(63,185,80,.15);color:var(--down);border:1px solid var(--down);}
.badge.dl{background:rgba(163,113,247,.15);color:var(--purple);border:1px solid var(--purple);}
.db-file{font-family:Consolas,monospace;font-size:12px;color:var(--muted);margin-top:6px;}
.kv{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:11px;margin-bottom:16px;}
.kv .cell{background:var(--panel2);border:1px solid var(--border);border-radius:7px;padding:10px 13px;}
.kv .cell .k{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}
.kv .cell .v{font-size:18px;font-weight:700;color:var(--gold);margin-top:3px;}
.kv .cell .v small{font-size:11px;color:var(--muted);font-weight:400;}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;}
.btn{display:inline-block;padding:9px 16px;border-radius:6px;font-size:12px;font-weight:600;text-decoration:none;}
.btn.primary{background:var(--accent2);color:#04121e;}
.btn.primary:hover{background:var(--accent);text-decoration:none;}
.btn.sec{background:var(--panel2);color:var(--txt);border:1px solid var(--border);}
.btn.sec:hover{border-color:var(--accent);text-decoration:none;}
h3{font-size:13px;color:var(--accent);margin:14px 0 8px;font-weight:600;}
.schema{background:#0a0e14;border:1px solid var(--border);border-radius:7px;padding:12px 14px;font-family:Consolas,monospace;font-size:11px;color:#9ecbff;line-height:1.9;max-height:190px;overflow:auto;word-break:break-all;}
pre.code{background:#0a0e14;border:1px solid var(--border);border-radius:7px;padding:12px 14px;font-family:Consolas,monospace;font-size:11.5px;color:#a9d3a0;line-height:1.7;overflow:auto;}
.foot{text-align:center;color:var(--muted);font-size:11px;padding:18px 0 30px;border-top:1px solid var(--border);margin-top:10px;}
@media(max-width:640px){.wrap{padding:14px;}.db{padding:15px;}}
</style>
</head>
<body>
<div class="header">
  <h1>CUL 航运数据库公告中心</h1>
  <div class="sub">船舶报告全字段库 &amp; Abstract Log 航次库 — 在线查看 / 下载 / 查询说明</div>
  <div class="stamp">数据快照更新：__STAMP__ (Asia/Shanghai)</div>
</div>
<div class="wrap">

  <div class="notice">
    <b>保密提示：</b>本页面与仓库为公开访问，库内含真实船名、航次、油耗与港口数据（<b>不含燃油价格</b>）。仅限_internal_使用，请勿外传或公开引用。
  </div>

  <div class="intro">
    这里是两份结构化数据库的统一入口。你可以<b>直接在浏览器在线查看</b>分析结果，或<b>下载 .db 文件</b>到本地用 SQLite 工具查询。
    数据每日随 CHANG-REPORT 自动化管线更新；本页统计数字在读页时从库内实时提取。
  </div>

  <!-- DATABASE 1: reports.db -->
  <div class="db">
    <div class="db-head">
      <div>
        <div class="db-title">① 船舶报告全字段库 <span class="badge dl">可下载</span></div>
        <div class="db-file">rob_data/reports.db &nbsp;·&nbsp; __SZ_S__ MB &nbsp;·&nbsp; SQLite</div>
      </div>
    </div>
    <div class="kv">
      <div class="cell"><div class="k">报告邮件</div><div class="v">__N_REPORTS__</div></div>
      <div class="cell"><div class="k">字段记录</div><div class="v">__N_FIELDS__</div></div>
      <div class="cell"><div class="k">船舶数</div><div class="v">__N_VES_S__</div></div>
      <div class="cell"><div class="k">覆盖时间</div><div class="v" style="font-size:13px">__S_FROM__<br>→ __S_TO__</div></div>
    </div>
    <div style="font-size:12px;color:var(--muted);margin-bottom:6px">报告类型：__TYPE_STR__</div>
    <div class="actions">
      <a class="btn primary" href="__REL__/reports.db" download>⬇ 下载 reports.db (__SZ_S__ MB)</a>
      <a class="btn sec" href="https://github.com/LeahLiuL/Claw-Report/releases/tag/db-snapshot-2026-11-23" target="_blank">Release 页面</a>
    </div>
    <h3>表结构</h3>
    <div class="schema">
reports (邮件级)  : mail_id, received, subject, sender, folder, attach_name, file_path,
                     vessel, code, report_type, field_count, parsed_at, collected_at
fields (字段级)  : mail_id, sheet, cell, stage, label, field, value_num, value_text, unit
                     ↑ 每封报告的「标签→值」逐行入库，stage 标记 POB/FWE/ANCHOR 等阶段
    </div>
    <h3>本地查询示例</h3>
    <pre class="code">-- 某船某字段的全历史曲线
SELECT received, field, value_num, stage
FROM fields f JOIN reports r USING(mail_id)
WHERE r.vessel='STAR BLESSING' AND f.field='ROB_LSFO'
ORDER BY received;

-- BERTH 报告六阶段 ROB 对比（诊断填报异常）
SELECT stage, value_num FROM fields
WHERE mail_id=(SELECT mail_id FROM reports
               WHERE vessel='STAR BLESSING' AND report_type='BERTH'
                 AND received LIKE '2026-07-22%')
  AND field='ROB_LSFO';</pre>
  </div>

  <!-- DATABASE 2: abstract_log.db -->
  <div class="db">
    <div class="db-head">
      <div>
        <div class="db-title">② Abstract Log 航次库 <span class="badge online">在线查看</span> <span class="badge dl">可下载</span></div>
        <div class="db-file">rob_data/abstract_log.db &nbsp;·&nbsp; __SZ_A__ MB &nbsp;·&nbsp; SQLite</div>
      </div>
    </div>
    <div class="kv">
      <div class="cell"><div class="k">航次数</div><div class="v">__N_VOY__</div></div>
      <div class="cell"><div class="k">航段数</div><div class="v">__N_LEGS__</div></div>
      <div class="cell"><div class="k">船舶数</div><div class="v">__N_VES_A__</div></div>
      <div class="cell"><div class="k">覆盖时间</div><div class="v" style="font-size:13px">__A_FROM__<br>→ __A_TO__</div></div>
    </div>
    <div style="font-size:12px;color:var(--muted);margin-bottom:6px">每港→下一港的 distance / speed / 各油种油耗；航次级含总距离、加油/结存/总消耗。</div>
    <div style="font-size:12px;background:rgba(76,175,120,.12);border:1px solid rgba(76,175,120,.35);color:#8fe0b0;padding:8px 11px;border-radius:6px;margin-bottom:10px">🔄 <b>每日自动更新</b>：每天定时从 Outlook 抓新邮件并覆盖本库（last_parse: __LAST_PARSE__）。另一台电脑每天重新下载此文件即为最新数据；reports.db（全字段库）不在此每日更新范围内。</div>
    <div class="actions">
      <a class="btn primary" href="__PAGES__/abstract_log_report.html" target="_blank">🔍 在线分析（不用下载）</a>
      <a class="btn sec" href="__REL__/abstract_log.db" download>⬇ 下载 abstract_log.db</a>
    </div>
    <h3>表结构</h3>
    <div class="schema">
voyages (航次级, 40 列) : __VOY_COLS__
legs    (航段级, 45 列) : __LEG_COLS__
    </div>
    <h3>本地查询示例</h3>
    <pre class="code">-- 每港到下一港的距离/航速/油耗
SELECT from_port, to_port, dist_sea_gps, avg_sog, fo_me_sea, fo_ae_sea
FROM legs WHERE vessel='M. MARINER' AND voyage_no='2643N'
ORDER BY seq;

-- 航次级油效 (距离优先 GPS, 为 0 回退 LOG)
SELECT vessel, voyage_no, voy_ttl_distance, tot_cons_lsfo,
       ROUND(voy_ttl_distance/tot_cons_lsfo,1) AS nm_per_mt
FROM voyages WHERE voy_ttl_distance>100 AND tot_cons_lsfo>5
ORDER BY nm_per_mt DESC LIMIT 10;</pre>
  </div>

  <div class="db">
    <div class="db-title" style="font-size:15px;">其他在线页面</div>
    <div class="actions" style="margin-top:12px;">
      <a class="btn sec" href="__PAGES__/abstract_log_report.html" target="_blank">Abstract Log 航次分析</a>
      <a class="btn sec" href="__PAGES__/rob_oil_report.html" target="_blank">ROB 盘油报告</a>
      <a class="btn sec" href="__PAGES__/cul_daily_movement.html" target="_blank">Daily Movement</a>
    </div>
  </div>

  <div class="foot">CUL 操作中心 · 数据仅供内部运营分析 · 快照 tag: db-snapshot-2026-11-23</div>
</div>
</body>
</html>'''

    rep = {
        "__STAMP__": stamp,
        "__SZ_S__": "%.0f" % d["sz_s"], "__SZ_A__": "%.2f" % d["sz_a"],
        "__N_REPORTS__": "{:,}".format(d["n_reports"]),
        "__N_FIELDS__": "{:,}".format(d["n_fields"]),
        "__N_VES_S__": str(d["n_vess_s"]),
        "__S_FROM__": d["s_from"], "__S_TO__": d["s_to"],
        "__TYPE_STR__": type_str,
        "__N_VOY__": str(d["n_voy"]), "__N_LEGS__": str(d["n_legs"]),
        "__N_VES_A__": str(d["n_vess_a"]),
        "__A_FROM__": d["a_from"], "__A_TO__": d["a_to"],
        "__LAST_PARSE__": d["last_parse"] or "—",
        "__VOY_COLS__": voy_cols_js, "__LEG_COLS__": leg_cols_js,
        "__REL__": REL, "__PAGES__": PAGES,
    }
    for k, v in rep.items():
        html = html.replace(k, v)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print("生成:", OUT)
    print("大小: %.1f KB" % (os.path.getsize(OUT) / 1024))
    print("统计: reports %s / fields %s / abstract voy %s legs %s" % (
        rep["__N_REPORTS__"], rep["__N_FIELDS__"], rep["__N_VOY__"], rep["__N_LEGS__"]))


if __name__ == "__main__":
    main()
