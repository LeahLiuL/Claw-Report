"""
gen_html.py  —  CUL Daily Movement HTML Generator
每次运行会从 Excel 提取最新数据，将 TODAY_DATA 和 FULL_SCHEDULE_DATA 注入到 HTML 中，
保留历史快照机制，输出 cul_daily_movement.html（含 Summary + 完整船期 两个视图）。

新增功能 (2026-07-01):
  - 解析全部列（PORT, man in, wait, Proforma, ltm eta/etd, Voy, date, ETA, ETB, ETD,
    run, Port Stay, fsp distance, speed, ETA Delay, ETD Delay）
  - 两个视图均支持"显示列"下拉选择器（可勾选显示/隐藏列）
  - Full Schedule 默认展示 Proforma 列

用法:
    python gen_html.py
    python gen_html.py --excel "P:/path/to/CUL DAILY MOVEMENT.xlsx"
    python gen_html.py --out "P:/path/to/output/cul_daily_movement.html"
"""

import openpyxl, json, re, sys, os, argparse
from datetime import datetime, date

# ── Defaults ──────────────────────────────────────────────────────────────
DEFAULT_EXCEL = r"P:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新\CUL DAILY MOVEMENT.xlsx"
DEFAULT_HTML  = r"P:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新\cul_daily_movement.html"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Column definitions ─────────────────────────────────────────────────────
# Full column list (all columns in the Excel schedule rows)
# Used for Full Schedule view; Summary view uses a subset.
FULL_COLUMNS = [
    ('route',       'Route',            True,   True),
    ('vessel',      'Vessel',           True,   True),
    ('code',        'Code',             True,   True),
    ('pic',         'PIC',              True,   True),
    ('port',        'Port',             True,   True),
    ('manIn',       'Man In',           True,   True),   # C2
    ('wait',        'Wait',             True,   True),   # C3
    ('proforma',    'Proforma',         False,  True),   # C4
    ('ltmEta',      'LTM ETA / LTS ETB',False,  False),
    ('ltmEtd',      'LTM ETD / LTS ETD',False,  False),
    ('voy',         'Voy. No',          True,   True),   # C7
    ('date',        'Date',             True,   True),   # C8
    ('eta',         'ETA',              True,   True),   # C9
    ('etb',         'ETB',              True,   True),   # C10
    ('etd',         'ETD',              True,   True),   # C11
    ('run',         'Run',              True,   True),   # C12
    ('portStay',    'Port Stay(hr)',    True,   True),   # C13
    ('fspDistance', 'FSP Distance',     True,   True),   # C14
    ('speed',       'Speed',            True,   True),   # C15
    ('etaDelay',    'ETA Delay',        True,   True),   # C16
    ('etdDelay',    'ETD Delay',        True,   False),  # C17
]
# Summary view columns (subset)
SUMMARY_COLUMNS = [
    ('_idx',       '#',                True),   # Row number, not sortable
    ('route',      'Route',            True),
    ('vessel',     'Vessel',           True),
    ('code',       'Code',             True),
    ('port',       'Port',             True),
    ('wait',       'Wait',             True),
    ('voy',        'Voy. No',          True),
    ('eta',        'ETA',              True),
    ('etb',        'ETB',              True),
    ('etd',        'ETD',              True),
    ('portStay',   'Port Stay(hr)',    True),
    ('etaDelay',   'ETA Delay',        True),
    ('etdDelay',   'ETD Delay',        True),
    ('pic',        'PIC',              True),
    ('remark',     'Remark',           True),
]

# ── Helpers ────────────────────────────────────────────────────────────────
def fmt_dt(v):
    if v is None: return ''
    if isinstance(v, datetime): return v.strftime('%m/%d %H:%M')
    return str(v).strip()

def get_str(v):
    if v is None: return ''
    if isinstance(v, datetime): return v.strftime('%m/%d %H:%M')
    return str(v).strip()

# ── Extract data ──────────────────────────────────────────────────────────
def extract(excel_path):
    today = date.today()
    wb_src = openpyxl.load_workbook(excel_path)
    ws_src = wb_src.active

    vessel_blocks = []
    rows_total = ws_src.max_row
    i = 1
    while i <= rows_total:
        c16 = ws_src.cell(i, 16).value
        if c16 and isinstance(c16, str) and 'PIC' in c16:
            route       = get_str(ws_src.cell(i, 1).value)
            vessel_full = get_str(ws_src.cell(i, 4).value)
            vessel_code = get_str(ws_src.cell(i, 9).value)
            pic = c16.replace('PIC:', '').replace('PIC :', '').strip()

            schedule_rows = []
            remark = ''
            j = i + 2
            while j <= rows_total:
                c1_j = ws_src.cell(j, 1).value
                if c1_j and isinstance(c1_j, str) and c1_j.strip().startswith('Remark'):
                    remark = c1_j.strip().replace('Remark:', '').replace('Remark :', '').strip()
                    i = j + 1
                    break
                c16_j = ws_src.cell(j, 16).value
                if c16_j and isinstance(c16_j, str) and 'PIC' in c16_j:
                    i = j
                    break
                if isinstance(ws_src.cell(j, 9).value, datetime):
                    schedule_rows.append(j)
                j += 1
            else:
                i = rows_total + 1

            vessel_blocks.append({'route': route, 'vessel_full': vessel_full,
                                   'vessel_code': vessel_code, 'pic': pic,
                                   'schedule_rows': schedule_rows, 'remark': remark})
        else:
            i += 1

    # ── Summary: nearest ETB per vessel ──
    results = []
    for vb in vessel_blocks:
        best_row, best_etb = None, None
        for r in vb['schedule_rows']:
            etb_v = ws_src.cell(r, 10).value
            if isinstance(etb_v, datetime):
                etb_d = etb_v.date()
                if etb_d >= today and (best_etb is None or etb_d < best_etb):
                    best_etb, best_row = etb_d, r
        if best_row is None and vb['schedule_rows']:
            for r in reversed(vb['schedule_rows']):
                if isinstance(ws_src.cell(r, 10).value, datetime):
                    best_row = r; break

        if best_row:
            r = best_row
            rec = {
                'route':       vb['route'],
                'vessel':      vb['vessel_full'],
                'code':        vb['vessel_code'],
                'pic':         vb['pic'],
                'port':        get_str(ws_src.cell(r, 1).value),
                'manIn':       get_str(ws_src.cell(r, 2).value),
                'wait':        get_str(ws_src.cell(r, 3).value),
                'proforma':    get_str(ws_src.cell(r, 4).value),
                'voy':         get_str(ws_src.cell(r, 7).value),
                'ltmEta':      fmt_dt(ws_src.cell(r, 5).value),
                'ltmEtd':      fmt_dt(ws_src.cell(r, 6).value),
                'date':        fmt_dt(ws_src.cell(r, 8).value),
                'eta':         fmt_dt(ws_src.cell(r, 9).value),
                'etb':         fmt_dt(ws_src.cell(r, 10).value),
                'etd':         fmt_dt(ws_src.cell(r, 11).value),
                'run':         get_str(ws_src.cell(r, 12).value),
                'portStay':    get_str(ws_src.cell(r, 13).value),
                'fspDistance': get_str(ws_src.cell(r, 14).value),
                'speed':       get_str(ws_src.cell(r, 15).value),
                'etaDelay':    get_str(ws_src.cell(r, 16).value),
                'etdDelay':    get_str(ws_src.cell(r, 17).value),
                'remark':      vb['remark'],
            }
        else:
            rec = {'route': vb['route'], 'vessel': vb['vessel_full'],
                   'code': vb['vessel_code'], 'pic': vb['pic'],
                   'port':'','manIn':'','wait':'','proforma':'','voy':'','ltmEta':'','ltmEtd':'',
                   'date':'','eta':'','etb':'','etd':'','run':'',
                   'portStay':'','fspDistance':'','speed':'','etaDelay':'','etdDelay':'','remark': vb['remark']}
        results.append(rec)

    # ── Full Schedule: ALL port rows for ALL vessels ──
    full_schedule = []
    for vb in vessel_blocks:
        for r in vb['schedule_rows']:
            full_schedule.append({
                'route':       vb['route'],
                'vessel':      vb['vessel_full'],
                'code':        vb['vessel_code'],
                'pic':         vb['pic'],
                'port':        get_str(ws_src.cell(r, 1).value),
                'manIn':       get_str(ws_src.cell(r, 2).value),
                'wait':        get_str(ws_src.cell(r, 3).value),
                'proforma':    get_str(ws_src.cell(r, 4).value),
                'voy':         get_str(ws_src.cell(r, 7).value),
                'ltmEta':      fmt_dt(ws_src.cell(r, 5).value),
                'ltmEtd':      fmt_dt(ws_src.cell(r, 6).value),
                'date':        fmt_dt(ws_src.cell(r, 8).value),
                'eta':         fmt_dt(ws_src.cell(r, 9).value),
                'etb':         fmt_dt(ws_src.cell(r, 10).value),
                'etd':         fmt_dt(ws_src.cell(r, 11).value),
                'run':         get_str(ws_src.cell(r, 12).value),
                'portStay':    get_str(ws_src.cell(r, 13).value),
                'fspDistance': get_str(ws_src.cell(r, 14).value),
                'speed':       get_str(ws_src.cell(r, 15).value),
                'etaDelay':    get_str(ws_src.cell(r, 16).value),
                'etdDelay':    get_str(ws_src.cell(r, 17).value),
                'remark':      vb['remark'],
            })

    return {
        'date': today.strftime('%Y-%m-%d'),
        'vessels': results,
        'fullSchedule': full_schedule,
    }

# ── HTML Template ────────────────────────────────────────────────────────
# JS: COLUMN_DEFS_SUMMARY and COLUMN_DEFS_FULL are injected from Python.
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CUL Daily Movement</title>
<!-- Lazy-load xlsx-js-style -->
<script>
var _xlsxReady=false,_xlsxStyled=true;
function _loadXlsx(cb){
  if(window.XLSX){cb();return;}
  var s=document.createElement('script');
  s.src='https://unpkg.com/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
  s.onerror=function(){
    _xlsxStyled=false;
    s.src='https://cdn.bootcdn.net/ajax/libs/SheetJS/xlsx.full.min.js';
    s.onerror=function(){
      s.src='https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js';
      s.onerror=function(){alert('Failed to load Excel library. Please check your network.');};
      document.head.appendChild(s);
    };
    document.head.appendChild(s);
  };
  s.onload=cb;
  document.head.appendChild(s);
}
</script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; color: #1a2332; }

  /* ── Header ─────────────────────────────────────────────────────────── */
  .header {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    color: #fff; padding: 14px 28px 10px;
    box-shadow: 0 3px 12px rgba(31,78,121,.35);
  }
  .header-top { display: flex; align-items: center; justify-content: space-between; }
  .header-left h1 { font-size: 18px; font-weight: 700; letter-spacing: .5px; }
  .header-left .sub { font-size: 11px; opacity: .75; margin-top: 2px; }
  .header-right { display: flex; gap: 10px; align-items: center; }
  .btn { padding: 7px 16px; border: none; border-radius: 5px; font-size: 13px; font-weight: 600; cursor: pointer; transition: .15s; }
  .btn-export { background: #F6A623; color: #fff; }
  .btn-export:hover { background: #d4891a; }
  .btn-history { background: rgba(255,255,255,.18); color: #fff; border: 1px solid rgba(255,255,255,.4); }
  .btn-history:hover { background: rgba(255,255,255,.30); }

  /* ── Tabs ────────────────────────────────────────────────────────────── */
  .tabs { display: flex; gap: 0; padding: 0 28px; background: #fff; border-bottom: 2px solid #dde4ed; }
  .tab-btn {
    padding: 10px 24px; font-size: 13px; font-weight: 600; cursor: pointer;
    border: none; background: none; color: #5a6e82; border-bottom: 3px solid transparent;
    transition: .15s; margin-bottom: -2px;
  }
  .tab-btn:hover { color: #1F4E79; }
  .tab-btn.active { color: #1F4E79; border-bottom-color: #1F4E79; }
  .tab-content { display: none; }
  .tab-content.active { display: block; }

  /* ── Controls ───────────────────────────────────────────────────────── */
  .controls {
    padding: 14px 28px; background: #fff; border-bottom: 1px solid #dde4ed;
    display: flex; gap: 14px; align-items: center; flex-wrap: wrap;
  }
  .controls input, .controls select {
    padding: 6px 12px; border: 1px solid #c9d5e2; border-radius: 5px;
    font-size: 13px; outline: none; height: 34px;
  }
  .controls input:focus, .controls select:focus { border-color: #2E75B6; box-shadow: 0 0 0 2px rgba(46,117,182,.15); }
  .controls input { width: 220px; }
  .controls select { min-width: 130px; }
  .col-toggle-btn {
    padding: 6px 14px; border: 1px solid #c9d5e2; border-radius: 5px;
    font-size: 12px; cursor: pointer; background: #fff; color: #1F4E79;
    font-weight: 600; transition: .15s; height: 34px;
  }
  .col-toggle-btn:hover { background: #EBF3FB; border-color: #2E75B6; }
  .filter-btn {
    padding: 6px 14px; border: 1px solid #c9d5e2; border-radius: 5px;
    font-size: 12px; cursor: pointer; background: #fff; color: #1F4E79;
    font-weight: 600; transition: .15s; min-width: 130px; text-align: left;
    height: 34px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 180px;
  }
  .filter-btn:hover { background: #EBF3FB; border-color: #2E75B6; }
  .filter-btn.has-selection { background: #EBF3FB; border-color: #2E75B6; color: #0d3b5e; }
  .filter-dropdown { min-width: 200px; }
  .filter-dropdown .filter-actions {
    display: flex; gap: 0; border-bottom: 1px solid #e4ecf5; padding: 4px 8px;
    margin-bottom: 4px;
  }
  .filter-dropdown .filter-actions button {
    flex: 1; padding: 4px 8px; border: none; background: none; font-size: 11px;
    color: #2E75B6; cursor: pointer; font-weight: 600; border-radius: 3px; transition:.1s;
  }
  .filter-dropdown .filter-actions button:hover { background: #EBF3FB; }
  .filter-dropdown label { font-size: 12px; }
  .col-dropdown {
    display: none; position: absolute; background: #fff; border: 1px solid #c9d5e2;
    border-radius: 8px; padding: 10px 0; min-width: 200px;
    box-shadow: 0 8px 30px rgba(0,0,0,.18); z-index: 50; max-height: 400px; overflow-y: auto;
  }
  .col-dropdown.open { display: block; }
  .col-dropdown label {
    display: flex; align-items: center; gap: 8px; padding: 6px 16px;
    font-size: 12.5px; cursor: pointer; transition: .1s;
  }
  .col-dropdown label:hover { background: #f0f7ff; }
  .col-dropdown input[type="checkbox"] { accent-color: #1F4E79; width: 15px; height: 15px; }
  .stat-chip { margin-left: auto; background: #EBF3FB; border: 1px solid #c3d9f0; border-radius: 20px; padding: 4px 14px; font-size: 12px; color: #1F4E79; font-weight: 600; }
  .delay-chip { background: #fff0f0; border: 1px solid #f5c6c6; border-radius: 20px; padding: 4px 14px; font-size: 12px; color: #c00000; font-weight: 600; }

  /* ── Tables ──────────────────────────────────────────────────────────── */
  .table-wrap { overflow-x: auto; padding: 0 28px 28px; position: relative; }
  table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 12.5px; min-width: 1200px; }
  th {
    background: #1F4E79; color: #fff; font-weight: 600; padding: 9px 8px;
    text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 2;
    cursor: pointer; user-select: none;
  }
  th:hover { background: #163b5f; }
  th .sort-arrow { display: inline-block; margin-left: 4px; opacity: .5; font-size: 10px; }
  th.sort-asc .sort-arrow::after { content: '\25b2'; opacity: 1; }
  th.sort-desc .sort-arrow::after { content: '\25bc'; opacity: 1; }
  th:not(.sort-asc):not(.sort-desc) .sort-arrow::after { content: '\21c5'; }
  td { padding: 7px 9px; border-bottom: 1px solid #e4ecf5; vertical-align: middle; }

  /* Summary table row striping */
  #summaryView tr:nth-child(even) td { background: #f5f9fe; }
  #summaryView tr:nth-child(odd)  td { background: #fff; }
  #summaryView tr:hover td { background: #dcedf9 !important; }

  /* Full schedule: color-band by vessel group */
  .vessel-group-even td { background: #f5f9fe !important; }
  .vessel-group-odd  td { background: #fff !important; }
  .vessel-group-even:hover td, .vessel-group-odd:hover td { background: #dcedf9 !important; }
  .vessel-group-first td { border-top: 3px solid #2E75B6; }
  .vessel-group-last  td { border-bottom: 2px solid #c3d9f0; }

  .td-center { text-align: center; }
  .td-idx { text-align: center; color: #8fa3b8; font-size: 11px; font-weight: 500; width: 40px; }
  .td-mono { font-family: 'Consolas', 'Courier New', monospace; font-size: 12px; }
  .badge-route { display: inline-block; padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; background: #1F4E79; color: #fff; letter-spacing: .5px; }
  .badge-code { display: inline-block; padding: 2px 7px; border-radius: 4px; font-size: 11px; font-weight: 600; background: #E8F4FD; color: #1F4E79; border: 1px solid #a8cfe8; }
  .delay-tag { display: inline-block; padding: 2px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; background: #fff0f0; color: #c00000; border: 1px solid #f5c6c6; }
  .ahead-tag { display: inline-block; padding: 2px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; background: #f0fff4; color: #1a7340; border: 1px solid #b7dfca; }
  .no-data { text-align: center; padding: 40px; color: #8a9bb0; font-size: 14px; }
  .remark-cell { max-width: 260px; white-space: normal; line-height: 1.4; color: #5a6e82; font-size: 11.5px; }
  .vessel-label {
    font-weight: 700; color: #1F4E79; font-size: 12px;
    display: inline-block; padding: 1px 6px; border-radius: 3px;
    background: #E8F4FD; border: 1px solid #a8cfe8;
  }
  .proforma-cell { font-family: 'Consolas', monospace; font-size: 11px; color: #5a6e82; background: #f8fafc; border-radius: 3px; padding: 1px 5px; }

  /* ── Modal ───────────────────────────────────────────────────────────── */
  .modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.5); z-index: 100; align-items: center; justify-content: center; }
  .modal-overlay.open { display: flex; }
  .modal { background: #fff; border-radius: 10px; width: 92%; max-width: 980px; max-height: 85vh; display: flex; flex-direction: column; box-shadow: 0 20px 60px rgba(0,0,0,.3); }
  .modal-header { padding: 16px 22px; background: #1F4E79; color: #fff; border-radius: 10px 10px 0 0; display: flex; justify-content: space-between; align-items: center; }
  .modal-header h3 { font-size: 15px; }
  .modal-close { cursor: pointer; font-size: 20px; opacity: .7; background: none; border: none; color: #fff; }
  .modal-close:hover { opacity: 1; }
  .modal-body { overflow-y: auto; padding: 16px 22px; }
  .history-date-list { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 16px; }
  .history-date-btn { padding: 5px 14px; border: 1px solid #c9d5e2; border-radius: 20px; font-size: 12px; cursor: pointer; background: #fff; transition: .15s; }
  .history-date-btn:hover, .history-date-btn.active { background: #1F4E79; color: #fff; border-color: #1F4E79; }
  .history-table-wrap { overflow-x: auto; }
  .history-table-wrap table { font-size: 11.5px; min-width: 900px; }
  .footer { text-align: center; padding: 12px; color: #8a9bb0; font-size: 11px; }
</style>
</head>
<body>

<!-- ── Header ──────────────────────────────────────────────────────────── -->
<div class="header">
  <div class="header-top">
    <div class="header-left">
      <h1>&#9875; CUL VESSEL DAILY MOVEMENT</h1>
      <div class="sub" id="headerDate">Loading&hellip;</div>
    </div>
    <div class="header-right">
      <button class="btn btn-history" onclick="openHistory()">&#128203; History</button>
      <button class="btn btn-export" id="btnExport" onclick="exportExcel()">&#8595; Export Excel</button>
    </div>
  </div>
</div>

<!-- ── Tabs ────────────────────────────────────────────────────────────── -->
<div class="tabs">
  <button class="tab-btn active" data-tab="summaryView" onclick="switchTab('summaryView',this)">&#128202; Summary</button>
  <button class="tab-btn" data-tab="fullScheduleView" onclick="switchTab('fullScheduleView',this)">&#128203; Full Schedule</button>
</div>

<!-- ═══════════════════════════════════════════════════════════════════
     VIEW 1: Summary (one row per vessel, nearest ETA)
     ═════════════════════════════════════════════════════════════════════ -->
<div id="summaryView" class="tab-content active">
  <div class="controls">
    <input type="text" id="searchBox" placeholder="&#128269; Search vessel / port / PIC&hellip;" oninput="renderSummary()">
    <div style="position:relative;">
      <button class="filter-btn" id="filterRouteBtn1" onclick="toggleFilterDropdown('route','1')">All Routes</button>
      <div class="filter-dropdown col-dropdown" id="filterDropdownRoute1"></div>
    </div>
    <div style="position:relative;">
      <button class="filter-btn" id="filterVesselBtn1" onclick="toggleFilterDropdown('vessel','1')">All Vessels</button>
      <div class="filter-dropdown col-dropdown" id="filterDropdownVessel1"></div>
    </div>
    <div style="position:relative;">
      <button class="filter-btn" id="filterPicBtn1" onclick="toggleFilterDropdown('pic','1')">All PIC</button>
      <div class="filter-dropdown col-dropdown" id="filterDropdownPic1"></div>
    </div>
    <select id="filterDelay" onchange="renderSummary()">
      <option value="">All Status</option>
      <option value="delay">Delay Only</option>
      <option value="ahead">Ahead Only</option>
      <option value="normal">No Delay</option>
    </select>
    <div style="position:relative;">
      <button class="col-toggle-btn" id="colToggleBtn1" onclick="toggleColDropdown('1')">&#9881; Columns</button>
      <div class="col-dropdown" id="colDropdown1"></div>
    </div>
    <span class="stat-chip" id="statTotal">&#8212; vessels</span>
    <span class="delay-chip" id="statDelay">&#8212; delayed</span>
  </div>
  <div class="table-wrap">
    <table id="summaryTable">
      <thead><tr id="summaryThead"></tr></thead>
      <tbody id="summaryTbody"></tbody>
    </table>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════════════════
     VIEW 2: Full Schedule (all ports for all vessels)
     ═════════════════════════════════════════════════════════════════════ -->
<div id="fullScheduleView" class="tab-content">
  <div class="controls">
    <input type="text" id="searchBox2" placeholder="&#128269; Search vessel / port / PIC / voyage&hellip;" oninput="renderFullSchedule()">
    <div style="position:relative;">
      <button class="filter-btn" id="filterRouteBtn2" onclick="toggleFilterDropdown('route','2')">All Routes</button>
      <div class="filter-dropdown col-dropdown" id="filterDropdownRoute2"></div>
    </div>
    <div style="position:relative;">
      <button class="filter-btn" id="filterVesselBtn2" onclick="toggleFilterDropdown('vessel','2')">All Vessels</button>
      <div class="filter-dropdown col-dropdown" id="filterDropdownVessel2"></div>
    </div>
    <div style="position:relative;">
      <button class="filter-btn" id="filterPicBtn2" onclick="toggleFilterDropdown('pic','2')">All PIC</button>
      <div class="filter-dropdown col-dropdown" id="filterDropdownPic2"></div>
    </div>
    <div style="position:relative;">
      <button class="col-toggle-btn" id="colToggleBtn2" onclick="toggleColDropdown('2')">&#9881; Columns</button>
      <div class="col-dropdown" id="colDropdown2"></div>
    </div>
    <span class="stat-chip" id="statTotal2">&#8212; rows</span>
  </div>
  <div class="table-wrap">
    <table id="fullTable">
      <thead><tr id="fullThead"></tr></thead>
      <tbody id="fullTbody"></tbody>
    </table>
  </div>
</div>

<!-- ── History Modal ──────────────────────────────────────────────────── -->
<div class="modal-overlay" id="historyModal">
  <div class="modal">
    <div class="modal-header">
      <h3>&#128203; Historical Records</h3>
      <button class="modal-close" onclick="closeHistory()">&#10005;</button>
    </div>
    <div class="modal-body">
      <div class="history-date-list" id="historyDateList"></div>
      <div class="history-table-wrap" id="historyTableWrap"></div>
    </div>
  </div>
</div>

<div class="footer">CUL Daily Movement Dashboard &nbsp;|&nbsp; Data from CUL DAILY MOVEMENT.xlsx &nbsp;|&nbsp; <span id="footerTs"></span></div>

<!-- ── JavaScript ───────────────────────────────────────────────────────── -->
<script>
const SNAPSHOTS = {};
const TODAY_DATA = __TODAY_DATA__;
const COLUMN_DEFS_SUMMARY = __COLUMN_DEFS_SUMMARY__;
const COLUMN_DEFS_FULL    = __COLUMN_DEFS_FULL__;

// Visible columns state (key = viewId '1' or '2', value = Set of colKeys)
var visibleCols = {
  '1': new Set(COLUMN_DEFS_SUMMARY.filter(c=>c.defaultVisible).map(c=>c.key)),
  '2': new Set(COLUMN_DEFS_FULL.filter(c=>c.defaultVisible).map(c=>c.key)),
};

// Multi-select filter state: {route1: Set, pic1: Set, route2: Set, pic2: Set}
// Default: null means "show all" (empty Set also means show all, used after init)
var filterSelections = {route1: null, vessel1: null, pic1: null, route2: null, vessel2: null, pic2: null};

function setAllFilterOptions(type, viewId){
  var key = type + viewId;
  // Build set of all possible values
  var allValues;
  if(viewId==='1'){
    allValues = new Set(summaryData.map(function(r){return r[type];}));
  } else {
    allValues = new Set(fullData.map(function(r){return r[type];}));
  }
  filterSelections[key] = new Set(allValues);
}

function getFilterSelected(type, viewId){
  var key = type + viewId;
  var sel = filterSelections[key];
  if(sel == null || sel.size === 0) return null;             // null / undefined / empty = show all
  return sel;
}

function buildFilterDropdown(type, viewId){
  var dd = document.getElementById('filterDropdown'+type.charAt(0).toUpperCase()+type.slice(1)+viewId);
  dd.innerHTML = '';

  // Get all unique values
  var dataArr = viewId==='1' ? summaryData : fullData;
  var values = [...new Set(dataArr.map(function(r){return r[type];}))].sort();

  // Select All / Clear All actions
  var actions = document.createElement('div');
  actions.className = 'filter-actions';
  var selAll = document.createElement('button');
  selAll.textContent = 'Select All';
  selAll.onclick = function(e){ e.stopPropagation(); setAllFilterOptions(type, viewId); buildFilterDropdown(type, viewId); updateFilterButton(type, viewId); rerender(viewId); };
  var clrAll = document.createElement('button');
  clrAll.textContent = '全不选';
  clrAll.onclick = function(e){ e.stopPropagation(); filterSelections[type+viewId]=new Set(['\x00NONE\x00']); buildFilterDropdown(type, viewId); updateFilterButton(type, viewId); rerender(viewId); };
  actions.appendChild(selAll);
  actions.appendChild(clrAll);
  dd.appendChild(actions);

  var sel = getFilterSelected(type, viewId);
  var isAll = sel===null;
  values.forEach(function(v){
    var label = document.createElement('label');
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = isAll || sel.has(v);
    cb.onchange = function(e){
      e.stopPropagation();
      if(!filterSelections[type+viewId]) filterSelections[type+viewId] = new Set();
      if(isAll){
        // Transition from "all" to explicit set
        filterSelections[type+viewId] = new Set(values);
        isAll = false;
        sel = filterSelections[type+viewId];
      }
      if(this.checked) filterSelections[type+viewId].add(v);
      else filterSelections[type+viewId].delete(v);
      // If all are deselected, show nothing (not "show all")
      if(filterSelections[type+viewId].size===0) filterSelections[type+viewId]=new Set(['\x00NONE\x00']);
      updateFilterButton(type, viewId);
      rerender(viewId);
    };
    label.appendChild(cb);
    label.appendChild(document.createTextNode(' '+v));
    dd.appendChild(label);
  });
}

function updateFilterButton(type, viewId){
  var btn = document.getElementById('filter'+type.charAt(0).toUpperCase()+type.slice(1)+'Btn'+viewId);
  var key = type + viewId;
  var fs = filterSelections[key];
  var dataArr = viewId==='1' ? summaryData : fullData;
  var totalValues = new Set(dataArr.map(function(r){return r[type];})).size;
  var isNone = fs && fs.has('\x00NONE\x00');
  if(fs===null || (!isNone && fs.size===totalValues)){
    btn.textContent = type==='route' ? 'All Routes' : type==='vessel' ? 'All Vessels' : 'All PIC';
    btn.classList.remove('has-selection');
  } else if(isNone){
    btn.textContent = 'None';
    btn.classList.add('has-selection');
  } else {
    btn.textContent = fs.size + ' selected';
    btn.classList.add('has-selection');
  }
}

function toggleFilterDropdown(type, viewId){
  var ddId = 'filterDropdown'+type.charAt(0).toUpperCase()+type.slice(1)+viewId;
  var dd = document.getElementById(ddId);
  // Rebuild to ensure checkboxes reflect current state
  buildFilterDropdown(type, viewId);
  var isOpen = dd.classList.contains('open');
  document.querySelectorAll('.col-dropdown,.filter-dropdown').forEach(function(d){d.classList.remove('open');});
  if(!isOpen) dd.classList.add('open');
}

function rerender(viewId){
  if(viewId==='1') renderSummary();
  else renderFullSchedule();
}

function loadSnapshots(){try{const s=localStorage.getItem('cul_movement_history');if(s)Object.assign(SNAPSHOTS,JSON.parse(s));}catch(e){}}
function saveSnapshot(d){SNAPSHOTS[d.date]=d.vessels;try{localStorage.setItem('cul_movement_history',JSON.stringify(SNAPSHOTS));}catch(e){}}

/* ── Column toggle dropdown ──────────────────────────────────────────── */
function buildColDropdown(viewId){
  var defs = viewId==='1' ? COLUMN_DEFS_SUMMARY : COLUMN_DEFS_FULL;
  var dd = document.getElementById('colDropdown'+viewId);
  dd.innerHTML = '';
  defs.forEach(function(col){
    var label = document.createElement('label');
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = visibleCols[viewId].has(col.key);
    cb.onchange = function(){
      if(this.checked) visibleCols[viewId].add(col.key);
      else visibleCols[viewId].delete(col.key);
      if(viewId==='1') renderSummary();
      else renderFullSchedule();
    };
    label.appendChild(cb);
    label.appendChild(document.createTextNode(' '+col.label));
    dd.appendChild(label);
  });
}
function toggleColDropdown(viewId){
  var dd = document.getElementById('colDropdown'+viewId);
  var isOpen = dd.classList.contains('open');
  // Close all dropdowns first
  document.querySelectorAll('.col-dropdown').forEach(function(d){d.classList.remove('open');});
  if(!isOpen) dd.classList.add('open');
}
// Close dropdowns when clicking outside
document.addEventListener('click', function(e){
  if(!e.target.closest('.col-toggle-btn') && !e.target.closest('.col-dropdown') && !e.target.closest('.filter-btn') && !e.target.closest('.filter-dropdown')){
    document.querySelectorAll('.col-dropdown,.filter-dropdown').forEach(function(d){d.classList.remove('open');});
  }
});

/* ── Build table header row (respecting visibleCols) ───────────────── */
function buildTableHeader(viewId, sortState){
  // sortState = {col: idx, dir: 1/-1} or null
  var defs = viewId==='1' ? COLUMN_DEFS_SUMMARY : COLUMN_DEFS_FULL;
  var html = '';
  defs.forEach(function(col, idx){
    if(!visibleCols[viewId].has(col.key)) return;
    var cls = '';
    if(sortState && sortState.col===idx){
      cls = sortState.dir===1 ? 'sort-asc' : 'sort-desc';
    }
    // Index column (#) is not sortable
    if(col.key==='_idx'){
      html += '<th style="width:40px">'+col.label+'</th>';
    } else {
      var onclick = 'onclick="sort' + (viewId==='1'?'Summary':'Full') + '(' + idx + ')"';
      html += '<th class="'+cls+'" '+onclick+'>'+col.label+'<span class="sort-arrow"></span></th>';
    }
  });
  return html;
}

/* ── Tab switching ─────────────────────────────────────────────────────── */
function switchTab(viewId, btn){
  document.querySelectorAll('.tab-content').forEach(el=>el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el=>el.classList.remove('active'));
  document.getElementById(viewId).classList.add('active');
  btn.classList.add('active');
  document.getElementById('btnExport').textContent =
    viewId==='fullScheduleView' ? '\u2193 Export Full Schedule' : '\u2193 Export Excel';
  // Close any open column/filter dropdown
  document.querySelectorAll('.col-dropdown,.filter-dropdown').forEach(function(d){d.classList.remove('open');});
}

/* ═══════════════════════════════════════════════════════════════════
   SUMMARY VIEW
   ═══════════════════════════════════════════════════════════════════ */
let summaryData=[], summarySortCol=-1, summarySortDir=1;

function initSummary(){
  summaryData = TODAY_DATA.vessels;
  buildColDropdown('1');
  updateFilterButton('route', '1');
  updateFilterButton('vessel', '1');
  updateFilterButton('pic', '1');
  renderSummary();
}

function getFilteredSummary(){
  const q=document.getElementById('searchBox').value.toLowerCase();
  const selRoute = getFilterSelected('route', '1');
  const selVessel = getFilterSelected('vessel', '1');
  const selPic = getFilterSelected('pic', '1');
  const delay=document.getElementById('filterDelay').value;
  let data=summaryData.filter(r=>{
    if(selRoute && !selRoute.has(r.route)) return false;
    if(selVessel && !selVessel.has(r.vessel)) return false;
    if(selPic && !selPic.has(r.pic)) return false;
    if(q && !`${r.vessel} ${r.port} ${r.wait} ${r.pic} ${r.code} ${r.remark}`.toLowerCase().includes(q)) return false;
    if(delay==='delay') return r.etaDelay.toLowerCase().includes('delay');
    if(delay==='ahead') return r.etaDelay.toLowerCase().includes('ahead');
    if(delay==='normal') return !r.etaDelay;
    return true;
  });
  if(summarySortCol>=0){
    const defs = COLUMN_DEFS_SUMMARY;
    const colDef = defs[summarySortCol];
    if(colDef && colDef.key!=='_idx' && visibleCols['1'].has(colDef.key)){
      data.sort((a,b)=>((a[colDef.key]||'').localeCompare(b[colDef.key]||''))*summarySortDir);
    }
  }
  return data;
}

function delayTag(v){
  if(!v) return '';
  if(v.toLowerCase().startsWith('ahead')) return '<span class="ahead-tag">'+v+'</span>';
  if(v.toLowerCase().startsWith('delay')) return '<span class="delay-tag">'+v+'</span>';
  return v;
}

function renderSummary(){
  const data=getFilteredSummary();
  // Build header
  document.getElementById('summaryThead').innerHTML =
    buildTableHeader('1', {col: summarySortCol, dir: summarySortDir});

  const tbody=document.getElementById('summaryTbody');
  if(!data.length){tbody.innerHTML='<tr><td colspan="'+visibleCols['1'].size+'" class="no-data">No matching records found.</td></tr>';document.getElementById('statTotal').textContent='0 vessels';document.getElementById('statDelay').textContent='0 delayed';return;}
  const dc=data.filter(r=>r.etaDelay && r.etaDelay.toLowerCase().includes('delay')).length;
  document.getElementById('statTotal').textContent=data.length+' vessels';
  document.getElementById('statDelay').textContent=dc+' delayed';

  const defs = COLUMN_DEFS_SUMMARY;
  tbody.innerHTML = data.map(function(r, rowIdx){
    var cells = defs.map(function(col){
      if(!visibleCols['1'].has(col.key)) return null;
      // Row number column
      if(col.key==='_idx') return '<td class="td-center td-idx">'+(rowIdx+1)+'</td>';
      var v = r[col.key] || '';
      // Special formatting per column
      if(col.key==='route')     return '<td class="td-center"><span class="badge-route">'+v+'</span></td>';
      if(col.key==='vessel')    return '<td><strong>'+v+'</strong></td>';
      if(col.key==='code')      return '<td class="td-center"><span class="badge-code">'+v+'</span></td>';
      if(col.key==='pic')       return '<td>'+v+'</td>';
      if(col.key==='port')      return '<td class="td-center td-mono"><strong>'+v+'</strong></td>';
      if(col.key==='wait')      return '<td class="td-center">'+v+'</td>';
      if(col.key==='proforma')  return '<td class="proforma-cell">'+v+'</td>';
      if(col.key==='voy')       return '<td class="td-center td-mono">'+v+'</td>';
      if(col.key==='ltmEta' || col.key==='ltmEtd') return '<td class="td-center td-mono" style="font-size:11px">'+v+'</td>';
      if(col.key==='eta' || col.key==='etb' || col.key==='etd') return '<td class="td-center td-mono">'+v+'</td>';
      if(col.key==='portStay')  return '<td class="td-center">'+v+'</td>';
      if(col.key==='etaDelay')  return '<td class="td-center">'+delayTag(v)+'</td>';
      if(col.key==='etdDelay')  return '<td class="td-center">'+delayTag(v)+'</td>';
      if(col.key==='remark')    return '<td class="remark-cell">'+v+'</td>';
      return '<td>'+v+'</td>';
    });
    return '<tr>'+cells.filter(c=>c!==null).join('')+'</tr>';
  }).join('');
}

function sortSummary(col){
  const defs = COLUMN_DEFS_SUMMARY;
  if(defs[col].key==='_idx') return;
  if(!visibleCols['1'].has(defs[col].key)) return;
  if(summarySortCol===col){ summarySortDir*=-1; } else { summarySortCol=col; summarySortDir=1; }
  renderSummary();
}

/* ═══════════════════════════════════════════════════════════════════
   FULL SCHEDULE VIEW
   ═══════════════════════════════════════════════════════════════════ */
let fullData=[], fullSortCol=-1, fullSortDir=1;
let vesselGroupMap = {};

function initFullSchedule(){
  fullData = TODAY_DATA.fullSchedule || [];
  let seen = {}, gi = 0;
  fullData.forEach(function(r){
    if(!(r.vessel in seen)){ seen[r.vessel] = gi%2; gi++; }
  });
  vesselGroupMap = seen;

  buildColDropdown('2');
  updateFilterButton('route', '2');
  updateFilterButton('vessel', '2');
  updateFilterButton('pic', '2');
  renderFullSchedule();
}

function getFilteredFull(){
  const q=document.getElementById('searchBox2').value.toLowerCase();
  const selRoute = getFilterSelected('route', '2');
  const selVessel = getFilterSelected('vessel', '2');
  const selPic = getFilterSelected('pic', '2');
  let data=fullData.filter(r=>{
    if(selRoute && !selRoute.has(r.route)) return false;
    if(selVessel && !selVessel.has(r.vessel)) return false;
    if(selPic && !selPic.has(r.pic)) return false;
    if(q && !`${r.vessel} ${r.port} ${r.wait} ${r.manIn} ${r.pic} ${r.code} ${r.voy} ${r.date}`.toLowerCase().includes(q)) return false;
    return true;
  });
  if(fullSortCol>=0){
    const defs = COLUMN_DEFS_FULL;
    const colDef = defs[fullSortCol];
    if(colDef && visibleCols['2'].has(colDef.key)){
      data.sort((a,b)=>((a[colDef.key]||'').localeCompare(b[colDef.key]||''))*fullSortDir);
    }
  }
  return data;
}

function renderFullSchedule(){
  const data=getFilteredFull();
  // Build header
  document.getElementById('fullThead').innerHTML =
    buildTableHeader('2', {col: fullSortCol, dir: fullSortDir});

  const tbody=document.getElementById('fullTbody');
  if(!data.length){tbody.innerHTML='<tr><td colspan="'+visibleCols['2'].size+'" class="no-data">No matching records found.</td></tr>';document.getElementById('statTotal2').textContent='0 rows';return;}
  document.getElementById('statTotal2').textContent=data.length+' rows';

  const defs = COLUMN_DEFS_FULL;
  let prevVessel = '';
  tbody.innerHTML = data.map(function(r, idx){
    var gc = vesselGroupMap[r.vessel]===0 ? 'vessel-group-even' : 'vessel-group-odd';
    var boundaryCls = '';
    if(r.vessel !== prevVessel){
      boundaryCls = ' vessel-group-first';
      if(idx>0) boundaryCls += ' vessel-group-last-prev';
      prevVessel = r.vessel;
    }
    var endCls = '';
    if(idx===data.length-1 || data[idx+1].vessel !== r.vessel){
      endCls = ' vessel-group-last';
    }
    var cls = gc+boundaryCls+endCls;

    var cells = defs.map(function(col){
      if(!visibleCols['2'].has(col.key)) return null;
      var v = r[col.key] || '';
      if(col.key==='route')      return '<td class="td-center"><span class="badge-route">'+v+'</span></td>';
      if(col.key==='vessel')     return '<td><span class="vessel-label">'+v+'</span></td>';
      if(col.key==='code')       return '<td class="td-center"><span class="badge-code">'+v+'</span></td>';
      if(col.key==='pic')        return '<td>'+v+'</td>';
      if(col.key==='port')       return '<td class="td-center td-mono"><strong>'+v+'</strong></td>';
      if(col.key==='manIn' || col.key==='wait' || col.key==='run' || col.key==='fspDistance' || col.key==='speed') return '<td class="td-center">'+v+'</td>';
      if(col.key==='proforma')   return '<td class="proforma-cell">'+v+'</td>';
      if(col.key==='voy')        return '<td class="td-center td-mono">'+v+'</td>';
      if(col.key==='ltmEta' || col.key==='ltmEtd' || col.key==='date') return '<td class="td-center td-mono" style="font-size:11px">'+v+'</td>';
      if(col.key==='eta' || col.key==='etb' || col.key==='etd') return '<td class="td-center td-mono">'+v+'</td>';
      if(col.key==='portStay')   return '<td class="td-center">'+v+'</td>';
      if(col.key==='etaDelay')   return '<td class="td-center">'+delayTag(v)+'</td>';
      if(col.key==='etdDelay')   return '<td class="td-center">'+delayTag(v)+'</td>';
      return '<td>'+v+'</td>';
    });
    return '<tr class="'+cls+'">'+cells.filter(c=>c!==null).join('')+'</tr>';
  }).join('');
}

function sortFull(col){
  const defs = COLUMN_DEFS_FULL;
  if(!visibleCols['2'].has(defs[col].key)) return;
  if(fullSortCol===col){ fullSortDir*=-1; } else { fullSortCol=col; fullSortDir=1; }
  renderFullSchedule();
}

/* ═══════════════════════════════════════════════════════════════════
   EXPORT EXCEL (matches Summary Excel format exactly)
   ═══════════════════════════════════════════════════════════════════ */
function exportExcel(){
  var activeTab = document.querySelector('.tab-btn.active').getAttribute('data-tab');
  _loadXlsx(function(){
    if(activeTab==='fullScheduleView'){
      exportFullScheduleExcel();
    } else {
      exportSummaryExcel();
    }
  });
}

function exportSummaryExcel(){
  var data=getFilteredSummary(), todayStr=TODAY_DATA.date;
  var headers = COLUMN_DEFS_SUMMARY.filter(c=>visibleCols['1'].has(c.key) && c.key!=='_idx').map(c=>c.label);

  function thinBorder(){var s={style:'thin',color:{rgb:'BFBFBF'}};return{top:s,bottom:s,left:s,right:s};}
  var B=thinBorder();
  function F(rgb){return{patternType:'solid',fgColor:{rgb:rgb}};}
  function A(h,v,wrap){var o={horizontal:h,vertical:v};if(wrap)o.wrapText=true;return o;}

  var tS={font:{name:'Arial',bold:true,color:{rgb:'FFFFFF'},sz:14},fill:F('2E75B6'),alignment:A('center','center'),border:B};
  var hS={font:{name:'Arial',bold:true,color:{rgb:'FFFFFF'},sz:10},fill:F('1F4E79'),alignment:A('center','center',true),border:B};

  var sheetData=[];
  var tr=[];
  var numCols = headers.length;
  for(var c=0;c<numCols;c++) tr[c]={v:(c===0?'CUL VESSEL DAILY MOVEMENT SUMMARY  --  As of '+todayStr:''),s:tS};
  sheetData.push(tr);
  sheetData.push(headers.map(function(h){return{v:h,s:hS};}));

  var defs = COLUMN_DEFS_SUMMARY;
  for(var i=0;i<data.length;i++){
    var r=data[i], fc=(i%2===0)?'EBF3FB':'FFFFFF';
    var nS={font:{name:'Arial',sz:9},fill:F(fc),border:B,alignment:A('left','center')};
    var bS={font:{name:'Arial',bold:true,sz:9},fill:F(fc),border:B,alignment:A('left','center')};
    var cS={font:{name:'Arial',sz:9},fill:F(fc),border:B,alignment:A('center','center')};
    var rS={font:{name:'Arial',sz:9},fill:F(fc),border:B,alignment:A('left','center',true)};
    var dS={font:{name:'Arial',bold:true,color:{rgb:'C00000'},sz:9},fill:F(fc),border:B,alignment:A('center','center')};
    function ds(v){if(!v)return cS;return(v.toLowerCase().indexOf('delay')>=0)?dS:cS;}
    var row = [];
    defs.forEach(function(col){
      if(!visibleCols['1'].has(col.key) || col.key==='_idx') return;
      var v = r[col.key]||'';
      if(col.key==='remark') row.push({v:v,s:rS});
      else if(col.key==='vessel') row.push({v:v,s:bS});
      else if(col.key==='etaDelay'||col.key==='etdDelay') row.push({v:v,s:ds(v)});
      else if(['route','port','voy','eta','etb','etd','portStay','proforma','ltmEta','ltmEtd','wait','date','run','fspDistance','speed','manIn'].includes(col.key)) row.push({v:v,s:cS});
      else row.push({v:v,s:nS});
    });
    sheetData.push(row);
  }

  var ws=XLSX.utils.aoa_to_sheet(sheetData);
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:numCols-1}}];
  ws['!cols'] = headers.map(function(h){return{wch:h.length+4};});
  ws['!rows']=[{hpt:28},{hpt:30}];for(var j=0;j<data.length;j++)ws['!rows'].push({hpt:18});
  ws['!autofilter']={ref:'A2:'+XLSX.utils.encode_col(numCols-1)+(data.length+2)};

  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Daily Movement Summary');
  XLSX.writeFile(wb,'CUL Daily Movement Summary '+todayStr+'.xlsx');
}

function exportFullScheduleExcel(){
  var data=getFilteredFull(), todayStr=TODAY_DATA.date;
  var headers = COLUMN_DEFS_FULL.filter(c=>visibleCols['2'].has(c.key)).map(c=>c.label);

  function thinBorder(){var s={style:'thin',color:{rgb:'BFBFBF'}};return{top:s,bottom:s,left:s,right:s};}
  var B=thinBorder();
  function F(rgb){return{patternType:'solid',fgColor:{rgb:rgb}};}
  function A(h,v,wrap){var o={horizontal:h,vertical:v};if(wrap)o.wrapText=true;return o;}

  var tS={font:{name:'Arial',bold:true,color:{rgb:'FFFFFF'},sz:14},fill:F('2E75B6'),alignment:A('center','center'),border:B};
  var hS={font:{name:'Arial',bold:true,color:{rgb:'FFFFFF'},sz:10},fill:F('1F4E79'),alignment:A('center','center',true),border:B};

  var sheetData=[];
  var tr=[];
  var numCols = headers.length;
  for(var c=0;c<numCols;c++) tr[c]={v:(c===0?'CUL VESSEL FULL SCHEDULE  --  As of '+todayStr:''),s:tS};
  sheetData.push(tr);
  sheetData.push(headers.map(function(h){return{v:h,s:hS};}));

  var defs = COLUMN_DEFS_FULL;
  var prevVessel='', rowGroup=0;
  for(var i=0;i<data.length;i++){
    var r=data[i];
    if(r.vessel!==prevVessel){rowGroup=(rowGroup+1)%2;prevVessel=r.vessel;}
    var fc=rowGroup===0?'EBF3FB':'FFFFFF';
    var nS={font:{name:'Arial',sz:9},fill:F(fc),border:B,alignment:A('left','center')};
    var bS={font:{name:'Arial',bold:true,sz:9},fill:F(fc),border:B,alignment:A('left','center')};
    var cS={font:{name:'Arial',sz:9},fill:F(fc),border:B,alignment:A('center','center')};
    var dS={font:{name:'Arial',bold:true,color:{rgb:'C00000'},sz:9},fill:F(fc),border:B,alignment:A('center','center')};
    function ds(v){if(!v)return cS;return(v.toLowerCase().indexOf('delay')>=0)?dS:cS;}
    var row = [];
    defs.forEach(function(col){
      if(!visibleCols['2'].has(col.key)) return;
      var v = r[col.key]||'';
      if(col.key==='vessel') row.push({v:v,s:bS});
      else if(col.key==='etaDelay'||col.key==='etdDelay') row.push({v:v,s:ds(v)});
      else if(['route','port','voy','eta','etb','etd','portStay','proforma','ltmEta','ltmEtd','wait','date','run','fspDistance','speed','manIn'].includes(col.key)) row.push({v:v,s:cS});
      else row.push({v:v,s:nS});
    });
    sheetData.push(row);
  }

  var ws=XLSX.utils.aoa_to_sheet(sheetData);
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:numCols-1}}];
  ws['!cols'] = headers.map(function(h){return{wch:Math.max(h.length+4, 10)};});
  ws['!rows']=[{hpt:28},{hpt:30}];for(var j=0;j<data.length;j++)ws['!rows'].push({hpt:16});
  ws['!autofilter']={ref:'A2:'+XLSX.utils.encode_col(numCols-1)+(data.length+2)};

  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Full Schedule');
  XLSX.writeFile(wb,'CUL Daily Movement Full Schedule '+todayStr+'.xlsx');
}

/* ═══════════════════════════════════════════════════════════════════
   HISTORY MODAL
   ═══════════════════════════════════════════════════════════════════ */
function openHistory(){
  const dates=Object.keys(SNAPSHOTS).sort().reverse();
  const listEl=document.getElementById('historyDateList');
  listEl.innerHTML='';
  dates.forEach((d,i)=>{
    const btn=document.createElement('button');
    btn.className='history-date-btn'+(i===0?' active':'');
    btn.textContent=d;
    btn.onclick=function(){document.querySelectorAll('.history-date-btn').forEach(b=>b.classList.remove('active'));btn.classList.add('active');renderHistoryTable(d);};
    listEl.appendChild(btn);
  });
  if(dates.length) renderHistoryTable(dates[0]);
  document.getElementById('historyModal').classList.add('open');
}
function renderHistoryTable(date){
  const vessels=SNAPSHOTS[date]||[];
  if(!vessels.length){document.getElementById('historyTableWrap').innerHTML='<p style="color:#8a9bb0;padding:20px">No data.</p>';return;}
  const headers=['Route','Vessel','Code','Port','Wait','Voy','ETA','ETB','ETD','Port Stay','ETA Delay','ETD Delay','PIC','Remark'];
  const keys=['route','vessel','code','port','wait','voy','eta','etb','etd','portStay','etaDelay','etdDelay','pic','remark'];
  document.getElementById('historyTableWrap').innerHTML='<table><thead><tr>'+headers.map(h=>'<th style="font-size:11px;padding:7px 8px">'+h+'</th>').join('')+'</tr></thead><tbody>'+vessels.map(r=>'<tr>'+keys.map(k=>'<td style="padding:6px 8px;border-bottom:1px solid #eee;font-size:11px">'+(r[k]||'')+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
}
function closeHistory(){document.getElementById('historyModal').classList.remove('open');}
document.getElementById('historyModal').addEventListener('click',function(e){if(e.target===document.getElementById('historyModal'))closeHistory();});

/* ═══════════════════════════════════════════════════════════════════
   INIT
   ═══════════════════════════════════════════════════════════════════ */
function init(){
  loadSnapshots(); saveSnapshot(TODAY_DATA);
  document.getElementById('headerDate').textContent='As of '+TODAY_DATA.date+'  |  '+TODAY_DATA.vessels.length+' vessels  |  '+TODAY_DATA.fullSchedule.length+' schedule rows';
  document.getElementById('footerTs').textContent='Generated: '+TODAY_DATA.date;
  initSummary();
  initFullSchedule();
}
init();
</script>
</body>
</html>"""

# ── Main ──────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--excel', default=DEFAULT_EXCEL)
    parser.add_argument('--out',   default=DEFAULT_HTML)
    args = parser.parse_args()

    excel_path = args.excel
    out_path   = args.out

    if not os.path.exists(os.path.dirname(excel_path)):
        excel_path = os.path.join(SCRIPT_DIR, 'CUL DAILY MOVEMENT.xlsx')
    if not os.path.exists(os.path.dirname(out_path)):
        out_path = os.path.join(SCRIPT_DIR, 'cul_daily_movement.html')

    print(f'Reading: {excel_path}')
    data = extract(excel_path)
    print(f'  -> {len(data["vessels"])} vessels (summary)')
    print(f'  -> {len(data["fullSchedule"])} schedule rows (full)')
    print(f'  -> date={data["date"]}')

    # Build column defs JSON for JS
    col_defs_summary = [{"key":c[0],"label":c[1],"defaultVisible":c[2]} for c in SUMMARY_COLUMNS]
    col_defs_full    = [{"key":c[0],"label":c[1],"defaultVisible":c[3]} for c in FULL_COLUMNS]

    html = HTML_TEMPLATE
    html = html.replace('__TODAY_DATA__',       json.dumps(data, ensure_ascii=False))
    html = html.replace('__COLUMN_DEFS_SUMMARY__', json.dumps(col_defs_summary, ensure_ascii=False))
    html = html.replace('__COLUMN_DEFS_FULL__',    json.dumps(col_defs_full,    ensure_ascii=False))

    os.makedirs(os.path.dirname(out_path) if os.path.dirname(out_path) else '.', exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Saved : {out_path}')

if __name__ == '__main__':
    main()
