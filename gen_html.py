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
from datetime import datetime, date, timedelta

# ── Defaults ──────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 数据更新目录: 本机(leahliu)=P:, 另一台(culadmin)=Z:。自动探测, 两机通用, 无需传参。
_BASES = [
    r"Z:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新",
    r"P:\04 上海操作中心\01 船期管理科\船期管理\VSL Daily Movement\更新",
]
UPD_DIR = next((b for b in _BASES if os.path.isdir(b)), _BASES[0])
from build_fleet_movement import GEN_DIR
DEFAULT_EXCEL = os.path.join(GEN_DIR, "CUL DAILY MOVEMENT.rebuilt.xlsx")
DEFAULT_HTML  = os.path.join(SCRIPT_DIR, "cul_daily_movement.html")

# ── Column definitions ─────────────────────────────────────────────────────
# Full column list (all columns in the Excel schedule rows)
# Used for Full Schedule view; Summary view uses a subset.
FULL_COLUMNS = [
    ('route',       'Route',            True,   True),
    ('vessel',      'Vessel',           True,   True),
    ('code',        'Code',             True,   True),
    ('port',        'Port',             True,   True),
    ('manIn',       'Man In',           True,   False),  # C2 - default hidden
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
    ('fspDistance', 'FSP Distance',     True,   False),  # C14 - default hidden
    ('speed',       'Speed',            True,   True),   # C15
    ('etaDelay',    'ETA Delay',        True,   True),   # C16
    ('etdDelay',    'ETD Delay',        True,   False),  # C17
    ('remark',      'Remark',           True,   False),
    ('pic',         'PIC',              True,   True),
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
    ('remark',     'Remark',           False),
    ('pic',        'PIC',              True),
]

# ── Helpers ────────────────────────────────────────────────────────────────
def fmt_dt(v):
    if v is None: return ''
    if isinstance(v, datetime):
        return v.strftime('%m/%d %H:%M')   # ETA/ETB/ETD 始终显示时间(含00:00)
    return str(v).strip()

def get_str(v):
    if v is None: return ''
    if isinstance(v, datetime): return v.strftime('%m/%d %H:%M')
    return str(v).strip()

# ── Extract data ──────────────────────────────────────────────────────────
def extract(excel_path):
    today = date.today()
    data_eta_min = None   # capture ALL rows from Excel (no date window)
    data_eta_max = None
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
            remarks_by_row = {}  # row_number -> remark for per-row association
            j = i + 2
            while j <= rows_total:
                c1_j = ws_src.cell(j, 1).value
                if c1_j and isinstance(c1_j, str) and c1_j.strip().startswith('Remark'):
                    remark_text = c1_j.strip().replace('Remark:', '').replace('Remark :', '').strip()
                    if remark_text:
                        # Parse remark to find target row: first two tokens = voyage + port
                        parts = remark_text.split()
                        if len(parts) >= 2 and schedule_rows:
                            target_voy, target_port = parts[0], parts[1]
                            matched = False
                            for sr in schedule_rows:
                                sr_voy = get_str(ws_src.cell(sr, 7).value)
                                sr_port = get_str(ws_src.cell(sr, 1).value)
                                if sr_voy == target_voy and sr_port == target_port:
                                    if sr in remarks_by_row:
                                        remarks_by_row[sr] = remarks_by_row[sr] + '; ' + remark_text
                                    else:
                                        remarks_by_row[sr] = remark_text
                                    matched = True
                                    break
                            if not matched:
                                # Fallback: assign to last schedule row
                                remarks_by_row[schedule_rows[-1]] = remark_text
                        elif schedule_rows:
                            remarks_by_row[schedule_rows[-1]] = remark_text
                        else:
                            remark = remark_text  # fallback: no schedule rows yet
                    j += 1
                    continue
                c16_j = ws_src.cell(j, 16).value
                if c16_j and isinstance(c16_j, str) and 'PIC' in c16_j:
                    i = j
                    break
                eta_val = ws_src.cell(j, 9).value
                if isinstance(eta_val, datetime):
                    eta_d = eta_val.date()
                    if data_eta_min is None or eta_d < data_eta_min:
                        data_eta_min = eta_d
                    if data_eta_max is None or eta_d > data_eta_max:
                        data_eta_max = eta_d
                    schedule_rows.append(j)
                j += 1
            else:
                i = rows_total + 1

            vessel_blocks.append({'route': route, 'vessel_full': vessel_full,
                                   'vessel_code': vessel_code, 'pic': pic,
                                   'schedule_rows': schedule_rows,
                                   'remarks_by_row': remarks_by_row})
        else:
            i += 1

    # ── Summary: nearest ETB per vessel ──
    results = []
    summary_row_set = set()
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
            summary_row_set.add(best_row)
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
                'etaRaw':      ws_src.cell(r, 9).value.strftime('%Y-%m-%d') if isinstance(ws_src.cell(r, 9).value, datetime) else '',
                'etbRaw':      ws_src.cell(r, 10).value.strftime('%Y-%m-%d') if isinstance(ws_src.cell(r, 10).value, datetime) else '',
                'etb':         fmt_dt(ws_src.cell(r, 10).value),
                'etd':         fmt_dt(ws_src.cell(r, 11).value),
                'run':         get_str(ws_src.cell(r, 12).value),
                'portStay':    get_str(ws_src.cell(r, 13).value),
                'fspDistance': get_str(ws_src.cell(r, 14).value),
                'speed':       get_str(ws_src.cell(r, 15).value),
                'etaDelay':    get_str(ws_src.cell(r, 16).value),
                'etdDelay':    get_str(ws_src.cell(r, 17).value),
                'remark':      vb['remarks_by_row'].get(r, ''),
            }
        else:
            rec = {'route': vb['route'], 'vessel': vb['vessel_full'],
                   'code': vb['vessel_code'], 'pic': vb['pic'],
                   'port':'','manIn':'','wait':'','proforma':'','voy':'','ltmEta':'','ltmEtd':'',
                   'date':'','eta':'','etaRaw':'','etb':'','etbRaw':'','etd':'','run':'',
                   'portStay':'','fspDistance':'','speed':'','etaDelay':'','etdDelay':'','remark': ''}
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
                'etaRaw':      ws_src.cell(r, 9).value.strftime('%Y-%m-%d') if isinstance(ws_src.cell(r, 9).value, datetime) else '',
                'etbRaw':      ws_src.cell(r, 10).value.strftime('%Y-%m-%d') if isinstance(ws_src.cell(r, 10).value, datetime) else '',
                'etb':         fmt_dt(ws_src.cell(r, 10).value),
                'etd':         fmt_dt(ws_src.cell(r, 11).value),
                'run':         get_str(ws_src.cell(r, 12).value),
                'portStay':    get_str(ws_src.cell(r, 13).value),
                'fspDistance': get_str(ws_src.cell(r, 14).value),
                'speed':       get_str(ws_src.cell(r, 15).value),
                'etaDelay':    get_str(ws_src.cell(r, 16).value),
                'etdDelay':    get_str(ws_src.cell(r, 17).value),
                'remark':      vb['remarks_by_row'].get(r, ''),
                'isSummary':   r in summary_row_set,
            })

    return {
        'date': today.strftime('%Y-%m-%d'),
        'generatedAt': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'vessels': results,
        'fullSchedule': full_schedule,
        'defaultEtaFrom': (today - timedelta(days=14)).strftime('%Y-%m-%d'),
        'defaultEtaTo':   (today + timedelta(days=30)).strftime('%Y-%m-%d'),
        'defaultEtbFrom': (today - timedelta(days=14)).strftime('%Y-%m-%d'),
        'defaultEtbTo':   (today + timedelta(days=30)).strftime('%Y-%m-%d'),
        'dataEtaMin':     data_eta_min.strftime('%Y-%m-%d') if data_eta_min else '',
        'dataEtaMax':     data_eta_max.strftime('%Y-%m-%d') if data_eta_max else '',
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
  .controls input[type="date"] {
    width: 138px; padding: 6px 10px; font-size: 12px;
    color: #1F4E79; font-weight: 500; font-family: inherit;
  }
  .eta-label { font-size: 12px; color: #5a6e82; font-weight: 600; margin-right: -6px; }
  .date-type-select { padding: 6px 8px; border: 1px solid #c8d6e5; border-radius: 6px; background: #fff; color: #333; font-size: 12px; font-weight: 600; cursor: pointer; }
  .eta-sep { color: #a8b8c8; font-weight: 400; margin: 0 -2px; }
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
  td.delay { background: #fff0f0 !important; color: #c00000; font-weight: 700; }
  td.ontime { color: #27ae60; font-weight: 600; }
  tr.port-row:hover { filter: brightness(0.96); }
  tr.detail-wrap td { border-bottom: 2px solid #d4e0eb; }
  tr.detail-row:hover { background: #f5f8fb !important; }

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
  .summary-highlight td { background: #fff3cd !important; }
  .summary-highlight:hover td { background: #ffe69c !important; }
  .vessel-group-first td { border-top: 3px solid #2E75B6; }
  .vessel-group-last  td { border-bottom: 2px solid #c3d9f0; }

  .td-center { text-align: center; }
  .td-idx { text-align: center; color: #8fa3b8; font-size: 11px; font-weight: 500; width: 40px; }
  .td-mono { font-family: 'Consolas', 'Courier New', monospace; font-size: 12px; }
  .badge-route { display: inline-block; padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; background: #1F4E79; color: #fff; letter-spacing: .5px; }
  .badge-code { display: inline-block; padding: 2px 7px; border-radius: 4px; font-size: 11px; font-weight: 600; background: #E8F4FD; color: #1F4E79; border: 1px solid #a8cfe8; }
  .weekday-tag { display: inline-block; padding: 2px 7px; border-radius: 3px; font-size: 11px; font-weight: 600; background: #1F4E79; color: #fff; }
  .delay-tag { display: inline-block; padding: 2px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; background: #fff0f0; color: #c00000; border: 1px solid #f5c6c6; }
  .ahead-tag { display: inline-block; padding: 2px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; background: #f0fff4; color: #1a7340; border: 1px solid #b7dfca; }
  .no-data { text-align: center; padding: 40px; color: #8a9bb0; font-size: 14px; }
  .remark-cell { max-width: 260px; white-space: normal; line-height: 1.4; color: #5a6e82; font-size: 11.5px; }
  .remark-note { display: block; max-width: 200px; white-space: normal; line-height: 1.3; color: #e67e22; font-size: 10.5px; font-weight: normal; margin-top: 2px; }
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
  <button class="tab-btn" data-tab="portView" onclick="switchTab('portView',this)">&#9889; Port Wait</button>
  <button class="tab-btn" data-tab="speedView" onclick="switchTab('speedView',this)">&#128168; Speed</button>
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
    <select id="dateFilterType2" class="date-type-select" onchange="onDateFilterTypeChange()">
      <option value="eta" selected>ETA</option>
      <option value="etb">ETB</option>
    </select>
    <input type="date" id="etaFrom2" title="ETA from" onchange="renderFullSchedule()">
    <span class="eta-sep">–</span>
    <input type="date" id="etaTo2" title="ETA to" onchange="renderFullSchedule()">
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

<!-- ═══════════════════════════════════════════════════════════════════
     VIEW 3: Port Wait Analytics
     ═══════════════════════════════════════════════════════════════════ -->
<div id="portView" class="tab-content">
  <div class="controls">
    <span style="font-weight:600;font-size:14px;margin-right:8px;">&#128205; Port Filter:</span>
    <div style="position:relative;">
      <button class="filter-btn" id="portFilterBtn" onclick="togglePortFilter()">All Ports</button>
      <div class="filter-dropdown col-dropdown" id="portFilterDropdown"></div>
    </div>
    <span style="font-weight:600;font-size:14px;margin:0 8px;">&#128205; Remark Filter:</span>
    <div style="position:relative;">
      <button class="filter-btn" id="remarkFilterBtn" onclick="toggleRemarkFilter()">All Remarks</button>
      <div class="filter-dropdown col-dropdown" id="remarkFilterDropdown"></div>
    </div>
    <span class="stat-chip" id="statPortWait" style="margin-left:12px;">&#8212; port entries</span>
  </div>

  <!-- Port Wait Analysis -->
  <h3 style="margin:16px 0 8px;color:#1F4E79;">&#9889; Port Wait Time Analysis</h3>
  <p style="font-size:11px;color:#8a9bb0;margin:0 0 8px;">Ports normalized (terminal suffixes merged). Bunkering-only calls excluded. Berth Rate (到靠率) = % of calls with wait &lt; 6 hours. Ranked best&#8594;worst by default.</p>
  <div class="table-wrap">
    <table id="portWaitTable">
      <thead><tr id="portWaitThead"></tr></thead>
      <tbody id="portWaitTbody"></tbody>
    </table>
  </div>

  <!-- Remark Category Wait Breakdown -->
  <div class="table-wrap" style="max-width:900px;">
    <h4 style="margin:6px 0 10px;color:#1F4E79;font-size:13px;">&#128202; Wait Time by Remark Category</h4>
    <div id="remarkSummary" style="display:flex;flex-direction:column;gap:6px;"></div>
  </div>
</div>

<!-- ═══════════════════════════════════════════════════════════════════
     VIEW 4: Speed Analytics
     ═══════════════════════════════════════════════════════════════════ -->
<div id="speedView" class="tab-content">
  <div class="controls">
    <span style="font-weight:600;font-size:14px;margin-right:8px;">&#128205; Port Filter:</span>
    <div style="position:relative;">
      <button class="filter-btn" id="speedPortFilterBtn" onclick="toggleSpeedPortFilter()">All Ports</button>
      <div class="filter-dropdown col-dropdown" id="speedPortFilterDropdown"></div>
    </div>
    <span class="stat-chip" id="statSpeed" style="margin-left:12px;">&#8212; vessels</span>
    <button class="filter-btn" style="margin-left:12px;" onclick="exportSpeedExcel()">&#8595; Export Speed</button>
  </div>

  <!-- Vessel Speed Analysis -->
  <h3 style="margin:16px 0 8px;color:#1F4E79;">&#128168; Vessel Speed Analysis</h3>
  <p style="font-size:11px;color:#8a9bb0;margin:0 0 8px;">Speed values from source Excel column "SPEED" (calculated as leg distance &#247; sailing time). Min/Max = extreme values per vessel; values &#8804;0 or >20kn excluded as dirty data. Hover values to see raw source.</p>
  <div class="table-wrap">
    <table id="speedTable">
      <thead><tr id="speedThead"></tr></thead>
      <tbody id="speedTbody"></tbody>
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

// Default route display order (user-specified 2026-08-05). Unknown routes sort to the end.
// Expanded from combined tokens: NP2-REX -> NP2,REX | RES-CGX -> RES,CGX | CGS-AEM-IMR -> CGS,AEM,IMR
var ROUTE_ORDER = ['ST3','NSCT1','HDT','NSX','CST','CCT','NP2','REX','RTS','SGX','RES','CGX','HLX','CGS','AEM','IMR','NAX','JPS','SJA'];
function routeOrderKey(r){ var i = ROUTE_ORDER.indexOf(r); return i<0 ? 9999 : i; }

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
  var values = [...new Set(dataArr.map(function(r){return r[type];}))];
  if(type==='route'){
    values.sort(function(a,b){return routeOrderKey(a)-routeOrderKey(b);});
  } else {
    values.sort();
  }

  // Select All / Clear All actions
  var actions = document.createElement('div');
  actions.className = 'filter-actions';
  var selAll = document.createElement('button');
  selAll.textContent = 'Select All';
  selAll.onclick = function(e){ e.stopPropagation(); setAllFilterOptions(type, viewId); buildFilterDropdown(type, viewId); updateFilterButton(type, viewId); rerender(viewId); };
  var clrAll = document.createElement('button');
  clrAll.textContent = 'None';
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
  if(viewId==='fullScheduleView'){
    document.getElementById('btnExport').textContent='\u2193 Export Full Schedule';
    document.getElementById('btnExport').style.display='';
  } else if(viewId==='summaryView'){
    document.getElementById('btnExport').textContent='\u2193 Export Excel';
    document.getElementById('btnExport').style.display='';
  } else {
    document.getElementById('btnExport').style.display='none';
  }
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
      if(colDef.key==='route'){
        data.sort((a,b)=>(routeOrderKey(a.route)-routeOrderKey(b.route))*summarySortDir);
      } else {
        data.sort((a,b)=>((a[colDef.key]||'').localeCompare(b[colDef.key]||''))*summarySortDir);
      }
    }
  } else {
    data.sort((a,b)=>routeOrderKey(a.route)-routeOrderKey(b.route));
  }
  return data;
}

function delayTag(v){
  if(!v) return '';
  if(v.toLowerCase().startsWith('ahead')) return '<span class="ahead-tag">'+v+'</span>';
  if(v.toLowerCase().startsWith('delay')) return '<span class="delay-tag">'+v+'</span>';
  return v;
}

function dateWithWeekday(dateStr){
  if(dateStr == null || dateStr === '') return '';
  var s = String(dateStr).trim();
  // Excel serial number (e.g. 46210.2291666667)
  var num = parseFloat(s);
  if(!isNaN(num) && num > 40000 && num < 100000) {
    var d = new Date((num - 25569) * 86400000);
    var days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    return '<span class="weekday-tag">' + days[d.getUTCDay()] + '</span>';
  }
  // Format: "07/14 07:00"
  var m = s.match(/(\d{1,2})\/(\d{1,2})\s+(\d{2}:\d{2})/);
  if(m) {
    var d = new Date(new Date().getFullYear(), parseInt(m[1])-1, parseInt(m[2]), parseInt(m[3].split(':')[0]), parseInt(m[3].split(':')[1]));
    var days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    return '<span class="weekday-tag">' + days[d.getDay()] + '</span>';
  }
  // Already weekday name: "Mon", "Thu", "TUE0800"
  var wd = s.match(/^(MON|TUE|WED|THU|FRI|SAT|SUN)/i);
  if(wd) return '<span class="weekday-tag">' + wd[1].toUpperCase() + '</span>';
  return '';
}

function renderSummary(){
  const data=getFilteredSummary();
  // Build header
  document.getElementById('summaryThead').innerHTML =
    buildTableHeader('1', {col: summarySortCol, dir: summarySortDir});

  const tbody=document.getElementById('summaryTbody');
  if(!data.length){tbody.innerHTML='<tr><td colspan="'+visibleCols['1'].size+'" class="no-data">No matching records found.</td></tr>';document.getElementById('statTotal').textContent='0 vessels';document.getElementById('statDelay').textContent='0 delayed';return;}
  const dc=data.filter(r=>r.etaDelay && r.etaDelay.toLowerCase().includes('delay'));
  document.getElementById('statTotal').textContent=new Set(data.map(r=>r.vessel)).size+' vessels';
  document.getElementById('statDelay').textContent=new Set(dc.map(r=>r.vessel)).size+' delayed';

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

  // Set default ETA date range
  document.getElementById('etaFrom2').value = TODAY_DATA.defaultEtaFrom || '';
  document.getElementById('etaTo2').value = TODAY_DATA.defaultEtaTo || '';
  document.getElementById('dateFilterType2').value = 'eta';

  buildColDropdown('2');
  updateFilterButton('route', '2');
  updateFilterButton('vessel', '2');
  updateFilterButton('pic', '2');
  renderFullSchedule();
}

function onDateFilterTypeChange(){
  const mode = document.getElementById('dateFilterType2').value;
  const key = mode === 'etb' ? 'defaultEtbFrom' : 'defaultEtaFrom';
  const keyTo = mode === 'etb' ? 'defaultEtbTo' : 'defaultEtaTo';
  document.getElementById('etaFrom2').value = TODAY_DATA[key] || '';
  document.getElementById('etaTo2').value = TODAY_DATA[keyTo] || '';
  renderFullSchedule();
}

function getFilteredFull(){
  const q=document.getElementById('searchBox2').value.toLowerCase();
  const selRoute = getFilterSelected('route', '2');
  const selVessel = getFilterSelected('vessel', '2');
  const selPic = getFilterSelected('pic', '2');
  const etaFrom = document.getElementById('etaFrom2').value;
  const etaTo   = document.getElementById('etaTo2').value;
  const dateMode = document.getElementById('dateFilterType2').value;
  const dateKey = dateMode === 'etb' ? 'etbRaw' : 'etaRaw';
  let data=fullData.filter(r=>{
    if(selRoute && !selRoute.has(r.route)) return false;
    if(selVessel && !selVessel.has(r.vessel)) return false;
    if(selPic && !selPic.has(r.pic)) return false;
    if(etaFrom && r[dateKey] < etaFrom) return false;
    if(etaTo   && r[dateKey] > etaTo)   return false;
    if(q && !`${r.route} ${r.vessel} ${r.port} ${r.wait} ${r.manIn} ${r.pic} ${r.code} ${r.voy} ${r.date} ${r.remark}`.toLowerCase().includes(q)) return false;
    return true;
  });
  if(fullSortCol>=0){
    const defs = COLUMN_DEFS_FULL;
    const colDef = defs[fullSortCol];
    if(colDef && visibleCols['2'].has(colDef.key)){
      if(colDef.key==='route'){
        data.sort((a,b)=>(routeOrderKey(a.route)-routeOrderKey(b.route))*fullSortDir);
      } else {
        data.sort((a,b)=>((a[colDef.key]||'').localeCompare(b[colDef.key]||''))*fullSortDir);
      }
    }
  } else {
    data.sort((a,b)=>routeOrderKey(a.route)-routeOrderKey(b.route));
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
    if(r.isSummary) cls += ' summary-highlight';

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
      if(col.key==='ltmEta' || col.key==='ltmEtd') return '<td class="td-center td-mono" style="font-size:11px">'+v+'</td>';
      if(col.key==='date') return '<td class="td-center td-mono" style="font-size:11px">'+dateWithWeekday(r.eta)+'</td>';
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
  var sepS={font:{name:'Arial',sz:6},fill:F('D9D9D9'),border:B};

  // Sort by route (custom order) then vessel for grouping (copy to avoid affecting UI sort state)
  var exportData = data.slice().sort(function(a,b){
    var dr = routeOrderKey(a.route)-routeOrderKey(b.route);
    if(dr!==0) return dr;
    return (a.vessel||'').localeCompare(b.vessel||'');
  });

  var sheetData=[];
  var tr=[];
  var numCols = headers.length;
  for(var c=0;c<numCols;c++) tr[c]={v:(c===0?'CUL VESSEL FULL SCHEDULE  --  As of '+todayStr:''),s:tS};
  sheetData.push(tr);
  sheetData.push(headers.map(function(h){return{v:h,s:hS};}));

  var defs = COLUMN_DEFS_FULL;
  var prevVessel='';
  for(var i=0;i<exportData.length;i++){
    var r=exportData[i];
    if(r.vessel!==prevVessel){
      // Insert separator row before new vessel group (skip for first vessel)
      if(prevVessel!==''){
        var sepRow = [];
        for(var c=0;c<numCols;c++) sepRow[c]={v:'',s:sepS};
        sheetData.push(sepRow);
      }
      prevVessel=r.vessel;
    }
    var fc='FFFFFF';
    if(r.isSummary) fc='FFF3CD';
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

  var totalRows = sheetData.length;
  var ws=XLSX.utils.aoa_to_sheet(sheetData);
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:numCols-1}}];
  ws['!cols'] = headers.map(function(h){return{wch:Math.max(h.length+4, 10)};});
  ws['!rows']=[{hpt:28},{hpt:30}];for(var j=2;j<totalRows;j++)ws['!rows'].push({hpt:16});
  ws['!autofilter']={ref:'A2:'+XLSX.utils.encode_col(numCols-1)+(totalRows-1)};

  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Full Schedule');
  XLSX.writeFile(wb,'CUL Daily Movement Full Schedule '+todayStr+'.xlsx');
}

/* ═══════════════════════════════════════════════════════════════════
   PORT & SPEED ANALYTICS
   ═══════════════════════════════════════════════════════════════════ */

function escapeHtml(s){
  if(!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// Remark classification: keyword -> category mapping
var REMARK_CATEGORIES = [
  {key:'congestion',  label:'Port Congestion / 塞港',    keywords:['congestion','塞港','congestion delay','congested','拥堵','拥挤','congesiton']},
  {key:'weather',     label:'Weather / 天气',            keywords:['typhoon','台风','避台','大风浪','weather','storm','swell','fog','雾','monsoon']},
  {key:'bunker',      label:'Bunker / 加油',             keywords:['bunker','加油','bunkering','fuel','LSFO','MT','BUNKER']},
  {key:'phase',       label:'Phase In/Out / 航线调整',   keywords:['phase in','phase out','slide','eco speed','rotation','P/O','P/I','omit','OMIT','改靠','shifted','suspension']},
  {key:'msa',         label:'MSA / 海事监管',            keywords:['msa','regulatory','MSA']},
  {key:'adhoc',       label:'Ad Hoc Call / 临时挂靠',    keywords:['ad hoc','adhoc','extra call','add call','private call','ADD CALL']},
  {key:'cargo',       label:'Trade / Cargo Balance / 备货配货',  keywords:['balance','load balance','connection','trade','备货','等货','wait cargo']},
  {key:'other',       label:'Other / 其他',              keywords:[]}  // fallback
];

function classifyRemark(remark){
  if(!remark) return null;
  var r=remark.toLowerCase();
  for(var i=0;i<REMARK_CATEGORIES.length-1;i++){
    var kw=REMARK_CATEGORIES[i].keywords;
    for(var j=0;j<kw.length;j++){
      if(r.indexOf(kw[j])>=0) return REMARK_CATEGORIES[i].key;
    }
  }
  return 'other';
}

// Selected remark categories filter (null = show all)
var selRemarkCats = null;

function buildRemarkFilterDropdown(){
  var dd=document.getElementById('remarkFilterDropdown');
  var usedCats={};
  TODAY_DATA.fullSchedule.forEach(function(sr){
    var cat=classifyRemark(sr.remark);
    if(cat) usedCats[cat]=true;
  });

  var html='';
  REMARK_CATEGORIES.forEach(function(cat){
    if(!usedCats[cat.key] && cat.key!=='other') return;
    var checked = !selRemarkCats || selRemarkCats.indexOf(cat.key)>=0;
    html+='<label style="display:flex;align-items:center;gap:6px;padding:4px 8px;cursor:pointer;white-space:nowrap;">';
    html+='<input type="checkbox" value="'+cat.key+'" '+(checked?'checked':'')+' onchange="onRemarkCatChange()">';
    html+=cat.label+'</label>';
  });
  dd.innerHTML=html;
}

function toggleRemarkFilter(){
  var dd=document.getElementById('remarkFilterDropdown');
  if(!dd.classList.contains('open')){ buildRemarkFilterDropdown(); dd.classList.add('open'); }
  else dd.classList.remove('open');
  document.querySelectorAll('.filter-dropdown').forEach(function(d){if(d!==dd) d.classList.remove('open');});
}

function onRemarkCatChange(){
  var checks=document.querySelectorAll('#remarkFilterDropdown input[type=checkbox]');
  var sel=[];
  checks.forEach(function(cb){if(cb.checked) sel.push(cb.value);});
  selRemarkCats = sel.length===REMARK_CATEGORIES.length ? null : sel;
  // Update button text
  var btn=document.getElementById('remarkFilterBtn');
  if(selRemarkCats===null){
    btn.textContent='All Remarks';
    btn.style.background='';
  } else {
    btn.textContent=sel.length+'/'+REMARK_CATEGORIES.length+' categories';
    btn.style.background='#fff3e0';
    btn.style.borderColor='#e67e22';
  }
  renderPortWaitTable();
  renderRemarkSummary();
}
var selPortFilter = null;  // null = show all ports

function buildPortFilterDropdown(){
  var dd=document.getElementById('portFilterDropdown');
  var allPorts={};
  TODAY_DATA.fullSchedule.forEach(function(sr){
    var p=normalizePort(sr.port);
    if(p && !isBunkeringPort(sr.port)) allPorts[p]=true;
  });
  var sorted=Object.keys(allPorts).sort();
  var html='<label style="display:flex;align-items:center;gap:6px;padding:4px 8px;cursor:pointer;white-space:nowrap;">';
  html+='<input type="checkbox" value="__all__" '+(selPortFilter===null?'checked':'')+' onchange="onPortFilterChange()">';
  html+='<b>All Ports</b></label>';
  sorted.forEach(function(p){
    var checked = selPortFilter===null || selPortFilter.indexOf(p)>=0;
    html+='<label style="display:flex;align-items:center;gap:6px;padding:4px 8px;cursor:pointer;white-space:nowrap;">';
    html+='<input type="checkbox" value="'+p+'" '+(checked?'checked':'')+' onchange="onPortFilterChange()">';
    html+=p+'</label>';
  });
  dd.innerHTML=html;
}

function togglePortFilter(){
  var dd=document.getElementById('portFilterDropdown');
  if(!dd.classList.contains('open')){ buildPortFilterDropdown(); dd.classList.add('open'); }
  else dd.classList.remove('open');
  document.querySelectorAll('.filter-dropdown').forEach(function(d){if(d!==dd) d.classList.remove('open');});
}

function onPortFilterChange(){
  var allCb=document.querySelector('#portFilterDropdown input[value="__all__"]');
  var checks=document.querySelectorAll('#portFilterDropdown input[type=checkbox]:not([value="__all__"])');
  var sel=[];
  var total=0;
  checks.forEach(function(cb){total++; if(cb.checked) sel.push(cb.value);});

  // "All Ports" checked = show all; unchecked = use individual selections
  if(allCb && allCb.checked){
    selPortFilter=null;
    checks.forEach(function(cb){cb.checked=true;});
  } else {
    selPortFilter = sel.length===total ? null : sel;
  }
  if(allCb) allCb.checked = (selPortFilter===null);
  var btn=document.getElementById('portFilterBtn');
  btn.textContent=selPortFilter===null?'All Ports':selPortFilter.length+' of '+total+' ports';
  renderPortWaitTable();
  renderRemarkSummary();
}

// ── Speed Port Filter (independent filter, shares selPortFilter) ─────
function buildSpeedPortFilterDropdown(){
  var dd=document.getElementById('speedPortFilterDropdown');
  var btn=document.getElementById('speedPortFilterBtn');
  var html='<label><input type="checkbox" value="__all__" onchange="onSpeedPortFilterChange()"'+(selPortFilter===null?' checked':'')+'> All Ports</label>';
  var ports=[]; var seen=new Set();
  TODAY_DATA.fullSchedule.forEach(function(sr){
    var p=normalizePort(sr.port);
    if(p && !isBunkeringPort(p) && !seen.has(p)){ seen.add(p); ports.push(p); }
  });
  var selArr=selPortFilter||[];
  ports.sort();
  ports.forEach(function(p){
    html+='<label><input type="checkbox" value="'+p+'" onchange="onSpeedPortFilterChange()"'+(selPortFilter===null||selArr.indexOf(p)>=0?' checked':'')+'> '+p+'</label>';
  });
  dd.innerHTML=html;
  btn.textContent=selPortFilter===null?'All Ports':selPortFilter.length+' of '+ports.length+' ports';
}

function toggleSpeedPortFilter(){
  var dd=document.getElementById('speedPortFilterDropdown');
  if(!dd.classList.contains('open')){ buildSpeedPortFilterDropdown(); dd.classList.add('open'); }
  else dd.classList.remove('open');
  document.querySelectorAll('.filter-dropdown').forEach(function(d){if(d!==dd) d.classList.remove('open');});
}

function onSpeedPortFilterChange(){
  var allCb=document.querySelector('#speedPortFilterDropdown input[value="__all__"]');
  var checks=document.querySelectorAll('#speedPortFilterDropdown input[type=checkbox]:not([value="__all__"])');
  var sel=[];
  var total=0;
  checks.forEach(function(cb){total++; if(cb.checked) sel.push(cb.value);});
  if(allCb && allCb.checked){
    selPortFilter=null;
    checks.forEach(function(cb){cb.checked=true;});
  } else {
    selPortFilter = sel.length===total ? null : sel;
  }
  if(allCb) allCb.checked = (selPortFilter===null);
  var btn=document.getElementById('speedPortFilterBtn');
  btn.textContent=selPortFilter===null?'All Ports':selPortFilter.length+' of '+total+' ports';
  renderSpeedTable();
}

// ── Port Wait Analysis ────────────────────────────────────────────────

var portWaitData=[];
var portWaitSortCol=-1, portWaitSortDir=1;
var remarkCatTotals={};

// Merge port name variants: AEJEA(T1)→AEJEA, CNSHK-CCT→CNSHK, DJJIB(DMP)→DJJIB,
// MYPKG (1st CALL)→MYPKG, THLCH (ESCO)→THLCH, SGSIN(Bunkering)→SGSIN(bunker) etc.
function normalizePort(p){
  if(!p) return '';
  var s=p.trim();
  // UN/LOCODE alias: CNNSA same as CNNAS (Nansha)
  if(s==='CNNSA') s='CNNAS';
  // UN/LOCODE alias: JED same as SAJED (Jeddah)
  if(s==='JED') s='SAJED';
  // Strip anything in parentheses/brackets: (T1), (DMP), (SGTD), (RSGT), (ESCO), (TIPS), (1st CALL), (2nd CALL), (Bunkering)
  s=s.replace(/\s*[\(\（][^)\）]*[\)\）]/g,'');
  // Strip "-suffix": -Shipyard, -CCT, -MCT
  s=s.replace(/\s*-\s*[A-Za-z0-9]+$/,'');
  // Strip " anchorage" / "anchoage" (typo)
  s=s.replace(/\s*anchorage/i,'').replace(/\s*anchoage/i,'');
  // Strip " 1st CALL" / " 2nd CALL"
  s=s.replace(/\s*\d+st\s*CALL/i,'').replace(/\s*\d+nd\s*CALL/i,'');
  // Special: "CJK & NGB anchoage" → "CNNGB" (merge into Ningbo)
  if(/CJK.*NGB/i.test(s)) s='CNNGB';
  return s.trim();
}

// Is this a bunkering-only call?
function isBunkeringPort(p){
  return /bunker/i.test(p||'');
}

var totalExcludedByRemark=0;
function buildPortWaitData(){
  var byPort={};
  totalExcludedByRemark=0;
  TODAY_DATA.fullSchedule.forEach(function(sr){
    var rawPort=sr.port||'';
    if(!rawPort) return;
    // Skip bunkering-only calls
    if(isBunkeringPort(rawPort)) return;
    var port=normalizePort(rawPort);
    if(!port) return;

    // Apply port filter
    if(selPortFilter && selPortFilter.indexOf(port)<0) return;

    var wait=parseFloat(sr.wait)||0;
    var remark=sr.remark||'';
    var cat=classifyRemark(remark)||'other';

    // If remark filter is active, ONLY include calls whose remark matches selected categories
    if(selRemarkCats){
      var match=false;
      for(var i=0;i<selRemarkCats.length;i++){
        if(cat===selRemarkCats[i]){match=true;break;}
        // 'other' catches calls with no remark
        if(selRemarkCats[i]==='other' && !remark){match=true;break;}
      }
      if(!match){
        // Track excluded call for UX feedback (show in port detail)
        if(!byPort[port]) byPort[port]={port:port, calls:[], totalWait:0, maxWait:0, longWaitCalls:0, berthCalls:0, remarks:{}, excludedCalls:[]};
        byPort[port].excludedCalls.push({wait:wait, remark:remark, cat:cat, vessel:sr.vessel, voy:sr.voy, eta:sr.eta, etb:sr.etb, rawPort:rawPort});
        totalExcludedByRemark++;
        return;  // skip this call
      }
    }

    if(!byPort[port]) byPort[port]={port:port, calls:[], totalWait:0, maxWait:0, longWaitCalls:0, berthCalls:0, remarks:{}, excludedCalls:[]};
    var rec=byPort[port];
    rec.calls.push({wait:wait, remark:remark, cat:cat, vessel:sr.vessel, voy:sr.voy, eta:sr.eta, etb:sr.etb, etd:sr.etd, dateKey:sr.etaRaw, rawPort:rawPort});
    rec.totalWait+=wait;
    if(wait>rec.maxWait) rec.maxWait=wait;
    if(wait>=24) rec.longWaitCalls++;
    if(wait<6) rec.berthCalls++;
    if(remark){
      if(!rec.remarks[cat]) rec.remarks[cat]=[];
      rec.remarks[cat].push(remark);
    }
  });

  var result=[];
  for(var p in byPort){
    var rec=byPort[p];
    rec.avgWait=rec.calls.length>0 ? (rec.totalWait/rec.calls.length) : 0;
    rec.berthRate=rec.calls.length>0 ? Math.round(rec.berthCalls/rec.calls.length*100) : 0;
    rec.catLabels=Object.keys(rec.remarks).map(function(k){
      var found=REMARK_CATEGORIES.find(function(c){return c.key===k;});
      return found ? found.label : k;
    }).join(', ');
    result.push(rec);
  }

  // Default sort: berth rate descending (best ports first)
  if(portWaitSortCol<0){
    result.sort(function(a,b){return b.berthRate-a.berthRate || a.avgWait-b.avgWait;});
  }

  // Aggregate wait time by remark category
  var catTotals={};
  result.forEach(function(rec){
    rec.calls.forEach(function(cl){
      var c=cl.cat||'other';
      if(!catTotals[c]) catTotals[c]={wait:0, calls:0};
      catTotals[c].wait+=cl.wait;
      catTotals[c].calls++;
    });
  });

  portWaitData=result;
  remarkCatTotals=catTotals;
}

var PORT_WAIT_COLS=[
  {key:'rank',      label:'#'},
  {key:'port',      label:'Port'},
  {key:'calls',     label:'Calls'},
  {key:'berthRate', label:'Berth Rate (到靠率)'},
  {key:'avgWait',   label:'Avg Wait (hrs)'},
  {key:'maxWait',   label:'Max Wait (hrs)'},
  {key:'longWaitCalls', label:'Calls > 24h'},
  {key:'catLabels', label:'Remark Categories'},
];

function berthRateColor(rate){
  if(rate>=80) return '#27ae60';  // green = good
  if(rate>=50) return '#e67e22';  // orange = medium
  return '#c0392b';               // red = bad
}

function renderPortWaitTable(){
  buildPortWaitData();
  var data=portWaitData;
  var thead=document.getElementById('portWaitThead');
  var tbody=document.getElementById('portWaitTbody');

  // Header
  var h='<th style="width:22px;"></th>';
  PORT_WAIT_COLS.forEach(function(col,i){
    var arrow='';
    if(portWaitSortCol===i) arrow=portWaitSortDir===1?' \u25B2':' \u25BC';
    h+='<th style="cursor:pointer" onclick="sortPortWait('+i+')">'+col.label+arrow+'</th>';
  });
  thead.innerHTML='<tr>'+h+'</tr>';

  // Sort
  if(portWaitSortCol>=0){
    var def=PORT_WAIT_COLS[portWaitSortCol];
    var k=def.key; var d=portWaitSortDir;
    data=data.slice().sort(function(a,b){
      if(k==='port' || k==='catLabels') return (a[k]||'').localeCompare(b[k]||'')*d;
      if(k==='rank') return 0;  // rank is display-only, re-sort by berthRate
      var va=a[k]||0, vb=b[k]||0;
      return (va-vb)*d;
    });
  }

  // Body
  var rows='';
  data.forEach(function(r,idx){
    var rate=r.berthRate;
    var c=berthRateColor(rate);
    var bar='<div style="display:inline-flex;align-items:center;gap:6px;">'+
      '<div style="width:80px;height:14px;background:#e8e8e8;border-radius:7px;overflow:hidden;">'+
      '<div style="width:'+rate+'%;height:100%;background:'+c+';border-radius:7px;"></div></div>'+
      '<b style="color:'+c+';min-width:36px;">'+rate+'%</b></div>';
    var cat=(r.catLabels||'') ? '<span style="color:#C00000;font-size:11px;">'+r.catLabels+'</span>' : '<span style="color:#8a9bb0;">—</span>';
    var rowBg=rate<50 ? ' style="background:#fff5f5;"' : (rate>=80 ? ' style="background:#f0faf3;"' : '');
    var pid='pw'+idx;
    // Build call detail rows (hidden by default)
    var callRows='';
    r.calls.forEach(function(cl,i2){
      var w=cl.wait;
      var wClass=w>=24?'delay':(w<6?'ontime':'');
      var wDisp=w.toFixed(1);
      if(wClass==='delay') wDisp='<b>'+wDisp+'</b>';
      callRows+='<tr class="detail-row" style="background:#fff;">'+
        '<td style="color:#8a9bb0;font-size:11px;">'+(i2+1)+'</td>'+
        '<td style="font-size:12px;">'+escapeHtml(cl.vessel||'')+'</td>'+
        '<td style="font-size:11px;color:#6a7b8d;">'+escapeHtml(cl.voy||'')+'</td>'+
        '<td style="font-size:11px;color:#6a7b8d;">'+escapeHtml(cl.eta||'')+'</td>'+
        '<td class="center '+wClass+'" style="font-size:12px;">'+wDisp+'</td>'+
        '<td style="font-size:11px;'+(cl.remark?'color:#C00000;':'color:#8a9bb0;')+'">'+(cl.remark?escapeHtml(cl.remark):'—')+'</td>'+
        '</tr>';
    });
    // Show excluded calls if remark filter is active
    var excludedRows='';
    if(selRemarkCats && r.excludedCalls && r.excludedCalls.length>0){
      excludedRows='<tr><td colspan="6" style="padding:6px 8px;color:#999;font-size:11px;border-top:1px dashed #e0e0e0;">'+
        '<span style="color:#e67e22;">&#9888;</span> Filtered out ('+r.excludedCalls.length+' calls excluded by remark filter):</td></tr>';
      r.excludedCalls.forEach(function(cl){
        var catLabel='';
        var found=REMARK_CATEGORIES.find(function(c){return c.key===cl.cat;});
        if(found) catLabel=found.label;
        excludedRows+='<tr class="excluded-row" style="background:#fafafa;color:#b0b0b0;text-decoration:line-through;">'+
          '<td style="color:#ccc;font-size:11px;">—</td>'+
          '<td style="font-size:11px;">'+escapeHtml(cl.vessel||'')+'</td>'+
          '<td style="font-size:10px;">'+escapeHtml(cl.voy||'')+'</td>'+
          '<td style="font-size:10px;">'+escapeHtml(cl.eta||'')+'</td>'+
          '<td class="center" style="font-size:11px;">'+(cl.wait||0).toFixed(1)+'</td>'+
          '<td style="font-size:10px;"><span style="background:#f5f5f5;color:#999;padding:1px 5px;border-radius:3px;">'+catLabel+'</span> '+(cl.remark?escapeHtml(cl.remark):'—')+'</td>'+
          '</tr>';
      });
    }
    var detailHtml='';
    if(callRows || excludedRows){
      detailHtml='<tr id="'+pid+'-detail" class="detail-wrap" style="display:none;"><td></td><td colspan="8" style="padding:0;">'+
        '<div style="padding:4px 0;">'+
        '<table style="width:100%;font-size:12px;border-collapse:collapse;">'+
        '<thead><tr style="background:#eef3f7;color:#5a697a;font-size:11px;">'+
        '<th style="width:24px;padding:3px 6px;">#</th>'+
        '<th style="padding:3px 6px;text-align:left;">Vessel</th>'+
        '<th style="padding:3px 6px;text-align:left;">Voy</th>'+
        '<th style="padding:3px 6px;text-align:left;">ETA</th>'+
        '<th style="padding:3px 6px;text-align:center;">Wait (hrs)</th>'+
        '<th style="padding:3px 6px;text-align:left;">Remark</th>'+
        '</tr></thead>'+
        '<tbody>'+callRows+excludedRows+'</tbody>'+
        '</table></div></td></tr>';
    }
    rows+='<tr'+rowBg+' class="port-row" onclick="togglePortWaitDetail(\''+pid+'\')" style="cursor:pointer;">'+
      '<td class="center" style="font-size:12px;color:#8a9bb0;">'+(callRows?'<span id="'+pid+'-icon">&#9654;</span>':'')+'</td>'+
      '<td class="center" style="color:#8a9bb0;font-size:12px;">'+(idx+1)+'</td>'+
      '<td><strong>'+r.port+'</strong></td>'+
      '<td class="center">'+r.calls.length+'</td>'+
      '<td>'+bar+'</td>'+
      '<td class="center">'+r.avgWait.toFixed(1)+'</td>'+
      '<td class="center">'+r.maxWait.toFixed(1)+'</td>'+
      '<td class="center'+(r.longWaitCalls>0?' delay':'')+'">'+(r.longWaitCalls>0?'<b>'+r.longWaitCalls+'</b>':'0')+'</td>'+
      '<td>'+cat+'</td>'+
      '</tr>'+detailHtml;
  });
  tbody.innerHTML=rows;
  var statText=data.length+' ports';
  if(selRemarkCats && totalExcludedByRemark>0){
    statText+=' · <span style="color:#e67e22;">'+totalExcludedByRemark+' calls filtered</span>';
  }
  statText+=' · 到靠率 = wait < 6h';
  document.getElementById('statPortWait').innerHTML=statText;
}

function togglePortWaitDetail(pid){
  var detail=document.getElementById(pid+'-detail');
  var icon=document.getElementById(pid+'-icon');
  if(!detail) return;
  if(detail.style.display==='none'){
    detail.style.display='';
    if(icon) icon.innerHTML='&#9660;';
  }else{
    detail.style.display='none';
    if(icon) icon.innerHTML='&#9654;';
  }
}

function sortPortWait(col){
  // col is PORT_WAIT_COLS index: 0=rank(not sortable), 1=port, 2=calls, ...
  if(col===0) return;
  if(portWaitSortCol===col) portWaitSortDir*=-1;
  else{portWaitSortCol=col;portWaitSortDir=1;}
  renderPortWaitTable();
  renderRemarkSummary();
}

function renderRemarkSummary(){
  var totals=remarkCatTotals;
  var totalWait=0, totalCalls=0;
  for(var k in totals){ totalWait+=totals[k].wait; totalCalls+=totals[k].calls; }

  var cont=document.getElementById('remarkSummary');
  if(totalWait===0){
    cont.innerHTML='<div style=\"color:#8a9bb0;font-size:12px;padding:12px;\">No wait data matches the selected remark filter.</div>';
    return;
  }

  // Build sorted list by wait desc
  var items=[];
  REMARK_CATEGORIES.forEach(function(cat){
    var t=totals[cat.key]||{wait:0,calls:0};
    if(t.calls>0 || cat.key==='other') items.push({key:cat.key, label:cat.label, wait:t.wait, calls:t.calls});
  });
  items.sort(function(a,b){return b.wait-a.wait;});

  // Color palette per category
  var colors={congestion:'#c0392b', weather:'#2980b9', bunker:'#8e44ad', phase:'#d35400', msa:'#e67e22', adhoc:'#f39c12', cargo:'#16a085', other:'#7f8c8d'};

  var html='';
  // Summary header
  html+='<div style="display:flex;align-items:center;gap:16px;margin-bottom:8px;font-size:12px;">';
  html+='<span style="font-weight:600;">Total Wait: <b style="color:#c00000;">'+totalWait.toFixed(1)+'h</b></span>';
  html+='<span>Avg per call: <b>'+ (totalCalls>0?(totalWait/totalCalls).toFixed(1):'0') +'h</b></span>';
  html+='<span>Calls: <b>'+totalCalls+'</b></span>';
  html+='</div>';

  // Bars
  var maxW=totalWait||1;
  items.forEach(function(it){
    var pct=Math.round(it.wait/maxW*100);
    var c=colors[it.key]||'#7f8c8d';
    if(pct===0){ html+='<div style="display:flex;align-items:center;gap:8px;font-size:12px;color:#8a9bb0;height:20px;">'+it.label+' <span>— 0h (0 calls)</span></div>'; return; }
    html+='<div style="display:flex;align-items:center;gap:8px;font-size:12px;">';
    html+='<span style="width:160px;text-align:right;font-weight:600;flex-shrink:0;">'+it.label+'</span>';
    html+='<div style="flex:1;height:20px;background:#e8e8e8;border-radius:4px;overflow:hidden;min-width:80px;">';
    html+='<div style="width:'+(pct||1)+'%;height:100%;background:'+c+';border-radius:4px;display:flex;align-items:center;justify-content:flex-end;padding-right:6px;font-size:10px;color:#fff;font-weight:600;min-width:'+(pct>0?pct*0.5:1)+'px;">'+(pct>=8?pct+'%':'')+'</div>';
    html+='</div>';
    html+='<span style="font-weight:700;color:'+c+';width:56px;text-align:right;flex-shrink:0;">'+it.wait.toFixed(1)+'h</span>';
    html+='<span style="color:#8a9bb0;width:48px;text-align:right;flex-shrink:0;">'+it.calls+' call'+(it.calls>1?'s':'')+'</span>';
    html+='</div>';
  });

  cont.innerHTML=html;
}

// ── Vessel Speed Analysis ─────────────────────────────────────────────

var vesselSpeedData=[];
var speedSortCol=-1, speedSortDir=1;

function buildVesselSpeedData(){
  // If port filter is active, first determine which vessels qualify (visit >=1 selected port)
  var allowedVessels=null;
  if(selPortFilter){
    allowedVessels={};
    TODAY_DATA.fullSchedule.forEach(function(sr){
      var p=normalizePort(sr.port);
      if(p && selPortFilter.indexOf(p)>=0 && sr.vessel) allowedVessels[sr.vessel]=true;
    });
  }

  var byVessel={};
  TODAY_DATA.fullSchedule.forEach(function(sr){
    var v=sr.vessel||'';
    if(allowedVessels && !allowedVessels[v]) return;
    var spd=parseFloat(sr.speed);
    if(!v || isNaN(spd) || spd<=0 || spd>20) return;  // exclude unrealistic: ≤0 or >20kn
    if(!byVessel[v]) byVessel[v]={vessel:v, route:sr.route, speeds:[], sum:0, min:spd, max:spd};
    var rec=byVessel[v];
    rec.speeds.push(spd);
    rec.sum+=spd;
    if(spd<rec.min) rec.min=spd;
    if(spd>rec.max) rec.max=spd;
  });
  var result=[];
  for(var v in byVessel){
    var rec=byVessel[v];
    rec.avg=rec.speeds.length>0 ? rec.sum/rec.speeds.length : 0;
    rec.legs=rec.speeds.length;
    result.push(rec);
  }
  vesselSpeedData=result;
}

var SPEED_COLS=[
  {key:'vessel', label:'Vessel'},
  {key:'route',  label:'Route'},
  {key:'legs',   label:'Legs'},
  {key:'avg',    label:'Avg Speed (kn)'},
  {key:'min',    label:'Min Speed (kn)'},
  {key:'max',    label:'Max Speed (kn)'},
];

function renderSpeedTable(){
  buildVesselSpeedData();
  var data=vesselSpeedData;
  var thead=document.getElementById('speedThead');
  var tbody=document.getElementById('speedTbody');

  var h='';
  SPEED_COLS.forEach(function(col,i){
    var arrow='';
    if(speedSortCol===i) arrow=speedSortDir===1?' \u25B2':' \u25BC';
    h+='<th style="cursor:pointer" onclick="sortSpeed('+i+')">'+col.label+arrow+'</th>';
  });
  thead.innerHTML='<tr>'+h+'</tr>';

  if(speedSortCol>=0){
    var def=SPEED_COLS[speedSortCol];
    var k=def.key; var d=speedSortDir;
    data=data.slice().sort(function(a,b){
      if(k==='vessel' || k==='route') return (a[k]||'').localeCompare(b[k]||'')*d;
      var va=a[k]||0, vb=b[k]||0;
      return (va-vb)*d;
    });
  }

  var rows='';
  data.forEach(function(r){
    rows+='<tr>'+
      '<td><strong>'+r.vessel+'</strong></td>'+
      '<td class="center">'+r.route+'</td>'+
      '<td class="center">'+r.legs+'</td>'+
      '<td class="center"><b>'+r.avg.toFixed(1)+'</b></td>'+
      '<td class="center">'+r.min.toFixed(1)+'</td>'+
      '<td class="center">'+r.max.toFixed(1)+'</td>'+
      '</tr>';
  });
  tbody.innerHTML=rows;
  document.getElementById('statSpeed').textContent=data.length+' vessels';
}

function sortSpeed(col){
  if(speedSortCol===col) speedSortDir*=-1;
  else{speedSortCol=col;speedSortDir=1;}
  renderSpeedTable();
}

function exportSpeedExcel(){
  buildVesselSpeedData();
  var data=vesselSpeedData;
  var todayStr=TODAY_DATA.date;
  var headers=SPEED_COLS.map(function(c){return c.label;});
  var numCols=headers.length;

  function thinBorder(){var s={style:'thin',color:{rgb:'BFBFBF'}};return{top:s,bottom:s,left:s,right:s};}
  var B=thinBorder();
  function F(rgb){return{patternType:'solid',fgColor:{rgb:rgb}};}
  var tS={font:{name:'Arial',bold:true,color:{rgb:'FFFFFF'},sz:14},fill:F('2E75B6'),alignment:{horizontal:'center',vertical:'center'},border:B};
  var hS={font:{name:'Arial',bold:true,color:{rgb:'FFFFFF'},sz:10},fill:F('1F4E79'),alignment:{horizontal:'center',vertical:'center',wrapText:true},border:B};

  var sheetData=[];
  var tr=[];
  for(var c=0;c<numCols;c++) tr[c]={v:(c===0?'CUL VESSEL SPEED ANALYSIS  --  As of '+todayStr:''),s:tS};
  sheetData.push(tr);
  sheetData.push(headers.map(function(h){return{v:h,s:hS};}));

  data.forEach(function(r){
    var nS={font:{name:'Arial',sz:9},fill:F('FFFFFF'),border:B,alignment:{horizontal:'left',vertical:'center'}};
    var cS={font:{name:'Arial',sz:9},fill:F('FFFFFF'),border:B,alignment:{horizontal:'center',vertical:'center'}};
    var bS={font:{name:'Arial',bold:true,sz:9},fill:F('FFFFFF'),border:B,alignment:{horizontal:'left',vertical:'center'}};
    sheetData.push([
      {v:r.vessel,s:bS},
      {v:r.route,s:cS},
      {v:r.legs,s:cS},
      {v:r.avg.toFixed(1),s:{font:{name:'Arial',bold:true,sz:9},fill:F('FFFFFF'),border:B,alignment:{horizontal:'center',vertical:'center'}}},
      {v:r.min.toFixed(1),s:cS},
      {v:r.max.toFixed(1),s:cS},
    ]);
  });

  var ws=XLSX.utils.aoa_to_sheet(sheetData);
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:numCols-1}}];
  ws['!cols']=headers.map(function(h){return{wch:Math.max(h.length+4,10)};});
  var totalRows=sheetData.length;
  ws['!rows']=[{hpt:28},{hpt:30}];for(var j=2;j<totalRows;j++)ws['!rows'].push({hpt:16});

  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Vessel Speed');
  XLSX.writeFile(wb,'CUL Vessel Speed '+todayStr+'.xlsx');
}

// ── Init ──────────────────────────────────────────────────────────────

function initPortView(){
  buildPortFilterDropdown();
  buildRemarkFilterDropdown();
  renderPortWaitTable();
  renderRemarkSummary();
}

function initSpeedView(){
  buildSpeedPortFilterDropdown();
  renderSpeedTable();
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
  document.getElementById('headerDate').textContent='Data as of '+TODAY_DATA.date+'  |  Updated '+TODAY_DATA.generatedAt+'  |  '+TODAY_DATA.vessels.length+' vessels  |  '+TODAY_DATA.fullSchedule.length+' schedule rows';
  document.getElementById('footerTs').textContent='Data updated: '+TODAY_DATA.generatedAt;
  initSummary();
  initFullSchedule();
  initPortView();
  initSpeedView();
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
