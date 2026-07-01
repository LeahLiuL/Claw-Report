"""
gen_html.py  —  CUL Daily Movement HTML Generator
每次运行会从 Excel 提取最新数据，将 TODAY_DATA 注入到 HTML 中，
保留历史快照机制，然后输出 cul_daily_movement.html。

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
# Fallback to same folder as this script when P: drive not available
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Helpers ───────────────────────────────────────────────────────────────
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

    results = []
    for vb in vessel_blocks:
        best_row, best_eta = None, None
        for r in vb['schedule_rows']:
            eta_v = ws_src.cell(r, 9).value
            if isinstance(eta_v, datetime):
                eta_d = eta_v.date()
                if eta_d >= today and (best_eta is None or eta_d < best_eta):
                    best_eta, best_row = eta_d, r
        if best_row is None and vb['schedule_rows']:
            for r in reversed(vb['schedule_rows']):
                if isinstance(ws_src.cell(r, 9).value, datetime):
                    best_row = r; break

        if best_row:
            r = best_row
            rec = {
                'route': vb['route'], 'vessel': vb['vessel_full'],
                'code': vb['vessel_code'], 'pic': vb['pic'],
                'port':      get_str(ws_src.cell(r, 1).value),
                'voy':       get_str(ws_src.cell(r, 7).value),
                'eta':       fmt_dt(ws_src.cell(r, 9).value),
                'etb':       fmt_dt(ws_src.cell(r, 10).value),
                'etd':       fmt_dt(ws_src.cell(r, 11).value),
                'portStay':  get_str(ws_src.cell(r, 13).value),
                'etaDelay':  get_str(ws_src.cell(r, 16).value),
                'etdDelay':  get_str(ws_src.cell(r, 17).value),
                'remark':    vb['remark'],
            }
        else:
            rec = {'route': vb['route'], 'vessel': vb['vessel_full'],
                   'code': vb['vessel_code'], 'pic': vb['pic'],
                   'port':'','voy':'','eta':'','etb':'','etd':'',
                   'portStay':'','etaDelay':'','etdDelay':'','remark': vb['remark']}
        results.append(rec)

    return {'date': today.strftime('%Y-%m-%d'), 'vessels': results}

# ── Inject into HTML ──────────────────────────────────────────────────────
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CUL Daily Movement</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"><\/script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; color: #1a2332; }
  .header {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    color: #fff; padding: 18px 28px 14px;
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 3px 12px rgba(31,78,121,.35);
  }
  .header-left h1 { font-size: 20px; font-weight: 700; letter-spacing: .5px; }
  .header-left .sub { font-size: 12px; opacity: .75; margin-top: 3px; }
  .header-right { display: flex; gap: 10px; align-items: center; }
  .btn { padding: 7px 16px; border: none; border-radius: 5px; font-size: 13px; font-weight: 600; cursor: pointer; transition: .15s; }
  .btn-export { background: #F6A623; color: #fff; }
  .btn-export:hover { background: #d4891a; }
  .btn-history { background: rgba(255,255,255,.18); color: #fff; border: 1px solid rgba(255,255,255,.4); }
  .btn-history:hover { background: rgba(255,255,255,.30); }
  .controls {
    padding: 14px 28px; background: #fff; border-bottom: 1px solid #dde4ed;
    display: flex; gap: 14px; align-items: center; flex-wrap: wrap;
  }
  .controls input, .controls select {
    padding: 6px 12px; border: 1px solid #c9d5e2; border-radius: 5px;
    font-size: 13px; outline: none; height: 34px;
  }
  .controls input:focus, .controls select:focus { border-color: #2E75B6; box-shadow: 0 0 0 2px rgba(46,117,182,.15); }
  .controls input { width: 200px; }
  .controls select { min-width: 130px; }
  .stat-chip { margin-left: auto; background: #EBF3FB; border: 1px solid #c3d9f0; border-radius: 20px; padding: 4px 14px; font-size: 12px; color: #1F4E79; font-weight: 600; }
  .delay-chip { background: #fff0f0; border: 1px solid #f5c6c6; border-radius: 20px; padding: 4px 14px; font-size: 12px; color: #c00000; font-weight: 600; }
  .table-wrap { overflow-x: auto; padding: 0 28px 28px; }
  table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 12.5px; min-width: 1100px; }
  th { background: #1F4E79; color: #fff; font-weight: 600; padding: 10px; text-align: center; white-space: nowrap; position: sticky; top: 0; z-index: 2; cursor: pointer; user-select: none; }
  th:hover { background: #163b5f; }
  th .sort-arrow { display: inline-block; margin-left: 4px; opacity: .5; font-size: 10px; }
  th.sort-asc .sort-arrow::after { content: '\25b2'; opacity: 1; }
  th.sort-desc .sort-arrow::after { content: '\25bc'; opacity: 1; }
  th:not(.sort-asc):not(.sort-desc) .sort-arrow::after { content: '\21c5'; }
  td { padding: 8px 10px; border-bottom: 1px solid #e4ecf5; vertical-align: middle; }
  tr:nth-child(even) td { background: #f5f9fe; }
  tr:nth-child(odd) td { background: #fff; }
  tr:hover td { background: #dcedf9 !important; }
  .td-center { text-align: center; }
  .td-mono { font-family: 'Consolas', monospace; }
  .badge-route { display: inline-block; padding: 2px 8px; border-radius: 4px; font-weight: 700; font-size: 11px; background: #1F4E79; color: #fff; letter-spacing: .5px; }
  .badge-code { display: inline-block; padding: 2px 7px; border-radius: 4px; font-size: 11px; font-weight: 600; background: #E8F4FD; color: #1F4E79; border: 1px solid #a8cfe8; }
  .delay-tag { display: inline-block; padding: 2px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; background: #fff0f0; color: #c00000; border: 1px solid #f5c6c6; }
  .ahead-tag { display: inline-block; padding: 2px 7px; border-radius: 3px; font-size: 11px; font-weight: 700; background: #f0fff4; color: #1a7340; border: 1px solid #b7dfca; }
  .no-data { text-align: center; padding: 40px; color: #8a9bb0; font-size: 14px; }
  .remark-cell { max-width: 280px; white-space: normal; line-height: 1.4; color: #5a6e82; font-size: 11.5px; }
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
<div class="header">
  <div class="header-left">
    <h1>&#9875; CUL VESSEL DAILY MOVEMENT</h1>
    <div class="sub" id="headerDate">Loading&hellip;</div>
  </div>
  <div class="header-right">
    <button class="btn btn-history" onclick="openHistory()">&#128203; History</button>
    <button class="btn btn-export" onclick="exportExcel()">&#8595; Export Excel</button>
  </div>
</div>
<div class="controls">
  <input type="text" id="searchBox" placeholder="&#128269; Search vessel / port / PIC&hellip;" oninput="renderTable()">
  <select id="filterRoute" onchange="renderTable()"><option value="">All Routes</option></select>
  <select id="filterPic" onchange="renderTable()"><option value="">All PIC</option></select>
  <select id="filterDelay" onchange="renderTable()">
    <option value="">All Status</option>
    <option value="delay">Delay Only</option>
    <option value="ahead">Ahead Only</option>
    <option value="normal">No Delay</option>
  </select>
  <span class="stat-chip" id="statTotal">&#8212; vessels</span>
  <span class="delay-chip" id="statDelay">&#8212; delayed</span>
</div>
<div class="table-wrap">
  <table id="mainTable">
    <thead><tr>
      <th onclick="sortBy(0)">Route<span class="sort-arrow"></span></th>
      <th onclick="sortBy(1)">Vessel<span class="sort-arrow"></span></th>
      <th onclick="sortBy(2)">Code<span class="sort-arrow"></span></th>
      <th onclick="sortBy(3)">PIC<span class="sort-arrow"></span></th>
      <th onclick="sortBy(4)">Port<span class="sort-arrow"></span></th>
      <th onclick="sortBy(5)">Voy. No<span class="sort-arrow"></span></th>
      <th onclick="sortBy(6)">ETA<span class="sort-arrow"></span></th>
      <th onclick="sortBy(7)">ETB<span class="sort-arrow"></span></th>
      <th onclick="sortBy(8)">ETD<span class="sort-arrow"></span></th>
      <th onclick="sortBy(9)">Port Stay(hr)<span class="sort-arrow"></span></th>
      <th onclick="sortBy(10)">ETA Delay<span class="sort-arrow"></span></th>
      <th onclick="sortBy(11)">ETD Delay<span class="sort-arrow"></span></th>
      <th>Remark</th>
    </tr></thead>
    <tbody id="tbody"></tbody>
  </table>
</div>
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
<script>
const SNAPSHOTS = {};
const TODAY_DATA = __TODAY_DATA__;

function loadSnapshots(){try{const s=localStorage.getItem('cul_movement_history');if(s)Object.assign(SNAPSHOTS,JSON.parse(s));}catch(e){}}
function saveSnapshot(d){SNAPSHOTS[d.date]=d.vessels;try{localStorage.setItem('cul_movement_history',JSON.stringify(SNAPSHOTS));}catch(e){}}

let allData=[],sortCol=-1,sortDir=1;
function init(){
  loadSnapshots(); saveSnapshot(TODAY_DATA);
  allData=TODAY_DATA.vessels;
  document.getElementById('headerDate').textContent='As of '+TODAY_DATA.date+'  |  '+allData.length+' vessels';
  document.getElementById('footerTs').textContent='Generated: '+TODAY_DATA.date;
  const routes=[...new Set(allData.map(r=>r.route))].sort();
  const pics=[...new Set(allData.map(r=>r.pic))].sort();
  const rSel=document.getElementById('filterRoute'),pSel=document.getElementById('filterPic');
  routes.forEach(r=>{const o=document.createElement('option');o.value=r;o.textContent=r;rSel.appendChild(o);});
  pics.forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;pSel.appendChild(o);});
  renderTable();
}
function getFilteredData(){
  const q=document.getElementById('searchBox').value.toLowerCase();
  const route=document.getElementById('filterRoute').value;
  const pic=document.getElementById('filterPic').value;
  const delay=document.getElementById('filterDelay').value;
  let data=allData.filter(r=>{
    if(route&&r.route!==route)return false;
    if(pic&&r.pic!==pic)return false;
    if(q&&!`${r.vessel} ${r.port} ${r.pic} ${r.code} ${r.remark}`.toLowerCase().includes(q))return false;
    if(delay==='delay')return r.etaDelay.toLowerCase().includes('delay');
    if(delay==='ahead')return r.etaDelay.toLowerCase().includes('ahead');
    if(delay==='normal')return!r.etaDelay;
    return true;
  });
  if(sortCol>=0){
    const keys=['route','vessel','code','pic','port','voy','eta','etb','etd','portStay','etaDelay','etdDelay'];
    data.sort((a,b)=>(a[keys[sortCol]]||'').localeCompare(b[keys[sortCol]]||'')*sortDir);
  }
  return data;
}
function delayTag(v){
  if(!v)return'';
  if(v.toLowerCase().startsWith('ahead'))return`<span class="ahead-tag">${v}</span>`;
  if(v.toLowerCase().startsWith('delay'))return`<span class="delay-tag">${v}</span>`;
  return v;
}
function renderTable(){
  const data=getFilteredData();
  const tbody=document.getElementById('tbody');
  if(!data.length){tbody.innerHTML='<tr><td colspan="13" class="no-data">No matching records found.</td></tr>';document.getElementById('statTotal').textContent='0 vessels';document.getElementById('statDelay').textContent='0 delayed';return;}
  const dc=data.filter(r=>r.etaDelay.toLowerCase().includes('delay')).length;
  document.getElementById('statTotal').textContent=data.length+' vessels';
  document.getElementById('statDelay').textContent=dc+' delayed';
  tbody.innerHTML=data.map(r=>`<tr>
    <td class="td-center"><span class="badge-route">${r.route}</span></td>
    <td><strong>${r.vessel}</strong></td>
    <td class="td-center"><span class="badge-code">${r.code}</span></td>
    <td>${r.pic}</td>
    <td class="td-center td-mono"><strong>${r.port}</strong></td>
    <td class="td-center td-mono">${r.voy}</td>
    <td class="td-center td-mono">${r.eta}</td>
    <td class="td-center td-mono">${r.etb}</td>
    <td class="td-center td-mono">${r.etd}</td>
    <td class="td-center">${r.portStay}</td>
    <td class="td-center">${delayTag(r.etaDelay)}</td>
    <td class="td-center">${delayTag(r.etdDelay)}</td>
    <td class="remark-cell">${r.remark}</td>
  </tr>`).join('');
}
function sortBy(col){
  const ths=document.querySelectorAll('#mainTable thead th');
  ths.forEach(th=>th.classList.remove('sort-asc','sort-desc'));
  if(sortCol===col){sortDir*=-1;}else{sortCol=col;sortDir=1;}
  ths[col].classList.add(sortDir===1?'sort-asc':'sort-desc');
  renderTable();
}
function exportExcel(){
  const data=getFilteredData();
  const headers=['Route','Vessel Name','Code','PIC','Port','Voy. No','ETA','ETB','ETD','Port Stay (hr)','ETA Delay','ETD Delay','Remark'];
  const rows=data.map(r=>[r.route,r.vessel,r.code,r.pic,r.port,r.voy,r.eta,r.etb,r.etd,r.portStay,r.etaDelay,r.etdDelay,r.remark]);
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet([headers,...rows]);
  ws['!cols']=[8,22,8,14,10,10,14,14,14,13,14,14,40].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb,ws,'Daily Movement');
  XLSX.writeFile(wb,'CUL Daily Movement '+TODAY_DATA.date+'.xlsx');
}
function openHistory(){
  const dates=Object.keys(SNAPSHOTS).sort().reverse();
  const listEl=document.getElementById('historyDateList');
  listEl.innerHTML='';
  dates.forEach((d,i)=>{
    const btn=document.createElement('button');
    btn.className='history-date-btn'+(i===0?' active':'');
    btn.textContent=d;
    btn.onclick=()=>{document.querySelectorAll('.history-date-btn').forEach(b=>b.classList.remove('active'));btn.classList.add('active');renderHistoryTable(d);};
    listEl.appendChild(btn);
  });
  if(dates.length)renderHistoryTable(dates[0]);
  document.getElementById('historyModal').classList.add('open');
}
function renderHistoryTable(date){
  const vessels=SNAPSHOTS[date]||[];
  if(!vessels.length){document.getElementById('historyTableWrap').innerHTML='<p style="color:#8a9bb0;padding:20px">No data.</p>';return;}
  const headers=['Route','Vessel','Code','PIC','Port','Voy','ETA','ETB','ETD','Port Stay','ETA Delay','ETD Delay','Remark'];
  const keys=['route','vessel','code','pic','port','voy','eta','etb','etd','portStay','etaDelay','etdDelay','remark'];
  document.getElementById('historyTableWrap').innerHTML=`<table><thead><tr>${headers.map(h=>`<th style="font-size:11px;padding:7px 8px">${h}</th>`).join('')}</tr></thead><tbody>${vessels.map(r=>`<tr>${keys.map(k=>`<td style="padding:6px 8px;border-bottom:1px solid #eee;font-size:11px">${r[k]||''}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
}
function closeHistory(){document.getElementById('historyModal').classList.remove('open');}
document.getElementById('historyModal').addEventListener('click',e=>{if(e.target===document.getElementById('historyModal'))closeHistory();});
init();
<\/script>
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

    # Fallback: use script dir when P: not available
    if not os.path.exists(os.path.dirname(excel_path)):
        excel_path = os.path.join(SCRIPT_DIR, 'CUL DAILY MOVEMENT.xlsx')
    if not os.path.exists(os.path.dirname(out_path)):
        out_path = os.path.join(SCRIPT_DIR, 'cul_daily_movement.html')

    print(f'Reading: {excel_path}')
    data = extract(excel_path)
    print(f'  -> {len(data["vessels"])} vessels extracted, date={data["date"]}')

    json_str = json.dumps(data, ensure_ascii=False)
    html = HTML_TEMPLATE.replace('__TODAY_DATA__', json_str)

    os.makedirs(os.path.dirname(out_path) if os.path.dirname(out_path) else '.', exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Saved : {out_path}')

if __name__ == '__main__':
    main()
