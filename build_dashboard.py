#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build a self-contained HTML dashboard from agg.json (Vessel Bapfile analysis)."""
import json, os, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
AGG = os.path.join(HERE, "agg.json")
OUT = os.path.join(HERE, "Vessel_Bapfile_Dashboard.html")
CHART_JS_LOCAL = os.path.join(HERE, "chart.umd.min.js")
CHART_CDN = "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"

def get_chartjs():
    if os.path.exists(CHART_JS_LOCAL) and os.path.getsize(CHART_JS_LOCAL) > 100000:
        with open(CHART_JS_LOCAL, "r", encoding="utf-8") as f:
            return f.read(), False
    try:
        req = urllib.request.Request(CHART_CDN, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as r:
            data = r.read().decode("utf-8")
        with open(CHART_JS_LOCAL, "w", encoding="utf-8") as f:
            f.write(data)
        return data, False
    except Exception as e:
        print("[WARN] could not fetch Chart.js, using CDN link:", e)
        return None, True

def main():
    with open(AGG, "r", encoding="utf-8") as f:
        agg = json.load(f)
    chartjs_src, use_cdn = get_chartjs()

    t = agg["totals"]
    rows = agg["total_rows"]
    dvvd = agg["distinct_vvd"]

    # month data sorted
    months = sorted(agg["month"]["rows"].keys())
    month_rows = [agg["month"]["rows"].get(m,0) for m in months]
    month_teu = [agg["month"]["teu"].get(m,0) for m in months]
    month_wt = [agg["month"]["weight_tons"].get(m,0) for m in months]

    lane_items = list(agg["lane"]["teu"].items())
    lane_items_top = sorted(lane_items, key=lambda x:-x[1])[:15]
    lane_names = [k for k,_ in lane_items_top][::-1]
    lane_teu = [v for _,v in lane_items_top][::-1]

    carrier_items = list(agg["carrier"]["teu"].items())
    carrier_items_top = sorted(carrier_items, key=lambda x:-x[1])[:15]
    carrier_names = [k for k,_ in carrier_items_top][::-1]
    carrier_teu = [v for _,v in carrier_items_top][::-1]

    vessel_items = list(agg["vessel"]["teu"].items())
    vessel_items_top = sorted(vessel_items, key=lambda x:-x[1])[:20]
    vessel_names = [k for k,_ in vessel_items_top][::-1]
    vessel_teu = [v for _,v in vessel_items_top][::-1]

    ctype = agg["cont_type"]["rows"]
    ctype_names = list(ctype.keys())
    ctype_vals = list(ctype.values())

    fe = agg["fe"]["rows"]
    fe_labels = list(fe.keys())
    fe_vals = list(fe.values())

    flags = agg["flags"]
    pol = agg["ports"]["pol_top"]
    pod = agg["ports"]["pod_top"]
    pol_names = list(pol.keys())[::-1]; pol_vals = list(pol.values())[::-1]
    pod_names = list(pod.keys())[::-1]; pod_vals = list(pod.values())[::-1]

    edi = agg["coc_soc"]["edi"]
    bkg = agg["coc_soc"]["bkg"]

    dirs = agg["dir"]

    data_js = {
        "months": months, "month_rows": month_rows, "month_teu": month_teu, "month_wt": month_wt,
        "lane_names": lane_names, "lane_teu": lane_teu,
        "carrier_names": carrier_names, "carrier_teu": carrier_teu,
        "vessel_names": vessel_names, "vessel_teu": vessel_teu,
        "ctype_names": ctype_names, "ctype_vals": ctype_vals,
        "fe_labels": fe_labels, "fe_vals": fe_vals,
        "flags": flags,
        "pol_names": pol_names, "pol_vals": pol_vals,
        "pod_names": pod_names, "pod_vals": pod_vals,
        "edi": edi, "bkg": bkg, "dirs": dirs,
    }
    data_json = json.dumps(data_js, ensure_ascii=False)

    cdn_tag = f'<script src="{CHART_CDN}"></script>' if use_cdn else ''
    inline_tag = f'<script>{chartjs_src}</script>' if not use_cdn else ''

    html = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CULines Vessel Bapfile 数据分析仪表盘</title>
__INLINE__
__CDN__
<style>
:root{--bg:#f5f7fa;--card:#fff;--ink:#1f2937;--muted:#6b7280;--line:#e5e7eb;
--blue:#2563eb;--teal:#0d9488;--amber:#d97706;--rose:#e11d48;--green:#059669;--purple:#7c3aed;}
*{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;}
body{background:var(--bg);color:var(--ink);padding:24px;line-height:1.5;}
h1{font-size:22px;margin-bottom:4px;}
.sub{color:var(--muted);font-size:13px;margin-bottom:20px;}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:24px;}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;}
.kpi .v{font-size:24px;font-weight:700;color:var(--blue);}
.kpi .l{font-size:12px;color:var(--muted);margin-top:4px;}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:18px;}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;min-height:300px;}
.card.full{grid-column:1 / -1;}
.card h3{font-size:15px;margin-bottom:12px;display:flex;align-items:center;gap:8px;}
.card h3::before{content:"";width:4px;height:16px;background:var(--blue);border-radius:2px;display:inline-block;}
.chart-wrap{position:relative;height:260px;}
@media(max-width:900px){.grid{grid-template-columns:1fr;}}
table{width:100%;border-collapse:collapse;font-size:12px;}
th,td{border:1px solid var(--line);padding:6px 8px;text-align:left;}
th{background:#f0f4f8;}
.note{font-size:12px;color:var(--muted);margin-top:8px;}
</style></head>
<body>
<h1>CULines · Vessel Bapfile 数据分析仪表盘</h1>
<div class="sub" id="subline"></div>
<div class="kpis" id="kpis"></div>
<div class="grid">
  <div class="card full"><h3>月度趋势（箱量 TEU / 票数 / 货重 吨）</h3><div class="chart-wrap"><canvas id="cMonth"></canvas></div></div>
  <div class="card"><h3>航线 TEU 排行 Top15</h3><div class="chart-wrap"><canvas id="cLane"></canvas></div></div>
  <div class="card"><h3>承运人 TEU 排行 Top15</h3><div class="chart-wrap"><canvas id="cCarrier"></canvas></div></div>
  <div class="card"><h3>船舶 TEU 排行 Top20</h3><div class="chart-wrap"><canvas id="cVessel"></canvas></div></div>
  <div class="card"><h3>箱型分布（票数）</h3><div class="chart-wrap"><canvas id="cCtype"></canvas></div></div>
  <div class="card"><h3>重箱 / 空箱占比</h3><div class="chart-wrap"><canvas id="cFe"></canvas></div></div>
  <div class="card"><h3>特殊箱占比（冷箱 / 危险品 / 超限）</h3><div class="chart-wrap"><canvas id="cFlags"></canvas></div></div>
  <div class="card"><h3>航向分布</h3><div class="chart-wrap"><canvas id="cDir"></canvas></div></div>
  <div class="card"><h3>TOP 装货港 (POL)</h3><div class="chart-wrap"><canvas id="cPol"></canvas></div></div>
  <div class="card"><h3>TOP 卸货港 (POD)</h3><div class="chart-wrap"><canvas id="cPod"></canvas></div></div>
  <div class="card full"><h3>COC / SOC 分布（EDI 来源 / BKG 来源）</h3>
    <div style="display:flex;gap:24px;flex-wrap:wrap;">
      <div class="chart-wrap" style="flex:1;min-width:300px;"><canvas id="cEdi"></canvas></div>
      <div class="chart-wrap" style="flex:1;min-width:300px;"><canvas id="cBkg"></canvas></div>
    </div>
  </div>
</div>
<div class="card full" style="margin-top:18px;"><h3>各 Sheet 数据行数</h3>
  <table id="sheetTable"></table>
  <div class="note">说明：数据来自 Vessel Bapfile.xlsx（Apache POI 导出），共 5 个分表；指标为逐箱（container-level）记录聚合。</div>
</div>
<script>
const D = __DATA__;
const fmt = n => n>=1e6 ? (n/1e6).toFixed(2)+'M' : n>=1e3 ? (n/1e3).toFixed(1)+'k' : (''+n);
// subline
document.getElementById('subline').textContent =
  '总记录数 '+D.months.length+' 个月覆盖 · 数据源：Vessel Bapfile.xlsx';
// KPIs
const kpis = [
  ['总记录数', D.__rows__],
  ['独立航次(VVD)', D.__dvvd__],
  ['总 TEU', fmt(D.__teu__)],
  ['总货重(吨)', fmt(D.__wt__)],
  ['总 UNIT', fmt(D.__unit__)],
  ['重箱数', fmt(D.__full__)],
  ['空箱数', fmt(D.__empty__)],
];
document.getElementById('kpis').innerHTML = kpis.map(k=>
  `<div class="kpi"><div class="v">${k[1]}</div><div class="l">${k[0]}</div></div>`).join('');
// sheet table
const rowsHtml = Object.entries(D.__sheets__).map(([k,v])=>
  `<tr><th>${k}</th><td>${v.data_rows.toLocaleString()}</td></tr>`).join('');
document.getElementById('sheetTable').innerHTML = '<tr><th>分表</th><th>数据行数</th></tr>'+rowsHtml;

const C = {blue:'#2563eb',teal:'#0d9488',amber:'#d97706',rose:'#e11d48',green:'#059669',purple:'#7c3aed',slate:'#64748b'};
function bar(ctx, labels, vals, color, horizontal=false, label=''){
  return new Chart(ctx,{type:horizontal?'bar':'bar',data:{labels,datasets:[{label,data:vals,backgroundColor:color,borderRadius:4}]},
    options:{indexAxis:horizontal?'y':'x',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:!!label},tooltip:{callbacks:{label:c=>c.parsed[horizontal?'x':'y'].toLocaleString()}}},
      scales:horizontal?{x:{ticks:{callback:v=>fmt(v)}}}:{y:{ticks:{callback:v=>fmt(v)}}}}});
}
// month
new Chart(document.getElementById('cMonth'),{type:'bar',
  data:{labels:D.months,datasets:[
    {label:'TEU',data:D.month_teu,backgroundColor:C.blue,yAxisID:'y',borderRadius:4},
    {label:'票数',data:D.month_rows,type:'line',borderColor:C.amber,yAxisID:'y1',tension:.3},
    {label:'货重(吨)',data:D.month_wt,type:'line',borderColor:C.teal,yAxisID:'y2',tension:.3}
  ]},
  options:{responsive:true,maintainAspectRatio:false,
    scales:{y:{position:'left',title:{display:true,text:'TEU'},ticks:{callback:v=>fmt(v)}},
      y1:{position:'right',title:{display:true,text:'票数'},grid:{drawOnChartArea:false},ticks:{callback:v=>fmt(v)}},
      y2:{display:false}},
    plugins:{tooltip:{}}}});
bar(document.getElementById('cLane'),D.lane_names,D.lane_teu,C.teal,true,'TEU');
bar(document.getElementById('cCarrier'),D.carrier_names,D.carrier_teu,C.purple,true,'TEU');
bar(document.getElementById('cVessel'),D.vessel_names,D.vessel_teu,C.blue,true,'TEU');
new Chart(document.getElementById('cCtype'),{type:'bar',
  data:{labels:D.ctype_names,datasets:[{label:'票数',data:D.ctype_vals,backgroundColor:C.amber,borderRadius:4}]},
  options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>fmt(v)}}}}});
new Chart(document.getElementById('cFe'),{type:'doughnut',
  data:{labels:D.fe_labels,datasets:[{data:D.fe_vals,backgroundColor:[C.green,C.slate,C.rose]}]},
  options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right'}}}});
const fl=D.flags;
new Chart(document.getElementById('cFlags'),{type:'bar',
  data:{labels:['冷箱 RF','危险品 DG','超限 AWK'],datasets:[{label:'箱数',data:[fl.reefer,fl.dg,fl.awk],backgroundColor:[C.teal,C.rose,C.amber],borderRadius:4}]},
  options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>fmt(v)}}}}});
const dirEntries=Object.entries(D.dirs);
new Chart(document.getElementById('cDir'),{type:'bar',
  data:{labels:dirEntries.map(e=>e[0]),datasets:[{label:'票数',data:dirEntries.map(e=>e[1]),backgroundColor:C.blue,borderRadius:4}]},
  options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{ticks:{callback:v=>fmt(v)}}}}});
bar(document.getElementById('cPol'),D.pol_names,D.pol_vals,C.teal,true,'票数');
bar(document.getElementById('cPod'),D.pod_names,D.pod_vals,C.purple,true,'票数');
const ediE=Object.entries(D.edi), bkgE=Object.entries(D.bkg);
new Chart(document.getElementById('cEdi'),{type:'doughnut',
  data:{labels:ediE.map(e=>e[0]),datasets:[{data:ediE.map(e=>e[1]),backgroundColor:['#2563eb','#0d9488','#d97706','#e11d48','#7c3aed','#64748b']}]},
  options:{responsive:true,maintainAspectRatio:false,plugins:{title:{display:true,text:'EDI COC/SOC'},legend:{position:'right'}}}});
new Chart(document.getElementById('cBkg'),{type:'doughnut',
  data:{labels:bkgE.map(e=>e[0]),datasets:[{data:bkgE.map(e=>e[1]),backgroundColor:['#2563eb','#0d9488','#d97706','#e11d48','#7c3aed','#64748b']}]},
  options:{responsive:true,maintainAspectRatio:false,plugins:{title:{display:true,text:'BKG COC/SOC'},legend:{position:'right'}}}});
</script>
</body></html>"""

    # inject scalar values
    html = (html
        .replace("__INLINE__", inline_tag)
        .replace("__CDN__", cdn_tag)
        .replace("__DATA__", data_json)
        .replace("D.__rows__", str(rows))
        .replace("D.__dvvd__", str(dvvd))
        .replace("D.__teu__", str(t["teu"]))
        .replace("D.__wt__", str(t["weight_tons"]))
        .replace("D.__unit__", str(t["unit"]))
        .replace("D.__full__", str(t["n_full"]))
        .replace("D.__empty__", str(t["n_empty"]))
        .replace("D.__sheets__", json.dumps(agg["rows_by_sheet"], ensure_ascii=False))
    )
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[DONE] dashboard -> {OUT} ({os.path.getsize(OUT)} bytes)", flush=True)

if __name__ == "__main__":
    main()
