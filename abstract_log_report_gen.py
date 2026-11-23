# -*- coding: utf-8 -*-
"""
从 abstract_log.db 生成自包含分析网页 abstract_log_report.html.

数据内嵌为 JSON, 纯静态可直接打开/部署. 深色主题, 含:
  - 概览卡片 (航次/航段/总距离/总油耗/船数)
  - 筛选 (船/航次/时间/港口)
  - 航段明细表 (每港->下一港 distance/speed/油耗, 可排序)
  - 图表 (船均油耗对比 / 距离分布 / 航速-油耗散点)
  - 航次汇总表 (总距离/总油耗/油效)

用法: python abstract_log_report_gen.py
"""
import os
import json
import sqlite3
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(HERE, "rob_data", "abstract_log.db")
OUT = os.path.join(HERE, "abstract_log_report.html")


def load_data():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    voyages = [dict(r) for r in c.execute("SELECT * FROM voyages ORDER BY received")]
    legs = [dict(r) for r in c.execute("SELECT * FROM legs ORDER BY received, seq")]
    c.close()
    return voyages, legs


def main():
    voyages, legs = load_data()
    # 数值清理: NaN -> None
    def clean(rows):
        out = []
        for r in rows:
            d = {}
            for k, v in r.items():
                if isinstance(v, float) and (v != v):
                    d[k] = None
                elif isinstance(v, str):
                    d[k] = v.strip()
                else:
                    d[k] = v
            out.append(d)
        return out
    voyages, legs = clean(voyages), clean(legs)
    generated = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    html = HTML_TEMPLATE
    html = html.replace("/*__VOYAGES__*/", json.dumps(voyages, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("/*__LEGS__*/", json.dumps(legs, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("__GENERATED__", generated)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML -> %s (%d bytes)" % (OUT, os.path.getsize(OUT)))
    print("voyages=%d legs=%d generated=%s" % (len(voyages), len(legs), generated))


HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CUL Abstract Log 航次数据分析</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"
        onerror="var s=document.createElement('script');s.src='https://cdn.bootcdn.net/ajax/libs/Chart.js/4.4.1/chart.umd.min.js';document.head.appendChild(s);"></script>
<style>
:root{
  --bg:#0d1117; --panel:#161b22; --panel2:#1c2330; --border:#2d3748;
  --txt:#e6edf3; --muted:#8b98a9; --accent:#4fc3f7; --accent2:#29b6f6;
  --up:#f85149; --down:#3fb950; --warn:#d29922; --gold:#e3b341;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI','PingFang SC','Microsoft YaHei',Arial,sans-serif;background:var(--bg);color:var(--txt);font-size:13px;line-height:1.5;}
.header{background:linear-gradient(135deg,#1a3a5c 0%,#0d5a7a 100%);padding:16px 28px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--border);}
.header h1{font-size:19px;font-weight:700;letter-spacing:.5px;}
.header .sub{font-size:11px;color:#a8c8e0;margin-top:3px;}
.header .stamp{font-size:12px;background:rgba(255,255,255,.1);padding:7px 14px;border-radius:6px;border:1px solid rgba(255,255,255,.2);}
.wrap{max-width:1600px;margin:0 auto;padding:20px 24px;}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin-bottom:20px;}
.card{background:var(--panel);border:1px solid var(--border);border-radius:9px;padding:16px 18px;}
.card .label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;}
.card .val{font-size:25px;font-weight:700;color:var(--accent);margin-top:6px;}
.card .val small{font-size:12px;color:var(--muted);font-weight:400;}
.filters{background:var(--panel);border:1px solid var(--border);border-radius:9px;padding:14px 18px;margin-bottom:18px;display:flex;flex-wrap:wrap;gap:14px;align-items:flex-end;}
.fg{display:flex;flex-direction:column;gap:5px;}
.fg label{font-size:11px;color:var(--muted);}
.fg select,.fg input{background:var(--panel2);color:var(--txt);border:1px solid var(--border);border-radius:5px;padding:7px 10px;font-size:12px;min-width:130px;outline:none;}
.fg select:focus,.fg input:focus{border-color:var(--accent);}
.btn{background:var(--accent2);color:#04121e;border:none;border-radius:5px;padding:8px 16px;font-size:12px;font-weight:600;cursor:pointer;}
.btn:hover{background:var(--accent);}
.btn-ghost{background:transparent;color:var(--muted);border:1px solid var(--border);}
.btn-ghost:hover{color:var(--txt);border-color:var(--accent);}
.tabs{display:flex;gap:4px;margin-bottom:14px;border-bottom:1px solid var(--border);}
.tab{background:transparent;color:var(--muted);border:none;border-bottom:2px solid transparent;padding:9px 18px;font-size:13px;cursor:pointer;}
.tab.active{color:var(--accent);border-bottom-color:var(--accent);font-weight:600;}
.view{display:none;}
.view.active{display:block;}
.panel{background:var(--panel);border:1px solid var(--border);border-radius:9px;padding:16px 18px;margin-bottom:18px;}
.panel h2{font-size:14px;color:var(--accent);margin-bottom:12px;font-weight:600;}
.tbl-wrap{overflow:auto;max-height:640px;border:1px solid var(--border);border-radius:6px;}
table{border-collapse:collapse;width:100%;font-size:12px;}
th{background:var(--panel2);color:var(--muted);font-weight:600;text-align:left;padding:9px 10px;position:sticky;top:0;cursor:pointer;white-space:nowrap;border-bottom:1px solid var(--border);z-index:2;}
th:hover{color:var(--accent);}
td{padding:7px 10px;border-bottom:1px solid #212937;white-space:nowrap;}
tr:hover td{background:#1a2230;}
.num{text-align:right;font-variant-numeric:tabular-nums;}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:600;}
.b-good{background:rgba(63,185,80,.15);color:var(--down);}
.b-warn{background:rgba(210,153,34,.15);color:var(--warn);}
.route{color:var(--accent);font-weight:600;}
.charts{display:grid;grid-template-columns:1fr 1fr;gap:18px;}
.chart-box{background:var(--panel);border:1px solid var(--border);border-radius:9px;padding:16px 18px;}
.chart-box h3{font-size:13px;color:var(--accent);margin-bottom:10px;}
.chart-box canvas{max-height:320px;}
.stat-line{font-size:11px;color:var(--muted);padding:4px 0;}
.stat-line b{color:var(--txt);}
.empty{text-align:center;color:var(--muted);padding:40px;font-size:14px;}
.info{font-size:11px;color:var(--muted);margin-top:8px;line-height:1.7;}
@media(max-width:900px){.charts{grid-template-columns:1fr;}}
</style>
</head>
<body>
<div class="header">
  <div><h1>CUL Abstract Log 航次数据分析</h1><div class="sub">每港→下一港 距离 / 航速 / 油耗 · 源自船长 Abstract Log 报告全字段解析</div></div>
  <div class="stamp">数据截至 __GENERATED__</div>
</div>
<div class="wrap">
  <div class="cards" id="cards"></div>
  <div class="filters">
    <div class="fg"><label>船舶</label><select id="fVessel"><option value="">全部船舶</option></select></div>
    <div class="fg"><label>航次</label><select id="fVoyage"><option value="">全部航次</option></select></div>
    <div class="fg"><label>起始日期</label><input type="date" id="fDateFrom"></div>
    <div class="fg"><label>结束日期</label><input type="date" id="fDateTo"></div>
    <div class="fg"><label>港口(起或讫)</label><input id="fPort" placeholder="如 THBKK / CNSHK"></div>
    <div class="fg"><label>仅有效航段</label><select id="fValid"><option value="1">是(距离>0)</option><option value="0">全部</option></select></div>
    <button class="btn btn-ghost" id="btnReset">重置</button>
  </div>
  <div class="tabs">
    <button class="tab active" data-v="legs">航段明细</button>
    <button class="tab" data-v="voyages">航次汇总</button>
    <button class="tab" data-v="charts">图表分析</button>
  </div>

  <div class="view active" id="v-legs">
    <div class="panel">
      <h2>每港 → 下一港 航段数据 <span id="legCount" style="color:var(--muted);font-weight:400"></span></h2>
      <div class="info">距离优先取 GPS，GPS=0 时回退 LOG；FO=燃油，DO=柴油；M/E=主机，A/E=辅机，BLR=锅炉。点击表头排序。</div>
      <div class="tbl-wrap" style="margin-top:10px"><table id="legTable"><thead></thead><tbody></tbody></table></div>
    </div>
  </div>

  <div class="view" id="v-voyages">
    <div class="panel">
      <h2>航次汇总 <span id="voyCount" style="color:var(--muted);font-weight:400"></span></h2>
      <div class="tbl-wrap" style="margin-top:10px"><table id="voyTable"><thead></thead><tbody></tbody></table></div>
    </div>
  </div>

  <div class="view" id="v-charts">
    <div class="charts">
      <div class="chart-box"><h3>各船平均航段油耗 (FO M/E 吨/段)</h3><canvas id="c1"></canvas></div>
      <div class="chart-box"><h3>各船总航次数</h3><canvas id="c2"></canvas></div>
      <div class="chart-box"><h3>航速 vs 主机油耗 (按距离着色)</h3><canvas id="c3"></canvas></div>
      <div class="chart-box"><h3>航段距离分布 (NM)</h3><canvas id="c4"></canvas></div>
    </div>
  </div>
</div>

<script>
const VOYAGES=/*__VOYAGES__*/;
const LEGS=/*__LEGS__*/;
const $=id=>document.getElementById(id);
const fmt=(v,d=1)=>v===null||v===undefined||v===''?'—':(typeof v==='number'?v.toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d}):v);
const dist=leg=>{let g=leg.dist_sea_gps,l=leg.dist_sea_log;return (g&&g>0)?g:(l&&l>0?l:null);};

function initFilters(){
  const vs=[...new Set(VOYAGES.map(v=>v.vessel).filter(Boolean))].sort();
  vs.forEach(v=>$('fVessel').add(new Option(v,v)));
}
function syncVoyages(){
  const vsel=$('fVessel').value;
  const vs=[...new Set(VOYAGES.filter(v=>!vsel||v.vessel===vsel).map(v=>v.voyage_no).filter(Boolean))].sort();
  const cur=$('fVoyage').value;
  $('fVoyage').innerHTML='<option value="">全部航次</option>';
  vs.forEach(v=>$('fVoyage').add(new Option(v,v)));
  if(vs.includes(cur))$('fVoyage').value=cur;
}
function filteredLegs(){
  const v=$('fVessel').value,vy=$('fVoyage').value,df=$('fDateFrom').value,dt=$('fDateTo').value,
        p=$('fPort').value.trim().toUpperCase(),valid=$('fValid').value==='1';
  return LEGS.filter(l=>{
    if(v&&l.vessel!==v)return false;
    if(vy&&l.voyage_no!==vy)return false;
    const d=(l.received||'').slice(0,10);
    if(df&&d<df)return false; if(dt&&d>dt)return false;
    if(p&&!((l.from_port||'').toUpperCase().includes(p)||(l.to_port||'').toUpperCase().includes(p)))return false;
    if(valid&&!(dist(l)>0))return false;
    return true;
  });
}
function filteredVoyages(){
  const v=$('fVessel').value,vy=$('fVoyage').value,df=$('fDateFrom').value,dt=$('fDateTo').value;
  return VOYAGES.filter(x=>{
    if(v&&x.vessel!==v)return false;
    if(vy&&x.voyage_no!==vy)return false;
    const d=(x.received||'').slice(0,10);
    if(df&&d<df)return false; if(dt&&d>dt)return false;
    return true;
  });
}

// ---- cards ----
function renderCards(){
  const legs=filteredLegs(),voys=filteredVoyages();
  const totDist=legs.reduce((s,l)=>s+(dist(l)||0),0);
  const totFo=legs.reduce((s,l)=>s+(l.fo_me_sea||0)+(l.fo_ae_sea||0)+(l.fo_blr_sea||0),0);
  const ships=new Set(voys.map(v=>v.vessel).filter(Boolean)).size;
  const spd=legs.filter(l=>l.avg_sog>0);
  const avgSpd=spd.length?spd.reduce((s,l)=>s+l.avg_sog,0)/spd.length:0;
  const data=[
    ['航次数',voys.length,''],['航段数',legs.length,''],
    ['总航行距离',fmt(totDist,0),'NM'],['总燃油消耗(主机)',fmt(totFo,0),'MT'],
    ['平均每港距离',fmt(legs.length?totDist/legs.length:0,0),'NM'],['平均航速',fmt(avgSpd,1),'kn'],
    ['涉及船舶',ships,'艘'],['数据覆盖',VOYAGES.length?'2026':'—','年']
  ];
  $('cards').innerHTML=data.map(d=>`<div class="card"><div class="label">${d[0]}</div><div class="val">${d[1]}${d[2]?' <small>'+d[2]+'</small>':''}</div></div>`).join('');
}

// ---- tables ----
const legCols=[
  ['vessel','船舶',null],['voyage_no','航次',null],['received','日期',x=>(x||'').slice(0,10)],
  ['from_port','起港',null],['to_port','讫港',null],
  ['dist_gps','距离(NM)',x=>fmt(dist(x),0)],['avg_sog','平均航速(kn)',x=>fmt(x.avg_sog,1)],
  ['avg_rpm','RPM',x=>fmt(x.avg_rpm,0)],['sea_hrs','海上(h)',x=>fmt(x.sea_hrs,1)],
  ['fo_me_sea','FO主机(MT)',x=>fmt(x.fo_me_sea,2)],['fo_ae_sea','FO辅机(MT)',x=>fmt(x.fo_ae_sea,2)],
  ['fo_blr_sea','FO锅炉(MT)',x=>fmt(x.fo_blr_sea,2)],['do_me_sea','DO主机(MT)',x=>fmt(x.do_me_sea,2)],
  ['arr_draft','到港吃水',x=>fmt(x.arr_draft,1)],['dep_draft','离港吃水',x=>fmt(x.dep_draft,1)]
];
const voyCols=[
  ['vessel','船舶',null],['voyage_no','航次',null],['voy_commenced','开航',x=>(x||'').slice(0,10)],
  ['voy_completed','完航',x=>(x||'').slice(0,10)],['leg_count','航段数',x=>fmt(x.leg_count,0)],
  ['voy_ttl_distance','总距离(NM)',x=>fmt(x.voy_ttl_distance,0)],
  ['ttl_sea_hrs','海上(h)',x=>fmt(x.ttl_sea_hrs,0)],
  ['tot_cons_lsfo','总耗LSFO(MT)',x=>fmt(x.tot_cons_lsfo,1)],
  ['tot_cons_mgo','总耗MGO(MT)',x=>fmt(x.tot_cons_mgo,1)],
  ['bunk_start_port','加油港',null],
  ['eff','油效(NM/MT)',x=>x.voy_ttl_distance&&x.tot_cons_lsfo?fmt(x.voy_ttl_distance/x.tot_cons_lsfo,1):'—']
];
let legSort={k:'received',dir:-1},voySort={k:'voy_completed',dir:-1};
function makeTable(tbl,cols,rows,sort,countEl){
  $(countEl).textContent='('+rows.length+' 条)';
  rows=[...rows].sort((a,b)=>{let va=a[sort.k],vb=b[sort.k];if(va===null||va===undefined)va=-1e18;if(vb===null||vb===undefined)vb=-1e18;
    if(typeof va==='string')return va.localeCompare(vb)*sort.dir;return (va-vb)*sort.dir;});
  tbl.querySelector('thead').innerHTML='<tr>'+cols.map(c=>`<th data-k="${c[0]}">${c[1]}</th>`).join('')+'</tr>';
  tbl.querySelector('thead').querySelectorAll('th').forEach(th=>th.onclick=()=>{
    const k=th.dataset.k;if(sort.k===k)sort.dir*=-1;else{sort.k=k;sort.dir=1;}renderTables();});
  const body=rows.map(r=>'<tr>'+cols.map(c=>{
    let val=c[2]?c[2](r):(r[c[0]]??'');
    const isNum=typeof val==='number'||(c[2]&&/^\d/.test(String(val)));
    return `<td class="${isNum?'num':''}">${val}</td>`;
  }).join('')+'</tr>').join('');
  tbl.querySelector('tbody').innerHTML=body||`<tr><td colspan="${cols.length}" class="empty">无数据</td></tr>`;
}
function renderTables(){
  makeTable($('legTable'),legCols,filteredLegs(),legSort,'legCount');
  makeTable($('voyTable'),voyCols,filteredVoyages(),voySort,'voyCount');
}

// ---- charts ----
let charts={};
function renderCharts(){
  if(typeof Chart==='undefined')return;
  Object.values(charts).forEach(c=>c.destroy());charts={};
  const legs=filteredLegs(),voys=filteredVoyages();
  Chart.defaults.color='#8b98a9';Chart.defaults.borderColor='#2d3748';
  // c1: 各船平均航段油耗
  const byShip={};
  legs.forEach(l=>{if(!l.vessel)return;const f=(l.fo_me_sea||0)+(l.fo_ae_sea||0)+(l.fo_blr_sea||0);
    (byShip[l.vessel]=byShip[l.vessel]||[]).push(f);});
  const ships=Object.keys(byShip).filter(k=>byShip[k].length>=2).sort((a,b)=>{
    const av=x=>byShip[x].reduce((s,v)=>s+v,0)/byShip[x].length;return av(b)-av(a);}).slice(0,14);
  charts.c1=new Chart($('c1'),{type:'bar',data:{labels:ships,datasets:[{label:'平均油耗 MT/段',
    data:ships.map(s=>byShip[s].reduce((a,b)=>a+b,0)/byShip[s].length),backgroundColor:'#29b6f6',borderRadius:3}]},
    options:{indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true}}}});
  // c2: 各船航次数
  const vc={};voys.forEach(v=>{if(v.vessel)vc[v.vessel]=(vc[v.vessel]||0)+1;});
  const vk=Object.keys(vc).sort((a,b)=>vc[b]-vc[a]).slice(0,14);
  charts.c2=new Chart($('c2'),{type:'bar',data:{labels:vk,datasets:[{label:'航次数',
    data:vk.map(k=>vc[k]),backgroundColor:'#66bb6a',borderRadius:3}]},
    options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{precision:0}}}}});
  // c3: 航速 vs 主机油耗 (散点)
  const pts=legs.filter(l=>l.avg_sog>0&&l.fo_me_sea!==null&&dist(l)>0).map(l=>({x:l.avg_sog,y:l.fo_me_sea,d:dist(l)}));
  charts.c3=new Chart($('c3'),{type:'scatter',data:{datasets:[{label:'航段',data:pts,
    backgroundColor:pts.map(p=>p.d>1000?'#e3b341':p.d>300?'#4fc3f7':'#66bb6a'),pointRadius:3}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`航速 ${c.parsed.x.toFixed(1)}kn · 主机耗 ${c.parsed.y.toFixed(1)}MT · ${c.raw.d.toFixed(0)}NM`}}},
      scales:{x:{title:{display:true,text:'平均航速 kn'}},y:{title:{display:true,text:'主机油耗 MT'}}}}});
  // c4: 距离分布直方图
  const ds=legs.map(dist).filter(d=>d>0);
  const bins=[0,100,300,600,1000,2000,4000,1e9],labels=['0-100','100-300','300-600','600-1k','1k-2k','2k-4k','4k+'];
  const cnt=labels.map((_,i)=>ds.filter(d=>d>=bins[i]&&d<bins[i+1]).length);
  charts.c4=new Chart($('c4'),{type:'bar',data:{labels:labels.map(l=>l+' NM'),datasets:[{label:'航段数',
    data:cnt,backgroundColor:'#ab47bc',borderRadius:3}]},
    options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{precision:0}}}}});
}
function renderAll(){renderCards();renderTables();renderCharts();}

// events
['fVessel','fVoyage','fDateFrom','fDateTo','fPort','fValid'].forEach(id=>$(id).addEventListener('input',()=>{
  if(id==='fVessel')syncVoyages();renderAll();}));
$('btnReset').onclick=()=>{['fVessel','fVoyage','fDateFrom','fDateTo','fPort'].forEach(id=>$(id).value='');$('fValid').value='1';syncVoyages();renderAll();};
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
  document.querySelectorAll('.view').forEach(x=>x.classList.remove('active'));
  t.classList.add('active');$('v-'+t.dataset.v).classList.add('active');
  if(t.dataset.v==='charts')setTimeout(renderCharts,50);
});
initFilters();syncVoyages();renderAll();
</script>
</body>
</html>
'''

if __name__ == "__main__":
    main()
