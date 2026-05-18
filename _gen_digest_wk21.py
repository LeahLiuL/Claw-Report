
# Script to generate the Week 21 2026 Alphaliner Digest HTML
content = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CUL &mdash; Alphaliner Intelligence Digest | Week 21, 2026</title>
<style>
:root {
  --bg: #0d1117; --card: #161b22; --border: #21262d;
  --text: #e6edf3; --muted: #8b949e; --accent: #58a6ff;
  --green: #3fb950; --orange: #d29922; --red: #f85149;
  --purple: #bc8cff; --teal: #39d353;
  --tag-ops: #0d4429; --tag-ops-t: #3fb950;
  --tag-fin: #1a2f45; --tag-fin-t: #58a6ff;
  --tag-leg: #2d1b4e; --tag-leg-t: #bc8cff;
  --tag-eq: #3d2a0a; --tag-eq-t: #d29922;
  --tag-net: #0f3240; --tag-net-t: #39d353;
  --tag-trd: #3d1515; --tag-trd-t: #f85149;
  --tag-rnd: #1a2a1a; --tag-rnd-t: #56d364;
  --tag-com: #2a1a3d; --tag-com-t: #e79dff;
  --tag-it: #0f2040; --tag-it-t: #79c0ff;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }
a { color: var(--accent); text-decoration: none; }
.top-bar { background: #010409; border-bottom: 1px solid var(--border); padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; z-index: 100; }
.logo { display: flex; align-items: center; gap: 10px; }
.logo-icon { background: linear-gradient(135deg, #1f6feb, #58a6ff); border-radius: 8px; width: 36px; height: 36px; display: flex; align-items: center; justify-content: center; font-size: 18px; }
.logo-text h1 { font-size: 16px; font-weight: 700; color: var(--text); }
.logo-text p { font-size: 11px; color: var(--muted); }
.lang-btn { background: var(--card); border: 1px solid var(--border); border-radius: 20px; padding: 5px 14px; font-size: 12px; cursor: pointer; color: var(--text); transition: all 0.2s; }
.lang-btn:hover { border-color: var(--accent); color: var(--accent); }
.lang-btn.active { background: var(--accent); color: #0d1117; border-color: var(--accent); font-weight: 700; }
.lang-controls { display: flex; gap: 8px; }
.hero { background: linear-gradient(135deg, #0d1117 0%, #161b22 100%); border-bottom: 1px solid var(--border); padding: 28px 24px 20px; }
.hero-inner { max-width: 1200px; margin: 0 auto; }
.hero h2 { font-size: 24px; font-weight: 800; margin-bottom: 6px; }
.hero h2 .hl { color: var(--accent); }
.hero p { font-size: 13px; color: var(--muted); margin-bottom: 16px; }
.hero-stats { display: flex; gap: 20px; flex-wrap: wrap; }
.stat-pill { background: var(--card); border: 1px solid var(--border); border-radius: 20px; padding: 5px 14px; font-size: 12px; }
.stat-pill .sv { color: var(--accent); font-weight: 700; }
.filter-bar { background: var(--card); border-bottom: 1px solid var(--border); padding: 12px 24px; position: sticky; top: 61px; z-index: 90; }
.filter-inner { max-width: 1200px; margin: 0 auto; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
.filter-label { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; margin-right: 4px; }
.ftag { background: #21262d; border: 1px solid var(--border); border-radius: 20px; padding: 4px 12px; font-size: 12px; cursor: pointer; transition: all 0.15s; user-select: none; }
.ftag:hover { border-color: var(--accent); color: var(--accent); }
.ftag.active { background: var(--accent); color: #0d1117; border-color: var(--accent); font-weight: 700; }
.result-count { margin-left: auto; font-size: 12px; color: var(--muted); }
.result-count span { color: var(--accent); font-weight: 700; }
.feed { max-width: 1200px; margin: 0 auto; padding: 20px 24px; }
.week-label { font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: var(--muted); margin: 24px 0 12px; border-left: 3px solid var(--accent); padding-left: 10px; }
.article-card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px; margin-bottom: 10px; cursor: pointer; transition: border-color 0.15s, background 0.15s; position: relative; }
.article-card:hover { border-color: #388bfd; background: #1c2128; }
.article-card.featured { border-left: 3px solid var(--orange); }
.article-card.cul-direct { border-left: 3px solid var(--green); }
.card-top { display: flex; gap: 10px; align-items: flex-start; flex-wrap: wrap; margin-bottom: 8px; }
.card-title { font-size: 14px; font-weight: 600; flex: 1; min-width: 200px; line-height: 1.5; }
.card-badges { display: flex; gap: 5px; flex-wrap: wrap; }
.badge { font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 10px; white-space: nowrap; }
.badge-star { background: rgba(210,153,34,0.2); color: var(--orange); border: 1px solid rgba(210,153,34,0.4); }
.badge-cul { background: rgba(63,185,80,0.15); color: var(--green); border: 1px solid rgba(63,185,80,0.3); }
.badge-ops { background: var(--tag-ops); color: var(--tag-ops-t); border: 1px solid rgba(63,185,80,0.2); }
.badge-fin { background: var(--tag-fin); color: var(--tag-fin-t); border: 1px solid rgba(88,166,255,0.2); }
.badge-leg { background: var(--tag-leg); color: var(--tag-leg-t); border: 1px solid rgba(188,140,255,0.2); }
.badge-eq  { background: var(--tag-eq);  color: var(--tag-eq-t);  border: 1px solid rgba(210,153,34,0.2); }
.badge-net { background: var(--tag-net); color: var(--tag-net-t); border: 1px solid rgba(57,211,83,0.2); }
.badge-trd { background: var(--tag-trd); color: var(--tag-trd-t); border: 1px solid rgba(248,81,73,0.2); }
.badge-rnd { background: var(--tag-rnd); color: var(--tag-rnd-t); border: 1px solid rgba(86,211,100,0.2); }
.badge-com { background: var(--tag-com); color: var(--tag-com-t); border: 1px solid rgba(231,157,255,0.2); }
.badge-it  { background: var(--tag-it);  color: var(--tag-it-t);  border: 1px solid rgba(121,192,255,0.2); }
.badge-region { background: #21262d; color: var(--muted); border: 1px solid var(--border); }
.card-summary { font-size: 13px; color: var(--muted); line-height: 1.6; margin-bottom: 8px; }
.card-impact { font-size: 12px; background: rgba(63,185,80,0.06); border: 1px solid rgba(63,185,80,0.15); border-radius: 8px; padding: 8px 12px; margin-top: 6px; color: #8fe898; line-height: 1.5; }
.card-impact strong { color: var(--green); }
.card-bottom { display: flex; gap: 10px; align-items: center; margin-top: 10px; flex-wrap: wrap; }
.card-meta { font-size: 11px; color: var(--muted); }
.card-actions { margin-left: auto; display: flex; gap: 8px; }
.btn-sm { font-size: 11px; padding: 3px 10px; border-radius: 6px; border: 1px solid var(--border); background: #21262d; color: var(--text); cursor: pointer; transition: all 0.15s; }
.btn-sm:hover { border-color: var(--accent); color: var(--accent); }
.sidebar-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 200; }
.sidebar-overlay.open { display: block; }
.sidebar { position: fixed; right: -500px; top: 0; bottom: 0; width: 480px; background: #161b22; border-left: 1px solid #30363d; z-index: 210; overflow-y: auto; transition: right 0.3s ease; padding: 0; }
.sidebar.open { right: 0; }
.sidebar-header { background: #010409; border-bottom: 1px solid var(--border); padding: 16px 20px; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; }
.sidebar-header h3 { font-size: 14px; font-weight: 700; }
.sidebar-close { background: none; border: none; color: var(--muted); font-size: 20px; cursor: pointer; padding: 0 4px; line-height: 1; }
.sidebar-close:hover { color: var(--text); }
.sidebar-body { padding: 20px; }
.sidebar-section { margin-bottom: 20px; }
.sidebar-section h4 { font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); margin-bottom: 8px; }
.sidebar-section ul { padding-left: 16px; }
.sidebar-section li { font-size: 13px; color: var(--muted); line-height: 1.7; }
.sidebar-section li strong { color: var(--text); }
.impact-box { background: rgba(63,185,80,0.07); border: 1px solid rgba(63,185,80,0.2); border-radius: 8px; padding: 12px; }
.impact-box p { color: #8fe898; font-size: 13px; line-height: 1.6; }
.data-tabs { max-width: 1200px; margin: 0 auto; padding: 0 24px 30px; }
.tabs-nav { display: flex; gap: 4px; border-bottom: 1px solid var(--border); margin-bottom: 20px; }
.tab-btn { background: none; border: none; color: var(--muted); font-size: 13px; padding: 8px 16px; cursor: pointer; border-bottom: 2px solid transparent; transition: all 0.15s; }
.tab-btn:hover { color: var(--text); }
.tab-btn.active { color: var(--accent); border-bottom-color: var(--accent); font-weight: 600; }
.tab-pane { display: none; }
.tab-pane.active { display: block; }
.data-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 14px; margin-bottom: 20px; }
.data-card { background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; }
.dc-label { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
.dc-value { font-size: 24px; font-weight: 800; color: var(--text); margin-bottom: 4px; }
.dc-change { font-size: 12px; font-weight: 600; }
.dc-change.up { color: var(--red); }
.dc-change.down { color: var(--green); }
.dc-sub { font-size: 11px; color: var(--muted); margin-top: 4px; }
.data-table { background: var(--card); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; margin-bottom: 16px; }
.data-table table { width: 100%; border-collapse: collapse; font-size: 13px; }
.data-table th { background: #21262d; color: var(--muted); padding: 10px 14px; text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1px solid var(--border); }
.data-table td { padding: 10px 14px; border-bottom: 1px solid #21262d; color: var(--text); }
.data-table tr:last-child td { border-bottom: none; }
.data-table tr:hover td { background: #1c2128; }
.td-up { color: var(--red); font-weight: 600; }
.td-dn { color: var(--green); font-weight: 600; }
.td-pr { font-weight: 700; color: var(--text); }
.section-sub { font-size: 12px; color: var(--muted); margin-bottom: 12px; }
footer { border-top: 1px solid var(--border); padding: 20px 24px; text-align: center; font-size: 11px; color: var(--muted); }
.zh { display: none; }
.en { display: block; }
body.lang-zh .zh { display: block; }
body.lang-zh .en { display: none; }
span.zh { display: none; }
span.en { display: inline; }
body.lang-zh span.zh { display: inline; }
body.lang-zh span.en { display: none; }
</style>
</head>
<body id="pageBody">
<div class="top-bar">
  <div class="logo">
    <div class="logo-icon">&#x26F2;</div>
    <div class="logo-text">
      <h1>CUL &mdash; Alphaliner Intelligence Digest</h1>
      <p>Week 21, 2026 (May 12 &ndash; May 18) &middot; Powered by Alphaliner &amp; Market Intelligence</p>
    </div>
  </div>
  <div class="lang-controls">
    <button class="lang-btn active" id="btnEN" onclick="setLang('en')">EN</button>
    <button class="lang-btn" id="btnZH" onclick="setLang('zh')">&#x4E2D;&#x6587;</button>
  </div>
</div>
<div class="hero">
  <div class="hero-inner">
    <h2 class="en">Alphaliner <span class="hl">Weekly Intelligence</span> &mdash; Week 21</h2>
    <h2 class="zh">Alphaliner <span class="hl">&#x6BCF;&#x5468;&#x60C5;&#x62A5;&#x6458;&#x8981;</span> &mdash; &#x7B2C;21&#x5468;</h2>
    <p class="en">AI-curated shipping intelligence with CUL impact assessment | May 12 &ndash; May 18, 2026 | Source: Alphaliner / Drewry WCI / SCFI / Seavantage</p>
    <p class="zh">AI&#x7CBE;&#x9009;&#x822A;&#x8FD0;&#x60C5;&#x62A5;&#xFF0C;&#x9644;CUL&#x5F71;&#x54CD;&#x8BC4;&#x4F30; | 2026&#x5E745&#x6708;12&#x65E5;&ndash;5&#x6708;18&#x65E5; | &#x6570;&#x636E;&#x6765;&#x6E90;&#xFF1A;Alphaliner / Drewry / SCFI / Seavantage</p>
    <div class="hero-stats">
      <div class="stat-pill"><span class="sv">10</span> <span class="en">Articles This Week</span><span class="zh">&#x672C;&#x5468;&#x6761;&#x76EE;</span></div>
      <div class="stat-pill"><span class="sv">6</span> <span class="en">CUL Direct</span><span class="zh">CUL&#x76F4;&#x63A5;&#x76F8;&#x5173;</span></div>
      <div class="stat-pill"><span class="sv">3</span> <span class="en">Featured Items</span><span class="zh">&#x91CD;&#x70B9;&#x9879;&#x76EE;</span></div>
      <div class="stat-pill en"><span class="sv" style="color:var(--red);">WCI +12%</span> Largest Surge in 2026</div>
      <div class="stat-pill zh"><span class="sv" style="color:var(--red);">WCI +12%</span> 2026年最大单周涨幅</div>
      <div class="stat-pill en"><span class="sv">SCFI 2,140</span> Above 2,000 pts</div>
      <div class="stat-pill zh"><span class="sv">SCFI 2,140</span> 突破2,000点</div>
    </div>
  </div>
</div>
<div class="filter-bar">
  <div class="filter-inner">
    <span class="filter-label en">Dept</span>
    <span class="filter-label zh">&#x90E8;&#x95E8;</span>
    <span class="ftag active" data-dept="all">All</span>
    <span class="ftag" data-dept="ops"><span class="en">Ops</span><span class="zh">&#x8FD0;&#x8425;</span></span>
    <span class="ftag" data-dept="fin"><span class="en">Finance</span><span class="zh">&#x8D22;&#x52A1;</span></span>
    <span class="ftag" data-dept="legal"><span class="en">Legal</span><span class="zh">&#x6CD5;&#x52A1;</span></span>
    <span class="ftag" data-dept="equip"><span class="en">Equipment</span><span class="zh">&#x8BBE;&#x5907;</span></span>
    <span class="ftag" data-dept="network"><span class="en">Network</span><span class="zh">&#x822A;&#x7EBF;</span></span>
    <span class="ftag" data-dept="trade"><span class="en">Trade</span><span class="zh">&#x8D38;&#x6613;</span></span>
    <span class="ftag" data-dept="commercial"><span class="en">Commercial</span><span class="zh">&#x5546;&#x4E1A;</span></span>
    <div class="result-count en">Showing <span id="countEN">10</span> of 10 items</div>
    <div class="result-count zh">&#x663E;&#x793A; <span id="countZH">10</span> / 10 &#x6761;</div>
  </div>
</div>
<div class="feed" id="mainFeed">
  <div class="week-label">&#x1F5D3; Week 21 &middot; May 12 &ndash; May 18, 2026 &middot; Alphaliner / Drewry / SSE / Seavantage</div>

  <div class="article-card featured cul-direct" data-depts="commercial,trade,fin">
    <div class="card-top">
      <div class="card-title en">&#x2B50; WCI Surges 12% to $2,553/FEU (May 14) &mdash; Largest Single-Week Jump in 2026, Transpacific &amp; Asia-Europe Both Rally</div>
      <div class="card-title zh">&#x2B50; WCI 单周飙升 12% 至 $2,553/FEU（5月14日）——2026年最大单周涨幅，跨太平洋与亚欧齐涨</div>
      <div class="card-badges">
        <span class="badge badge-star">&#x2B50; Featured</span><span class="badge badge-cul">&#x26A1; CUL Direct</span>
        <span class="badge badge-com">Commercial</span><span class="badge badge-trd">Trade</span><span class="badge badge-fin">Finance</span>
        <span class="badge badge-region">Global</span>
      </div>
    </div>
    <div class="card-summary en">The Drewry World Container Index (WCI) surged 12% to $2,553/FEU in the week ending May 14, 2026 &mdash; the largest single-week increase recorded in 2026. Driven simultaneously by Transpacific (TPEB) and Asia-Europe rate increases, this surge reflects: (1) GRI implementation by Maersk, MSC, and CMA CGM; (2) US importer front-loading ahead of the US-China tariff truce mid-July 2026 expiry; (3) systematic blank sailing capacity withdrawal (30 blanks in Wk21-25 per Drewry); (4) Strait of Hormuz remaining effectively closed, compounding Red Sea/Suez avoidance to absorb ~10-12% of global effective capacity. Transpacific rates (US East Coast) have surged to approximately $3,800/FEU (Seavantage). WCI is now up 14.33% YoY and 13.67% over the past month. <span style="color:#8b949e;font-size:11px;">(Source: Drewry WCI, 2026-05-14)</span></div>
    <div class="card-summary zh">Drewry 世界集装箱运价指数（WCI）在截至 2026 年 5 月 14 日的一周内飙升 12%，至 $2,553/FEU——创下 2026 年最大单周涨幅。跨太平洋（TPEB）和亚欧航线同步大涨，驱动因素包括：(1) 马士基、MSC、达飞 GRI 落地执行；(2) 美国进口商在中美关税休战 7 月中旬到期前大规模前置备货；(3) 系统性空班运力削减（Drewry 追踪第 21-25 周 30 个空班）；(4) 霍尔木兹海峡仍实际关闭，叠加红海/苏伊士绕行，共吸收约 10-12% 全球有效运力。跨太平洋运价（美东）激增至约 $3,800/FEU（Seavantage）。WCI 同比上涨 14.33%，近一个月上涨 13.67%。 <span style="color:#8b949e;font-size:11px;">(数据来源: Drewry WCI, 2026-05-14)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> The +12% WCI surge is a major positive signal — use this data in Q3 contract negotiations immediately. Blank sailing support (30 Wk21-25) + front-load demand suggests sustained rates through June. However, prepare for mid-July demand cliff when tariff truce window closes. TPEB rates near $3,800/FEU — benchmark CUL Transpacific pricing. Alert chartering team: improved rate environment warrants reviewing Q3 open vessel positions.</span>
      <span class="zh"> +12% 的 WCI 激增是重大积极信号——立即在三季度合同谈判中使用此数据。空班支撑（第21-25周30个）+ 前置备货需求表明运价可持续至整个6月。但须为7月中旬关税休战到期后的需求断崖做准备。跨太平洋接近 $3,800/FEU——以此为 CUL 跨太平洋定价基准。提醒租船团队审查三季度开放仓位。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">May 14, 2026 &middot; <a href="https://www.drewry.co.uk/supply-chain-advisors/supply-chain-expertise/world-container-index-assessed-by-drewry" target="_blank">Drewry WCI Wk21</a> / <a href="https://tradingeconomics.com/commodity/world-container-index" target="_blank">TradingEconomics</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s1')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card featured cul-direct" data-depts="commercial,fin">
    <div class="card-top">
      <div class="card-title en">&#x2B50; SCFI Jumps to 2,140.66 pts (May 15) &mdash; +186 Points WoW, First Above 2,000 Since Q1; SCFI/CCFI Divergence Signals Spot Premium</div>
      <div class="card-title zh">&#x2B50; SCFI 升至 2,140.66 点（5月15日）——周环比+186点，一季度以来首次破2,000；SCFI/CCFI 分化显示即期溢价</div>
      <div class="card-badges">
        <span class="badge badge-star">&#x2B50; Featured</span><span class="badge badge-cul">&#x26A1; CUL Direct</span>
        <span class="badge badge-fin">Finance</span><span class="badge badge-com">Commercial</span>
        <span class="badge badge-region">Asia / Global</span>
      </div>
    </div>
    <div class="card-summary en">SCFI published by the Shanghai Shipping Exchange (SSE) on May 15, 2026: composite at 2,140.66 pts, up 186.45 pts from 1,954 pts (May 8). First time above 2,000 pts since Q1 2026. The CCFI (contract-rate focused index) reached only 1,280.46 pts, up just 0.1% WoW &mdash; highlighting that spot rates are surging far ahead of long-term contract rates. The SCFI/CCFI divergence (+12% spot vs. +0.1% contract) is characteristic of demand surge or capacity tightening phases and is a strong signal to accelerate contract negotiations. The SCFI gain from Wk19 (1,911 pts) to Wk21 (2,140 pts) represents a +12% move in two weeks, consistent with the WCI surge. <span style="color:#8b949e;font-size:11px;">(Source: SSE/GMT Eight, 2026-05-15)</span></div>
    <div class="card-summary zh">上海航交所（SSE）2026 年 5 月 15 日发布 SCFI 综合指数 2,140.66 点，较 5 月 8 日约 1,954 点上涨 186.45 点，一季度以来首次突破 2,000 点。CCFI（合同运价指数）仅报 1,280.46 点，周涨幅仅 0.1%——凸显即期运价远超长期合同运价。SCFI/CCFI 分化（即期 +12% vs. 合同 +0.1%）是需求激增或运力趋紧阶段的典型特征，是加快合同谈判的强烈信号。SCFI 从第 19 周（1,911 点）到第 21 周（2,140 点）两周涨幅约 12%，与 WCI 飙升同步。 <span style="color:#8b949e;font-size:11px;">(数据来源: 上海航交所/GMT Eight, 2026-05-15)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> SCFI above 2,000 is a key psychological milestone — use in commercial strategy. SCFI/CCFI divergence is urgent: accelerate Q3 contract renewals before the spot premium narrows. CUL should benchmark weekly booked rates vs. SCFI. Brief management team on strongest SCFI weekly move in 2026 (+186 pts).</span>
      <span class="zh"> SCFI 突破 2,000 是关键心理里程碑——用于商业策略。SCFI/CCFI 分化紧迫：在即期溢价收窄前加快三季度合同续约。CUL 应每周对比已订运价与 SCFI。向管理团队通报 2026 年最强 SCFI 单周涨幅（+186点）。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">May 15, 2026 &middot; <a href="https://gmteight.com/flash/detail/1366490" target="_blank">GMT EIGHT</a> / <a href="https://www.sse.net.cn/index/singleIndex?indexType=scfi" target="_blank">SSE (SCFI)</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s2')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card featured cul-direct" data-depts="ops,legal,network,commercial">
    <div class="card-top">
      <div class="card-title en">&#x2B50; &#x26A0;&#xFE0F; Strait of Hormuz Effectively Closed &mdash; Iran Limits Transits, Dual Disruption (Red Sea + Hormuz) Absorbs ~10-12% Global Capacity</div>
      <div class="card-title zh">&#x2B50; &#x26A0;&#xFE0F; 霍尔木兹海峡实际关闭——伊朗限制通行，双重中断（红海+霍尔木兹）吸收约 10-12% 全球运力</div>
      <div class="card-badges">
        <span class="badge badge-star">&#x2B50; Featured</span><span class="badge badge-cul">&#x26A1; CUL Direct</span>
        <span class="badge badge-ops">Ops</span><span class="badge badge-leg">Legal</span>
        <span class="badge badge-net">Network</span><span class="badge badge-com">Commercial</span>
        <span class="badge badge-region">Middle East / Global</span>
      </div>
    </div>
    <div class="card-summary en">Strait of Hormuz remains effectively closed as of Wk21, following Operation Epic Fury (US-Israel coalition, March 2026). Iran is actively limiting transits and reportedly imposing controlled routing and tolls. Container carriers suspended Hormuz transits in March 2026, rerouting around the southern tip of Africa. As of April 12, 2026 (Seavantage), the ceasefire was fragile; no material improvement as of Wk21. Combined with ongoing Red Sea/Suez avoidance (since 2024), the dual disruption absorbs an estimated 10-12% of global effective container capacity — a structural supply constraint that is a key driver of the WCI +12% surge. War risk insurance for Gulf/Hormuz transits remains extremely elevated. <span style="color:#8b949e;font-size:11px;">(Source: Seavantage / CNBC / HS Today, 2026)</span></div>
    <div class="card-summary zh">截至第 21 周，霍尔木兹海峡自 2026 年 3 月"史诗怒火行动"后仍实际关闭。伊朗积极限制通行并据报对船舶征收通行费。集装箱船公司已于 2026 年 3 月暂停霍尔木兹通行，绕道非洲南端。截至 2026 年 4 月 12 日（Seavantage），停火脆弱；第 21 周无实质改善。叠加自 2024 年以来持续的红海/苏伊士绕行，双重中断估计吸收全球约 10-12% 有效集装箱运力——这是 WCI +12% 飙升的关键结构性供给约束因素。波斯湾/霍尔木兹战争风险保险维持极高水平。 <span style="color:#8b949e;font-size:11px;">(数据来源: Seavantage / CNBC, 2026)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> Hormuz closure is CUL's most critical geopolitical risk in 2026. Immediately review: (1) CUL Middle East service schedules; (2) EBS/EFS validity for affected routes; (3) war risk insurance adequacy; (4) customer advisories. Model three re-opening scenarios: Suez-only, Hormuz-only, both. Any Hormuz re-opening = market shock comparable to Suez re-opening — pre-position rate strategy now.</span>
      <span class="zh"> 霍尔木兹关闭是 CUL 2026 年最关键地缘政治风险。立即审查：(1) CUL 中东服务班期；(2) 受影响航线 EBS/EFS 有效性；(3) 战争风险保险充足性；(4) 客户通知。建立三个重开情景模型：仅苏伊士、仅霍尔木兹、同时重开。任何霍尔木兹重开 = 与苏伊士重开相当的市场冲击——现在预设运价策略。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">Wk21, 2026 &middot; <a href="https://www.seavantage.com/blog/strait-of-hormuz-crisis-2026-shipping-disruption-timeline" target="_blank">Seavantage Hormuz Timeline</a> / <a href="https://www.cnbc.com/2026/03/02/strait-of-hormuz-crisis-us-iran-israel-war-shipping-trade-oil.html" target="_blank">CNBC</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s3')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card cul-direct" data-depts="network,ops,equip">
    <div class="card-top">
      <div class="card-title en">&#x26A1; MSC Makes History: World&#39;s First Carrier to Operate 1,000 Container Vessels &mdash; Alphaliner May 5 Confirms Milestone</div>
      <div class="card-title zh">&#x26A1; MSC 创历史：全球首家运营 1,000 艘集装箱船——Alphaliner 5 月 5 日确认里程碑</div>
      <div class="card-badges">
        <span class="badge badge-cul">&#x26A1; CUL Direct</span><span class="badge badge-net">Network</span>
        <span class="badge badge-ops">Ops</span><span class="badge badge-eq">Equipment</span>
        <span class="badge badge-region">Global</span>
      </div>
    </div>
    <div class="card-summary en">MSC reached a historic milestone in early May 2026: the world&#39;s first container shipping company to operate 1,000 vessels, confirmed by Alphaliner data (May 5-6, 2026) and reported by Seatrade Maritime, Puentedemando, and others. This reflects MSC&#39;s extraordinary expansion since 2020. MSC holds ~20.4% global market share and currently has ~132 vessels on firm order. Alphaliner TOP 100 (May 16, 2026) also highlights MSC&#39;s ~50% capacity share on North Europe-Mediterranean trade — a level of market concentration unprecedented even by today&#39;s standards (with only 10 carriers total on this route). <span style="color:#8b949e;font-size:11px;">(Source: Alphaliner TOP 100 / Seatrade Maritime, 2026-05-06)</span></div>
    <div class="card-summary zh">MSC 于 2026 年 5 月初达到历史性里程碑：成为全球首家运营 1,000 艘集装箱船的班轮公司，由 Alphaliner（2026 年 5 月 5-6 日）确认，Seatrade Maritime 等报道。这反映了 MSC 自 2020 年以来的非凡扩张。MSC 持有约 20.4% 全球市场份额，目前约有 132 艘船舶在确定订单中。Alphaliner TOP 100（2026 年 5 月 16 日）还指出 MSC 在北欧—地中海航线约 50% 的运力份额——即使以当今标准衡量也前所未有（该航线仅有 10 家船公司运营）。 <span style="color:#8b949e;font-size:11px;">(数据来源: Alphaliner TOP 100 / Seatrade Maritime, 2026-05-06)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> MSC&#39;s 1,000-vessel milestone underscores the growing competitive gap. CUL must leverage niche service excellence and regional expertise. MSC&#39;s 50% N. Europe-Med share creates competitive pressure. The rise of NOO vessel activity (Seatrade/Alphaliner) signals more spot charter availability for CUL&#39;s chartering team. Update competitive intelligence deck with May 16 TOP 100 data.</span>
      <span class="zh"> MSC 1,000 艘里程碑凸显竞争差距拉大。CUL 须发挥利基服务优势和区域专业知识。MSC 50% 北欧—地中海份额形成竞争压力。NOO 船舶活动增加（Seatrade/Alphaliner）预示 CUL 租船团队有更多即期租船可用性。用 5 月 16 日 TOP 100 数据更新竞争情报材料。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">May 6-16, 2026 &middot; <a href="https://www.seatrade-maritime.com/keyword/alphaliner" target="_blank">Seatrade Maritime</a> / <a href="https://alphaliner.axsmarine.com/PublicTop100/" target="_blank">Alphaliner TOP 100 (May 16)</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s4')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card cul-direct" data-depts="ops,network,commercial">
    <div class="card-top">
      <div class="card-title en">&#x26A1; Drewry Blank Sailing Tracker Wk21-25: 30 Cancelled Sailings &mdash; Systematic 4.3% Capacity Withdrawal Underpins Rate Surge</div>
      <div class="card-title zh">&#x26A1; Drewry 空班追踪第21-25周：30个取消航次——系统性 4.3% 运力削减支撑运价飙升</div>
      <div class="card-badges">
        <span class="badge badge-cul">&#x26A1; CUL Direct</span><span class="badge badge-ops">Ops</span>
        <span class="badge badge-net">Network</span><span class="badge badge-com">Commercial</span>
        <span class="badge badge-region">Global E-W</span>
      </div>
    </div>
    <div class="card-summary en">Drewry&#39;s Cancelled Sailings Tracker confirms 30 blank sailings on major East-West trades from Week 21 (May 18-24) through Week 25 (June 15-21), out of 698 total planned sailings (~4.3% capacity reduction). This coordinated blank sailing strategy by Gemini (Maersk/Hapag), Ocean Alliance (CMA CGM/COSCO/Evergreen) and Premier (ONE/HMM/Yang Ming) supports GRI execution, compensates for demand normalization, and manages vessel deployment efficiency in the Cape routing environment. This supply-side intervention is clearly a primary contributor to the WCI +12% and SCFI +186 pts surge in Wk21. <span style="color:#8b949e;font-size:11px;">(Source: Drewry Cancelled Sailings Tracker, Wk21-25 2026)</span></div>
    <div class="card-summary zh">Drewry 空班追踪器确认，第 21 周（5 月 18-24 日）至第 25 周（6 月 15-21 日），主要东西方航线共有 30 个空班，占计划总班次 698 班的约 4.3%。Gemini（马士基/赫伯罗特）、海洋联盟（达飞/中远/长荣）和 Premier 联盟（ONE/现代/阳明）协调实施空班，支持 GRI 落地，补偿需求正常化并在好望角绕行环境下管理运力部署效率。这一供给侧干预明显是第 21 周 WCI +12% 和 SCFI +186 点飙升的主要贡献因素。 <span style="color:#8b949e;font-size:11px;">(数据来源: Drewry空班追踪器, 2026年第21-25周)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> 30 blank sailings Wk21-25 create a tight capacity window: use this to execute rate increases and fill slots. Communicate space constraints proactively to customers. Track Drewry weekly blank sailing data as a rate leading indicator — zero blanks = rate softening risk.</span>
      <span class="zh"> 第 21-25 周 30 个空班形成运力偏紧窗口：利用此机会推动涨价并填满舱位。主动向客户传达舱位限制。将 Drewry 周度空班数据作为运价领先指标跟踪——空班清零 = 运价走软风险。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">Wk21-25, 2026 &middot; <a href="https://www.drewry.co.uk/supply-chain-advisors/supply-chain-expertise/cancelled-sailings-tracker" target="_blank">Drewry Cancelled Sailings Tracker</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s5')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card" data-depts="ops,trade">
    <div class="card-top">
      <div class="card-title en">Qingdao Port Congestion Wk21: ~4-Day Delays &mdash; Tariff-Truce Cargo Rush Hits Northern China Ports</div>
      <div class="card-title zh">第21周青岛港拥堵：约4天延误——关税休战前货物激增冲击中国北方港口</div>
      <div class="card-badges">
        <span class="badge badge-ops">Ops</span><span class="badge badge-trd">Trade</span>
        <span class="badge badge-region">China / Asia</span>
      </div>
    </div>
    <div class="card-summary en">Port of Qingdao (CNTAO) is experiencing ~4-day vessel delays in Week 21 (Seavantage May 2026 Update), driven by the pre-tariff-truce expiry cargo surge. US importers front-loading China-origin goods ahead of mid-July 2026 have generated exceptional TPEB booking volumes at Qingdao, Tianjin, and other northern China ports. In April, Qingdao handled 731,476 TEU exports and 178,646 TEU imports across 564 vessel calls. Tianjin and Dalian report similar congestion pressure. The congestion is expected to persist through June (tariff truce driven). <span style="color:#8b949e;font-size:11px;">(Source: Seavantage / EconDB, 2026-05)</span></div>
    <div class="card-summary zh">第 21 周青岛港（CNTAO）出现约 4 天船舶延误（Seavantage 2026 年 5 月更新），由关税休战到期前的货物激增驱动。美国进口商在 7 月中旬前大规模前置备货，在青岛、天津等北方港口产生异常高的跨太平洋订舱量。4 月青岛处理出口 731,476 TEU、进口 178,646 TEU，共 564 次靠泊。天津、大连同样面临拥堵压力。此次拥堵预计持续至 6 月（关税休战驱动）。 <span style="color:#8b949e;font-size:11px;">(数据来源: Seavantage / EconDB, 2026-05)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> Qingdao 4-day delays impact CUL Transpacific services. Adjust cargo cutoff dates; issue customer advisories for potential rollovers; pre-position equipment. Monitor Tianjin overflow. Congestion expected through June — plan proactively.</span>
      <span class="zh"> 青岛4天延误影响 CUL 跨太平洋服务。调整截关日期；就潜在滚装向客户发布通知；预先调配设备。监控天津外溢影响。拥堵预计持续至6月——提前规划。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">May 2026 &middot; <a href="https://www.seavantage.com/blog/ocean-freight-market-update-may-2026" target="_blank">Seavantage May 2026</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s6')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card cul-direct" data-depts="network,ops">
    <div class="card-top">
      <div class="card-title en">&#x26A1; Alphaliner TOP 100 (May 16, 2026): MSC Controls ~50% North Europe-Mediterranean Capacity; NOO Activity Rising</div>
      <div class="card-title zh">&#x26A1; Alphaliner TOP 100（2026年5月16日）：MSC 控制北欧—地中海约50%运力；NOO 活动增加</div>
      <div class="card-badges">
        <span class="badge badge-cul">&#x26A1; CUL Direct</span><span class="badge badge-net">Network</span>
        <span class="badge badge-ops">Ops</span><span class="badge badge-region">Global</span>
      </div>
    </div>
    <div class="card-summary en">Alphaliner TOP 100 (May 16, 2026) highlights: MSC controls nearly 50% of capacity on the North Europe-Mediterranean corridor, with only 10 carriers on this trade (7 major MLOs dominating) — an unprecedented concentration level (Ship&amp;Bunker / Safety4Sea, May 13-14). The survey also notes a notable rise in non-operating owner (NOO) vessel activity, reflecting charter market activity growth in the current rate environment. CULines remains at #45 (Top 5% global). <span style="color:#8b949e;font-size:11px;">(Source: Alphaliner TOP 100, 2026-05-16)</span></div>
    <div class="card-summary zh">Alphaliner TOP 100（2026 年 5 月 16 日）要点：MSC 控制北欧—地中海通道近 50% 运力，该航线仅有 10 家船公司（7 家主干航线运营商主导）——市场集中度史无前例（Ship&Bunker / Safety4Sea, 5月13-14日）。调查还指出非运营船东（NOO）船舶活动显著增加，反映当前运价环境下租船市场活动增长。CULines 维持第 45 位（全球前 5%）。 <span style="color:#8b949e;font-size:11px;">(数据来源: Alphaliner TOP 100, 2026-05-16)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> Update competitive intelligence with May 16 data. MSC 50% Med share is a structural competitive threat on this corridor — assess CUL Mediterranean strategy. NOO rise = more spot charter availability for CUL chartering team. CUL #45 remains a key commercial asset in customer pitches.</span>
      <span class="zh"> 用 5 月 16 日数据更新竞争情报。MSC 50% 地中海份额是结构性竞争威胁——评估 CUL 地中海策略。NOO 增加 = CUL 租船团队更多即期租船可用性。CUL 第 45 名仍是客户推介的关键商业资产。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">May 16, 2026 &middot; <a href="https://alphaliner.axsmarine.com/PublicTop100/" target="_blank">Alphaliner TOP 100</a> / <a href="https://shipandbunker.com/news/emea/212863-msc-controls-nearly-half-of-north-europe-mediterranean-container-capacity" target="_blank">Ship&amp;Bunker</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s7')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card" data-depts="trade,commercial">
    <div class="card-top">
      <div class="card-title en">US-China Tariff Truce Front-Loading at Peak &mdash; TPEB Near $3,800/FEU (US East), Mid-July 2026 Expiry Window Looms</div>
      <div class="card-title zh">中美关税休战前置备货达峰值——跨太平洋（美东）接近 $3,800/FEU，7月中旬到期窗口临近</div>
      <div class="card-badges">
        <span class="badge badge-trd">Trade</span><span class="badge badge-com">Commercial</span>
        <span class="badge badge-region">Asia / Americas</span>
      </div>
    </div>
    <div class="card-summary en">The US-China 90-day tariff truce (Apr 22) is now in peak front-loading phase: US importers rush to ship before mid-July 2026 expiry. TPEB rates have surged to ~$3,800/FEU (US East Coast, Seavantage). Average US tariff on Chinese goods: ~33% trade-weighted blend (MS Advisory). Front-loading is bifurcated: large retailers aggressively pre-stock; smaller importers and those with diversified chains moderate China orders in favor of Vietnam, India, Bangladesh. Market is also watching for early signals of a potential truce extension. <span style="color:#8b949e;font-size:11px;">(Source: Seavantage / MS Advisory, 2026-05)</span></div>
    <div class="card-summary zh">中美 90 天关税休战（4 月 22 日）现已处于前置备货潮顶峰：美国进口商争相在 7 月中旬到期前发货。TPEB 运价激增至约 $3,800/FEU（美东，Seavantage）。中国商品美国平均关税约 33%（贸易加权混合，MS Advisory）。前置备货呈分化态势：大型零售商积极预备货；中小进口商和有多元化供应链的企业减少中国订单，转向越南、印度、孟加拉国。市场也在关注关税休战潜在延期的早期信号。 <span style="color:#8b949e;font-size:11px;">(数据来源: Seavantage / MS Advisory, 2026-05)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> Front-loading peak = near-term revenue window. Maximize TPEB utilization over next 6-8 weeks. Prepare July demand cliff contingency. $3,800/FEU TPEB is a commercial pricing reference. Q3 scenario planning (truce extension vs. expiry) is now the most critical commercial planning exercise for CUL.</span>
      <span class="zh"> 前置备货峰值 = 近期收入窗口。未来 6-8 周最大化跨太平洋利用率。准备 7 月需求断崖应急预案。$3,800/FEU 跨太平洋是商业定价参考。三季度情景规划（延期 vs. 到期）现在是 CUL 最关键的商业规划工作。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">May 2026 &middot; <a href="https://www.seavantage.com/blog/ocean-freight-market-update-may-2026" target="_blank">Seavantage May 2026</a> / <a href="https://msadvisory.com/china-us-tariffs-guide/" target="_blank">MS Advisory</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s8')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card" data-depts="ops,legal,network">
    <div class="card-top">
      <div class="card-title en">Red Sea Wk21: Cape Routing Unchanged &mdash; Dual Disruption (Red Sea + Hormuz) Reshapes Global Routes, War Risk Elevated</div>
      <div class="card-title zh">第21周红海：好望角绕行持续不变——双重中断（红海+霍尔木兹）重塑全球航线，战争风险高企</div>
      <div class="card-badges">
        <span class="badge badge-ops">Ops</span><span class="badge badge-leg">Legal</span>
        <span class="badge badge-net">Network</span><span class="badge badge-region">Middle East / Global</span>
      </div>
    </div>
    <div class="card-summary en">As of Week 21, all major container carriers continue Cape of Good Hope routing. Houthi threats remain active with renewed vessel-specific targeting. The combined Red Sea/Suez avoidance (since 2024) and Hormuz closure (since March 2026) absorbs ~10-12% of global effective container capacity — providing structural freight rate support. Red Sea shipping has partially resumed on some routes amid reduced attacks per S&amp;P Global (Feb-Mar 2026), but renewed threats deter full-scale carrier return. War risk insurance elevated across Red Sea, Gulf of Oman, Persian Gulf. <span style="color:#8b949e;font-size:11px;">(Source: S&amp;P Global / Seavantage, 2026)</span></div>
    <div class="card-summary zh">截至第 21 周，所有主要集装箱船公司仍经好望角绕行。胡塞武装威胁持续活跃并有新的船舶定向打击威胁。红海/苏伊士绕行（自 2024 年）与霍尔木兹关闭（自 2026 年 3 月）叠加，吸收约 10-12% 全球有效集装箱运力，为运价提供结构性支撑。据标普全球（2-3 月），红海航运在袭击减少情况下部分恢复，但持续威胁阻碍全面回归。战争风险保险在红海、阿曼湾、波斯湾维持高位。 <span style="color:#8b949e;font-size:11px;">(数据来源: S&amp;P Global / Seavantage, 2026)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> Dual disruption is CUL's most important structural macro factor. EBS/EFS remain valid. Bunker plan for Cape+Hormuz extended routings. Prepare three re-opening scenarios (Suez-only, Hormuz-only, both). Both re-opening simultaneously would be the largest market event of 2026 — prepare rate impact analysis now.</span>
      <span class="zh"> 双重中断是 CUL 最重要的结构性宏观因素。EBS/EFS 仍然有效。为好望角+霍尔木兹延长航线安排燃油规划。准备三个重开情景（仅苏伊士、仅霍尔木兹、同时重开）。同时重开将是 2026 年最大市场事件——现在准备运价影响分析。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">Wk21, 2026 &middot; <a href="https://www.spglobal.com/market-intelligence/en/news-insights/research/2026/02/red-sea-shipping-reopens" target="_blank">S&amp;P Global</a> / <a href="https://www.seavantage.com/blog/ocean-freight-market-update-may-2026" target="_blank">Seavantage</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s9')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

  <div class="article-card" data-depts="network,trade,commercial">
    <div class="card-top">
      <div class="card-title en">Supply Chain Diversification Accelerates: Vietnam, India &amp; Bangladesh Gain Container Volume as China-US Trade Friction Reshapes Flows</div>
      <div class="card-title zh">供应链多元化加速：越南、印度和孟加拉国获益，中美贸易摩擦重塑集装箱流向</div>
      <div class="card-badges">
        <span class="badge badge-net">Network</span><span class="badge badge-trd">Trade</span>
        <span class="badge badge-com">Commercial</span><span class="badge badge-region">SE Asia / South Asia</span>
      </div>
    </div>
    <div class="card-summary en">The structural acceleration of supply chain diversification from China intensifies in May 2026. Vietnam remains primary China+1 destination (electronics, garments, furniture); India emerges for higher-value manufacturing; Bangladesh RMG absorbs textile trade flows; Thailand and Indonesia also gain share. This creates growing demand on intra-Asia feeder routes and SE Asia-origin main lane services. Market reports (goCubic, vizionapi) characterize this as a permanent structural shift reshaping container flows for the rest of the decade, regardless of Q3 tariff truce outcomes. <span style="color:#8b949e;font-size:11px;">(Source: Seavantage / goCubic / vizionapi, 2026-05)</span></div>
    <div class="card-summary zh">从中国向外的供应链多元化结构性加速在 2026 年 5 月持续强化。越南继续是主要中国+1 目的地（电子、服装、家具）；印度在高附加值制造业崭露头角；孟加拉国成衣（RMG）吸收纺织品贸易流；泰国、印尼同样获益。这在亚洲内部支线和东南亚始发主干航线上创造了增长需求。市场报告（goCubic、vizionapi）将此定性为永久性结构转变，将在本十年余下时间重塑集装箱流向，无论三季度关税休战结果如何。 <span style="color:#8b949e;font-size:11px;">(数据来源: Seavantage / goCubic / vizionapi, 2026-05)</span></div>
    <div class="card-impact"><strong>&#x1F3AF; CUL Impact（AI 解读，仅供参考）:</strong>
      <span class="en"> CUL network strategy must prioritize SE Asia and South Asia connectivity for 3-5 years. Review Vietnam-US, India-US, Bangladesh-Europe service coverage. These markets offer CUL regional expertise advantage over mega-carriers. Consider commercial investment in Ho Chi Minh City, Hai Phong, JNPT, Chennai as priority growth markets.</span>
      <span class="zh"> CUL 网络战略须在 3-5 年内优先布局东南亚和南亚连通性。审查越南—美国、印度—美国、孟加拉国—欧洲服务覆盖。这些市场是 CUL 区域专业知识优于超大型船公司的领域。考虑在胡志明市、海防港、JNPT、金奈加大商业投入作为优先增长市场。</span>
    </div>
    <div class="card-bottom">
      <span class="card-meta">May 2026 &middot; <a href="https://www.seavantage.com/blog/ocean-freight-market-update-may-2026" target="_blank">Seavantage</a> / <a href="https://www.gocubic.io/guides/market-intelligence/2026-freight-market-outlook" target="_blank">goCubic 2026 Outlook</a></span>
      <div class="card-actions"><button class="btn-sm" onclick="openSidebar('s10')">&#x1F4D6; <span class="en">Details</span><span class="zh">&#x8BE6;&#x60C5;</span></button></div>
    </div>
  </div>

</div>

<div class="data-tabs">
  <div class="tabs-nav">
    <button class="tab-btn active" onclick="switchTab('rates', this)"><span class="en">Freight Rates</span><span class="zh">运价指数</span></button>
    <button class="tab-btn" onclick="switchTab('capacity', this)"><span class="en">Capacity</span><span class="zh">运力市场</span></button>
    <button class="tab-btn" onclick="switchTab('gri', this)"><span class="en">GRI / Blank Sail</span><span class="zh">GRI / 空班</span></button>
  </div>
  <div class="tab-pane active" id="tab-rates">
    <p class="section-sub en">Data as of May 14-15, 2026. Source: Drewry WCI (May 14) / SCFI SSE (May 15) / Seavantage. All freight in USD/40ft unless noted.</p>
    <p class="section-sub zh">数据截至 2026 年 5 月 14-15 日。来源：Drewry WCI（5月14日）/ SCFI 上海航交所（5月15日）/ Seavantage。运价单位 USD/40ft，另有注明除外。</p>
    <div class="data-grid">
      <div class="data-card">
        <div class="dc-label">Drewry WCI (Global Composite)</div>
        <div class="dc-value">$2,553</div>
        <div class="dc-change up">&#x25B2; +12% WoW &mdash; 2026 Record</div>
        <div class="dc-sub">Source: Drewry, May 14, 2026</div>
      </div>
      <div class="data-card">
        <div class="dc-label">SCFI Composite (Shanghai)</div>
        <div class="dc-value">2,140.66</div>
        <div class="dc-change up">&#x25B2; +186.45 pts vs. May 8</div>
        <div class="dc-sub">First above 2,000 since Q1 &middot; SSE, May 15</div>
      </div>
      <div class="data-card">
        <div class="dc-label">TPEB Est. (SHA &rarr; US East Coast)</div>
        <div class="dc-value">~$3,800</div>
        <div class="dc-change up">&#x25B2; Front-load demand surge</div>
        <div class="dc-sub">Per FEU &middot; Seavantage May 2026</div>
      </div>
      <div class="data-card">
        <div class="dc-label">Dual Disruption Capacity Absorbed</div>
        <div class="dc-value" style="color:var(--red);">~10-12%</div>
        <div class="dc-sub">Red Sea + Hormuz &middot; Global effective fleet</div>
      </div>
    </div>
    <div class="data-table">
      <table>
        <thead><tr>
          <th><span class="en">Index / Route</span><span class="zh">指数 / 航线</span></th>
          <th><span class="en">Value (Wk21)</span><span class="zh">数值（第21周）</span></th>
          <th><span class="en">WoW Change</span><span class="zh">周环比</span></th>
          <th><span class="en">Source</span><span class="zh">来源</span></th>
        </tr></thead>
        <tbody>
          <tr><td>Drewry WCI Global Composite</td><td class="td-pr">$2,553/FEU</td><td class="td-up">&#x25B2; +12%</td><td>Drewry, May 14</td></tr>
          <tr><td>SCFI Composite</td><td class="td-pr">2,140.66 pts</td><td class="td-up">&#x25B2; +186 pts (+10.6%)</td><td>SSE, May 15</td></tr>
          <tr><td>CCFI Composite</td><td class="td-pr">1,280.46 pts</td><td class="td-up">&#x25B2; +0.1%</td><td>SSE, May 15</td></tr>
          <tr><td>TPEB (est. SHA &rarr; US East)</td><td class="td-pr">~$3,800/FEU</td><td class="td-up">&#x25B2; Surging</td><td>Seavantage est.</td></tr>
          <tr><td>WCI YoY Change</td><td class="td-pr">+14.33%</td><td class="td-up">&#x25B2; YoY</td><td>TradingEconomics</td></tr>
          <tr><td>WCI 1-Month Change</td><td class="td-pr">+13.67%</td><td class="td-up">&#x25B2; 1M</td><td>TradingEconomics</td></tr>
        </tbody>
      </table>
    </div>
    <p class="section-sub" style="margin-top:8px;"><span class="en">Note: TPEB $3,800/FEU is Seavantage estimate for US East Coast as of May 2026. SCFI sub-indices by route pending full publication. Dual disruption (Red Sea + Hormuz) structurally absorbing 10-12% effective capacity.</span><span class="zh">注：跨太平洋（美东）$3,800/FEU 为 Seavantage 2026 年 5 月估算。SCFI 分航线指数待完整发布。双重中断（红海+霍尔木兹）在结构上吸收约 10-12% 有效运力。</span></p>
  </div>
  <div class="tab-pane" id="tab-capacity">
    <p class="section-sub en">Source: Alphaliner TOP 100 (May 16, 2026). MSC reaches 1,000-vessel historic milestone (May 5-6, 2026).</p>
    <p class="section-sub zh">来源：Alphaliner TOP 100（2026 年 5 月 16 日）。MSC 达到 1,000 艘历史性里程碑（2026 年 5 月 5-6 日）。</p>
    <div class="data-table">
      <table>
        <thead><tr>
          <th>Rank</th>
          <th><span class="en">Carrier</span><span class="zh">船公司</span></th>
          <th><span class="en">Alliance</span><span class="zh">联盟</span></th>
          <th><span class="en">Est. Market Share</span><span class="zh">预估市场份额</span></th>
          <th><span class="en">Key Development Wk21</span><span class="zh">第21周重要动态</span></th>
        </tr></thead>
        <tbody>
          <tr><td>1</td><td>MSC</td><td>Independent</td><td class="td-pr">~20.4%</td><td style="color:var(--orange);font-weight:600;">&#x1F3C6; 1,000 vessels milestone (May 5); ~50% N.Europe-Med capacity</td></tr>
          <tr><td>2</td><td>Maersk</td><td>Gemini</td><td class="td-pr">~14.4%</td><td>Cape routing maintained; Gemini ops; GRI announced</td></tr>
          <tr><td>3</td><td>CMA CGM</td><td>Ocean Alliance</td><td class="td-pr">~11.9%</td><td>Ocean Alliance; Jun 1 GRI</td></tr>
          <tr><td>4</td><td>COSCO</td><td>Ocean Alliance</td><td class="td-pr">~10.8%</td><td>Ocean Alliance member</td></tr>
          <tr><td>5</td><td>Hapag-Lloyd</td><td>Gemini</td><td class="td-pr">~7.0%</td><td>24 newbuilds on order; Gemini alliance</td></tr>
          <tr><td>6</td><td>Evergreen</td><td>Ocean Alliance</td><td class="td-pr">~5%</td><td>&nbsp;</td></tr>
          <tr><td>7</td><td>ONE</td><td>Premier</td><td class="td-pr">~4%</td><td>&nbsp;</td></tr>
          <tr><td>8</td><td>HMM</td><td>Premier</td><td class="td-pr">~3%</td><td>&nbsp;</td></tr>
          <tr><td>9</td><td>Yang Ming</td><td>Premier</td><td class="td-pr">~2.5%</td><td>&nbsp;</td></tr>
          <tr><td>10</td><td>ZIM</td><td>Independent</td><td class="td-pr">~2%</td><td>&nbsp;</td></tr>
          <tr><td style="color:var(--green);font-weight:700;">45</td><td style="color:var(--green);font-weight:700;">CULines</td><td>Independent</td><td class="td-pr" style="color:var(--green);">&mdash;</td><td style="color:var(--green);">Top 5% global carrier &middot; Alphaliner TOP 100</td></tr>
        </tbody>
      </table>
    </div>
    <div class="data-grid" style="margin-top:14px;">
      <div class="data-card">
        <div class="dc-label">MSC Fleet (Historic Milestone)</div>
        <div class="dc-value">1,000+</div>
        <div class="dc-sub">Vessels &middot; First carrier globally &middot; May 5, 2026</div>
      </div>
      <div class="data-card">
        <div class="dc-label">MSC N. Europe-Med Share</div>
        <div class="dc-value">~50%</div>
        <div class="dc-sub">Capacity dominance &middot; Alphaliner May 2026</div>
      </div>
      <div class="data-card">
        <div class="dc-label">Suez Return Oversupply Risk</div>
        <div class="dc-value" style="color:var(--red);">14-15%</div>
        <div class="dc-sub">vs. current ~3-4% effective &middot; Braemar</div>
      </div>
    </div>
  </div>
  <div class="tab-pane" id="tab-gri">
    <p class="section-sub en">GRI / Blank Sailing Summary &mdash; Week 21, 2026. Source: Drewry Cancelled Sailings Tracker / Carrier Advisories.</p>
    <p class="section-sub zh">GRI / 空班汇总 &mdash;&mdash; 2026 年第 21 周。来源：Drewry 空班追踪器 / 各船公司公告。</p>
    <div class="data-table">
      <table>
        <thead><tr>
          <th><span class="en">Carrier/Program</span><span class="zh">船公司/计划</span></th>
          <th><span class="en">Trade</span><span class="zh">航线</span></th>
          <th><span class="en">Effective</span><span class="zh">生效日期</span></th>
          <th><span class="en">Status</span><span class="zh">状态</span></th>
          <th><span class="en">Source</span><span class="zh">来源</span></th>
        </tr></thead>
        <tbody>
          <tr><td>Maersk</td><td>Transpacific / Asia-Europe</td><td>Jun 1, 2026</td><td class="td-up">Announced</td><td>Carrier Advisory</td></tr>
          <tr><td>MSC</td><td>Asia-North America (EFS)</td><td>May 1, 2026</td><td class="td-up">Active</td><td>MSC Advisory (May 11)</td></tr>
          <tr><td>CMA CGM</td><td>Transpacific / Asia-Europe</td><td>Jun 1, 2026</td><td class="td-up">Announced</td><td>Carrier Advisory</td></tr>
          <tr><td>Major Alliances (all)</td><td>East-West trades</td><td>Wk21-Wk25</td><td style="color:var(--orange);font-weight:700;">30 Blank Sailings</td><td>Drewry Tracker</td></tr>
          <tr><td>Market Result</td><td>All trades</td><td>Wk21</td><td class="td-up" style="color:var(--green);">WCI +12% / SCFI +186 pts confirmed</td><td>AI&#x89E3;&#x8BFB;&#xFF0C;&#x4EC5;&#x4F9B;&#x53C2;&#x8003;</td></tr>
        </tbody>
      </table>
    </div>
    <div class="data-grid" style="margin-top:14px;">
      <div class="data-card">
        <div class="dc-label">Blank Sailings Wk21-Wk25</div>
        <div class="dc-value">30</div>
        <div class="dc-sub">Out of 698 total &middot; Drewry</div>
      </div>
      <div class="data-card">
        <div class="dc-label">Effective Capacity Reduction</div>
        <div class="dc-value">~4.3%</div>
        <div class="dc-sub">East-West trades &middot; Drewry est.</div>
      </div>
    </div>
    <p class="section-sub"><span class="en">&#x26A0; 30 blank sailings in Wk21-25 combined with tariff-truce front-load demand is the primary driver of WCI +12% surge. Monitor Drewry&#39;s weekly blank sailing releases. Zero blanks = rate softening signal.</span><span class="zh">&#x26A0; 第 21-25 周 30 个空班与关税休战前置备货需求叠加是 WCI +12% 的主要驱动因素。关注 Drewry 周度空班发布。空班清零 = 运价走软信号。</span></p>
  </div>
</div>

<div class="sidebar-overlay" id="sOverlay" onclick="closeSidebar()"></div>
<div class="sidebar" id="sidebarS1">
  <div class="sidebar-header"><h3>WCI Wk21 &mdash; $2,553 (+12%)</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Key Data</h4><ul>
      <li><strong>WCI Composite (May 14):</strong> $2,553/FEU &mdash; +12% WoW (largest gain 2026)</li>
      <li><strong>vs. Wk19:</strong> $2,286 &rarr; $2,553 (+11.7% in 2 weeks)</li>
      <li><strong>YoY:</strong> +14.33% &middot; 1-Month: +13.67%</li>
      <li><strong>Primary drivers:</strong> GRI execution + tariff front-loading + 30 blank sailings + dual disruption</li>
      <li><strong>TPEB US East Coast:</strong> ~$3,800/FEU (Seavantage est.)</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Use WCI $2,553 (+12%) as primary market reference in Q3 contract negotiations. Rate surge is real and data-confirmed. Monitor next WCI release (~May 21) for sustainability assessment.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS2">
  <div class="sidebar-header"><h3>SCFI 2,140.66 &mdash; May 15</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Key Data</h4><ul>
      <li><strong>SCFI (May 15):</strong> 2,140.66 pts &mdash; +186.45 pts vs. May 8</li>
      <li><strong>CCFI (May 15):</strong> 1,280.46 pts &mdash; only +0.1% WoW</li>
      <li><strong>Divergence:</strong> SCFI +10.6% vs. CCFI +0.1% = large spot premium</li>
      <li><strong>First above 2,000:</strong> Since Q1 2026</li>
      <li><strong>Risk:</strong> Post-mid-July demand cliff if truce expires</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> SCFI above 2,000 is significant. Accelerate Q3 contract renewals. Spot/contract divergence is urgent signal. Monitor next Friday SSE release.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS3">
  <div class="sidebar-header"><h3>Hormuz Crisis &mdash; Wk21</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Status</h4><ul>
      <li><strong>Status:</strong> Effectively closed since Operation Epic Fury (Mar 2026)</li>
      <li><strong>Iran:</strong> Limiting transits, imposing controlled routing &amp; tolls</li>
      <li><strong>Container carriers:</strong> All suspended Hormuz; rerouting Africa</li>
      <li><strong>Combined disruption:</strong> Red Sea + Hormuz absorbing ~10-12% global capacity</li>
      <li><strong>War risk:</strong> Extremely elevated for Gulf/Hormuz transits</li>
      <li><strong>Ceasefire:</strong> Fragile, no material improvement since Apr 12</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Review all Middle East service schedules. Verify EBS/EFS validity. Prepare Hormuz re-opening rate scenarios. Check war risk insurance coverage for chartered vessels.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS4">
  <div class="sidebar-header"><h3>MSC 1,000 Vessels Milestone</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Milestone Details</h4><ul>
      <li><strong>Achievement:</strong> 1,000 vessels operated &mdash; world first</li>
      <li><strong>Confirmed:</strong> Alphaliner TOP 100, May 5-6, 2026</li>
      <li><strong>Market share:</strong> ~20.4% global &middot; ~50% N.Europe-Med</li>
      <li><strong>On order:</strong> ~132 vessels (ongoing expansion)</li>
      <li><strong>NOO activity:</strong> Rising &mdash; more charter market availability</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Update competitive intelligence with May 16 TOP 100 data. Assess MSC Med dominance impact on CUL strategy. NOO rise = more spot charter options for chartering team.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS5">
  <div class="sidebar-header"><h3>Blank Sailings Wk21-25</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Blank Sailing Data</h4><ul>
      <li><strong>Total (Wk21-Wk25):</strong> 30 out of 698 scheduled sailings</li>
      <li><strong>Capacity reduction:</strong> ~4.3% on East-West trades</li>
      <li><strong>Period:</strong> May 18 &ndash; June 21, 2026</li>
      <li><strong>Purpose:</strong> Support GRI execution + manage demand normalization</li>
      <li><strong>Impact:</strong> Key contributor to WCI +12% / SCFI +186 pts</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Leverage tight capacity window Wk21-25 for rate execution. Communicate space constraints to customers. Track Drewry blank sailing data weekly as rate leading indicator.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS6">
  <div class="sidebar-header"><h3>Qingdao Port Congestion Wk21</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Port Status</h4><ul>
      <li><strong>Delay:</strong> ~4 days (Seavantage, May 2026)</li>
      <li><strong>Cause:</strong> Tariff-truce front-loading on TPEB</li>
      <li><strong>April volumes:</strong> 731,476 TEU exports; 178,646 TEU imports (EconDB)</li>
      <li><strong>Also affected:</strong> Tianjin, Dalian</li>
      <li><strong>Duration:</strong> Expected through June 2026</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Adjust TPEB cargo cutoffs at Qingdao. Issue proactive customer advisories for potential rollovers. Pre-position empty equipment. Monitor Tianjin overflow.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS7">
  <div class="sidebar-header"><h3>Alphaliner TOP 100 &mdash; May 16</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Key Highlights</h4><ul>
      <li><strong>MSC:</strong> 1,000-vessel milestone; ~50% N.Europe-Med capacity</li>
      <li><strong>N. Europe-Med carriers:</strong> Only 10 total; 7 MLOs dominate</li>
      <li><strong>NOO activity:</strong> Rising (charter market + S&amp;H activity growth)</li>
      <li><strong>CULines:</strong> #45 (Top 5% global)</li>
      <li><strong>Global fleet:</strong> ~33.8M+ TEU fully cellular</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Update competitive intelligence with May 16 TOP 100 data. CUL #45 is a key marketing asset. NOO rise means more tonnage availability in spot charter market for CUL.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS8">
  <div class="sidebar-header"><h3>Tariff Truce Front-Loading Peak</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Key Data</h4><ul>
      <li><strong>Truce period:</strong> Apr 22 &ndash; mid-Jul 2026 (90 days)</li>
      <li><strong>US tariff on China:</strong> ~33% trade-weighted (MS Advisory)</li>
      <li><strong>TPEB US East Coast:</strong> ~$3,800/FEU (Seavantage est.)</li>
      <li><strong>Phase:</strong> Peak front-loading NOW in Wk21</li>
      <li><strong>Qingdao:</strong> ~4-day delay from booking surge</li>
      <li><strong>Post-Jul demand cliff:</strong> Risk if truce expires without extension</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Maximize TPEB utilization next 6-8 weeks. Prepare July demand cliff contingency plans. Q3 scenario planning (extension vs. expiry) is critical. Monitor US-China trade talks for July outcome signals.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS9">
  <div class="sidebar-header"><h3>Red Sea + Dual Disruption Wk21</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Status Summary</h4><ul>
      <li><strong>Red Sea:</strong> Cape routing maintained by all majors</li>
      <li><strong>Houthi:</strong> Active threats, vessel-specific targeting renewed</li>
      <li><strong>Hormuz:</strong> Effectively closed since March 2026</li>
      <li><strong>Combined absorption:</strong> ~10-12% global effective capacity</li>
      <li><strong>Rate impact:</strong> Structural support for WCI/SCFI at elevated levels</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Model dual disruption scenario impacts on network utilization. EBS/EFS valid. Bunker plan for extended routings. Three re-opening scenarios: Suez-only, Hormuz-only, both simultaneously.</p></div>
  </div>
</div>
<div class="sidebar" id="sidebarS10">
  <div class="sidebar-header"><h3>SE Asia Supply Chain Shift</h3><button class="sidebar-close" onclick="closeSidebar()">&#x2715;</button></div>
  <div class="sidebar-body">
    <div class="sidebar-section"><h4>Key Trends</h4><ul>
      <li><strong>Vietnam:</strong> Primary China+1 destination (electronics, garments, furniture)</li>
      <li><strong>India:</strong> Higher-value manufacturing, rapidly growing US exports</li>
      <li><strong>Bangladesh:</strong> RMG/textile trade flow absorption</li>
      <li><strong>Thailand/Indonesia:</strong> Gaining manufacturing share</li>
      <li><strong>Nature:</strong> Structural/permanent shift (goCubic, vizionapi 2026)</li>
      <li><strong>CUL opportunity:</strong> Intra-Asia feeder + SE Asia-origin main lane growth</li>
    </ul></div>
    <div class="impact-box"><p><strong>CUL Action:</strong> Prioritize SE Asia/South Asia network in 3-5 year strategy. Review Vietnam-US, India-US, Bangladesh-Europe service coverage. Target Ho Chi Minh, Hai Phong, JNPT, Chennai as growth markets.</p></div>
  </div>
</div>

<footer>
  <p>&#x1F99E; <strong>CUL &mdash; Alphaliner Intelligence Digest</strong> | Week 21, 2026 (May 12 &ndash; May 18)</p>
  <p style="margin-top:6px;">Generated: May 18, 2026 09:00 CST &middot; Powered by Claw Intelligence System &middot; <a href="https://LeahLiuL.github.io/Claw-Report/reports/alphaliner-digest.html" target="_blank">View Online</a></p>
  <p style="margin-top:4px;">Data Sources: <a href="https://www.drewry.co.uk/supply-chain-advisors/supply-chain-expertise/world-container-index-assessed-by-drewry" target="_blank">Drewry WCI</a> | <a href="https://alphaliner.axsmarine.com/PublicTop100/" target="_blank">Alphaliner TOP100</a> | <a href="https://gmteight.com/flash/detail/1366490" target="_blank">GMT Eight (SCFI)</a> | <a href="https://www.sse.net.cn" target="_blank">SSE (SCFI)</a> | <a href="https://www.seavantage.com/blog/ocean-freight-market-update-may-2026" target="_blank">Seavantage</a> | <a href="https://www.drewry.co.uk/supply-chain-advisors/supply-chain-expertise/cancelled-sailings-tracker" target="_blank">Drewry Blank Sailings</a></p>
  <p style="margin-top:4px; color: #6e7681;">&#x26A0;&#xFE0F; AI &#x5206;&#x6790;&#x9879;&#x76EE;&#x6807;&#x6CE8;&#x300C;&#xFF08;AI &#x89E3;&#x8BFB;&#xFF0C;&#x4EC5;&#x4F9B;&#x53C2;&#x8003;&#xFF09;&#x300D; | &#x8FD0;&#x529B;&#x6392;&#x540D;&#x6570;&#x636E;&#x6765;&#x6E90; Alphaliner | &#x8FD0;&#x4EF7;&#x6570;&#x636E;&#x6765;&#x6E90; Drewry WCI / SCFI&#xFF08;&#x4E0A;&#x6D77;&#x822A;&#x4EA4;&#x6240;&#xFF09; | &#x6700;&#x65B0;&#x6570;&#x636E;&#x622A;&#x81F3; 2026&#x5E745&#x6708;15&#x65E5;</p>
</footer>

<script>
function setLang(l) {
  document.getElementById('pageBody').className = l === 'zh' ? 'lang-zh' : '';
  document.getElementById('btnEN').className = 'lang-btn' + (l === 'en' ? ' active' : '');
  document.getElementById('btnZH').className = 'lang-btn' + (l === 'zh' ? ' active' : '');
}
function switchTab(id, btn) {
  document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
  btn.classList.add('active');
}
function openSidebar(id) {
  closeSidebar();
  var map = {s1:'sidebarS1',s2:'sidebarS2',s3:'sidebarS3',s4:'sidebarS4',s5:'sidebarS5',s6:'sidebarS6',s7:'sidebarS7',s8:'sidebarS8',s9:'sidebarS9',s10:'sidebarS10'};
  var el = document.getElementById(map[id]);
  if (el) { el.classList.add('open'); document.getElementById('sOverlay').classList.add('open'); }
}
function closeSidebar() {
  document.querySelectorAll('.sidebar').forEach(s => s.classList.remove('open'));
  document.getElementById('sOverlay').classList.remove('open');
}
document.querySelectorAll('.ftag').forEach(tag => {
  tag.addEventListener('click', function() {
    document.querySelectorAll('.ftag').forEach(t => t.classList.remove('active'));
    this.classList.add('active');
    var dept = this.dataset.dept;
    var cards = document.querySelectorAll('.article-card');
    var count = 0;
    cards.forEach(c => {
      var depts = c.dataset.depts || '';
      if (dept === 'all' || depts.split(',').indexOf(dept) !== -1) { c.style.display = ''; count++; }
      else { c.style.display = 'none'; }
    });
    document.getElementById('countEN').textContent = count;
    document.getElementById('countZH').textContent = count;
  });
});
</script>
</body>
</html>"""

with open(r'C:\\Users\\culadmin\\Claw-Report\\reports\\alphaliner-digest.html', 'w', encoding='utf-8') as f:
    f.write(content)
print("File written successfully!")
