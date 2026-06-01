# Alphaliner Intelligence Digest — Automation Memory

## Execution History

### 2026-05-25 (Week 22)
- **Status**: SUCCESS
- **Report**: `alphaliner-intelligence-digest-wk22-2026.html`
- **Path**: `C:\Users\culadmin\Claw-Report\reports\alphaliner-intelligence-digest-wk22-2026.html`
- **GitHub**: Committed & pushed to LeahLiuL/Claw-Report main branch (commit: 6785150)
- **GitHub Pages URL**: https://LeahLiuL.github.io/Claw-Report/reports/alphaliner-intelligence-digest-wk22-2026.html
- **Email**: Sent to leahliu@culines.com via AgentMail (subject: [Claw] Alphaliner Intelligence Digest - Week 22 2026)
- **Email Script**: `C:\Users\culadmin\WorkBuddy\Claw\send_alphaliner_wk22_email.ps1`
- **Data Sources**: Drewry WCI (6% surge to $2,712), OilMonster (VLSFO $873/MT), public market data
- **Outlook inbox search**: MATON_API_KEY not configured — used public news search instead
- **Articles**: 10 articles, 6 CUL-direct, 3 featured
- **Key themes**: WCI +6% (3rd consecutive weekly rise), PGSA Hormuz transit authority, OOCL WISDOM 24K TEU methanol ship, June GRI collective announcement, 41 blank sailings, Singapore VLSFO $873/MT

### 2026-05-18 (Week 20)
- **Status**: SUCCESS
- **Report**: `alphaliner-intelligence-digest-wk20-2026.html`
- **Path**: `C:\Users\culadmin\Claw-Report\reports\alphaliner-intelligence-digest-wk20-2026.html`
- **GitHub**: Committed & pushed to LeahLiuL/Claw-Report main branch (commit: 19a2525)
- **GitHub Pages URL**: https://LeahLiuL.github.io/Claw-Report/reports/alphaliner-intelligence-digest-wk20-2026.html
- **Email**: Sent to leahliu@culines.com via AgentMail (subject: [Claw] Alphaliner Intelligence Digest - Week 20 2026)
- **Email Script**: `C:\Users\culadmin\WorkBuddy\Claw\send_alphaliner_wk20_email.ps1`
- **Data Sources**: Drewry WCI (12% surge to $2,553), OilMonster (VLSFO $843), Alphaliner/MarineLink/SeaVantage public data
- **Outlook inbox search**: 401 Unauthorized (MATON_API_KEY not configured in env) — used public news search instead
- **Articles**: 10 articles, 6 CUL-direct, 3 featured
- **Key themes**: WCI +12% surge, Hormuz Week 12, PSS/EFS stacking, Qingdao congestion, VLSFO crisis high

### Notes for future runs
- Outlook IMAP search requires node.js (not installed) — use Outlook skill (needs MATON_API_KEY)
- AgentMail API key: `am_us_fe0842c00141bdc1b872d0b9a1dda62163b350b6d3bb522e7cf3ba806237812b`
- Email script pattern: `send_alphaliner_wkXX_email.ps1` in `C:\Users\culadmin\WorkBuddy\Claw\`
- Run PS1 with: `powershell -ExecutionPolicy Bypass -File "...ps1"`
- If Alphaliner email not found, use public news search (Drewry WCI, Hellenic Shipping News, Container-News.com)
- Shared folder search (Z:\04 上海操作中心) may timeout — do not wait more than 30s

### 2026-06-01 (Week 23)
- **Status**: SUCCESS
- **Report**: `alphaliner-intelligence-digest-wk23-2026.html`
- **Path**: `C:\Users\culadmin\Claw-Report\reports\alphaliner-intelligence-digest-wk23-2026.html`
- **GitHub**: Committed & pushed to LeahLiuL/Claw-Report main branch (commit: 59e2e4f)
- **GitHub Pages URL**: https://LeahLiuL.github.io/Claw-Report/reports/alphaliner-intelligence-digest-wk23-2026.html
- **Email**: Sent to leahliu@culines.com via AgentMail (subject: [Claw] Alphaliner Intelligence Digest - Week 23 2026)
- **Email Script**: `C:\Users\culadmin\WorkBuddy\Claw\send_alphaliner_wk23_email.ps1`
- **Data Sources**: Drewry WCI (3% rise to $2,800/FEU, 4th consecutive), SCFI PGSA $4,462/TEU, OilMonster VLSFO $782/MT, Alphaliner fleet data, IRGC public statement
- **Outlook inbox search**: MATON_API_KEY confirmed invalid (401) — used public news search instead
- **Articles**: 10 articles, 6 CUL-direct, 3 featured
- **Key themes**: IRGC "Full Control" Hormuz (May 30), WCI $2,800 +3% (4th consecutive rise), Early peak season erupts (1-2 months early), Bulk carriers being converted to boxships, Gulf rates $4,462 4-week rise, India West Coast +50% equipment crisis, Brent drops to $91 on ceasefire extension

### Notes update (2026-06-01)
- AgentMail API key `am_us_fe0842c00141bdc1b872d0b9a1dda62163b350b6d3bb522e7cf3ba806237812b` is SEPARATE from MATON — use `https://api.agentmail.to/v0/inboxes/` (not MATON gateway)
- MATON key (same value) is 401 unauthorized — MATON account may be expired/different service
- AgentMail inbox: `claw-report-leah@agentmail.to` — confirmed working for sending

### Prior runs
- 2026-05-14: Week 19 — SUCCESS (Newsletter No.19 PDF integrated from shared folder)
