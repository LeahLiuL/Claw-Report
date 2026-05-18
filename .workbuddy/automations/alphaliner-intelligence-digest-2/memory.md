# Alphaliner Intelligence Digest — Automation Memory

## Execution History

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
- AgentMail API key: in `C:\Users\culadmin\WorkBuddy\Claw\send_alphaliner_wk19_email.ps1` (line 5)
- Email template for weekly: create `send_alphaliner_wkXX_email.ps1` based on wk19 template
- If Alphaliner email not found, use public news search (Drewry WCI, MarineLink, SeaVantage, etc.)
- Shared folder search (Z:\04 上海操作中心) may timeout — do not wait more than 30s

### Prior runs
- 2026-05-14: Week 19 — SUCCESS (Newsletter No.19 PDF integrated from shared folder)
