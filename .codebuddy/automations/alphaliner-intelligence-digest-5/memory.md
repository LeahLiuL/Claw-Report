# Alphaliner Intelligence Digest - Automation Memory

## Execution History

### WK25 - 2026-06-18 (Thursday 08:30)
- **Status:** ✅ Complete
- **Outlook Fetch:** 2 emails matched (WK21 May 27, WK20 May 22). Latest = WK21 (25 pages, 95,937 chars)
- **Note:** No new Alphaliner newsletters received since WK21. Outlook search covers last 60 days. WK22-25 not yet received.
- **Public Data Sources Used:**
  - SCFI: 2,985.22 (+9.5%, Jun 12) — 6th consecutive weekly rise
  - Drewry WCI: $3,549/40ft (+3%, Jun 11)
  - Ship&Bunker/HandyBulk: VLSFO Singapore $652/MT (-4.82%), Rotterdam $577/MT (-5.64%), Jun 17
  - Alphaliner TOP 100: 7,547 ships / 34.26M TEU (Jun 17)
  - Linerlytica/Xeneta: Peak season surge, Iran ceasefire MoU signed Jun 14
- **Key Market Event:** US-Iran 60-day ceasefire MoU signed Jun 14 — Hormuz reopening imminent (formal signing Jun 19)
- **Report Generated:** alphaliner-intelligence-digest-wk25-2026.html (14 articles, 5 CUL Direct, 3 Featured)
- **GitHub Push:** ✅ c294b34 → main (2 files changed, 952 insertions, 564 deletions)
- **Email Sent:** ✅ to leahliu@culines.com, Subject: "[Claw] Alphaliner Intelligence Digest - Week 25 2026"
- **Script Modified:** send_alphaliner_via_win32com.py — changed week calculation from (iso_week - 1) to iso_week (Thursday execution covers current week)

### WK25 - 2026-06-16 (Monday, prior run)
- Report generated with same Alphaliner WK21 data + public sources as of Jun 12-15

### WK24 - 2026-06-11 (prior run)
- Report generated with Alphaliner WK21 data + public sources as of Jun 5-11

## Configuration
- **Automation ID:** alphaliner-intelligence-digest-5
- **Schedule:** FREQ=WEEKLY;BYDAY=TH;BYHOUR=8;BYMINUTE=30
- **Fetch Script:** C:\Users\culadmin\WorkBuddy\Claw\fetch_alphaliner_email.py
- **Email Script:** C:\Users\culadmin\WorkBuddy\Claw\automation\send_alphaliner_via_win32com.py
- **Report Path:** C:\Users\culadmin\Claw-Report\reports\
- **Fixed Link:** alphaliner-intelligence-digest.html (always latest)
- **GitHub:** https://github.com/LeahLiuL/Claw-Report.git
- **GitHub Pages:** https://LeahLiuL.github.io/Claw-Report/reports/alphaliner-intelligence-digest.html

## Data Freshness Notes
- Alphaliner newsletters arrive via Outlook from Mabel Ong (CUL-SG/NPD)
- Typical cadence: Weekly, usually Tuesday/Wednesday
- WK21 (May 27) is the latest received as of Jun 18 — 3 weeks gap
- When new newsletters arrive, fetch script will pick them up automatically (--last 2 flag)
- Report always notes "基于最近可用 Alphaliner 数据 (WKxx, date)" when data is stale
