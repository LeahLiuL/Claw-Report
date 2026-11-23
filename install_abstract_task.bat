@echo off
REM ============================================================
REM  One-time setup for Abstract Log daily auto-update (run once on culadmin)
REM  Registers ONE scheduled task: daily 06:00
REM  (avoids ROB tasks at 01:00 / 13:00)
REM
REM  What it does daily:
REM    Outlook vesselreport -> abstract_log.db -> GitHub Release
REM    -> DATABASES.html -> git push
REM ============================================================
echo Registering scheduled task: ClawReport_Abstract_0600 (daily 06:00) ...
schtasks /Create /TN "ClawReport_Abstract_0600" /TR "\"%~dp0abstract_daily_update.bat\"" /SC DAILY /ST 06:00 /F

echo.
echo ============================================================
echo Setup done. Task: ClawReport_Abstract_0600
echo Test now:  schtasks /Run /TN "ClawReport_Abstract_0600"
echo Check log: rob_data\abstract_daily.log
echo DB download:
echo   https://github.com/LeahLiuL/Claw-Report/releases/download/db-snapshot-2026-11-23/abstract_log.db
echo Notice page:
echo   https://leahliul.github.io/Claw-Report/DATABASES.html
echo ============================================================
pause
