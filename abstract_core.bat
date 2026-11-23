@echo off
REM ============================================================
REM  Abstract Log daily CORE step - called from ROB bat, NOT scheduled itself.
REM  Does NOT touch git (the caller's git_safe_push pushes DATABASES.html).
REM  Steps: 1) Outlook collect -> abstract_log.db
REM         2) checkpoint WAL + upload db to GitHub Release
REM         3) regenerate DATABASES.html (tracked -> picked up by caller push)
REM  ASCII-ONLY ON PURPOSE.
REM ============================================================
cd /d %~dp0
set LOG=rob_data\abstract_daily.log
set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe
if not exist "%PYEXE%" set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe

echo ----- abstract_core %date% %time% ----- >> %LOG%

REM [1] collect + parse (exit 2 = outlook cache not synced -> skip gracefully)
"%PYEXE%" abstract_daily_sync.py --days 5 --apply >> %LOG% 2>&1
if errorlevel 2 (
    echo [abstract][SKIP] outlook cache not synced - next run backfills >> %LOG%
    exit /b 0
)
if errorlevel 1 echo [abstract][WARN] collect failed, continue >> %LOG%

REM [2] upload db (overwrites old asset; reports.db untouched)
"%PYEXE%" upload_abstract_db.py --apply >> %LOG% 2>&1
if errorlevel 1 echo [abstract][WARN] upload failed, next run retries >> %LOG%

REM [3] regenerate notice page (tracked file; caller push will include it)
"%PYEXE%" databases_page_gen.py >> %LOG% 2>&1
if errorlevel 1 echo [abstract][WARN] page gen failed, continue >> %LOG%

echo [abstract] CORE DONE >> %LOG%
exit /b 0
