@echo off
REM ============================================================
REM  ROB FAST daily update  (scheduled task entry, runs 01:00 + 13:00)
REM
REM  Step 1: original crew report emails (rob_mail_sync)  = primary source
REM          (the 1data.info summary emails miss many reports)
REM  Step 2: pre-built vessel management summary emails    = supplement
REM  Step 3: rebuild page (safe: preserves rob_draft_* fields)
REM  Step 4: safe commit + push
REM
REM  ASCII-ONLY ON PURPOSE.  UTF-8 Chinese inside a .bat makes cmd.exe
REM  mis-parse the file (it reads bytes as GBK) and the whole script dies
REM  on line 1 without even writing the log.  Never add Chinese here.
REM ============================================================
cd /d %~dp0
set LOG=rob_data\fastupdate.log
set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe
if not exist "%PYEXE%" set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe

echo ===== %date% %time% ===== >> %LOG%

REM ---- [1/4] git operation mutex lock ----
"%PYEXE%" git_op_lock.py acquire --owner rob_fast --wait 300 >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] git op lock held by another job, skip this run >> %LOG%
    exit /b 1
)

REM ---- [2/4] git fetch ----
echo [2/4] git fetch ... >> %LOG%
git fetch origin >> %LOG% 2>&1
if errorlevel 1 echo [WARN] git fetch failed, reconcile via safe-push >> %LOG%

REM ---- [3/4a] original crew report mail sync (primary) ----
echo [3/4] rob_mail_sync (original crew reports) ... >> %LOG%
"%PYEXE%" rob_mail_sync.py --days 3 --apply >> %LOG% 2>&1
if errorlevel 1 echo [WARN] rob_mail_sync failed, continue with summary sync >> %LOG%

REM ---- [3/4b] summary email sync (supplement) ----
echo [3/4b] rob_summary_sync ... >> %LOG%
"%PYEXE%" rob_summary_sync.py --days 4 --apply >> %LOG% 2>&1
if errorlevel 1 echo [WARN] rob_summary_sync failed, continue >> %LOG%

REM ---- [3/4c] safe rebuild ----
"%PYEXE%" rob_fast_rebuild.py >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] rob_fast_rebuild failed >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner rob_fast >> %LOG% 2>&1
    exit /b 1
)

REM ---- [4/4] safe push ----
echo [4/4] safe commit and push ... >> %LOG%
set CLAW_LOCK_OWNER=rob_fast
"%PYEXE%" git_safe_push.py -m "ROB fast update %date%" >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] git_safe_push failed >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner rob_fast >> %LOG% 2>&1
    exit /b 1
)

echo DONE >> %LOG%
"%PYEXE%" git_op_lock.py release --owner rob_fast >> %LOG% 2>&1
exit /b 0
