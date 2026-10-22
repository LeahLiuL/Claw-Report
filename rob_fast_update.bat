@echo off
REM ============================================================
REM  ROB FAST daily update  (scheduled task entry)
REM  Step 1: original crew report emails (rob_mail_sync) -- primary source,
REM          the 1data.info summary misses a large share of reports.
REM  Step 2: pre-built "[中联航运-船舶管理报告]-报表汇总" emails (fast, ~15s).
REM  Step 3: rebuild page (safe: preserves rob_draft_* fields).
REM  ASCII-only: UTF-8 Chinese + chcp crashes cmd.
REM ============================================================
cd /d %~dp0
set LOG=rob_data\fastupdate.log
set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe
if not exist "%PYEXE%" set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe

echo ===== %date% %time% ===== >> %LOG%

REM ---- [1/4] git 操作互斥锁 ----
"%PYEXE%" git_op_lock.py acquire --owner rob_fast --wait 300 >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] git op lock held by another job, skip this run >> %LOG%
    exit /b 1
)

REM ---- [2/4] git fetch ----
echo [2/4] git fetch ... >> %LOG%
git fetch origin >> %LOG% 2>&1
if errorlevel 1 ( echo [WARN] git fetch failed, will reconcile via safe-push >> %LOG% )

REM ---- [3/5] original crew report mail sync ----
echo [3/5] rob_mail_sync (original reports) ... >> %LOG%
"%PYEXE%" rob_mail_sync.py --days 3 --apply >> %LOG% 2>&1
if errorlevel 1 (
    echo [WARN] rob_mail_sync failed, continue with summary sync >> %LOG%
)

REM ---- [4/5] summary sync + rebuild ----
echo [4/5] rob_summary_sync + fast rebuild ... >> %LOG%
"%PYEXE%" rob_summary_sync.py --days 4 --apply >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] rob_summary_sync failed >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner rob_fast >> %LOG% 2>&1
    exit /b 1
)
"%PYEXE%" rob_fast_rebuild.py >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] rob_fast_rebuild failed >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner rob_fast >> %LOG% 2>&1
    exit /b 1
)

REM ---- [5/5] safe push ----
echo [5/5] safe commit and push ... >> %LOG%
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
