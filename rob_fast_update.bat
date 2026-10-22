@echo off
REM ============================================================
REM  ROB FAST daily update  (scheduled task entry)
REM  Parses the pre-built "[中联航运-船舶管理报告]-报表汇总" emails.
REM  ~20-40 seconds per run (vs 8-13 min for the full mailbox scan).
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

REM ---- [3/4] summary sync + rebuild ----
echo [3/4] rob_summary_sync + fast rebuild ... >> %LOG%
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
