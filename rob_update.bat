@echo off
REM ============================================================
REM  ROB FULL deep-scan update entry (manual / fallback use).
REM  The daily 01:00 + 13:00 scheduled tasks use rob_fast_update.bat instead.
REM  Unattended: no pause; all output -> rob_data\autoupdate.log
REM
REM  ASCII-ONLY ON PURPOSE.  UTF-8 Chinese in a .bat makes cmd.exe mis-parse
REM  the file (it reads bytes as GBK) and the script dies on line 1 without
REM  writing any log.  That silently broke the daily task for 4 days
REM  (2026-10-23..26).  NEVER add Chinese here.  Put Chinese in .py files.
REM ============================================================
cd /d %~dp0
set LOG=rob_data\autoupdate.log
echo ===== %date% %time% ===== >> %LOG%

REM interpreter: managed venv (win32com / openpyxl / pycryptodome)
set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe
if not exist "%PYEXE%" set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe

REM ---- [0/5] git operation mutex lock ----
REM 7 scheduled tasks touch this repo; the .git object store was corrupted
REM twice (2026-09-22 / 09-24). Hold the lock for the whole run (6-8 min).
REM If someone else holds it, wait up to 5 min, else skip (next run retries).
"%PYEXE%" git_op_lock.py acquire --owner rob_update --wait 300 >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] git op lock held by another job, skip this run ^(next scheduled run will retry^) >> %LOG%
    exit /b 1
)

REM ---- [1/5] git fetch (never pull --rebase: rebase corrupted this repo twice) ----
REM Worktree alignment is handled by git_safe_push.py at the end
REM (commit -> ls-remote -> align -> push). Here we only fetch, never touch
REM the worktree, to avoid racing the other automations.
echo [1/5] git fetch ...
git fetch origin >> %LOG% 2>&1
if errorlevel 1 ( echo [WARN] git fetch failed, will reconcile via safe-push >> %LOG% )

echo [2/5] sync_bunkering.py ^(bunkering VOLUME only - prices never leave this PC^) ...
if exist sync_bunkering.py (
    "%PYEXE%" sync_bunkering.py >> %LOG% 2>&1
    if errorlevel 1 (
        echo [WARN] sync_bunkering failed - network drive not mounted or openpyxl missing, keeping last bunkering.json >> %LOG%
    )
) else (
    echo [WARN] sync_bunkering.py not found, skip >> %LOG%
)

echo [3/5] rob_refresh.py ...
"%PYEXE%" rob_refresh.py >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] rob_refresh.py failed >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner rob_update >> %LOG% 2>&1
    exit /b 1
)

echo [4/5] safe commit and push ^(git_safe_push.py^) ...
set CLAW_LOCK_OWNER=rob_update
"%PYEXE%" git_safe_push.py -m "ROB auto update %date%" >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] git_safe_push failed >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner rob_update >> %LOG% 2>&1
    exit /b 1
)

echo [5/5] DONE >> %LOG%
"%PYEXE%" git_op_lock.py release --owner rob_update >> %LOG% 2>&1
exit /b 0
