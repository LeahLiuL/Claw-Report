@echo off
REM ============================================================
REM  ROB daily auto-update entry (run by scheduled tasks 13:00 & 01:00)
REM  Unattended: no pause; all output -> rob_data\autoupdate.log
REM ============================================================
cd /d %~dp0
set LOG=rob_data\autoupdate.log
echo ===== %date% %time% ===== >> %LOG%

REM 解释器: 统一管理版 venv(含 win32com / openpyxl / pycryptodome)。
set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe
if not exist "%PYEXE%" set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe

REM ---- [0/5] report archive (harvest) ----
REM Report-level archive: one line per report instead of one row per vessel-day.
REM Daily refresh keeps only the latest report of the day, dropping BERTH/SAILING,
REM which are exactly what voyage fuel consumption needs for its start/end bounds.
REM Runs BEFORE the git lock on purpose: it never touches git, and putting it inside
REM the lock would stretch lock hold time by ~8 min and starve other scheduled jobs.
echo [0/5] harvest_reports.py ^(report level archive, incremental^) ...
if exist harvest_reports.py (
    "%PYEXE%" harvest_reports.py --days 3 >> %LOG% 2>&1
    if errorlevel 1 echo [WARN] harvest_reports failed - keep existing archive >> %LOG%
) else (
    echo [WARN] harvest_reports.py not found, skip >> %LOG%
)

REM ---- [0/5] git 操作互斥锁 ----
REM 本机 7 个计划任务并发操作同一仓库, 2026-09-22/09-24 两次 .git 对象库损坏。
REM 全程持锁(含刷新的 6-8 分钟), 他人持锁最多等 5 分钟, 等不到跳过本轮(下轮重试)。
"%PYEXE%" git_op_lock.py acquire --owner rob_update --wait 300 >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] git op lock held by another job, skip this run ^(next scheduled run will retry^) >> %LOG%
    exit /b 1
)

REM ---- [1/5] git fetch (不再 pull --rebase: rebase 在本机两次损坏仓库) ----
REM 工作树对齐交给末尾 git_safe_push.py 的安全协议(commit->ls-remote->对齐->push),
REM 生成前只 fetch 不动工作树, 避免与其它自动化抢工作树。
echo [1/5] git fetch ...
git fetch origin >> %LOG% 2>&1
if errorlevel 1 ( echo [WARN] git fetch failed, will reconcile via safe-push >> %LOG% )

REM ---- [1b/5] sync templates from FETCH_HEAD ----
REM Why: `git fetch` moves FETCH_HEAD only, NOT the working tree. Before this
REM step every run rendered with the template synced at the END of the PREVIOUS
REM run, so a pushed code change always took effect one cycle late.
REM sync_templates.py writes ONLY templates/ (path whitelist + HTML validation +
REM backup + atomic replace). It never touches rob_data/, *.csv or *.json.
REM Bootstrap: on a machine that has not synced this script yet, pull the script
REM itself straight out of FETCH_HEAD without touching the git index.
echo [1b/5] sync_templates.py ^(templates/ + voyage_consumption.py, never data^) ...
REM Refresh the sync tool itself from FETCH_HEAD EVERY run (was: only if missing).
REM Otherwise a machine still holding the first copy would never learn about a
REM newly whitelisted file (voyage_consumption.py, added 2026-11-12) and would
REM keep rendering voyage consumption with the previous cycle's algorithm.
REM Write to .new then move: a failed cat-file can never truncate the script.
git cat-file -e FETCH_HEAD:sync_templates.py 2>nul && git cat-file blob FETCH_HEAD:sync_templates.py > sync_templates.py.new 2>nul && move /y sync_templates.py.new sync_templates.py >nul 2>&1
if exist sync_templates.py.new del /q sync_templates.py.new >nul 2>&1
if exist sync_templates.py (
    "%PYEXE%" sync_templates.py >> %LOG% 2>&1
    if errorlevel 1 echo [WARN] template sync failed - keep current templates, continue >> %LOG%
) else (
    echo [WARN] sync_templates.py not available, keep current templates >> %LOG%
)

echo [2/5] sync_bunkering.py ^(加油量, 仅体积, 价格不上网^) ...
if exist sync_bunkering.py (
    "%PYEXE%" sync_bunkering.py >> %LOG% 2>&1
    if errorlevel 1 (
        echo [WARN] sync_bunkering failed - 网络盘未挂载或 openpyxl 缺失, 沿用上次 bunkering.json >> %LOG%
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
