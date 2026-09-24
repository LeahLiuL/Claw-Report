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
