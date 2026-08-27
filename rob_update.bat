@echo off
REM ============================================================
REM  ROB daily auto-update entry (run by scheduled tasks 13:00 & 01:00)
REM  Unattended: no pause; all output -> rob_data\autoupdate.log
REM ============================================================
cd /d %~dp0
set LOG=rob_data\autoupdate.log
echo ===== %date% %time% ===== >> %LOG%

echo [1/4] git pull ...
git pull --rebase --autostash >> %LOG% 2>&1
if errorlevel 1 ( echo [WARN] git pull failed, will reconcile via safe-push >> %LOG% )

echo [2/4] rob_refresh.py ...
"C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12.old.14596\python.exe" rob_refresh.py >> %LOG% 2>&1
if errorlevel 1 ( echo [ERROR] rob_refresh.py failed >> %LOG% & exit /b 1 )

echo [3/4] safe commit and push (git_safe_push.py) ...
"C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe" git_safe_push.py -m "ROB auto update %date%" >> %LOG% 2>&1
if errorlevel 1 ( echo [ERROR] git_safe_push failed >> %LOG% & exit /b 1 )

echo [4/4] DONE >> %LOG%
exit /b 0
