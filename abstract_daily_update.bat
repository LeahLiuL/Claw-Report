@echo off
REM ============================================================
REM  Abstract Log daily update (scheduled task entry)
REM
REM  Step 1: collect Abstract Log mails from Outlook vesselreport
REM          -> incremental parse into rob_data/abstract_log.db
REM  Step 2: checkpoint WAL + upload ONLY abstract_log.db to GitHub
REM          Release (replaces old asset; reports.db untouched)
REM  Step 3: regenerate DATABASES.html (fresh timestamp) + safe push
REM
REM  ASCII-ONLY ON PURPOSE. Chinese inside .bat breaks cmd.exe.
REM  Exit 2 from collect = Outlook cache not synced -> skip this
REM  round; next scheduled run backfills (--days 5 window).
REM ============================================================
cd /d %~dp0
set LOG=rob_data\abstract_daily.log
set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe
if not exist "%PYEXE%" set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe

echo ===== %date% %time% ===== >> %LOG%

REM ---- [1/3] git op mutex (share the same lock as ROB) ----
"%PYEXE%" git_op_lock.py acquire --owner abstract --wait 300 >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] git lock held, skip this run >> %LOG%
    exit /b 1
)

REM ---- [2/3] collect + parse into abstract_log.db ----
echo [2/3] abstract_daily_sync --apply ... >> %LOG%
"%PYEXE%" abstract_daily_sync.py --days 5 --apply >> %LOG% 2>&1
if errorlevel 2 (
    echo [SKIP] outlook cache not synced - exit 2 - next run backfills >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner abstract >> %LOG% 2>&1
    exit /b 1
)
if errorlevel 1 (
    echo [ERROR] abstract_daily_sync failed >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner abstract >> %LOG% 2>&1
    exit /b 1
)

REM ---- [3/4] upload abstract_log.db to GitHub Release ----
echo [3/4] upload_abstract_db --apply ... >> %LOG%
"%PYEXE%" upload_abstract_db.py --apply >> %LOG% 2>&1
if errorlevel 1 (
    echo [ERROR] upload_abstract_db failed >> %LOG%
    "%PYEXE%" git_op_lock.py release --owner abstract >> %LOG% 2>&1
    exit /b 1
)

REM ---- [4/4] regenerate DATABASES.html with fresh timestamp + push ----
echo [4/4] databases_page_gen + safe push ... >> %LOG%
"%PYEXE%" databases_page_gen.py >> %LOG% 2>&1
if errorlevel 1 echo [WARN] databases_page_gen failed, continue >> %LOG%
set CLAW_LOCK_OWNER=abstract
"%PYEXE%" git_safe_push.py -m "abstract daily update %date%" >> %LOG% 2>&1
if errorlevel 1 echo [WARN] git_safe_push failed, will retry next run >> %LOG%

echo DONE >> %LOG%
"%PYEXE%" git_op_lock.py release --owner abstract >> %LOG% 2>&1
exit /b 0
