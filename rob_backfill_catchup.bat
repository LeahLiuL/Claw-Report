@echo off
rem RPA history catch-up: scan vesselreport mailbox only, all folders, 324-day window
rem Pure ASCII only. Never add Chinese text to this file.
set PYEXE=C:\Users\culadmin\.workbuddy\binaries\python\envs\default\Scripts\python.exe
cd /d C:\Users\culadmin\Claw-Report
set CLAW_LOCK_OWNER=rob_catchup
set ROB_FOLDER_CAP=30000
set ROB_PER_VESSEL=2000
echo ===== %DATE% %TIME% catchup start ===== >> rob_data\catchup.log
"%PYEXE%" git_op_lock.py acquire --owner rob_catchup --wait 120 >> rob_data\catchup.log 2>&1
if errorlevel 1 exit /b 1
"%PYEXE%" rob_mail_sync.py --days 324 --apply >> rob_data\catchup.log 2>&1
"%PYEXE%" rob_fast_rebuild.py >> rob_data\catchup.log 2>&1
"%PYEXE%" git_safe_push.py -m "RPA history catch-up backfill" >> rob_data\catchup.log 2>&1
"%PYEXE%" git_op_lock.py release --owner rob_catchup >> rob_data\catchup.log 2>&1
echo catchup DONE >> rob_data\catchup.log
