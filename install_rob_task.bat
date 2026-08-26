@echo off
REM ============================================================
REM  One-time setup for ROB daily auto-update (run once on culadmin)
REM  1) install python deps  2) register scheduled tasks 13:00 & 01:00
REM ============================================================
echo [1/2] Installing Python dependencies (pywin32 / pycryptodome / openpyxl) ...
"C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe" -m pip show pywin32 >nul 2>&1 || "C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe" -m pip install pywin32
"C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe" -m pip show pycryptodome >nul 2>&1 || "C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe" -m pip install pycryptodome
"C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe" -m pip show openpyxl >nul 2>&1 || "C:\Users\culadmin\.workbuddy\binaries\python\versions\3.13.12\python.exe" -m pip install openpyxl

echo [2/2] Registering scheduled tasks (daily 13:00 and 01:00) ...
schtasks /Create /TN "ClawReport_ROB_1300" /TR "\"%~dp0rob_update.bat\"" /SC DAILY /ST 13:00 /F
schtasks /Create /TN "ClawReport_ROB_0100" /TR "\"%~dp0rob_update.bat\"" /SC DAILY /ST 01:00 /F

echo.
echo ============================================================
echo Setup done. Tasks: ClawReport_ROB_1300 / ClawReport_ROB_0100
echo Test now:   schtasks /Run /TN "ClawReport_ROB_1300"
echo Check log:  rob_data\autoupdate.log
echo Webpage:    https://leahliul.github.io/Claw-Report/rob_oil_report.html
echo ============================================================
pause
