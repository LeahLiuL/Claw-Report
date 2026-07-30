@echo off
REM ============================================================
REM  CUL Daily Movement - daily update (run on culadmin by double-click)
REM  One-time setup: git clone this repo + install Python + pip install openpyxl
REM  Daily: just double-click this file.
REM    pull latest -> refresh PIC -> rebuild big Excel -> regenerate webpage -> push
REM ============================================================
cd /d %~dp0

echo [1/5] Pull latest scripts and vessel.csv ...
git pull --ff-only
if errorlevel 1 (
  echo [ERROR] git pull failed (local uncommitted changes?). Fix then rerun.
  pause
  exit /b 1
)

echo [2/5] Refresh PIC summary table (keeps manually edited PICs) ...
python gen_pic.py
if errorlevel 1 ( echo [ERROR] gen_pic.py failed & pause & exit /b 1 )

echo [3/5] Rebuild big Excel from current 2026 fleet ...
python build_fleet_movement.py
if errorlevel 1 ( echo [ERROR] build_fleet_movement.py failed & pause & exit /b 1 )

echo [4/5] Regenerate webpage ...
python gen_html.py
if errorlevel 1 ( echo [ERROR] gen_html.py failed & pause & exit /b 1 )

echo [5/5] Push webpage and vessel.csv if changed ...
git add cul_daily_movement.html vessel.csv
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "daily update %date% %time%"
  git push
) else (
  echo No changes, skip commit.
)

echo ============================================================
echo DONE. Webpage updated in local repo.
echo (If GitHub Pages publishes from gh-pages branch, one extra push step is needed - see notes)
echo ============================================================
pause
