@echo off
setlocal
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

cd /d "%~dp0"

echo [%date% %time%] Starting stock calendar update... >> update_log.txt

git pull --rebase --autostash origin main >> update_log.txt 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: git pull failed. >> update_log.txt
    exit /b 1
)

python update_stocks.py >> update_log.txt 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: update_stocks.py failed. >> update_log.txt
    exit /b 1
)

git add calendar.html events.json >> update_log.txt 2>&1
git diff --cached --quiet
if %errorlevel% == 0 (
    echo [%date% %time%] No changes to commit. >> update_log.txt
    exit /b 0
)

git commit -m "update stocks %date% %time%" >> update_log.txt 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: git commit failed. >> update_log.txt
    exit /b 1
)

git push origin main >> update_log.txt 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: git push failed. >> update_log.txt
    exit /b 1
)

echo [%date% %time%] Stock calendar update completed. >> update_log.txt
exit /b 0
