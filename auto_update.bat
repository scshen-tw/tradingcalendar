@echo off
setlocal
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

cd /d "%~dp0"

echo [%date% %time%] Starting CB calendar update... >> update_log.txt
python log_update_status.py CB START "scheduled update started" --counts --commit >> update_log.txt 2>&1

git pull --rebase --autostash origin main >> update_log.txt 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: git pull failed. >> update_log.txt
    python log_update_status.py CB ERROR "git pull failed" --counts --commit >> update_log.txt 2>&1
    exit /b 1
)

python extract_outlook.py >> update_log.txt 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: extract_outlook.py failed. >> update_log.txt
    python log_update_status.py CB ERROR "extract_outlook.py failed" --counts --commit >> update_log.txt 2>&1
    exit /b 1
)

git add calendar.html events.json >> update_log.txt 2>&1
git diff --cached --quiet
if %errorlevel% == 0 (
    echo [%date% %time%] No changes to commit. >> update_log.txt
    python log_update_status.py CB NO_CHANGE "no file changes after update" --counts --commit >> update_log.txt 2>&1
    exit /b 0
)

git commit -m "auto update %date% %time%" >> update_log.txt 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: git commit failed. >> update_log.txt
    python log_update_status.py CB ERROR "git commit failed" --counts --commit >> update_log.txt 2>&1
    exit /b 1
)

git push origin main >> update_log.txt 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: git push failed. >> update_log.txt
    python log_update_status.py CB ERROR "git push failed" --counts --commit >> update_log.txt 2>&1
    exit /b 1
)

echo [%date% %time%] CB calendar update completed. >> update_log.txt
python log_update_status.py CB SUCCESS "updated and pushed" --counts --commit >> update_log.txt 2>&1
exit /b 0
