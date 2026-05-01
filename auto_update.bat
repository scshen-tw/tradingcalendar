@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set GIT_TERMINAL_PROMPT=0
set GCM_INTERACTIVE=never

cd /d "%~dp0"

python ensure_utf8_log.py update_log.txt
if exist graph_config.bat call graph_config.bat
if exist "C:\Users\User\.ssh\id_ed25519_tradingcalendar" set "GIT_SSH_COMMAND=ssh -i C:/Users/User/.ssh/id_ed25519_tradingcalendar -o IdentitiesOnly=yes -o StrictHostKeyChecking=accept-new"

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
    set AHEAD=0
    for /f %%A in ('git rev-list --count origin/main..HEAD 2^>nul') do set AHEAD=%%A
    if not "!AHEAD!" == "0" (
        echo [%date% %time%] No file changes, but !AHEAD! local commits pending push. >> update_log.txt
        git push origin main >> update_log.txt 2>&1
        if errorlevel 1 (
            echo [%date% %time%] ERROR: pending git push failed. >> update_log.txt
            python log_update_status.py CB ERROR "pending git push failed" --counts --commit >> update_log.txt 2>&1
            exit /b 1
        )
        echo [%date% %time%] Pending local commits pushed. >> update_log.txt
        python log_update_status.py CB SUCCESS "pushed pending local commit" --counts --commit >> update_log.txt 2>&1
        exit /b 0
    )
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
