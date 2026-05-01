@echo off
setlocal
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

cd /d "%~dp0"

if exist graph_config.bat call graph_config.bat

python -c "from graph_mail import fetch_latest_cbas_email; fetch_latest_cbas_email()"
pause
