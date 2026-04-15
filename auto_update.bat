@echo off
chcp 65001 >nul
echo [%date% %time%] 開始更新 CB 行事曆...
cd /d "%~dp0"
python extract_outlook.py >> update_log.txt 2>&1
if %errorlevel% == 0 (
    echo [%date% %time%] ✅ 更新成功
) else (
    echo [%date% %time%] ❌ 更新失敗，請查看 update_log.txt
)
