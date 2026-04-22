@echo off
chcp 65001 >nul
echo [%date% %time%] 開始更新股票競拍行事曆...
cd /d "%~dp0"
python update_stocks.py >> update_log.txt 2>&1
if %errorlevel% == 0 (
    echo [%date% %time%] ✅ 股票競拍更新成功
    git add calendar.html events.json >> update_log.txt 2>&1
    git commit -m "update stocks %date% %time%" >> update_log.txt 2>&1
    git push origin main >> update_log.txt 2>&1
    echo [%date% %time%] ✅ 已推送至 GitHub
) else (
    echo [%date% %time%] ❌ 更新失敗，請查看 update_log.txt
)
