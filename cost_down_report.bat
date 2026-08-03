@echo off
rem 執行 Cost Down Project 預估報表 (供雙擊或 Windows 工作排程器使用)
rem 設定 (SERVER / DBFILE / VIEWNAME / 收件人 / DRY_RUN 等) 在 CostDownReport.vbs 頂端

cd /d "%~dp0"
echo [%date% %time%] === 開始產出報表 ===>> cost_down_report_log.txt
cscript //nologo CostDownReport.vbs >> cost_down_report_log.txt 2>&1
echo [%date% %time%] === 結束 (exit=%errorlevel%) ===>> cost_down_report_log.txt
