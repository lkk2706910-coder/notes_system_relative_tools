@echo off
rem 抓取 ECMS P58 Online Status 信件附件並更新機況表資料 (供 Windows 工作排程器使用)
rem 設定 (信箱 / 主旨 / 篩選 / 輸出資料夾) 在 FetchEcmsStatus.vbs 頂端

cd /d "%~dp0"
echo [%date% %time%] === 開始抓取 ===>> fetch_ecms_status_log.txt
cscript //nologo FetchEcmsStatus.vbs >> fetch_ecms_status_log.txt 2>&1
echo [%date% %time%] === 結束 (exit=%errorlevel%) ===>> fetch_ecms_status_log.txt
