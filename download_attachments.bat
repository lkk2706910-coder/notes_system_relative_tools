@echo off
rem 執行 Notes 附件自動下載 (供雙擊或 Windows 工作排程器使用)
rem 設定 (文件編號 / 存檔資料夾等) 在 DownloadNotesAttachments.vbs 頂端

cd /d "%~dp0"
echo [%date% %time%] === 開始下載 ===>> download_log.txt
cscript //nologo DownloadNotesAttachments.vbs >> download_log.txt 2>&1
echo [%date% %time%] === 結束 (exit=%errorlevel%) ===>> download_log.txt
