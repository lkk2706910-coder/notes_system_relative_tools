@echo off
rem 把 fetch_ecms_status.py 打包成單一 FetchEcmsStatus.exe（輸出在 dist\）
rem ★ 這裡用的 python 位元數必須跟 Notes client 一致（Notes 多半是 32 位元 → 用 32 位元 Python）
rem    檢查：python -c "import struct;print(struct.calcsize('P')*8)"  → 印 32 或 64

cd /d "%~dp0"
python -m pip install -r requirements.txt || goto :err
python -m PyInstaller --onefile --console --clean --name FetchEcmsStatus fetch_ecms_status.py || goto :err
echo.
echo 完成：dist\FetchEcmsStatus.exe
echo 請把 fetch_ecms_status.ini（由 .ini.example 複製修改）放在 exe 同資料夾。
goto :eof

:err
echo 打包失敗，請看上方訊息。
exit /b 1
