@echo off
rem 把 match_pms_docs.py 打包成單一 MatchPmsDocs.exe（輸出在 dist\）
rem ★ 打包用的 python 要能連得上 Notes；機器上有多個 Python 時改下面的 PY，例如  set PY=py -3-32
setlocal
set PY=python
cd /d "%~dp0"
%PY% -m pip install --upgrade pywin32 openpyxl pyinstaller || goto :err
%PY% -c "import pythoncom, pywintypes, win32com.client, openpyxl; import struct; print('deps OK,', struct.calcsize('P')*8, 'bit')" || goto :err
%PY% -m PyInstaller --onefile --noconsole --clean --name MatchPmsDocs --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32com --hidden-import win32com.client --hidden-import win32com.client.dynamic --hidden-import win32com.client.gencache --hidden-import win32timezone --collect-submodules win32com --collect-submodules openpyxl match_pms_docs.py || goto :err
echo.
echo 完成：dist\MatchPmsDocs.exe   （雙擊開視窗；或把 xlsx 拖到 exe 上；排程用 MatchPmsDocs.exe 檔案.xlsx --cli）
endlocal
goto :eof
:err
echo 打包失敗，請看上方訊息。
endlocal
exit /b 1
