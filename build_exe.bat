@echo off
rem 把 fetch_ecms_status.py 打包成單一 FetchEcmsStatus.exe（輸出在 dist\）
rem ★ 用來打包的 python 位元數要能連得上 Notes（先跑 python -c "import win32com.client as w;w.Dispatch('Notes.NotesSession')" 測）
rem ★ 若機器上有多個 Python，請把下面的 PY 改成你要用的那一個，例如  set PY=py -3-32

setlocal
set PY=python

cd /d "%~dp0"
echo === 1. 安裝套件（到打包用的同一個 Python） ===
%PY% -m pip install --upgrade pywin32 pyinstaller || goto :err

echo === 2. 確認 pywin32 在這個 Python 可以 import ===
%PY% -c "import pythoncom, pywintypes, win32com.client; import struct; print('pywin32 OK,', struct.calcsize('P')*8, 'bit')" || goto :err

echo === 3. 打包（pywin32 的動態模組要用 hidden-import 明確帶入） ===
%PY% -m PyInstaller --onefile --console --clean --name FetchEcmsStatus ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import win32com.client.dynamic ^
  --hidden-import win32com.client.gencache ^
  --hidden-import win32timezone ^
  --collect-submodules win32com ^
  fetch_ecms_status.py || goto :err

echo.
echo 完成：dist\FetchEcmsStatus.exe
echo 請把 fetch_ecms_status.ini（由 .ini.example 複製修改）放在 exe 同資料夾。
endlocal
goto :eof

:err
echo.
echo 打包失敗，請看上方訊息。
endlocal
exit /b 1
