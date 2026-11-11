@echo off
chcp 65001 >nul
setlocal

rem 找 Python：先試 python，沒有再試 py 啟動器
set "PY_CMD="
where python >nul 2>nul && set "PY_CMD=python"
if not defined PY_CMD (where py >nul 2>nul && set "PY_CMD=py -3")
if not defined PY_CMD goto nopython

rem 有拖檔案就用拖的；直接雙擊就找同資料夾裡最新的 GVOexcel*.xlsx（備份檔除外）
set "TARGET=%~1"
if defined TARGET goto run
for /f "delims=" %%F in ('dir /b /o-d "%~dp0GVOexcel*.xlsx" 2^>nul ^| findstr /v /i /c:".bak.xlsx"') do (
    set "TARGET=%~dp0%%F"
    goto run
)
echo 找不到 GVOexcel 檔。請把檔案拖到這個 bat 上，或把 bat 放在 GVOexcel 檔的資料夾裡再雙擊。
pause
exit /b 1

:run
echo 目標檔案：%TARGET%
%PY_CMD% "%~dp0gvo_add_all_sheet.py" "%TARGET%"
echo.
pause
exit /b 0

:nopython
echo 找不到 Python，請先安裝 Python，安裝時勾選 Add python.exe to PATH。
pause
exit /b 1
