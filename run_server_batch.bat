@echo off
title Production Test Report Production Report Generator Server
color 0B
echo ========================================================
echo   IPA Production Test Report System - IPA Private Limited
echo ========================================================
echo.
echo Cleaning up any old background server processes on port 5005...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr /r /c:":5005 " ^| findstr LISTENING 2^>nul') do (
    taskkill /f /pid %%a >nul 2>&1
)
echo.
echo Starting Flask Application Server...
echo Portal Hub URL: http://localhost:5005/portal
echo.
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" "http://localhost:5005/portal"
) else if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "http://localhost:5005/portal"
) else (
    start "" "http://localhost:5005/portal"
)

REM Support UNC network shares (\\192.168.100.248) and local drives seamlessly
pushd "%~dp0"
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" backend_python.py
) else if exist "%~dp0venv\Scripts\python.exe" (
    "%~dp0venv\Scripts\python.exe" backend_python.py
) else if exist "C:\Users\Admin\AppData\Local\Programs\Python\Python312\python.exe" (
    "C:\Users\Admin\AppData\Local\Programs\Python\Python312\python.exe" backend_python.py
) else (
    python backend_python.py
)
popd
pause
