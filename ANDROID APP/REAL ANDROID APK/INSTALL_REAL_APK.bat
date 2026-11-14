@echo off
title Install Real Production Test Report APK
color 0B

echo.
echo  ==============================================================
echo    PRODUCTION TEST REPORT - REAL ANDROID APK INSTALLER
echo    Brand: Production Test Report (IPA Weighing & Automation)
echo  ==============================================================
echo.

set "ADB_PATH=%~dp0..\ANDROID APP\platform-tools\adb.exe"
if not exist "%ADB_PATH%" set "ADB_PATH=%~dp0..\platform-tools\adb.exe"
if not exist "%ADB_PATH%" set "ADB_PATH=adb"

echo  Checking for connected Android phone via USB...
"%ADB_PATH%" devices | findstr /r /c:"\<device\>" > nul
if errorlevel 1 (
    echo.
    echo  [!] No Android phone detected in USB Debugging mode.
    echo.
    echo  Quick steps on your phone:
    echo    1. Connect your phone to this laptop with a USB cable.
    echo    2. Go to Settings ^> Developer Options.
    echo    3. Enable "USB Debugging".
    echo    4. Tap "Always Allow" on your phone screen.
    echo    5. Run this file again!
    echo.
    pause
    exit /b
)

echo  [OK] Phone detected!
echo.
echo  Setting up high-speed local USB port forwarding (5005)...
"%ADB_PATH%" reverse tcp:5005 tcp:5005 > nul 2>&1
echo  [OK] Port 5005 mapped directly to laptop!
echo.

echo  Installing ProductionReportApp.apk on your phone...
echo  (Unlock your phone screen and tap "Install" or "Allow" if prompted)
echo.
"%ADB_PATH%" install -r "%~dp0ProductionReportApp.apk"

if errorlevel 1 (
    echo.
    echo  [!] Installation failed.
    echo      Make sure your phone screen is unlocked and "Install via USB" is allowed.
) else (
    echo.
    echo  ==============================================================
    echo    SUCCESS! "Production Test Report" installed on your phone!
    echo  ==============================================================
    echo.
    echo  Launching app on phone now...
    "%ADB_PATH%" shell am start -n com.production.reportapp/.MainActivity
    echo.
    echo  Enjoy your full production testing app!
)

echo.
pause
