@echo off
title Connect Phone over USB - Production Report App
color 0A

pushd "%~dp0"

echo ===================================================================
echo       CONNECTING ANDROID PHONE TO PRODUCTION SERVER OVER USB
echo ===================================================================
echo.

set ADB_EXE=C:\Users\Admin\AppData\Local\Android\Sdk\platform-tools\adb.exe
if not exist "%ADB_EXE%" (
    set ADB_EXE=%~dp0ANDROID APP\platform-tools\adb.exe
)

if not exist "%ADB_EXE%" (
    echo [ERROR] adb.exe not found! Please ensure Android SDK is installed.
    pause
    exit /b 1
)

echo [...] Checking connected devices...
"%ADB_EXE%" devices
echo.

echo [...] Reversing port 5005 to your Android phone...
"%ADB_EXE%" reverse tcp:5005 tcp:5005
if errorlevel 1 (
    echo [ERROR] Failed to map port. Make sure USB Debugging is authorized on your phone screen!
    pause
    exit /b 1
)

echo.
echo ===================================================================
echo   SUCCESS! Phone is now connected directly to your laptop server.
echo   * Zero internet required.
echo   * Zero cloud tunnel drops.
echo   * All reports save directly to \\192.168.100.248\prdndata\TEST REPORT
echo ===================================================================
echo.
echo Starting app on phone...
"%ADB_EXE%" shell am start -n com.production.reportapp/.MainActivity

popd
pause
