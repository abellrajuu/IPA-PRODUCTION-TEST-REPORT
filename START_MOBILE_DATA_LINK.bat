@echo off
title Production Test Report - Global Mobile Server & Cloud Registry
color 0A

REM Support UNC network shares (\\192.168.100.248) and local drives seamlessly
pushd "%~dp0"

echo ===================================================================
echo     STARTING SERVER AND CONNECTING GLOBAL MOBILE DATA LINK
echo ===================================================================
echo.
echo   * Auto-starts Flask backend on port 5005
echo   * Connects secure public HTTPS tunnel (Works on 4G/5G mobile data)
echo   * Automatically publishes live URL to GitHub Cloud Registry
echo   * Zero link pasting required by employees on mobile devices!
echo.
echo ===================================================================
echo.

python -u tunnel_publisher.py

popd
pause
