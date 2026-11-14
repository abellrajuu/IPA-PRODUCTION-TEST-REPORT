@echo off
title Mobile Wi-Fi Connector
pushd "%~dp0"
python show_wifi_qr.py
popd
pause
