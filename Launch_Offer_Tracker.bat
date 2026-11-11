@echo off
title ATI Offer Tracker Launcher
cd /d "%~dp0"
echo Starting ATI Offer Tracker background server...
start /b py server.py >nul 2>&1
timeout /t 1 /nobreak >nul
echo Opening Dashboard in your browser...
start "" "http://localhost:8000/"
exit

