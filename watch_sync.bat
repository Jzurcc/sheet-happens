@echo off
title SheetHappens - Live Watcher
cd /d "%~dp0"
python sync_sheets.py --watch
echo.
pause
