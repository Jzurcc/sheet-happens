@echo off
title SheetHappens - Push to Google Sheets
cd /d "%~dp0"
python sync_sheets.py --push
echo.
pause
