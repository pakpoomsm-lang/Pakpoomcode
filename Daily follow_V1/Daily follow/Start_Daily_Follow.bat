@echo off
REM Double-click to start the Daily Follow dashboard with the SAP ZPP0059 button.
REM Requires: Python installed, and SAP GUI open + logged in (with scripting enabled).
cd /d "%~dp0"

REM --- First-lot check sheet folder -------------------------------------------
REM Set FIRSTLOT_DIR to the Server_firstlot folder on this PC.
REM The program will also search W:\ Y:\ J:\ automatically, so this line is
REM only needed when the drive letter differs from those defaults.
REM Example:  set FIRSTLOT_DIR=X:\PD\2.HEAT INDOOR\13.Suphamat P\Program First lot all process
set FIRSTLOT_DIR=W:\PD\2.HEAT INDOOR\13.Suphamat P\Program First lot all process

where py >nul 2>nul && (py serve_daily_follow.py & goto :eof)
python serve_daily_follow.py
pause
