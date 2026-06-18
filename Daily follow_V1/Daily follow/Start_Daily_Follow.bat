@echo off
REM Double-click to start the Daily Follow dashboard with the SAP ZPP0059 button.
REM Requires: Python installed, and SAP GUI open + logged in (with scripting enabled).
cd /d "%~dp0"
where py >nul 2>nul && (py serve_daily_follow.py & goto :eof)
python serve_daily_follow.py
pause
