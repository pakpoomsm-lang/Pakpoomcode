@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
title MECP + Cutting Server
cd /d "%~dp0"
echo.
echo  ================================================
echo   MECP + Cutting Check
echo   Port 3000
echo  ================================================
echo.
echo  Starting Server...
echo  Stop Server: Ctrl+C
echo.
python server.py
echo.
echo  Server stopped.
pause
