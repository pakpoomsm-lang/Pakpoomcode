@echo off
chcp 65001 >nul
title Expander First Lot + Cutting Server :3001

cd /d "W:\PD\2.HEAT INDOOR\13.Suphamat P\Server_firstlot"

echo.
echo  ================================================
echo   First Lot Check Server — port 3001
echo  ================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] ไม่พบ Python — กรุณาติดตั้ง Python ก่อน
    pause
    exit
)

echo  กำลังตรวจสอบ packages...
pip install -r requirements.txt --quiet 2>nul
echo  packages พร้อมแล้ว
echo.
echo  กำลังเริ่ม Server...
echo  หยุด Server กด Ctrl+C
echo.
echo  ================================================
echo   Expander First Lot  :  /expander/firstlot
echo   Fin Press First Lot :  /fp/firstlot
echo   FP Dashboard        :  /fp/dashboard
echo   Cutting Check       :  /cutting/firstlot
echo   HP Check            :  /hp/firstlot
echo   Hairpin Insert      :  /hp_insert/firstlot
echo   Shared Dashboard    :  /dashboard/firstlot
echo   Layout HEI          :  /Layout_HEI
echo  ================================================
echo.
python server_expander.py
echo.
echo  Server หยุดทำงาน
pause
