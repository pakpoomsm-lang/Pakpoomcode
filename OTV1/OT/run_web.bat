@echo off
echo ========================================
echo  ระบบบันทึกการทำ OT พนักงาน (Web Version)
echo  Heat Exchange Indoor
echo ========================================
echo.
echo กำลังเริ่มเว็บเซิร์ฟเวอร์...
echo.
echo เปิดเว็บเบราว์เซอร์แล้วไปที่: http://localhost:5000
echo.
echo กด Ctrl+C เพื่อหยุดเซิร์ฟเวอร์
echo ========================================
echo.

python app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================
    echo  เกิดข้อผิดพลาด!
    echo ========================================
    echo.
    echo กรุณาตรวจสอบว่าติดตั้ง Python libraries ครบถ้วนแล้ว
    echo รันคำสั่ง: pip install flask pandas openpyxl
    echo.
    pause
)
