@echo off
REM ดึง ZPP0059 จาก SAP อัตโนมัติทุก ๆ ~2 นาที เข้ามาอัปเดต sqlite ของ Daily Follow
REM ต้องมี SAP GUI เปิด + login ค้างไว้ + เปิด SAP GUI Scripting บนเครื่องนี้
cd /d "%~dp0"
python update_zpp0059.py --interval 120
pause
