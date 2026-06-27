# MECP Part Order System — Python Version
**FastAPI + WebSocket + Uvicorn**

---

## โครงสร้างไฟล์

```
mecp-python/
├── server.py          ← Python Backend (FastAPI + WebSocket)
├── requirements.txt   ← Python packages ที่ต้องการ
├── start.bat          ← ดับเบิลคลิกเพื่อเริ่ม Server
├── orders.json        ← ข้อมูล Orders (สร้างอัตโนมัติ)
└── public/
    └── index.html     ← Frontend
```

---

## ขั้นตอนติดตั้ง

### 1. ติดตั้ง Python
- ดาวน์โหลดที่ https://www.python.org/downloads/
- เลือก **Python 3.11** หรือใหม่กว่า
- ⚠️ **สำคัญ:** ติ๊ก ✅ **"Add Python to PATH"** ตอนติดตั้ง

### 2. วางโฟลเดอร์ mecp-python ไว้ที่ C:\
```
C:\mecp-python\
```

### 3. เริ่ม Server
**ดับเบิลคลิกที่ `start.bat`** — จบเลย!

หรือถ้าจะใช้ CMD:
```cmd
cd C:\mecp-python
pip install -r requirements.txt
python server.py
```

---

## เปิดใช้งาน

| เครื่อง | URL |
|---------|-----|
| Computer 1 (สั่งผลิต) | `http://192.168.x.x:3000/?role=planner` |
| Computer 2 (ฝ่ายผลิต) | `http://192.168.x.x:3000/?role=floor` |

IP จะแสดงใน CMD หลัง start server

---

## ตั้งให้รันอัตโนมัติเมื่อเปิดเครื่อง (ไม่ต้องเปิด CMD)

1. กด `Win + R` พิมพ์ `shell:startup` แล้ว Enter
2. Copy ไฟล์ `start.bat` ไปวางใน Folder นั้น
3. เสร็จ — เปิดเครื่องมา Server รันอัตโนมัติเลย

---

## แก้ปัญหาเบื้องต้น

| ปัญหา | วิธีแก้ |
|-------|--------|
| `python` not found | ติดตั้ง Python ใหม่ ติ๊ก Add to PATH |
| Port 3000 ถูกใช้ | แก้ `port=3000` ใน server.py เป็น 3001 |
| Computer อื่นเข้าไม่ได้ | เปิด Firewall Port 3000 |
