#!/usr/bin/env python3
"""ZPP0059 auto-updater — ดึงข้อมูล ZPP0059 จาก SAP อัตโนมัติเป็นรอบ ๆ แล้ว
upsert ลง sqlite เดิม (zpp0059.db ของโปรแกรม Daily Follow) เพื่อให้หน้า
dashboard เห็นยอดผลิตล่าสุดโดยไม่ต้องกดปุ่มเขียวเอง

โปรแกรมนี้รัน "แยก process" กับ serve_daily_follow.py แต่ *ใช้ฟังก์ชันเดิมทั้งหมด*
(ไม่ก๊อปโค้ดขับ SAP มาซ้ำ) และเขียนลง DB ไฟล์เดียวกัน

กันชนกันยังไง
  - SAP : _drive_sap_export ถือ file-lock กลาง (_sap_automation.lock) อยู่แล้ว
          → รอบดึงของโปรแกรมนี้จะไม่ชนกับปุ่มเขียว/ZPP0022 ของหน้าเว็บ
  - DB  : _get_db เปิด WAL + busy_timeout → หน้าเว็บ (อ่าน) กับตัวนี้ (เขียน)
          ทำงานพร้อมกันได้ปลอดภัย

จังหวะการดึงเป็นแบบ gap-based: ดึงเสร็จ -> พัก N วินาที -> ดึงใหม่ จึงไม่มีทาง
ซ้อนกันไม่ว่า SAP จะช้าแค่ไหน
"""
import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

# โฟลเดอร์ "Daily follow" อยู่ข้าง ๆ กัน — เอาเข้ามาใน path เพื่อ import โมดูลเดิม
HERE = Path(__file__).resolve().parent
DAILY_FOLLOW_DIR = HERE.parent / "Daily follow"
if not (DAILY_FOLLOW_DIR / "serve_daily_follow.py").exists():
    sys.exit(f"หาโฟลเดอร์ 'Daily follow' ไม่เจอที่ {DAILY_FOLLOW_DIR}\n"
             f"วางโฟลเดอร์ 'ZPP0059 Updater' ไว้ข้าง ๆ 'Daily follow' นะครับ")
sys.path.insert(0, str(DAILY_FOLLOW_DIR))

import serve_daily_follow as srv   # noqa: E402  (ต้องมาหลัง sys.path)

DEFAULT_INTERVAL = 120   # วินาทีที่ "พักหลังดึงเสร็จ" ก่อนเริ่มรอบใหม่
MIN_INTERVAL = 60        # กันตั้งถี่จนเป็นภาระ SAP (1 รอบใช้เวลาราว 15–60 วิ อยู่แล้ว)


def log(msg: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def one_cycle() -> str:
    """ดึง ZPP0059 หนึ่งครั้ง (ใช้ช่วงวันแบบ rolling ตามที่ตั้งไว้เดิม) แล้ว
    upsert ลง DB. คืนข้อความสรุปผล."""
    t0 = time.time()
    _mat, _lot, fname, stats = srv.run_zpp0059()
    dt = time.time() - t0
    new = stats.get("new_rows", 0)
    upd = stats.get("updated_rows", 0)
    total = stats.get("total_rows", "?")
    return f"{fname} · +{new} ใหม่ / {upd} อัปเดต · DB รวม {total} แถว · {dt:.1f}s"


def main() -> int:
    ap = argparse.ArgumentParser(
        description="ดึง ZPP0059 จาก SAP เข้ามาอัปเดต sqlite ของ Daily Follow เป็นรอบ ๆ")
    ap.add_argument("--interval", type=int, default=DEFAULT_INTERVAL,
                    help=f"วินาทีที่พักหลังดึงเสร็จก่อนเริ่มรอบใหม่ (ขั้นต่ำ {MIN_INTERVAL}, ค่าเริ่มต้น {DEFAULT_INTERVAL})")
    ap.add_argument("--once", action="store_true",
                    help="ดึงครั้งเดียวแล้วออก (ไว้ทดสอบ)")
    args = ap.parse_args()
    interval = max(args.interval, MIN_INTERVAL)

    log("=== ZPP0059 auto-updater ===")
    log(f"DB ปลายทาง : {srv.DB_FILE}")
    log(f"พักรอบละ   : {interval} วินาที (นับหลังดึงเสร็จ)")
    log("กด Ctrl+C เพื่อหยุด")

    while True:
        try:
            log("OK  " + one_cycle())
        except srv.SapNotReady as exc:
            log(f"SKIP SAP ยังไม่พร้อม — {exc}")
        except srv.SapBusy as exc:
            log(f"SKIP {exc}")
        except KeyboardInterrupt:
            log("หยุดโดยผู้ใช้ (Ctrl+C)")
            return 0
        except Exception as exc:                    # noqa: BLE001 — รอบหน้าค่อยลองใหม่
            log(f"ERROR {type(exc).__name__}: {exc}")

        if args.once:
            return 0
        try:
            time.sleep(interval)
        except KeyboardInterrupt:
            log("หยุดโดยผู้ใช้ (Ctrl+C)")
            return 0


if __name__ == "__main__":
    sys.exit(main())
