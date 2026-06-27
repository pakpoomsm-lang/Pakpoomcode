# Concept — คอลัมน์ "สถานะ Check Sheet" ในหน้า Daily Follow

ร่างแนวคิด (ยังไม่ลงมือเขียนโค้ด) สำหรับเพิ่ม **1 คอลัมน์** ในตาราง index
ที่บอกว่าแต่ละ lot ลง check sheet ยืนยัน "ครบทุก process แล้วหรือยัง"
โดยดึงข้อมูลจากโฟลเดอร์ `Server_firstlot`

---

## 1. เป้าหมาย
ให้คนดูตาราง Daily Follow เห็นได้ในแวบเดียวว่า lot ไหน "ยืนยัน check sheet
ครบแล้ว" และ lot ไหน "ยังขาด process อะไร" โดยไม่ต้องเปิดระบบ first-lot แยก

## 2. แหล่งข้อมูล (Server_firstlot)
แต่ละ process มี SQLite 1 ไฟล์ คีย์ร่วมคือ `line`, `seq`, `lot` (+ `prod_month` บางตัว)
และมีฟิลด์ผลตรวจ OK/NG:

| ไฟล์ DB                | table             | process          | ฟิลด์ผล   |
|------------------------|-------------------|------------------|-----------|
| `cutting_records.db`   | `cutting_records` | Cutting          | `status`  |
| `fp_records.db`        | `fp_records`      | FP (Insert/ฟิน)  | `final`   |
| `hp_records.db`        | `hp_records`      | H/P bender       | `result`  |
| `hp_insert_records.db` | `records`         | HP insert (ประกอบ)| `final_check` |
| `oven_records.db`      | `oven_records`    | Oven/Brazing     | `temp_result`/`status` |
| `expander_records.db`  | (ว่าง)            | Expander         | —         |

> หมายเหตุ: ปัจจุบัน `oven_records` กับ `expander_records` ยังว่าง — concept ต้อง
> เผื่อกรณี DB ว่าง/ไม่มีตาราง โดยไม่ทำให้ทั้งระบบพัง

## 3. การ join กับแถวใน Daily Follow
แถวใน index ถูกคีย์ด้วย `line | month | seq` (ดู `build_daily_follow.py` →
`load_progress`, key = `f"{line}|{month_display}|{seq_key}"`).
check sheet คีย์ด้วย `line | seq | lot | prod_month`.

**คีย์ที่เสนอให้ใช้ join:** `line + seq + month` เป็นหลัก
- `seq`: normalize เป็นตัวเลข (เทียบ `"110"` กับ `110` ให้ตรงกัน เหมือน `seq_key`)
- `month`: แปลง `prod_month` ของ check sheet (`"06/2026"`) ให้ตรงรูปแบบ
  `month_display` (`"6.2026"`)
- `line`: normalize ตัวพิมพ์ (พบค่า `a`,`b`,`h` ปนกับ `A`,`B`,`H`) → uppercase + trim

> **คำถามต้องตัดสินใจ (สำคัญสุด): granularity ของ `lot`**
> check sheet มีฟิลด์ `lot` (เช่น 200, 131) แต่แถว Daily Follow ไม่มี lot ตรง ๆ
> มีแต่ `assyOrderNo`. เลือกได้ 2 ทาง:
> - **(ก) นับระดับ seq/เดือน** — ขอแค่มี check sheet ที่ confirm ของ process นั้น
>   ภายใน line+seq+month เดียวกัน ก็ถือว่า "ครบ" (ง่าย, ตรงกับที่ตารางแสดงอยู่)
> - **(ข) นับระดับ lot** — ต้อง map `assyOrderNo` ↔ `lot` ให้ได้ก่อน (แม่นกว่า
>   แต่ตอนนี้ยังไม่มีตัวเชื่อม ต้องหา field กลาง)
> ผมแนะนำเริ่มที่ **(ก)** ก่อน เพราะตารางทำงานระดับ seq อยู่แล้ว แล้วค่อยยกระดับเป็น (ข) ทีหลัง

## 4. นิยาม "ครบทุก process"
process ที่ "ต้องมี" ของแต่ละแถว = process ที่ lot นั้นต้องผ่านจริงตาม route
ไม่ใช่ทั้ง 6 เสมอ (สินค้าต่างชนิดผ่านไม่เท่ากัน). ดึง route ได้จากข้อมูล ZPP0059
ที่มีอยู่แล้ว (operation → process):
- มียอด `cutting` → ต้องมี cutting check sheet
- มียอด `fp` (Insert) → ต้องมี fp check sheet
- มียอด `auto` (Brazing) → ต้องมี oven check sheet
- มียอด `hp` (H/P bender) → ต้องมี hp + hp_insert check sheet

**สถานะของแถว:**
- `required` = เซ็ต process ที่แถวนั้นต้องมี (จาก route/ยอด ZPP0059)
- `confirmed` = เซ็ต process ที่เจอ record ผล **OK** ใน DB (NG = ยังไม่ผ่าน ไม่นับว่าครบ)
- ถ้า `confirmed ⊇ required` → **ครบ**; ไม่งั้น → **ขาด (required − confirmed)**

## 5. การแสดงผล (คอลัมน์ใหม่ "Check Sheet")
แสดงเป็น badge สั้น ๆ อ่านแวบเดียวรู้:
- 🟢 `ครบ` — ทุก process ที่ต้องมี confirm OK แล้ว
- 🟡 `2/4` — ครบบางส่วน (hover เห็นว่าขาด: เช่น "ขาด: Oven, HP insert")
- 🔴 `NG` — มี process ที่ผล NG
- ⚪ `—` — ไม่มี process ใดต้องตรวจ / ยังไม่เริ่ม

ปรับ sort/filter ให้กรองเฉพาะ "ยังไม่ครบ" ได้ จะช่วยตามงานได้เร็ว

## 6. แนวทางลงมือ (ทำให้เรียบง่ายตาม CLAUDE.md)
1. ใน `build_daily_follow.py` เพิ่มฟังก์ชัน `load_checksheets()` อ่าน DB ทั้ง 6
   คืน dict `by_key[line|month|seq] = {process: "OK"/"NG"}` (อ่าน read-only,
   กรองเอา record ล่าสุด/ผลดีที่สุดต่อ process)
2. ตอนสร้างแต่ละแถว คำนวณ `required` จากยอด ZPP0059 ที่มีอยู่ แล้วเทียบกับ
   check sheet → ใส่ field ใหม่ เช่น `"checkSheet": {"state": "partial",
   "have": 2, "need": 4, "missing": ["oven","hp_insert"]}`
3. ใน `daily_follow_template.html` เพิ่ม `<th>` + cell render badge ตาม state
4. path ของ Server_firstlot ทำเป็น config ตัวแปรเดียว (เหมือน `SAP_EXPORT_BASE`)
   เผื่อเครื่องจริง path ต่างกัน

## 7. ความเสี่ยง/ข้อควรระวังที่เจอจากข้อมูลจริง
- ค่า `line` ไม่สม่ำเสมอ (`a/b/h` ตัวเล็ก, `6A`, `0`) → ต้อง normalize
- `prod_month` มีค่าเพี้ยน (`60/D202`, `/`) → ต้อง parse แบบกันพัง ข้ามตัวเสีย
- `cutting_records` ไม่มี `prod_month` → join ด้วย line+seq อย่างเดียว (อาจชนข้ามเดือน
  ถ้าใช้ทางเลือก (ก) ต้องยอมรับจุดนี้ หรือใช้ `saved_at` ช่วยกรองช่วงเวลา)
- DB ว่าง/ไม่มี table (oven, expander) → ต้อง try/except ไม่ให้ทั้ง build ล้ม
- DB อยู่บนเครื่อง first-lot server — ถ้าอ่านพร้อมระบบนั้นเขียน ให้เปิดแบบ read-only
  (`file:...?mode=ro`) กัน lock

## 8. การตัดสินใจ (สรุป) + สิ่งที่ทำจริงแล้ว
ผู้ใช้ตัดสินใจ:
1. **granularity** = ระดับ `line + seq + prod_month` เท่านั้น (ไม่สน lot)
2. **NG = ยังไม่ครบ**
3. ไม่ดึง required จาก ZPP0059 — map ทันทีตอน load ZPP0022

**ที่ implement แล้ว:**
- `serve_daily_follow.py` → `load_checksheets()` อ่าน 4 DB (read-only) +
  endpoint `GET /api/checksheets` คืน `{byKey, byLineSeq}`
- `daily_follow_template.html` → คอลัมน์ใหม่ **"Check Sheet"** (หลัง Assy Status),
  ฟังก์ชัน `csStatus(row)` + badge: 🟢 ครบ / 🟡 x/4 (hover เห็นว่าขาดอะไร) /
  🔴 NG / ว่าง = ยังไม่เริ่ม. โหลดผ่าน `loadChecksheets()` ตอน setup แล้ว re-render
- path Server_firstlot ตั้งผ่าน env `FIRSTLOT_DIR`

**ข้อจำกัดที่ยังเหลือ (ปรับได้ทีหลัง):** "required" ถูก fix เป็น 4 process
(`cutting, fp, hp, hp_insert` — แก้ได้ที่ const `CHECK_PROCS`). สินค้าที่ไม่ผ่าน
ครบทั้ง 4 จะขึ้นเป็น "partial" แม้จริง ๆ ลงครบตาม route แล้ว — ถ้าต้องการให้
required อิงตาม route จริงต่อรุ่น ค่อยเพิ่มภายหลัง
