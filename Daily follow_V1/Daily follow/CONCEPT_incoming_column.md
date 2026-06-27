# Concept — คอลัมน์ "Part Incoming" ในหน้า Daily Follow

ร่างแนวคิด (ยังไม่ลงมือเขียนโค้ด) สำหรับเพิ่ม **1 คอลัมน์** ในตาราง index
ที่บอกว่าแต่ละ lot มี **part item อะไรเข้ามา (รับเข้าคลังจริง) แล้วบ้าง**
โดยดึงข้อมูลจากระบบ **HEI Smart Stock Management** (`incoming.db`)

> เขียนตามโครงเดียวกับ `CONCEPT_checksheet_column.md` เพื่อให้ทีมอ่านต่อได้ง่าย

---

## 1. เป้าหมาย
ให้คนดูตาราง Daily Follow เห็นได้ในแวบเดียวว่า lot ไหน "part เริ่มทยอยเข้าคลัง
แล้ว" และเข้ามากี่รายการ/ยอดเท่าไหร่ โดยไม่ต้องเปิดระบบ HEI Smart Stock แยก

**ต่างจากคอลัมน์ที่มีอยู่อย่างไร:** คอลัมน์ `sourceReady` (METAL/PIPE/HEX…) ที่มี
อยู่เดิมมาจาก SAP (ZPP0022) ซึ่งเป็นสถานะ "ตามแผน/ระบบ". ส่วนคอลัมน์ใหม่นี้คือ
**การยืนยันรับของจริงหน้างาน** (สแกน QR ตอนรับเข้าคลัง) — เป็นคนละชั้นข้อมูลและ
ใช้เสริมกัน

## 2. แหล่งข้อมูล (HEI Smart Stock Management)
ไฟล์ `incoming.db` ตาราง `incoming` (ปัจจุบัน ~5,762 แถว) เป็น log การสแกนรับ
part เข้าคลังทีละครั้ง ฟิลด์ที่ใช้:

| ฟิลด์          | ตัวอย่าง          | ใช้ทำอะไร                               |
|----------------|-------------------|----------------------------------------|
| `part_no`      | `DQ02V939G02-F`   | รหัส part ที่รับเข้า (โชว์ในรายการ)      |
| `item_code`    | `100303`          | รหัส item (สำรองไว้ join/แสดง)          |
| `qty`          | `200`             | จำนวนที่รับเข้า                          |
| `unit`         | `PC`              | หน่วย                                   |
| `line_num`     | `A`               | line ปลายทาง → **คีย์ join**            |
| `seq`          | `0025`            | sequence → **คีย์ join**                |
| `pro_month`    | `062026`          | เดือนผลิต (MMYYYY) → **คีย์ join**       |
| `due_date`     | `29/05/2026`      | กำหนดต้องเข้า (ใช้เทียบ on-time ได้)     |
| `receive_date` | `31/05/2026`      | วันที่รับเข้าจริง                        |
| `receive_time` | `20:58:37`        | เวลารับเข้าจริง                         |
| `location`     | `32C4`            | ตำแหน่งจัดเก็บ (โชว์เสริมได้)            |

> ตารางเสริมที่ใช้แต่งหน้าได้: `item_descriptions(item_code → description,
> part_type)` ใช้แปลงรหัสเป็นชื่อ part ที่อ่านง่าย (เช่น "Side Plate")

## 3. การ join กับแถวใน Daily Follow
แถวใน index ถูกคีย์ด้วย `line | month | seq` (ดู `build_daily_follow.py` →
`load_progress`, key = `f"{line}|{month_display}|{seq_key}"`).
ตาราง `incoming` คีย์ด้วย `line_num | seq | pro_month`.

**คีย์ที่เสนอให้ใช้ join:** `line + seq + month` (แนวเดียวกับ Check Sheet)
- `line`: uppercase + trim (พบค่าปน `A`,`T`,`G`,`H` และค่าขยะ `None`,`PC`,`0061`)
- `seq`: normalize เป็นตัวเลขด้วย `seq_key` (`"0025"` → `25` ให้ตรงกับตาราง)
- `month`: แปลง `pro_month` `"062026"` (MMYYYY) → `month_display` `"6.2026"`

**ข้อมูลจริง:** จาก 5,762 แถว มี **4,634 แถว (~80%)** ที่มี `line+seq+month`
ครบ join ได้. อีก ~20% เป็นการสแกนรับที่ยังไม่ผูก line/seq (เช่นรับเข้า stock
กลาง) — join ไม่ได้ ก็แค่ไม่นับเข้า lot ใด (ไม่ทำให้พัง)

## 4. นิยาม "Part เข้าแล้ว" (จะ aggregate อะไรต่อแถว)
หนึ่งแถว Daily Follow = หนึ่ง `item` ภายใต้ `line+seq+month`. ต่อแถวเรา aggregate
incoming ทั้งหมดที่ตรง `line+seq+month` นั้นได้ค่า:
- `parts` = จำนวน **part_no ที่ไม่ซ้ำ** ที่รับเข้าแล้ว
- `qty` = ยอดรวม (sum `qty`)
- `lastReceive` = `receive_date`/`time` ล่าสุด (บอกว่าของเพิ่งเข้าเมื่อไหร่)
- `items` = รายการ `[{part_no, qty, receive_date, location}]` (ไว้โชว์ตอน hover)

> **คำถามต้องตัดสินใจ (สำคัญสุด): granularity / ระดับการ match**
> เลือกได้ 2 ทาง:
> - **(ก) ระดับ lot รวม** — โชว์ว่า "lot นี้มี part เข้ามากี่รายการ" โดยไม่สน
>   ว่าตรงกับ `item` ของแถวนั้นหรือไม่ (ทุกแถวที่ seq เดียวกันจะเห็นเลขชุดเดียวกัน)
>   → ง่าย, ตรงกับที่ตารางทำงานระดับ seq อยู่แล้ว (เหมือนที่ Check Sheet เลือก)
> - **(ข) ระดับ item** — match `incoming.item_code`/`part_no` กับ `item` ของแถว
>   ตรง ๆ เพื่อบอก "part ของแถวนี้เข้าแล้วยัง" → แม่นกว่าแต่ต้องมั่นใจว่ารหัส
>   ฝั่ง Daily Follow (`item`) กับฝั่ง incoming เป็นชุดเดียวกัน (ต้อง verify ก่อน)
> **แนะนำเริ่มที่ (ก)** ให้ใช้งานได้เร็วก่อน แล้วค่อยยกระดับเป็น (ข) เมื่อยืนยัน
> การ map รหัสได้แล้ว

## 5. การแสดงผล (คอลัมน์ใหม่ "Part เข้า")
แสดงเป็น badge สั้น ๆ อ่านแวบเดียวรู้ (วางหลังกลุ่ม `sourceReady`):
- 🟢 `5 รายการ` — มี part เข้าแล้ว (hover เห็นรายการ part_no + qty + วันรับ)
- 🟡 `2 รายการ` — เริ่มทยอยเข้า (สีเตือนถ้าต่ำกว่าที่ควร / ถ้าทำระดับ item ได้)
- ⚪ `—` — ยังไม่มี part เข้าเลยสำหรับ lot นี้
- (ถ้า `receive_date` > `due_date`) แสดงจุดแดงเล็ก ๆ ว่า "เข้าช้ากว่ากำหนด"

เพิ่ม filter "ยังไม่มี part เข้า" ช่วยให้ planner ไล่ตาม lot ที่ของยังไม่มาได้เร็ว

## 6. แนวทางลงมือ (ทำให้เรียบง่ายตาม CLAUDE.md)
อิงแพทเทิร์น `load_checksheets()` ที่ทำไว้แล้วใน `serve_daily_follow.py`:
1. เพิ่ม `load_incoming()` ใน `serve_daily_follow.py` — เปิด `incoming.db` แบบ
   **read-only** (`file:...?mode=ro`) คืน dict
   `by_key["LINE|month|seq"] = {parts, qty, lastReceive, items[]}`
2. เพิ่ม endpoint `GET /api/incoming` คืน `{byKey}` (เหมือน `/api/checksheets`)
   แล้วให้หน้าเว็บ `loadIncoming()` ตอน setup → re-render
3. ใน `daily_follow_template.html` เพิ่ม `<th>` + cell render badge ตาม state
   (ทำฟังก์ชัน `incomingStatus(row)` คล้าย `csStatus(row)`)
4. path ของ `incoming.db` ตั้งผ่านตัวแปร/่ env ตัวเดียว (เหมือน `FIRSTLOT_DIR`)
   เช่น `STOCK_DB` — เผื่อเครื่องจริง path ต่างกัน

> ทำฝั่ง serve (อ่านสด realtime) เหมือน Check Sheet ดีกว่าฝัง build ตอน gen html
> เพราะ part เข้าตลอดเวลา อยากให้เห็นค่าล่าสุดโดยไม่ต้อง re-build

## 7. ความเสี่ยง/ข้อควรระวังที่เจอจากข้อมูลจริง
- `line_num` ไม่สม่ำเสมอ (มีค่าขยะ `PC`, `0061`, `20260602`, `None`) → normalize
  + ตัวที่ join ไม่ได้ให้ข้ามเฉย ๆ ไม่ throw
- `pro_month` มีค่าเพี้ยน (`''`, `'0'`, `None`) → parse แบบกันพัง ข้ามตัวเสีย
- `seq` เป็น string zero-padded (`'0025'`) → ต้องผ่าน `seq_key` ให้ตรงกับตาราง
- `receive_date`/`due_date` เป็น `DD/MM/YYYY` (คนละ format กับ `month_display`)
  → parse ระวังสลับวัน/เดือน
- DB อยู่บนเครื่อง HEI Smart Stock — อ่านพร้อมระบบนั้นเขียน ให้เปิด read-only
  กัน lock; ถ้า DB หาย/ล็อก ให้ try/except คืนค่าว่าง ไม่ให้ทั้งหน้าเว็บพัง
- มี `deductions` (ตัด stock) แยกตาราง — v1 ยังไม่หักลบ โชว์ยอด "รับเข้า" ล้วน ๆ
  ก่อน (จะหัก net ทีหลังค่อยว่ากัน)

## 8. การตัดสินใจ (สรุป) + สิ่งที่ทำจริงแล้ว
ผู้ใช้ตัดสินใจ:
1. **granularity** = map ตาม `prod_month + line + seq` (ทางเลือก ก) โชว์ว่ามี
   item ไหนเข้ามาบ้าง + จำนวน (qty) — ไม่ match รายตัว item
2. **คอลัมน์โชว์แค่ `Stock` / ไม่มี** — รายละเอียด (part + qty + วันรับ) กดเข้าดู
3. **สี** = 🟢 `Stock` (มี part เข้า) / ⚪ ว่าง (ยังไม่มี)
4. ใช้ concept เดียวกับ Check Sheet (อ่านสดฝั่ง serve + endpoint + badge + env path)

**ที่ implement แล้ว:**
- `serve_daily_follow.py` → `load_incoming()` อ่าน `incoming.db` (read-only) รวม
  qty ต่อ part ต่อคีย์ `LINE|month|seq` + endpoint `GET /api/incoming` คืน `{byKey}`
- `daily_follow_template.html` → คอลัมน์ใหม่ **"Part เข้า"** (หลัง Check Sheet),
  ฟังก์ชัน `incStatus(row)` + badge 🟢 `Stock` (คลิกเปิด `openIncDetail` เห็น
  รายการ part + qty + วันรับ) / ว่าง = ยังไม่มี. โหลดผ่าน `loadIncoming()` ตอน
  setup แล้ว poll รอบเดียวกับ Check Sheet (ทุก 10 วิ)
- path `incoming.db` ตั้งผ่าน env `STOCK_DB` (มี candidate path สำรอง)

**ข้อจำกัดที่ยังเหลือ (ปรับได้ทีหลัง):** ~20% ของ record ที่ไม่มี line/seq/month
จะ join ไม่ได้ (รับเข้า stock กลาง) — v1 ไม่นับเข้า lot ใด. ยังไม่หักลบ
`deductions` (โชว์ยอดรับเข้าล้วน). ถ้าต้องการ match ระดับ item รายตัว หรือ
ตรวจ "เข้าช้ากว่า due_date" ค่อยเพิ่มภายหลัง
