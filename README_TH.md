TVC Desktop Bot v0.5.1

แก้เฉพาะ Flow บันทึกจาก v0.5

สาเหตุ: Dialog ของ T.V.C แสดง Yes/No แต่ control เป็น custom UI ทำให้ pywinauto และ Win32 อ่านปุ่มไม่เจอ

Flow ใหม่:
1. กรอกทุกช่องจาก Excel
2. เพิ่มบริการทีละรายการ
3. กด F10 จริง
4. รอ Dialog บันทึก
5. ส่งคีย์ Y = Yes
6. รอ Dialog พิมพ์
7. ส่งคีย์ N = No
8. เขียน DONE กลับ Excel

วิธีใช้:
- Copy .venv จาก v0.5 มาวางในโฟลเดอร์ v0.5.1
- เปิดหน้า เพิ่มใบงาน (JOB) แบบว่าง
- เปิด CMD แบบ Administrator
- รัน: .venv\Scripts\python.exe src\bot.py

ไฟล์ Excel ใน data ถูกรีเซ็ตเป็น TEST-001 และบริการ J จริง 3 รายการแล้ว

## GUI v0.8.0 และ Runtime Architecture

- Source/dev mode ใช้ `.venv\Scripts\python.exe src\bot.py` เป็น worker เหมือนเดิม
- Frozen mode ถูกออกแบบให้ `TVC Bot Control.exe` เรียก `TVC Bot Worker.exe`
  ซึ่งอยู่ในโฟลเดอร์เดียวกัน โดย worker เป็นเพียง entry point ที่เรียก `bot.main()`
- GUI และ worker ใช้ `src/runtime_paths.py` ร่วมกันสำหรับ config, field map,
  assets, logs และ screenshots; ไม่มี path ที่ผูกกับ `D:\BOT-PMS`
- หาก app directory เขียนไม่ได้ logs/screenshots จะใช้ user-data directory ที่เขียนได้
- รอบนี้เป็นการเตรียม architecture เท่านั้น ยังไม่มี PyInstaller spec และยังไม่ build `.exe`

## Safety Lock

- Pre-check อ่าน workbook ทุกครั้ง หากพบ `ERROR` ที่มี
  `UNCERTAIN_TVC_SAVE` ไฟล์จะเป็น `SAFETY LOCK` และ Start ไม่ได้
- Safety Lock ถูกสร้างใหม่จาก workbook หลังเปิด GUI ใหม่ และมี metadata สำหรับ
  abnormal/dirty-form lock เพื่อไม่ให้ restart ข้ามการตรวจสอบ
- Metadata มี health state `HEALTHY`, `MISSING`, `CORRUPT`, `UNREADABLE` และ
  `WRITE_FAILED`; สามสถานะท้ายเป็น `FAIL_CLOSED` และปิด Start จนกว่าผู้ใช้จะแก้
  แล้วกด `ตรวจสอบหลังแก้ไข` เพื่อ retry/reload โดย GUI จะไม่ overwrite ไฟล์ที่เสียเอง
- Registry ของ Safety Lock แยกจาก Excel Queue การนำไฟล์ออกจาก Queue ไม่ได้ลบ lock
  และการตรวจซ้ำจะอ่านทุก persisted workbook path ก่อนปลดเฉพาะ lock ที่ผ่านจริง
- ปุ่ม `ตรวจสอบหลังแก้ไข` อ่านสถานะ Excel/T.V.C เท่านั้น ไม่เปลี่ยน ERROR เป็น
  WAIT, ไม่ลบ ERROR และไม่ retry JOB อัตโนมัติ
- หลัง forced/abnormal pre-commit exit ผู้ใช้ต้องตรวจสอบและปิด/ยกเลิก active
  JOB form ใน T.V.C ก่อนจึงจะปลด lock ได้

## Stop Semantics (ตั้งใจออกแบบ)

`หยุดหลังจบ JOB` หมายถึงให้ JOB ปัจจุบันทำ fields, services, duplicate flow,
F10 และ YES/NO จนครบ safe flow แล้วหยุดก่อน JOB ถัดไปและก่อน Excel ไฟล์ถัดไป
ไม่มี stop check แทรกกลาง JOB ส่วน force terminate ใช้เฉพาะ timeout/failure และจะเข้า
Safety Lock/recovery flow

## Runtime Safety ก่อน Build

- GUI ใช้ Windows named mutex ชื่อ
  `Local\TVC_JOB_BOT_V080_GUI_SINGLE_INSTANCE` จึงเปิดได้ครั้งละหนึ่ง instance
  ทั้ง source mode และ frozen mode; หาก process crash Windows จะคืน mutex handle
  อัตโนมัติและไม่เกิด stale PID lock
- `safety_locks.json` มี GUI เป็น writer เพียง process เดียว ส่วน worker ไม่เขียน
  registry นี้ การบังคับ single-instance จึงป้องกัน concurrent GUI writers
  และแต่ละ atomic write ใช้ไฟล์ชั่วคราวชื่อเฉพาะ PID/UUID
- writable data root ถูกตรวจด้วยการสร้าง/เขียน/flush/`fsync`/atomic replace/readback/
  delete จริง ไม่อ้างอิง `os.access` หาก app directory ไม่ผ่านจะ fallback ไปที่
  `%LOCALAPPDATA%\TVC_JOB_BOT`; ถ้าทั้งสองตำแหน่งใช้ไม่ได้ GUI จะหยุดแบบ fatal
- logs, screenshots, safety metadata และ cooperative stop files อยู่ใต้ writable
  data root เดียวกัน ส่วน `config.ini`, `field_map.json` และ assets ยังคงเป็น
  resource แบบ read-only และไม่มีการเขียนทับอัตโนมัติ
