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
