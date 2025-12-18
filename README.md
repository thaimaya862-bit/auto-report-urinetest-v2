# auto_report_urinetest_v1

ระบบหน้าเว็บง่าย ๆ สำหรับสร้างเอกสาร "ส่งตรวจปัสสาวะ" จากทะเบียนราษฎร (PDF) สูงสุด 6 คน

## วิธีรันบนเครื่อง (Windows / CMD)

```bat
cd auto_report_urinetest_v1
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

แล้วเข้าเว็บ:
http://127.0.0.1:5000

## หมายเหตุ
- อัปโหลด PDF ทะเบียนราษฎร 1–6 ไฟล์ → ระบบเลือก template 1–6 ให้เอง
- วันที่: ไม่กรอก = วันนี้, แสดงรูปแบบ `16 ธันวาคม พ.ศ.2568`
- เวลา: ไม่กรอก = เว้นว่าง
- ใส่รูปทะเบียนราษฎรลง Word ผ่านตัวแปร `{{HOUSE_REG_IMAGE_A}}` … `{{HOUSE_REG_IMAGE_F}}`
- PDF output: ระบบพยายามแปลงเป็น PDF (ถ้าเครื่องมี docx2pdf/LibreOffice)
