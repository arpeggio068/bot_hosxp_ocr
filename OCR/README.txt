Python Thai OCR API — AutoIt connects directly to http://127.0.0.1:3078
ใช้ PaddleOCR PP-OCRv5 Thai (th_PP-OCRv5_mobile_rec), CPU
ไม่ใช้ Node.js, ไม่ตรวจชื่อหน้าต่าง, ไม่มีคำสั่งคลิก

ติดตั้งบน Windows:
1. ใช้ Python 3.12 แบบ 64-bit (แยกจาก conda/env RL เพื่อไม่เปลี่ยน dependency ของงานเดิม)
2. แตก ZIP เช่น D:\Nodejs_Project\hosxp_python_ocr
3. ดับเบิลคลิก install.bat หรือเปิด CMD ในโฟลเดอร์แล้วรัน:
   py -3.12 -m venv .venv
   .venv\Scripts\python.exe -m pip install --upgrade pip
   .venv\Scripts\python.exe -m pip install -r requirements.txt
4. รัน start.bat รอ READY อาจใช้เวลาครั้งแรกเพราะดาวน์โหลดโมเดล
5. เปิด TestOCR.au3 แล้วกด F1 ที่หน้าจอเป้าหมาย เมาส์เลื่อนไปกลางกรอบข้อความที่ OCR อ่านได้ ไม่มีการคลิก
6. กด ESC จบ AutoIt / Ctrl+C จบ Python

ถ้า py -3.12 ไม่พบ ให้ใช้ python.exe ของ Python 3.12 64-bit ที่ติดตั้งไว้แทน py -3.12
ถ้า Windows แจ้ง DLL missing ของ Paddle ให้ติดตั้ง Microsoft Visual C++ Redistributable x64 ตามข้อความ error

การตั้งค่าใน AutoIt:
$kScreen = 1.25 สำหรับเครื่องพัฒนา 125%; = 1 สำหรับเครื่องจริง 100%
$aItemTxtPos = [241, 252, 719, 461] ใช้ area เดิม
แก้คำค้นใน RunTest() ปัจจุบันเป็น เคลือบฟลูออไรด์วานิช
confidence = 60 หมายถึง 0-100 ของ PaddleOCR ไม่สามารถเทียบตรงกับคะแนน Tesseract
ไม่มีโหมด blue/gray ให้เลือก โปรแกรมลองภาพเดิม/contrast/กลับสี/adaptive อัตโนมัติเมื่อยังไม่พบ
ไม่ตรวจจับสีฟ้า ไม่ต้องจับภาพ template เก็บใหม่
หากพบข้อความเดียวกันหลายกรอบหรือหลายครั้งในกรอบเดียว จะคืน AMBIGUOUS และไม่เลื่อนเมาส์
ค้นข้อความตรงกันหลังปรับ Unicode/ช่องว่าง รองรับคำค้นอยู่ในข้อความยาว เช่น ... (ในเด็ก)
ตำแหน่งที่คืนคือกลางกรอบข้อความเต็มที่อ่านได้ ไม่ใช่กรอบย่อยรายตัวอักษร
ไม่ลบสระ/วรรณยุกต์ ไม่ใช้ fuzzy match
ถ้า OCR แบ่งคำเป้าหมายเป็นคนละกรอบอาจหาไม่พบ ดู lines ใน terminal เพื่อวิเคราะห์

การทดสอบภาพที่บันทึกไว้ ไม่ต้องเปิด AutoIt:
.venv\Scripts\python.exe server.py --image "C:\Users\ACER_NB\AppData\Local\Temp\hosxp_ocr_23664.png" --target "เคลือบฟลูออไรด์วานิช"
แสดงข้อความ confidence กรอบตำแหน่ง และ variant ที่ใช้ใน terminal

Local / Offline:
ภาพส่งเฉพาะ localhost ไป Python ในเครื่อง ไม่ส่งไป OCR cloud
ต้องมี Internet ตอนติดตั้งแพ็กเกจและดาวน์โหลดโมเดลครั้งแรก
Paddle cache โมเดลไว้ตามปกติที่ %USERPROFILE%\.paddlex\official_models
หลังดาวน์โหลดแล้ว สามารถคัดลอกโฟลเดอร์โมเดลทั้งสองไป:
models\PP-OCRv5_mobile_det\
models\th_PP-OCRv5_mobile_rec\
server จะใช้โฟลเดอร์ local เหล่านี้ก่อน (ต้องเป็นโฟลเดอร์โมเดลสมบูรณ์ มี inference.json และไฟล์ประกอบ)
อย่าลบ cache หรือโมเดลเมื่อใช้งาน offline และต้องทดสอบ restart แบบตัด net หนึ่งครั้ง

ข้อจำกัดการตรวจสอบชุดนี้:
ตรวจ syntax Python, matching/coordinate logic และ HTTP ด้วย backend จำลอง
ยังไม่ได้รันโมเดล PaddleOCR จริงหรือ AutoIt บน Windows จึงยังไม่ยืนยันผลกับภาพ HOSxP
เลือกโมเดลไทยโดยตรงจากโครงการ PaddleOCR แต่ไม่รับรองว่าจะอ่านทุก font/สีถูกทุกครั้ง
การลองภาพหลายแบบเพิ่มเวลาประมวลผลเฉพาะเมื่อรอบก่อนยังไม่พบ
