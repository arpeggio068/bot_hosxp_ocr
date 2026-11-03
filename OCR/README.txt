ติดตั้ง: ปิด Python server แล้วแทนที่ OCR/server.py และ OCR/ocr_core.py ด้วยไฟล์ใน ZIP นี้ จากนั้นเปิด start.bat เดิม ไม่ต้องติดตั้ง dependencies เพิ่มบนเครื่องที่ใช้ PaddleOCR ชุดเดิม

การทำงาน:
1. ตรวจภาพปัจจุบันด้วย OpenCV template cache ใน RAM ก่อน
2. cache ไม่ผ่าน: อ่านภาพต้นฉบับ ตรวจ exact แล้ว fuzzy (รวมตรวจตำแหน่งใกล้เคียงที่อาจซ้ำ)
3. ยังไม่พบ: อ่านภาพ contrast อีกหนึ่งแบบ รวมสูงสุด 2 ครั้ง
4. พบตำแหน่งเดียว: เก็บภาพบรรทัดนั้นเป็น template พร้อมพิกัดกึ่งกลาง
5. หลายตำแหน่ง: AMBIGUOUS ไม่คลิก

cache แยกตามคำค้น confidence และขนาด area เก็บสูงสุด 32 รายการ ล้างเมื่อปิด Python เปลี่ยนตำแหน่งภายใน area ได้ แต่การเปลี่ยนสี/font อาจต้อง OCR ใหม่ CACHE_THRESHOLD=0.97 ใน server.py คือเกณฑ์ความเหมือนภาพ คนละค่ากับ FUZZY_MINIMUM=85.0 ใน ocr_core.py
cache_score คือคะแนน OpenCV; confidence และ similarity เมื่อ cache hit เป็นค่าจาก OCR ครั้งที่เรียนรู้ template ไม่ใช่การอ่าน OCR ครั้งใหม่
cache ใช้ตรวจภาพ ไม่สามารถยืนยันว่าไม่มีข้อความอื่นที่หน้าตาคล้ายแต่ไม่เหมือน template ทุกกรณี หากสงสัยผล cache ให้เปิด server ใหม่เพื่อบังคับ OCR
STRICT_NUMBERS=False ตามความต้องการผู้ใช้ เลขที่คลาดเคลื่อนเข้ารวมคะแนนความคล้าย

response 7 ช่องแรกเดิม: status x y confidence text similarity match_method
เพิ่มช่อง 8 timings_ms JSON และช่อง 9 source (ocr/opencv_cache) AutoIt เดิมอ่านช่องแรกได้เหมือนเดิม

log timings_ms หน่วย milliseconds:
upload_read_ms: รับ PNG จาก HTTP (ไม่รวม AutoIt จับภาพ/เขียนไฟล์)
decode_ms: แปลง PNG
cache_ms: OpenCV cache lookup
preprocess_ms: เตรียมภาพ
ocr_ms: PaddleOCR รวมทุกรอบ
exact_ms: เทียบตรงตัว
fuzzy_ms: เทียบข้อความใกล้เคียง
request_before_reply_ms: รับภาพถึงก่อนส่ง response ไม่รวม network ส่งกลับ
seconds: decode ถึงได้ผล หน่วยวินาที

วัดเวลารวมบน AutoIt เพิ่มรอบคำสั่งทดสอบ:
Local $hMeasure = TimerInit()
TestToClickOCR("การซักประวัติ", $aAddHPIPos[0], $aAddHPIPos[1], $aAddHPIPos[2], $aAddHPIPos[3], 60)
ConsoleWrite("AutoIt total_ms=" & Round(TimerDiff($hMeasure), 2) & @CRLF)

หากใช้บอท _FindOCR ที่มี cache ภาพตรงกันทั้ง area อยู่แล้ว อาจข้าม Python ไปเลย ทำให้ไม่เห็น request ใน Python log นั่นเป็นพฤติกรรมปกติ ใช้ TestToClickOCR ที่ส่งทุกครั้งเพื่อวัด Python cache โดยตรง

ทดสอบครั้งแรก source=ocr จากนั้นทดสอบคำเดิม source=opencv_cache เปรียบเทียบ seconds และ timings_ms การอ่านครั้งแรกยังขึ้นกับขนาด area และ CPU ไม่รับประกันเวลาคงที่
